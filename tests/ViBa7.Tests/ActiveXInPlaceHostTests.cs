using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using ViBa7.ActiveXHost;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Compilation;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Execution;
using ViBa7.ProjectModel.IntelliSense;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class ActiveXInPlaceHostTests
{
    [Fact]
    public void ActiveXHostFactory_TryResolveClsid_KnownControls()
    {
        Assert.True(ActiveXHostFactory.TryResolveClsid("MSFlexGridLib.MSFlexGrid", out string flexClsid));
        Assert.Equal("{6262D3A0-5316-11CE-BFC1-08002BE10318}", flexClsid);

        Assert.True(ActiveXHostFactory.TryResolveClsid("MSComctlLib.TreeCtrl", out string treeClsid));
        Assert.Equal("{C74190B6-8589-11D1-B16A-00C0F0283628}", treeClsid);

        Assert.True(ActiveXHostFactory.TryResolveClsid("RichTextLib.RichTextBox", out string rtfClsid));
        Assert.Equal("{3B7C8860-D78F-101B-B9B5-04021C009402}", rtfClsid);

        Assert.True(ActiveXHostFactory.TryResolveClsid("TabDlg.SSTab", out string ssTabClsid));
        Assert.Equal("{BDC217C8-ED16-11CD-956C-0000C04E4C0A}", ssTabClsid);

        Assert.True(ActiveXHostFactory.TryResolveClsid("MSComDlg.CommonDialog", out string dlgClsid));
        Assert.Equal("{F9043C88-F6F2-101A-A3C9-08002B2F49FB}", dlgClsid);
    }

    [Fact]
    public void ActiveXHostFactory_TryResolveClsid_ExplicitGuid()
    {
        string rawGuid = "{12345678-ABCD-EF01-2345-6789ABCDEF01}";
        Assert.True(ActiveXHostFactory.TryResolveClsid(rawGuid, out string resolved));
        Assert.Equal(rawGuid, resolved);
    }

    [Fact]
    public void ActiveXHostFactory_TryResolveClsid_InvalidReturnsFalse()
    {
        Assert.False(ActiveXHostFactory.TryResolveClsid("NotARealControl.NonExistentClsidXYZ123", out _));
        Assert.False(ActiveXHostFactory.TryResolveClsid("", out _));
    }

    [Fact]
    public void ActiveXHostFactory_TryCreateHost_IntrinsicControls_ReturnsNull()
    {
        var btn = new FormDesignerControl { TypeName = "VB.CommandButton", Name = "Command1" };
        var txt = new FormDesignerControl { TypeName = "VB.TextBox", Name = "Text1" };
        var lbl = new FormDesignerControl { TypeName = "VB.Label", Name = "Label1" };

        Assert.Null(ActiveXHostFactory.TryCreateHost(btn));
        Assert.Null(ActiveXHostFactory.TryCreateHost(txt));
        Assert.Null(ActiveXHostFactory.TryCreateHost(lbl));
    }

    [Fact]
    public void ActiveXHostFactory_EnableInPlaceActiveXHosting_CanToggle()
    {
        bool original = ActiveXHostFactory.EnableInPlaceActiveXHosting;
        try
        {
            ActiveXHostFactory.EnableInPlaceActiveXHosting = false;
            var flex = new FormDesignerControl { TypeName = "MSFlexGridLib.MSFlexGrid", Name = "MSFlexGrid1" };
            var host = ActiveXHostFactory.TryCreateHost(flex);
            Assert.Null(host);
        }
        finally
        {
            ActiveXHostFactory.EnableInPlaceActiveXHosting = original;
        }
    }

    private static void RunInSta(Action action)
    {
        Exception? exception = null;
        var thread = new System.Threading.Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });
        thread.SetApartmentState(System.Threading.ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception).Throw();
        }
    }

    [Fact]
    public void DynamicActiveXControl_NormalizesClsid()
    {
        RunInSta(() =>
        {
            // Un GUID sin llaves debe normalizarse con llaves
            string raw = "6262D3A0-5316-11CE-BFC1-08002BE10318";
            var ctrl = new DynamicActiveXControl(raw);
            Assert.Equal("{6262D3A0-5316-11CE-BFC1-08002BE10318}", ctrl.Clsid);
            Assert.True(ctrl.IsDesignMode);
        });
    }

    [Fact]
    public void ActiveXInPlaceHost_InitializationAndPropertySync()
    {
        RunInSta(() =>
        {
            var modelControl = new FormDesignerControl
            {
                TypeName = "MSFlexGridLib.MSFlexGrid",
                Name = "Grid1",
                Left = 120,
                Top = 240,
                Width = 3000,
                Height = 2000
            };
            modelControl.CustomProperties.Add(new FormPropertyItem("Cols", "5"));
            modelControl.CustomProperties.Add(new FormPropertyItem("Rows", "10"));

            var host = new ActiveXInPlaceHost("{6262D3A0-5316-11CE-BFC1-08002BE10318}", "MSFlexGridLib.MSFlexGrid")
            {
                BoundControl = modelControl,
                Width = modelControl.WidthPixels,
                Height = modelControl.HeightPixels
            };

            Assert.Equal("{6262D3A0-5316-11CE-BFC1-08002BE10318}", host.Clsid);
            Assert.Equal("MSFlexGridLib.MSFlexGrid", host.ProgId);
            Assert.Same(modelControl, host.BoundControl);
            Assert.True(host.IsDesignMode);
            Assert.Equal(200.0, host.Width); // 3000 / 15
            Assert.Equal(133.33333333333334, host.Height); // 2000 / 15

            // Validar que SyncFromDesignerModel no genere excepciones aún sin instanciar la ventana Win32 en pruebas headless
            host.SyncFromDesignerModel();
            host.Dispose();
        });
    }

    [Fact]
    public void ActiveXInPlaceHost_FallbackVisual_AutomaticWhenNativeNotRendered()
    {
        RunInSta(() =>
        {
            var modelControl = new FormDesignerControl
            {
                TypeName = "ThirdPartyLib.CustomGrid",
                Name = "CustomGrid1",
                Width = 2400,
                Height = 1500
            };

            var fallbackBorder = new System.Windows.Controls.Border
            {
                Width = 160,
                Height = 100
            };

            // CLSID ficticio que no existe en el registro: TryInitializeAndRender fallará limpiamente y activará FallbackVisual
            using var host = new ActiveXInPlaceHost("{99999999-8888-7777-6666-555555555555}", "ThirdPartyLib.CustomGrid", 160, 100)
            {
                BoundControl = modelControl,
                FallbackVisual = fallbackBorder
            };

            bool nativeOk = host.TryInitializeAndRender();
            Assert.False(nativeOk);
            Assert.False(host.HasValidNativeVisual);
            Assert.Same(fallbackBorder, host.Child);
        });
    }

    [Fact]
    public void DynamicActiveXControl_ClampsZeroDimensionsAndSetsDesignModeSite()
    {
        RunInSta(() =>
        {
            using var host = new ActiveXInPlaceHost("{6262D3A0-5316-11CE-BFC1-08002BE10318}", "MSFlexGridLib.MSFlexGrid", 0, 0);
            Assert.True(host.AxControl.Width >= 16);
            Assert.True(host.AxControl.Height >= 16);
            Assert.NotNull(host.DesignSite);
            Assert.True(host.DesignSite.DesignMode);
        });
    }

    [Fact]
    public void VbPropertyBag_ReturnsS_OkForDefinedAndE_InvalidArgForMissing()
    {
        var props = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
        {
            ["Cols"] = 4,
            ["Caption"] = "Título",
            ["Enabled"] = true
        };

        var bag = new VbPropertyBag(props);

        object colsVal = 0;
        int hrCols = bag.Read("Cols", ref colsVal, null);
        Assert.Equal(0, hrCols); // S_OK
        Assert.Equal(4, colsVal);

        object capVal = string.Empty;
        int hrCap = bag.Read("Caption", ref capVal, null);
        Assert.Equal(0, hrCap); // S_OK
        Assert.Equal("Título", capVal);

        // Propiedad no definida en el .frm: debe retornar E_INVALIDARG (0x80070057) para que msvbvm60.dll use el DefaultValue del autor del OCX
        object missingVal = 100;
        int hrMissing = bag.Read("UnsetDivisorProperty", ref missingVal, null);
        Assert.Equal(unchecked((int)0x80070057), hrMissing);
        Assert.Equal(100, missingVal);
    }

    [Fact]
    public void VbControlMemberCatalog_TryResolveOcxClsid_ResolvesFromRegisteredTypeLib()
    {
        var dummyLib = new ViBa7.ProjectModel.ActiveX.OcxLibraryInfo
        {
            LibraryName = "JControlsLib",
            FilePath = @"C:\Dummy\axJControls.ocx",
            Guid = "{11111111-2222-3333-4444-555555555555}"
        };
        dummyLib.Controls.Add(new ViBa7.ProjectModel.ActiveX.OcxControlInfo
        {
            Name = "axJColCombo",
            Clsid = "{AABBCCDD-1122-3344-5566-77889900AABB}",
            IsControl = true
        });

        ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.RegisterOcxLibrary(dummyLib);
        try
        {
            Assert.True(ActiveXHostFactory.TryResolveClsid("JControlsLib.axJColCombo", out string clsid1));
            Assert.Equal("{AABBCCDD-1122-3344-5566-77889900AABB}", clsid1);

            Assert.True(ActiveXHostFactory.TryResolveClsid("axJControls.axJColCombo", out string clsid2));
            Assert.Equal("{AABBCCDD-1122-3344-5566-77889900AABB}", clsid2);
        }
        finally
        {
            ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.UnregisterOcxLibrary("JControlsLib");
        }
    }

    [Fact]
    public void ActiveXHostFactory_TryCreateHost_NonVisualControls_ReturnsNull()
    {
        var dlg = new FormDesignerControl { TypeName = "MSComDlg.CommonDialog", Name = "CommonDialog1" };
        var imgList = new FormDesignerControl { TypeName = "MSComctlLib.ImageListCtrl", Name = "ImageList1" };

        Assert.Null(ActiveXHostFactory.TryCreateHost(dlg));
        Assert.Null(ActiveXHostFactory.TryCreateHost(imgList));
    }

    [Fact]
    public void NativeDialogGuard_And_MaskFpuExceptions_ExecuteCleanly()
    {
        RunInSta(() =>
        {
            ViBa7.ActiveXHost.Native.Ole32Interop.MaskFpuExceptions();
            using var guard = new ViBa7.ActiveXHost.Native.NativeDialogGuard();
            Assert.False(guard.DialogSuppressed);
        });
    }

    [Fact]
    public void DynamicActiveXControl_IsIndependentOfWinFormsAxHost_SupportsWindowlessControls()
    {
        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}")
            {
                IsDesignMode = true,
                Width = 120,
                Height = 40
            };

            // Verificar que DynamicActiveXControl NO hereda de System.Windows.Forms.AxHost ni Control,
            // evitando System.NotSupportedException en AxHost.EnsureWindowPresent() para controles OCX windowless.
            Assert.False(typeof(System.Windows.Forms.Control).IsAssignableFrom(typeof(DynamicActiveXControl)));

            var site = new OleDesignClientSite(ctrl);
            Assert.False(site.UserMode); // DesignMode = True -> UserMode = False
            Assert.False(site.UIDead); // UIDead debe ser False para que los controles no queden bloqueados o deshabilitados
            Assert.Equal("Twip", site.ScaleUnits);
            Assert.False(site.ShowGrabHandles);
            Assert.False(site.ShowHatching);

            Assert.Equal(0, site.CanInPlaceActivate()); // S_OK
        });
    }

    [Fact]
    public void OleDesignClientSite_ProvidesValidFramePointerAndFrameMethods()
    {
        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}")
            {
                IsDesignMode = true,
                Width = 200,
                Height = 80
            };

            var site = new OleDesignClientSite(ctrl);
            var frameInfo = new ViBa7.ActiveXHost.Native.OLEINPLACEFRAMEINFO
            {
                cb = System.Runtime.InteropServices.Marshal.SizeOf<ViBa7.ActiveXHost.Native.OLEINPLACEFRAMEINFO>()
            };

            int hr = site.GetWindowContext(out IntPtr ppFrame, out IntPtr ppDoc, out var posRect, out var clipRect, ref frameInfo);
            Assert.Equal(0, hr);
            Assert.NotEqual(IntPtr.Zero, ppFrame);
            Assert.Equal(200, posRect.right);
            Assert.Equal(80, posRect.bottom);

            // Liberar interfaz COM obtenida en GetWindowContext
            System.Runtime.InteropServices.Marshal.Release(ppFrame);
        });
    }

    [Fact]
    public void DynamicActiveXControl_ConvertPropertyStringToOcxType_PreservesEnumZerosAndNegatives()
    {
        // Propiedades de enumeración / enteras con 0 o -1 no deben ser convertidas a booleano
        object? borderVal = DynamicActiveXControl.ConvertPropertyStringToOcxType("BorderStyle", "0");
        Assert.IsType<int>(borderVal);
        Assert.Equal(0, borderVal);

        object? styleVal = DynamicActiveXControl.ConvertPropertyStringToOcxType("Style", "-1");
        Assert.IsType<int>(styleVal);
        Assert.Equal(-1, styleVal);

        // Propiedades conocidas booleanas con 0 o -1 sí deben convertirse a booleano
        object? enabledFalse = DynamicActiveXControl.ConvertPropertyStringToOcxType("Enabled", "0");
        Assert.IsType<bool>(enabledFalse);
        Assert.False((bool)enabledFalse);

        object? visibleTrue = DynamicActiveXControl.ConvertPropertyStringToOcxType("Visible", "-1");
        Assert.IsType<bool>(visibleTrue);
        Assert.True((bool)visibleTrue);
    }

    [Fact]
    public void VbPropertyBag_CoercionAndSavedProperties()
    {
        var initial = new Dictionary<string, object>
        {
            ["Enabled"] = "-1",
            ["BorderStyle"] = 1,
            ["ForeColor"] = "&H00FF0000&"
        };

        var bag = new VbPropertyBag(initial);
        Assert.Equal(3, bag.SavedProperties.Count);

        object pVarBool = false;
        int hrBool = bag.Read("Enabled", ref pVarBool, null);
        Assert.Equal(0, hrBool);
        Assert.True((bool)pVarBool);

        object pVarInt = 0;
        int hrInt = bag.Read("BorderStyle", ref pVarInt, null);
        Assert.Equal(0, hrInt);
        Assert.Equal(1, pVarInt);

        object pVarMissing = null!;
        int hrMissing = bag.Read("NonExistentProperty", ref pVarMissing, null);
        Assert.NotEqual(0, hrMissing); // E_INVALIDARG para usar valor por defecto del control
    }

    [Fact]
    public void FormDesignParser_And_Serializer_PreserveObjectHeadersIndexedPropsAndNestedBeginProperty()
    {
        string frmContent = @"VERSION 5.00
Object = ""{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0""; ""MSCOMCTL.OCX""
Object = ""{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0""; ""TABCTL32.OCX""
Begin VB.Form frmTestActiveX 
   Caption         =   ""Prueba Fase 1""
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Begin TabDlg.SSTab SSTab1 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5500
      _ExtentX        =   9701
      _ExtentY        =   5292
      Tabs            =   2
      TabCaption(0)   =   ""General""
      Tab(0).Control(0)=   ""Toolbar1""
      Tab(0).ControlCount=   1
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1058
         Object.Visible         =   -1  'True
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   ""Guardar""
               Key             =   ""btnSave""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = ""frmTestActiveX""
";

        var model = FormDesignParser.Parse(frmContent);

        Assert.Equal(2, model.ObjectDeclarations.Count);
        Assert.Contains(model.ObjectDeclarations, h => h.Contains("MSCOMCTL.OCX", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(model.ObjectDeclarations, h => h.Contains("TABCTL32.OCX", StringComparison.OrdinalIgnoreCase));

        Assert.Single(model.Children);
        var ssTab = model.Children[0];
        Assert.Equal("TabDlg.SSTab", ssTab.TypeName);
        Assert.Contains(ssTab.CustomProperties, p => p.Name == "TabCaption(0)" && p.Value.Trim('"') == "General");
        Assert.Contains(ssTab.CustomProperties, p => p.Name == "Tab(0).Control(0)" && p.Value.Trim('"') == "Toolbar1");

        Assert.Single(ssTab.Children);
        var toolbar = ssTab.Children[0];
        Assert.Equal("MSComctlLib.Toolbar", toolbar.TypeName);
        Assert.Contains(toolbar.CustomProperties, p => p.Name == "Object.Visible");

        var buttonsBlock = toolbar.CustomProperties.Find(p => p.IsBeginProperty && p.Name.StartsWith("Buttons", StringComparison.Ordinal));
        Assert.NotNull(buttonsBlock);
        string joinedBlock = string.Join("\n", buttonsBlock!.RawPropertyLines);
        Assert.Contains("BeginProperty Button1", joinedBlock);
        Assert.Contains("Caption         =   \"Guardar\"", joinedBlock);

        string serialized = FormDesignSerializer.Serialize(model);
        Assert.Contains("Object = \"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0\"; \"MSCOMCTL.OCX\"", serialized);
        Assert.Contains("Object = \"{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0\"; \"TABCTL32.OCX\"", serialized);
        Assert.Contains("TabCaption(0)", serialized);
        Assert.Contains("\"General\"", serialized);
        Assert.Contains("Tab(0).Control(0)", serialized);
        Assert.Contains("BeginProperty Button1", serialized);
    }

    [Fact]
    public void FrxResourceFile_RoundTripsRawActiveXBinaryStreamAndPreservesExistingFrxOffsets()
    {
        string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ViBa7_Fase1_FrxTest_" + Guid.NewGuid().ToString("N"));
        System.IO.Directory.CreateDirectory(tempDir);
        try
        {
            string frmPath = System.IO.Path.Combine(tempDir, "frmOcxBin.frm");
            string frxPath = System.IO.Path.Combine(tempDir, "frmOcxBin.frx");

            byte[] stateBlob1 = new byte[] { 0xDE, 0xAD, 0xBE, 0xEF, 0x01, 0x02, 0x03, 0x04 };
            byte[] customBlob2 = new byte[] { 0xCA, 0xFE, 0xBA, 0xBE, 0x10, 0x20 };

            int offset1 = ViBa7.ProjectModel.Resources.FrxResourceFile.AppendRawBlock(frxPath, stateBlob1);
            int offset2 = ViBa7.ProjectModel.Resources.FrxResourceFile.AppendRawBlock(frxPath, customBlob2);

            Assert.Equal(0, offset1);
            Assert.Equal(12, offset2); // 4 bytes length + 8 bytes payload

            string frmText = $@"VERSION 5.00
Object = ""{{6262D3A0-5316-11CE-BFC1-08002BE10318}}#1.0#0""; ""MSFLXGRD.OCX""
Begin VB.Form frmOcxBin 
   Caption         =   ""Test FRX""
   ClientHeight    =   3000
   ClientWidth     =   4000
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1800
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      Object.zyx      =   ""frmOcxBin.frx"":{offset1:X4}
      CustomData      =   ""frmOcxBin.frx"":{offset2:X4}
   End
End
Attribute VB_Name = ""frmOcxBin""
";
            System.IO.File.WriteAllText(frmPath, frmText);

            var model = FormDesignParser.Parse(frmText);
            FormDesignParser.ResolveControlPictures(model, tempDir);
            Assert.NotNull(model.OriginalFrxBytes);
            Assert.Single(model.Children);

            var grid = model.Children[0];
            Assert.Equal("Object.zyx", grid.ActiveXStatePropertyName);
            Assert.NotNull(grid.ActiveXStateBlob);
            Assert.Equal(stateBlob1, grid.ActiveXStateBlob);
            Assert.True(grid.BinaryProperties.ContainsKey("CustomData"));
            Assert.Equal(customBlob2, grid.BinaryProperties["CustomData"]);

            // Modificar únicamente el estado binario principal del OCX y guardar
            byte[] updatedStateBlob = new byte[] { 0xAA, 0xBB, 0xCC, 0xDD, 0xEE, 0xFF };
            grid.ActiveXStateBlob = updatedStateBlob;
            grid.HasModifiedActiveXState = true;

            ViBa7.ProjectModel.Resources.FrxResourceFile.EnsureFrxResourcesSaved(model, frmPath);

            // Verificar que el offset original de CustomData (:000C) sigue intacto y legible
            byte[]? preservedCustom = ViBa7.ProjectModel.Resources.FrxResourceFile.ReadRawBlock(frxPath, offset2);
            Assert.NotNull(preservedCustom);
            Assert.Equal(customBlob2, preservedCustom);

            // Verificar que el nuevo bloque de Object.zyx fue añadido y referenciado correctamente
            var zyxProp = grid.CustomProperties.Find(p => p.Name == "Object.zyx");
            Assert.NotNull(zyxProp);
            Assert.True(ViBa7.ProjectModel.Resources.FrxResourceFile.ParseReference(zyxProp!.Value, out _, out int newZyxOffset));
            byte[]? savedState = ViBa7.ProjectModel.Resources.FrxResourceFile.ReadRawBlock(frxPath, newZyxOffset);
            Assert.NotNull(savedState);
            Assert.Equal(updatedStateBlob, savedState);
        }
        finally
        {
            try { System.IO.Directory.Delete(tempDir, true); } catch { }
        }
    }

    [Fact]
    public void Ole32Interop_CreateComStreamFromBytes_And_ReadBytesFromComStream_RoundTrip()
    {
        RunInSta(() =>
        {
            byte[] expected = new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 255, 128, 64 };
            IntPtr pStream = ViBa7.ActiveXHost.Native.Ole32Interop.CreateComStreamFromBytes(expected);
            Assert.NotEqual(IntPtr.Zero, pStream);
            try
            {
                byte[]? actual = ViBa7.ActiveXHost.Native.Ole32Interop.ReadBytesFromComStream(pStream);
                Assert.NotNull(actual);
                Assert.Equal(expected, actual);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.Release(pStream);
            }
        });
    }

    [Fact]
    public void DynamicActiveXControl_LifecycleState_And_Diagnostics_TrackedCorrectly()
    {
        RunInSta(() =>
        {
            var ctrl = new DynamicActiveXControl("{99999999-8888-7777-6666-555555555555}");
            Assert.Equal(ActiveXLifecycleState.Uninitialized, ctrl.LifecycleState);

            bool ok = ctrl.CreateControl();
            Assert.False(ok);
            Assert.Equal(ActiveXLifecycleState.Faulted, ctrl.LifecycleState);
            Assert.NotEmpty(ctrl.Diagnostics);

            ctrl.Dispose();
            Assert.Equal(ActiveXLifecycleState.Disposed, ctrl.LifecycleState);
        });
    }

    [Fact]
    public void OleDesignClientSite_GetExtendedControl_ReturnsVbControlExtender_And_MoveUpdatesBounds()
    {
        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}", 100, 50)
            {
                Name = "GridExt1"
            };

            var site = new OleDesignClientSite(ctrl);
            int hr = site.GetExtendedControl(out object? ppDisp);
            Assert.Equal(0, hr);
            Assert.NotNull(ppDisp);
            var extender = Assert.IsType<VbControlExtender>(ppDisp);
            Assert.Equal("GridExt1", extender.Name);

            float notifiedL = 0, notifiedT = 0, notifiedW = 0, notifiedH = 0;
            ctrl.ExtenderBoundsChanged += () =>
            {
                notifiedL = ctrl.Extender.Left;
                notifiedT = ctrl.Extender.Top;
                notifiedW = ctrl.Extender.Width;
                notifiedH = ctrl.Extender.Height;
            };

            extender.Move(300, 450, 2400, 1200);
            Assert.Equal(300f, extender.Left);
            Assert.Equal(450f, extender.Top);
            Assert.Equal(2400f, extender.Width);
            Assert.Equal(1200f, extender.Height);
            Assert.Equal(300f, notifiedL);
            Assert.Equal(450f, notifiedT);
            Assert.Equal(2400f, notifiedW);
            Assert.Equal(1200f, notifiedH);
            Assert.Equal(160, ctrl.Width); // 2400 / 15
            Assert.Equal(80, ctrl.Height); // 1200 / 15
        });
    }

    [Fact]
    public void OleDesignClientSite_TransformCoords_ConvertsHimetricAndTwipsAccurately()
    {
        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}");
            var site = new OleDesignClientSite(ctrl);

            // 1 pulgada = 2540 HIMETRIC = 1440 Twips
            var ptl = new ViBa7.ActiveXHost.Native.POINT { x = 2540, y = 5080 };
            var ptf = new ViBa7.ActiveXHost.Native.POINTF();
            int hr1 = site.TransformCoords(
                ref ptl,
                ref ptf,
                ViBa7.ActiveXHost.Native.Ole32Interop.XFORMCOORDS_SIZE | ViBa7.ActiveXHost.Native.Ole32Interop.XFORMCOORDS_HIMETRICTOCONTAINER);

            Assert.Equal(0, hr1);
            Assert.Equal(1440f, ptf.x, 1);
            Assert.Equal(2880f, ptf.y, 1);

            // Conversión inversa: Twips del contenedor -> HIMETRIC
            var ptlBack = new ViBa7.ActiveXHost.Native.POINT();
            var ptfIn = new ViBa7.ActiveXHost.Native.POINTF { x = 1440f, y = 720f };
            int hr2 = site.TransformCoords(
                ref ptlBack,
                ref ptfIn,
                ViBa7.ActiveXHost.Native.Ole32Interop.XFORMCOORDS_POSITION | ViBa7.ActiveXHost.Native.Ole32Interop.XFORMCOORDS_CONTAINERTOHIMETRIC);

            Assert.Equal(0, hr2);
            Assert.Equal(2540, ptlBack.x);
            Assert.Equal(1270, ptlBack.y);
        });
    }

    [Fact]
    public void OleDesignClientSite_AmbientProperties_ExposesStdFontAndTranslatedColors()
    {
        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}");
            ctrl.ConfigureAmbientFont("Tahoma", 9.75f, bold: true, italic: false, underline: false, strikethrough: false, charset: 0);

            var site = new OleDesignClientSite(ctrl);
            object? stdFont = site.Font;
            Assert.NotNull(stdFont);
            Assert.True(System.Runtime.InteropServices.Marshal.IsComObject(stdFont));

            Assert.True(site.MessageReflect);
            Assert.True(site.AutoClip);
            Assert.Equal("Twip", site.ScaleUnits);
            Assert.Equal((short)1, site.Appearance);
            Assert.False(site.RightToLeft);

            // Verificar OleTranslateColor con color RGB y color de sistema OLE_COLOR
            var rgbColor = ViBa7.ActiveXHost.Native.Ole32Interop.TranslateOleColor(0x00FF8040u);
            Assert.Equal(0x40, rgbColor.R);
            Assert.Equal(0x80, rgbColor.G);
            Assert.Equal(0xFF, rgbColor.B);

            var sysColor = ViBa7.ActiveXHost.Native.Ole32Interop.TranslateOleColor(0x8000000Fu); // COLOR_BTNFACE
            Assert.Equal(255, sysColor.A);
        });
    }

    [Fact]
    public void DynamicActiveXControl_IndexedPropertyParsing_And_PropertyNotifySink_Events()
    {
        Assert.True(DynamicActiveXControl.TryParseIndexedPropertyExpression("TabCaption(0)", out string base1, out object[] idx1));
        Assert.Equal("TabCaption", base1);
        Assert.Single(idx1);
        Assert.Equal(0, idx1[0]);

        Assert.True(DynamicActiveXControl.TryParseIndexedPropertyExpression("Cell(2, 5)", out string base2, out object[] idx2));
        Assert.Equal("Cell", base2);
        Assert.Equal(2, idx2.Length);
        Assert.Equal(2, idx2[0]);
        Assert.Equal(5, idx2[1]);

        Assert.False(DynamicActiveXControl.TryParseIndexedPropertyExpression("Caption", out _, out _));

        RunInSta(() =>
        {
            using var ctrl = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}");
            var site = new OleDesignClientSite(ctrl);

            int changedDispId = 0;
            ctrl.PropertyChangedByOcx += id => changedDispId = id;

            Assert.Equal(0, site.OnRequestEdit(-518));
            Assert.Equal(0, site.OnChanged(-518));
            Assert.Equal(-518, changedDispId);
        });
    }

    [Fact]
    public void DesignerPropertyFactory_FiltersHiddenNonBrowsableReadOnlyAndIndexedProperties_FromTypeLib()
    {
        var lib = new ViBa7.ProjectModel.ActiveX.OcxLibraryInfo
        {
            LibraryName = "Fase2TestLib",
            FilePath = @"C:\Dummy\Fase2Test.ocx",
            Guid = "{22222222-3333-4444-5555-666666666666}"
        };
        var ctrlInfo = new ViBa7.ProjectModel.ActiveX.OcxControlInfo
        {
            Name = "SuperGrid",
            Clsid = "{77777777-8888-9999-AAAA-BBBBBBBBBBBB}",
            IsControl = true
        };
        ctrlInfo.Members.Add(new ViBa7.ProjectModel.ActiveX.OcxMemberInfo
        {
            Name = "GridLines",
            Kind = ViBa7.ProjectModel.IntelliSense.VbSymbolKind.Property,
            ReturnType = "Boolean",
            DispId = 101,
            IsHidden = false,
            IsBrowsable = true,
            IsPropertyGet = true,
            IsPropertyPut = true
        });
        ctrlInfo.Members.Add(new ViBa7.ProjectModel.ActiveX.OcxMemberInfo
        {
            Name = "InternalHandle",
            Kind = ViBa7.ProjectModel.IntelliSense.VbSymbolKind.Property,
            ReturnType = "Long",
            DispId = 102,
            IsHidden = true,
            IsBrowsable = false,
            IsPropertyGet = true,
            IsPropertyPut = true
        });
        ctrlInfo.Members.Add(new ViBa7.ProjectModel.ActiveX.OcxMemberInfo
        {
            Name = "RuntimeOnlyProp",
            Kind = ViBa7.ProjectModel.IntelliSense.VbSymbolKind.Property,
            ReturnType = "Long",
            DispId = 103,
            IsHidden = false,
            IsBrowsable = false, // FNONBROWSABLE
            IsPropertyGet = true,
            IsPropertyPut = true
        });
        ctrlInfo.Members.Add(new ViBa7.ProjectModel.ActiveX.OcxMemberInfo
        {
            Name = "RowCountReadOnly",
            Kind = ViBa7.ProjectModel.IntelliSense.VbSymbolKind.Property,
            ReturnType = "Long",
            DispId = 104,
            IsHidden = false,
            IsBrowsable = true,
            IsPropertyGet = true,
            IsPropertyPut = false
        });
        var indexedProp = new ViBa7.ProjectModel.ActiveX.OcxMemberInfo
        {
            Name = "ColWidth",
            Kind = ViBa7.ProjectModel.IntelliSense.VbSymbolKind.Property,
            ReturnType = "Long",
            DispId = 105,
            IsHidden = false,
            IsBrowsable = true,
            IsPropertyGet = true,
            IsPropertyPut = true
        };
        indexedProp.Parameters.Add(new ViBa7.ProjectModel.ActiveX.OcxParameterInfo { Name = "Index", Type = "Integer", IsOptional = false });
        ctrlInfo.Members.Add(indexedProp);

        lib.Controls.Add(ctrlInfo);
        ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.RegisterOcxLibrary(lib);
        try
        {
            Assert.True(ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.TryResolvePropertyNameByDispId("Fase2TestLib.SuperGrid", 101, out string resolvedName));
            Assert.Equal("GridLines", resolvedName);

            var formCtrl = new FormDesignerControl
            {
                TypeName = "Fase2TestLib.SuperGrid",
                Name = "SuperGrid1"
            };
            var gridProps = DesignerPropertyFactory.GetPropertiesForControl(formCtrl, null, () => { });

            Assert.Contains(gridProps, p => p.Name == "GridLines");
            Assert.DoesNotContain(gridProps, p => p.Name == "InternalHandle");
            Assert.DoesNotContain(gridProps, p => p.Name == "RuntimeOnlyProp");
            Assert.DoesNotContain(gridProps, p => p.Name == "RowCountReadOnly");
            Assert.DoesNotContain(gridProps, p => p.Name == "ColWidth");
        }
        finally
        {
            ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.UnregisterOcxLibrary("Fase2TestLib");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // FASE 3: Alojamiento Interactivo en Modo Ejecución (ActiveXRuntimeHost) y Puente de Automatización
    // ═══════════════════════════════════════════════════════════════════════════

    [Fact]
    public void ActiveXRuntimeHost_InitializesInUserMode_And_ImplementsSimpleFrameSite()
    {
        RunInSta(() =>
        {
            var ctrlModel = new FormDesignerControl
            {
                TypeName = "MSFlexGridLib.MSFlexGrid",
                Name = "GridRuntime1",
                Left = 900,
                Top = 1125,
                Width = 4500,
                Height = 3000,
                TabIndex = 4,
                TabStop = true,
                LicenseKey = "TestKey",
                ActiveXStateBlob = new byte[] { 0x10, 0x20, 0x30, 0x40 }
            };
            ctrlModel.SetCustomProperty("Tag", "TagRuntime");
            ctrlModel.SetCustomProperty("ToolTipText", "TipRuntime");
            ctrlModel.SetCustomProperty("Rows", "10");
            ctrlModel.SetCustomProperty("Cols", "5");

            using var runtimeHost = new ActiveXRuntimeHost("{6262D3A0-5316-11CE-BFC1-08002BE10318}", ctrlModel.TypeName, 300, 200)
            {
                BoundControl = ctrlModel
            };
            Assert.True(runtimeHost.TryInitializeRuntime());

            Assert.NotNull(runtimeHost.AxControl);
            Assert.False(runtimeHost.AxControl.IsDesignMode);
            Assert.Equal("GridRuntime1", runtimeHost.AxControl.Extender.Name);
            Assert.Equal(4, runtimeHost.AxControl.Extender.TabIndex);
            Assert.Equal("TagRuntime", runtimeHost.AxControl.Extender.Tag);
            Assert.Equal("TipRuntime", runtimeHost.AxControl.Extender.ToolTipText);
            Assert.Equal("TestKey", runtimeHost.AxControl.LicenseKey);
            Assert.Equal(new byte[] { 0x10, 0x20, 0x30, 0x40 }, runtimeHost.AxControl.InitialStreamBlob);

            var site = new OleDesignClientSite(runtimeHost.AxControl);
            Assert.True(site.UserMode);

            // Verificar implementación de ISimpleFrameSite (PreMessageFilter / PostMessageFilter)
            var simpleFrame = Assert.IsAssignableFrom<ViBa7.ActiveXHost.Native.ISimpleFrameSite>(site);
            Assert.Equal(0, simpleFrame.PreMessageFilter(IntPtr.Zero, 0x0100, IntPtr.Zero, IntPtr.Zero, out IntPtr lResultPre, out uint cookie));
            Assert.Equal(0, simpleFrame.PostMessageFilter(IntPtr.Zero, 0x0100, IntPtr.Zero, IntPtr.Zero, out IntPtr lResultPost, cookie));
        });
    }

    private sealed class TestRuntimeActiveXControlProxy : ViBa7.ProjectModel.Execution.ILiveControl
    {
        private readonly ActiveXRuntimeHost _host;
        private readonly Dictionary<string, ViBa7.ProjectModel.Execution.VbValue> _props = new(StringComparer.OrdinalIgnoreCase);

        public TestRuntimeActiveXControlProxy(ActiveXRuntimeHost host)
        {
            _host = host;
        }

        public string Name => _host.BoundControl?.Name ?? _host.AxControl.Name;
        public string TypeName => _host.ProgId;

        public ViBa7.ProjectModel.Execution.VbValue GetProperty(string propertyName)
        {
            if (propertyName.Equals("Left", StringComparison.OrdinalIgnoreCase))
                return ViBa7.ProjectModel.Execution.VbValue.FromDouble(_host.AxControl.Extender.Left);
            if (propertyName.Equals("Top", StringComparison.OrdinalIgnoreCase))
                return ViBa7.ProjectModel.Execution.VbValue.FromDouble(_host.AxControl.Extender.Top);
            if (propertyName.Equals("Width", StringComparison.OrdinalIgnoreCase))
                return ViBa7.ProjectModel.Execution.VbValue.FromDouble(_host.AxControl.Extender.Width);
            if (propertyName.Equals("Height", StringComparison.OrdinalIgnoreCase))
                return ViBa7.ProjectModel.Execution.VbValue.FromDouble(_host.AxControl.Extender.Height);
            if (propertyName.Equals("Tag", StringComparison.OrdinalIgnoreCase))
                return ViBa7.ProjectModel.Execution.VbValue.FromString(_host.AxControl.Extender.Tag);

            return _props.TryGetValue(propertyName, out var v) ? v : ViBa7.ProjectModel.Execution.VbValue.Empty;
        }

        public void SetProperty(string propertyName, ViBa7.ProjectModel.Execution.VbValue value)
        {
            if (propertyName.Equals("Tag", StringComparison.OrdinalIgnoreCase))
            {
                _host.AxControl.Extender.Tag = value.AsString();
                return;
            }
            _props[propertyName] = value;
            _host.AxControl.SetPropertyValue(propertyName, value.RawValue);
        }

        public void SetIndexedProperty(string propertyName, ViBa7.ProjectModel.Execution.VbValue[] indices, ViBa7.ProjectModel.Execution.VbValue value)
        {
            string key = $"{propertyName}({string.Join(",", Array.ConvertAll(indices, i => i.AsString()))})";
            _props[key] = value;
            object?[] rawIdx = Array.ConvertAll(indices, i => i.RawValue);
            _host.AxControl.SetIndexedPropertyValue(propertyName, value.RawValue, rawIdx!);
        }

        public ViBa7.ProjectModel.Execution.VbValue InvokeMethod(string methodName, ViBa7.ProjectModel.Execution.VbValue[] args)
        {
            if (methodName.Equals("Move", StringComparison.OrdinalIgnoreCase))
            {
                float left = args.Length >= 1 && !args[0].IsEmpty ? (float)args[0].AsDouble() : _host.AxControl.Extender.Left;
                float top = args.Length >= 2 && !args[1].IsEmpty ? (float)args[1].AsDouble() : _host.AxControl.Extender.Top;
                float width = args.Length >= 3 && !args[2].IsEmpty ? (float)args[2].AsDouble() : _host.AxControl.Extender.Width;
                float height = args.Length >= 4 && !args[3].IsEmpty ? (float)args[3].AsDouble() : _host.AxControl.Extender.Height;
                _host.AxControl.Extender.Move(left, top, width, height);
                return ViBa7.ProjectModel.Execution.VbValue.Empty;
            }

            string key = $"{methodName}({string.Join(",", Array.ConvertAll(args, a => a.AsString()))})";
            if (_props.TryGetValue(key, out var cached))
                return cached;

            return ViBa7.ProjectModel.Execution.VbValue.Empty;
        }
    }

    [Fact]
    public void VbStatementInterpreter_And_DynamicActiveXControl_ExecuteIndexedProperties_Move_And_ComSubObjectDispatch()
    {
        RunInSta(() =>
        {
            var ctrlModel = new FormDesignerControl
            {
                TypeName = "MSFlexGridLib.MSFlexGrid",
                Name = "MSFlexGrid1",
                Left = 150,
                Top = 300,
                Width = 3000,
                Height = 1800
            };

            using var runtimeHost = new ActiveXRuntimeHost("{6262D3A0-5316-11CE-BFC1-08002BE10318}", ctrlModel.TypeName, 200, 120)
            {
                BoundControl = ctrlModel
            };
            Assert.True(runtimeHost.TryInitializeRuntime());

            var proxy = new TestRuntimeActiveXControlProxy(runtimeHost);
            var ctx = new ViBa7.ProjectModel.Execution.VbExecutionContext();
            ctx.ModuleVariables["MSFlexGrid1"] = ViBa7.ProjectModel.Execution.VbValue.FromObject(proxy);

            var lines = new[]
            {
                @"MSFlexGrid1.TextMatrix(1, 2) = ""Celda_1_2""",
                @"MSFlexGrid1.ColWidth(2) = 1800",
                @"Dim val As String",
                @"val = MSFlexGrid1.TextMatrix(1, 2)",
                @"Dim cw As Long",
                @"cw = MSFlexGrid1.ColWidth(2)",
                @"MSFlexGrid1.Move 300, 450, 3000, 1500",
                @"MSFlexGrid1.Tag = val & ""|"" & CStr(cw)"
            };

            var interpreter = new ViBa7.ProjectModel.Execution.VbStatementInterpreter(ctx);
            interpreter.ExecuteBlock(lines);

            Assert.Equal("Celda_1_2|1800", proxy.GetProperty("Tag").AsString());
            Assert.Equal(300.0, proxy.GetProperty("Left").AsDouble());
            Assert.Equal(450.0, proxy.GetProperty("Top").AsDouble());
            Assert.Equal(3000.0, proxy.GetProperty("Width").AsDouble());
            Assert.Equal(1500.0, proxy.GetProperty("Height").AsDouble());

            // Verificar puente IDispatch sobre sub-objetos COM reales (StdFont via OleCreateFontIndirect)
            object? stdFontCom = ViBa7.ActiveXHost.Native.Ole32Interop.CreateStdFont("Arial", 10f, bold: false, italic: false, underline: false, strikethrough: false, charset: 0);
            Assert.NotNull(stdFontCom);
            try
            {
                Assert.Equal("Arial", DynamicActiveXControl.InvokeComDispatchMember(stdFontCom!, "Name")?.ToString());
                Assert.True(DynamicActiveXControl.SetComDispatchProperty(stdFontCom!, "Name", "Courier New"));
                Assert.True(DynamicActiveXControl.SetComDispatchProperty(stdFontCom!, "Bold", true));
                Assert.Equal("Courier New", DynamicActiveXControl.InvokeComDispatchMember(stdFontCom!, "Name")?.ToString());
                Assert.Equal(true, DynamicActiveXControl.InvokeComDispatchMember(stdFontCom!, "Bold"));
            }
            finally
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(stdFontCom!))
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(stdFontCom!);
            }
        });
    }

    [Fact]
    public void ActiveXEventSink_Invoke_UnpacksReverseArgumentsAndFiresEvent()
    {
        using var sink = new ActiveXEventSink(
            eventIid: Guid.NewGuid(),
            controlName: "MSFlexGrid1",
            typeName: "MSFlexGridLib.MSFlexGrid");

        ActiveXEventInvocation? received = null;
        sink.EventTriggered += inv => received = inv;

        // Simulamos DISPPARAMS con 2 argumentos en orden inverso:
        // Firma lógica COM: Event_100(arg0 As Long, arg1 As String)
        // OLE ABI orden inverso:
        // rgvarg[0] = "Hola Mundo" (arg1)
        // rgvarg[1] = 12345 (arg0)
        int argCount = 2;
        int variantSize = 16;
        IntPtr pRgvarg = Marshal.AllocHGlobal(argCount * variantSize);
        try
        {
            // rgvarg[0] = VT_BSTR "Hola Mundo"
            IntPtr pVar0 = pRgvarg;
            Marshal.GetNativeVariantForObject("Hola Mundo", pVar0);

            // rgvarg[1] = VT_I4 12345
            IntPtr pVar1 = pRgvarg + variantSize;
            Marshal.GetNativeVariantForObject(12345, pVar1);

            var dispParams = new DISPPARAMS
            {
                rgvarg = pRgvarg,
                cArgs = (uint)argCount,
                cNamedArgs = 0,
                rgdispidNamedArgs = IntPtr.Zero
            };

            Guid iidNull = Guid.Empty;
            int hr = sink.Invoke(100, ref iidNull, 0, 1, ref dispParams, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.Equal(0, hr);
            Assert.NotNull(received);
            Assert.Equal(100, received!.DispId);
            Assert.Equal(2, received.Arguments.Length);
            // Orden corregido: arg0 = 12345, arg1 = "Hola Mundo"
            Assert.Equal(12345, received.Arguments[0].AsLong());
            Assert.Equal("Hola Mundo", received.Arguments[1].AsString());
        }
        finally
        {
            Marshal.FreeHGlobal(pRgvarg);
        }
    }

    [Fact]
    public void ActiveXEventSink_ByRef_ModifiesNativeCallerPointer()
    {
        using var sink = new ActiveXEventSink(
            eventIid: Guid.NewGuid(),
            controlName: "Text1",
            typeName: "VB.TextBox");

        ActiveXEventInvocation? received = null;
        sink.EventTriggered += inv => received = inv;

        // Simulamos KeyPress(KeyAscii As Integer ByRef)
        // DISPID_KEYPRESS = -603
        IntPtr pNativeKey = Marshal.AllocHGlobal(sizeof(short));
        Marshal.WriteInt16(pNativeKey, 65); // ASCII 'A'

        int variantSize = 16;
        IntPtr pRgvarg = Marshal.AllocHGlobal(variantSize);
        try
        {
            // VARIANT con VT_BYREF | VT_I2
            Marshal.WriteInt16(pRgvarg, 0, unchecked((short)((ushort)VarEnum.VT_BYREF | (ushort)VarEnum.VT_I2)));
            Marshal.WriteIntPtr(pRgvarg, 8, pNativeKey);

            var dispParams = new DISPPARAMS
            {
                rgvarg = pRgvarg,
                cArgs = 1,
                cNamedArgs = 0,
                rgdispidNamedArgs = IntPtr.Zero
            };

            Guid iidNull = Guid.Empty;
            sink.Invoke(Ole32Interop.DISPID_KEYPRESS, ref iidNull, 0, 1, ref dispParams, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

            Assert.NotNull(received);
            Assert.Equal("KeyPress", received!.EventName);
            Assert.Single(received.Arguments);
            Assert.Equal(65, received.Arguments[0].AsLong());
            Assert.True(received.IsByRef[0]);

            // Simulamos cancelación en VB6: KeyAscii = 0
            received.UpdateArgumentValue(0, VbValue.FromLong(0));

            // Verificar que la memoria nativa del llamador se actualizó a 0
            Assert.Equal((short)0, Marshal.ReadInt16(pNativeKey));
        }
        finally
        {
            Marshal.FreeHGlobal(pRgvarg);
            Marshal.FreeHGlobal(pNativeKey);
        }
    }

    [Fact]
    public void ActiveXEventSink_ControlArray_InjectsIndexAsFirstArgument()
    {
        using var sink = new ActiveXEventSink(
            eventIid: Guid.NewGuid(),
            controlName: "Command1",
            typeName: "VB.CommandButton",
            controlIndex: 4);

        ActiveXEventInvocation? received = null;
        sink.EventTriggered += inv => received = inv;

        var dispParams = new DISPPARAMS
        {
            rgvarg = IntPtr.Zero,
            cArgs = 0,
            cNamedArgs = 0,
            rgdispidNamedArgs = IntPtr.Zero
        };

        Guid iidNull = Guid.Empty;
        sink.Invoke(Ole32Interop.DISPID_CLICK, ref iidNull, 0, 1, ref dispParams, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

        Assert.NotNull(received);
        Assert.Equal("Click", received!.EventName);
        Assert.Single(received.Arguments);
        // Argumento 0 inyectado: Index As Integer = 4
        Assert.Equal(4, received.Arguments[0].AsLong());
    }

    [Fact]
    public void LiveExecutionHost_InvokeProcedure_PropagatesByRefModifications()
    {
        var host = new LiveExecutionHost();
        string vbCode = @"
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
    Shift = 7
End Sub
";
        host.LoadSourceCode(vbCode, "Form1");

        var args = new[] { VbValue.FromLong(13), VbValue.FromLong(1) };
        host.InvokeProcedure("Grid1_KeyDown", "Form1", args);

        Assert.Equal(0, args[0].AsLong());
        Assert.Equal(7, args[1].AsLong());
    }

    [Fact]
    public void VbControlMemberCatalog_DefaultEvents_GenerateExactSignatures()
    {
        Assert.True(VbControlMemberCatalog.TryGetDefaultEvent("MSFlexGridLib.MSFlexGrid", out var evFlex, out var pFlex));
        Assert.Equal("RowColChange", evFlex);
        Assert.Empty(pFlex);

        Assert.True(VbControlMemberCatalog.TryGetDefaultEvent("MSComctlLib.TreeCtrl", out var evTree, out var pTree));
        Assert.Equal("NodeClick", evTree);
        Assert.Single(pTree);
        Assert.Equal("Node", pTree[0].Name);
        Assert.Equal("Node", pTree[0].Type);
        Assert.False(pTree[0].IsByRef);

        Assert.True(VbControlMemberCatalog.TryGetDefaultEvent("MSComctlLib.ListViewCtrl", out var evList, out var pList));
        Assert.Equal("ItemClick", evList);
        Assert.Single(pList);
        Assert.Equal("Item", pList[0].Name);
        Assert.Equal("ListItem", pList[0].Type);

        Assert.True(VbControlMemberCatalog.TryGetDefaultEvent("TabDlg.SSTab", out var evTab, out var pTab));
        Assert.Equal("Click", evTab);
        Assert.Single(pTab);
        Assert.Equal("PreviousTab", pTab[0].Name);
        Assert.Equal("Integer", pTab[0].Type);

        Assert.True(VbControlMemberCatalog.TryGetDefaultEvent("MSWinsockLib.Winsock", out var evSock, out var pSock));
        Assert.Equal("DataArrival", evSock);
        Assert.Single(pSock);
        Assert.Equal("bytesTotal", pSock[0].Name);
        Assert.Equal("Long", pSock[0].Type);
    }

    [Fact]
    public void VbPropertyBagRuntime_ReadWriteAndSerialization()
    {
        var bag = new VbPropertyBagRuntime();
        bag.WriteProperty("Caption", VbValue.FromString("Botón de Prueba"));
        bag.WriteProperty("Value", VbValue.FromLong(100));

        Assert.Equal("Botón de Prueba", bag.ReadProperty("Caption").AsString());
        Assert.Equal(100, bag.ReadProperty("Value").AsLong());
        Assert.Equal("PorDefecto", bag.ReadProperty("NoExiste", VbValue.FromString("PorDefecto")).AsString());

        // Probar invocación dinámica como ILiveControl (sintaxis VB6 PropBag.ReadProperty / WriteProperty)
        var readResult = bag.InvokeMethod("ReadProperty", new[] { VbValue.FromString("Caption") });
        Assert.Equal("Botón de Prueba", readResult.AsString());

        bag.InvokeMethod("WriteProperty", new[] { VbValue.FromString("Color"), VbValue.FromLong(0xFF00) });
        Assert.Equal(0xFF00, bag.ReadProperty("Color").AsLong());

        // Serialización y deserialización binaria de Contents
        byte[] bytes = bag.SerializeToBytes();
        Assert.NotEmpty(bytes);

        var bag2 = new VbPropertyBagRuntime();
        bag2.DeserializeFromBytes(bytes);
        Assert.Equal("Botón de Prueba", bag2.ReadProperty("Caption").AsString());
        Assert.Equal(100, bag2.ReadProperty("Value").AsLong());
        Assert.Equal(0xFF00, bag2.ReadProperty("Color").AsLong());
    }

    [Fact]
    public void VbUserControlAmbient_And_Extender_Properties()
    {
        var ambient = new VbUserControlAmbient(isUserMode: true, displayName: "MiControl1");
        Assert.True(ambient.GetProperty("UserMode").AsBoolean());
        Assert.Equal("MiControl1", ambient.GetProperty("DisplayName").AsString());
        Assert.Equal("Twip", ambient.GetProperty("ScaleUnits").AsString());
        Assert.False(ambient.GetProperty("UIDead").AsBoolean());

        var ctrl = new FormDesignerControl
        {
            Name = "ctlCustom1",
            TypeName = "Proyecto1.MiControl",
            Left = 100,
            Top = 200,
            Width = 300,
            Height = 400,
            Visible = true,
            Enabled = true
        };

        var extender = new VbUserControlExtender(ctrl);
        Assert.Equal("ctlCustom1", extender.GetProperty("Name").AsString());
        Assert.Equal(100.0, extender.GetProperty("Left").AsDouble());
        Assert.Equal(200.0, extender.GetProperty("Top").AsDouble());

        extender.InvokeMethod("Move", new[] { VbValue.FromDouble(500), VbValue.FromDouble(600) });
        Assert.Equal(500, ctrl.Left);
        Assert.Equal(600, ctrl.Top);
    }

    [Fact]
    public void VbStatementInterpreter_RaiseEvent_Execution()
    {
        var host = new LiveExecutionHost();
        string vbCode = @"
Public Event ValueChanged(NewVal As Long, Msg As String)

Public Sub TestRaise()
    RaiseEvent ValueChanged(42, ""Exito"")
End Sub
";
        host.LoadSourceCode(vbCode, "UserControl1");

        string? firedEvent = null;
        VbValue[]? firedArgs = null;
        string? firedContainer = null;

        host.Context.OnRaiseEvent += (evName, args, container, activeInstance) =>
        {
            firedEvent = evName;
            firedArgs = args;
            firedContainer = container;
        };

        host.InvokeProcedure("TestRaise", "UserControl1");

        Assert.Equal("ValueChanged", firedEvent);
        Assert.NotNull(firedArgs);
        Assert.Equal(2, firedArgs!.Length);
        Assert.Equal(42, firedArgs[0].AsLong());
        Assert.Equal("Exito", firedArgs[1].AsString());
        Assert.Equal("UserControl1", firedContainer);
    }

    [Fact]
    public void Vb6ManifestGenerator_GeneratesValidSxSManifest()
    {
        var options = new Vb6CompileOptions(
            OutputExePath: @"C:\MiProyecto\App.exe",
            ProjectName: "AppPrueba",
            MajorVersion: 1,
            MinorVersion: 2,
            RevisionVersion: 3);

        var libInfo = new OcxLibraryInfo
        {
            FilePath = @"C:\MiProyecto\msflxgrd.ocx",
            LibraryName = "MSFlexGridLib",
            Guid = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",
            MajorVersion = 1,
            MinorVersion = 0
        };

        libInfo.Controls.Add(new OcxControlInfo
        {
            Name = "MSFlexGrid",
            Clsid = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",
            Description = "Microsoft FlexGrid Control 6.0"
        });

        string manifestXml = Vb6ManifestGenerator.GenerateManifestXml(options, new[] { libInfo });

        Assert.Contains("<?xml version=\"1.0\"", manifestXml);
        Assert.Contains("<assemblyIdentity", manifestXml);
        Assert.Contains("name=\"AppPrueba\"", manifestXml);
        Assert.Contains("version=\"1.2.3.0\"", manifestXml);
        Assert.Contains("<file name=\"msflxgrd.ocx\">", manifestXml);
        Assert.Contains("clsid=\"{6262D3A0-5316-11CE-BFC1-08002BE10318}\"", manifestXml);
        Assert.Contains("threadingModel=\"Apartment\"", manifestXml);
        Assert.Contains("<typelib tlbid=\"{6262D3A0-5316-11CE-BFC1-08002BE10318}\"", manifestXml);
    }

    [Fact]
    public void Vb6ProjectCompiler_EmitsExecutableAndSxSManifest()
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "ViBa7_Test_" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(tempDir);
        try
        {
            var project = new VB6Project
            {
                Name = "TestExeWithOcx",
                ProjectFilePath = Path.Combine(tempDir, "TestExeWithOcx.vbp"),
                Startup = "Form1"
            };

            var libInfo = new OcxLibraryInfo
            {
                FilePath = Path.Combine(tempDir, "testctl.ocx"),
                LibraryName = "TestLib",
                Guid = "{11223344-5566-7788-99AA-BBCCDDEEFF00}",
                MajorVersion = 1,
                MinorVersion = 0
            };
            libInfo.Controls.Add(new OcxControlInfo
            {
                Name = "TestCtrl",
                Clsid = "{11223344-5566-7788-99AA-BBCCDDEEFF00}",
                Description = "Control de Prueba"
            });

            project.AddActiveXControl(libInfo, "testctl.ocx");

            string exePath = Path.Combine(tempDir, "TestExeWithOcx.exe");
            var compiler = new Vb6ProjectCompiler();
            var res = compiler.Compile(project, exePath);

            Assert.True(res.Success);
            Assert.True(File.Exists(exePath));
            Assert.True(res.OutputSizeBytes > 0);

            // Verificar que se emitió el manifiesto SxS
            string manifestPath = exePath + ".manifest";
            Assert.True(File.Exists(manifestPath));
            string manifestText = File.ReadAllText(manifestPath);
            Assert.Contains("testctl.ocx", manifestText);
            Assert.Contains("{11223344-5566-7788-99AA-BBCCDDEEFF00}", manifestText);
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { }
        }
    }

    [Fact]
    public void FontName_And_Font_Properties_InferredCorrectly_WithTypographies()
    {
        // 1. FontName debe ser inferido como EnumChoice con lista de tipografías del sistema
        var (fontNameType, fontNameChoices) = DesignerPropertyFactory.InferPropertyType("FontName", "String", "Tahoma");
        Assert.Equal(DesignerPropertyType.EnumChoice, fontNameType);
        Assert.NotEmpty(fontNameChoices);
        Assert.Contains(fontNameChoices, c => c.DisplayText.Contains("Arial", StringComparison.OrdinalIgnoreCase) ||
                                              c.DisplayText.Contains("MS Sans Serif", StringComparison.OrdinalIgnoreCase) ||
                                              c.DisplayText.Contains("Tahoma", StringComparison.OrdinalIgnoreCase));

        // 2. Font / StdFont debe ser inferido como FontDescription
        var (fontType, fontChoices) = DesignerPropertyFactory.InferPropertyType("Font", "StdFont", null);
        Assert.Equal(DesignerPropertyType.FontDescription, fontType);
        Assert.Empty(fontChoices);

        // 3. SynchronizeStandardProperties debe actualizar ctrl.Font
        var ctrl = new FormDesignerControl { Name = "TestLabel", TypeName = "VB.Label" };
        DesignerPropertyFactory.SynchronizeStandardProperties(ctrl, "FontName", "Segoe UI");
        DesignerPropertyFactory.SynchronizeStandardProperties(ctrl, "FontSize", "12");
        DesignerPropertyFactory.SynchronizeStandardProperties(ctrl, "FontBold", "True");

        Assert.NotNull(ctrl.Font);
        Assert.Equal("Segoe UI", ctrl.Font.Name);
        Assert.Equal(12.0, ctrl.Font.Size);
        Assert.True(ctrl.Font.IsBold);
    }

    [Fact]
    public void VbControlExtender_SetBoundsDirect_DoesNotFireNotification_And_PreservesCoordinates()
    {
        var dummyAx = new DynamicActiveXControl("{00000000-0000-0000-0000-000000000000}", 100, 50);
        bool fired = false;
        dummyAx.ExtenderBoundsChanged += () => fired = true;

        dummyAx.Extender.SetBoundsDirect(1200f, 2400f, 3600f, 1800f);

        Assert.False(fired, "SetBoundsDirect no debe disparar ExtenderBoundsChanged para evitar ciclos destructivos");
        Assert.Equal(1200f, dummyAx.Extender.Left);
        Assert.Equal(2400f, dummyAx.Extender.Top);
        Assert.True(dummyAx.Extender.Width > 0);
        Assert.True(dummyAx.Extender.Height > 0);
    }

    [Fact]
    public void ActiveXInPlaceHost_ExtenderBoundsChanged_GuardedDuringSync_PreservesTopCoordinate()
    {
        RunInSta(() =>
        {
            var ctrl = new FormDesignerControl
            {
                Name = "Grid1",
                TypeName = "MSFlexGridLib.MSFlexGrid",
                Left = 500,
                Top = 3200,
                Width = 4000,
                Height = 2500
            };

            using var host = new ActiveXInPlaceHost("{6262D3A0-5316-11CE-B3E4-00AA006B1A69}", "MSFlexGridLib.MSFlexGrid", 200, 150)
            {
                BoundControl = ctrl
            };

            // Simular que el Extender actualiza coordenadas con SetBoundsDirect
            host.AxControl.Extender.SetBoundsDirect(ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height);

            // Las coordenadas del control de diseño deben permanecer intactas, especialmente Top != 0
            Assert.Equal(3200, ctrl.Top);
            Assert.Equal(500, ctrl.Left);
            Assert.Equal(4000, ctrl.Width);
            Assert.Equal(2500, ctrl.Height);
        });
    }
}



