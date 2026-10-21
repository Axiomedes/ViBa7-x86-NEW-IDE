using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using ViBa7.ActiveXHost.Diagnostics;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.IntelliSense;

namespace ViBa7.ActiveXHost;

/// <summary>
/// Fábrica de contenedores ActiveX In-Place para la superficie de diseño de ViBa7.
/// Resuelve ProgID, CLSID y estado de registro en el subsistema Win32 COM.
/// </summary>
public static class ActiveXHostFactory
{
    /// <summary>
    /// Habilita o deshabilita globalmente el hospedaje In-Place nativo de controles ActiveX.
    /// Si está deshabilitado, se utiliza el renderizador simulado de WPF.
    /// </summary>
    public static bool EnableInPlaceActiveXHosting { get; set; } = true;

    // Catálogo de CLSIDs conocidos de controles ActiveX clásicos de VB6
    private static readonly Dictionary<string, string> KnownClsids = new(StringComparer.OrdinalIgnoreCase)
    {
        // MSFlexGrid (MSFLXGRD.OCX)
        ["MSFlexGridLib.MSFlexGrid"] = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",
        ["MSFlexGrid.MSFlexGrid"] = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",
        ["MSFlexGrid"] = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",

        // MSHFlexGrid (MSHFLXGD.OCX)
        ["MSHierarchicalFlexGridLib.MSHFlexGrid"] = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}",
        ["MSHFlexGridLib.MSHFlexGrid"] = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}",
        ["MSHFlexGrid"] = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}",

        // Microsoft Windows Common Controls 6.0 (MSCOMCTL.OCX)
        ["MSComctlLib.TreeCtrl"] = "{C74190B6-8589-11D1-B16A-00C0F0283628}",
        ["MSComctlLib.TreeView"] = "{C74190B6-8589-11D1-B16A-00C0F0283628}",
        ["TreeView"] = "{C74190B6-8589-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.ListViewCtrl"] = "{BDD1F04B-858B-11D1-B16A-00C0F0283628}",
        ["MSComctlLib.ListView"] = "{BDD1F04B-858B-11D1-B16A-00C0F0283628}",
        ["ListView"] = "{BDD1F04B-858B-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.TabStrip"] = "{1EFB6596-857C-11D1-B16A-00C0F0283628}",
        ["TabStrip"] = "{1EFB6596-857C-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.StatusBar"] = "{8E3867A3-8586-11D1-B16A-00C0F0283628}",
        ["StatusBar"] = "{8E3867A3-8586-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.ProgBar"] = "{35053A22-8589-11D1-B16A-00C0F0283628}",
        ["MSComctlLib.ProgressBar"] = "{35053A22-8589-11D1-B16A-00C0F0283628}",
        ["ProgressBar"] = "{35053A22-8589-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.Slider"] = "{F08DF954-8592-11D1-B16A-00C0F0283628}",
        ["Slider"] = "{F08DF954-8592-11D1-B16A-00C0F0283628}",

        ["MSComctlLib.ImageListCtrl"] = "{2C247F23-8591-11D1-B16A-00C0F0283628}",
        ["ImageList"] = "{2C247F23-8591-11D1-B16A-00C0F0283628}",

        // Microsoft Windows Common Controls-2 6.0 (MSCOMCT2.OCX)
        ["MSComCtl2.DTPicker"] = "{20DD94A8-63C2-11CE-8028-00AA006A09FA}",
        ["DTPicker"] = "{20DD94A8-63C2-11CE-8028-00AA006A09FA}",

        ["MSComCtl2.MonthView"] = "{232E456A-87C3-11D1-8BE3-0000F8754DA1}",
        ["MonthView"] = "{232E456A-87C3-11D1-8BE3-0000F8754DA1}",

        ["MSComCtl2.UpDown"] = "{603C7E80-87C2-11D1-8BE3-0000F8754DA1}",
        ["UpDown"] = "{603C7E80-87C2-11D1-8BE3-0000F8754DA1}",

        ["MSComCtl2.Animation"] = "{B09DE715-87C1-11D1-8BE3-0000F8754DA1}",

        // RichTextBox (RICHTX32.OCX)
        ["RichTextLib.RichTextBox"] = "{3B7C8860-D78F-101B-B9B5-04021C009402}",
        ["RichTextBox"] = "{3B7C8860-D78F-101B-B9B5-04021C009402}",

        // Tabbed Dialog (TABCTL32.OCX)
        ["TabDlg.SSTab"] = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}",
        ["SSTab"] = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}",

        // CommonDialog (COMDLG32.OCX)
        ["MSComDlg.CommonDialog"] = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}",
        ["CommonDialog"] = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}",

        // MaskEdBox (MSMASK32.OCX)
        ["MSMask.MaskEdBox"] = "{C932BA85-4374-101B-A56C-00AA003668DC}",
        ["MaskEdBox"] = "{C932BA85-4374-101B-A56C-00AA003668DC}",

        // WebBrowser (SHDOCVW / IEFRAME)
        ["SHDocVw.WebBrowser"] = "{8856F961-340A-11D0-A96B-00C04FD705A2}",
        ["WebBrowser"] = "{8856F961-340A-11D0-A96B-00C04FD705A2}",

        // Winsock (MSWINSCK.OCX)
        ["MSWinsockLib.Winsock"] = "{248DD896-BB45-11CF-9ABC-0080C7E7B78D}",
        ["Winsock"] = "{248DD896-BB45-11CF-9ABC-0080C7E7B78D}",

        // MSComm (MSCOMM32.OCX)
        ["MSCommLib.MSComm"] = "{648A5600-2C6E-101B-82B6-000000000014}",
        ["MSComm"] = "{648A5600-2C6E-101B-82B6-000000000014}",

        // Internet Transfer Control (MSINET.OCX)
        ["InetCtlsObjects.Inet"] = "{48E59293-9880-11CF-9754-00AA00C00908}",
        ["Inet"] = "{48E59293-9880-11CF-9754-00AA00C00908}"
    };

    // Controles OCX que son intrínsecamente invisibles en tiempo de ejecución y se representan con un icono fijo en diseño
    private static readonly HashSet<string> NonVisualOcxTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "MSComDlg.CommonDialog",
        "CommonDialog",
        "MSComctlLib.ImageListCtrl",
        "ImageList",
        "MSWinsockLib.Winsock",
        "Winsock",
        "MSCommLib.MSComm",
        "MSComm",
        "InetCtlsObjects.Inet",
        "Inet"
    };

    /// <summary>
    /// Intenta resolver el CLSID de un control por su TypeName o ProgID.
    /// </summary>
    public static bool TryResolveClsid(string typeName, out string clsid)
    {
        clsid = string.Empty;
        if (string.IsNullOrWhiteSpace(typeName)) return false;

        // 1. Si ya tiene formato de GUID/CLSID {XXXXXXXX-...}
        if (typeName.StartsWith('{') && typeName.EndsWith('}') && Guid.TryParse(typeName, out _))
        {
            clsid = typeName;
            return true;
        }

        // 2. Búsqueda en catálogo de CLSIDs conocidos de VB6
        if (KnownClsids.TryGetValue(typeName, out var knownClsid))
        {
            clsid = knownClsid;
            return true;
        }

        // 3. Búsqueda en bibliotecas OCX cargadas dinámicamente en el proyecto (TypeLib real)
        if (VbControlMemberCatalog.TryResolveOcxClsid(typeName, out var ocxClsid))
        {
            clsid = ocxClsid;
            return true;
        }

        // 4. Búsqueda en el Registro de Windows de 32 bits (HKCR\<typeName>\CLSID)
        try
        {
            using var baseKey32 = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32);
            using var progIdKey = baseKey32.OpenSubKey($@"{typeName}\CLSID") ?? Registry.ClassesRoot.OpenSubKey($@"{typeName}\CLSID");
            if (progIdKey?.GetValue(null) is string regClsid && !string.IsNullOrWhiteSpace(regClsid))
            {
                clsid = regClsid;
                return true;
            }

            // Subclave CurVer
            using var curVerKey = baseKey32.OpenSubKey($@"{typeName}\CurVer") ?? Registry.ClassesRoot.OpenSubKey($@"{typeName}\CurVer");
            if (curVerKey?.GetValue(null) is string curVerProgId && !string.IsNullOrWhiteSpace(curVerProgId))
            {
                using var curVerClsidKey = baseKey32.OpenSubKey($@"{curVerProgId}\CLSID") ?? Registry.ClassesRoot.OpenSubKey($@"{curVerProgId}\CLSID");
                if (curVerClsidKey?.GetValue(null) is string curVerClsid && !string.IsNullOrWhiteSpace(curVerClsid))
                {
                    clsid = curVerClsid;
                    return true;
                }
            }
        }
        catch
        {
            // Ignorar errores de acceso al registro
        }

        return false;
    }

    /// <summary>
    /// Verifica si un CLSID está efectivamente registrado en el Registro de Windows
    /// y su archivo servidor (InprocServer32) existe en disco.
    /// </summary>
    public static bool IsControlRegistered(string clsid, out string serverPath)
    {
        serverPath = string.Empty;
        if (string.IsNullOrWhiteSpace(clsid)) return false;

        string cleanClsid = clsid.Trim();
        if (!cleanClsid.StartsWith('{')) cleanClsid = "{" + cleanClsid + "}";

        try
        {
            // Consultar vista de 32 bits (donde residen los OCX)
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry32);
            using var clsidKey = baseKey.OpenSubKey($@"CLSID\{cleanClsid}\InprocServer32");
            if (clsidKey?.GetValue(null) is string path && !string.IsNullOrWhiteSpace(path))
            {
                serverPath = Environment.ExpandEnvironmentVariables(path.Trim('\"'));
                return File.Exists(serverPath);
            }
        }
        catch
        {
            // Ignorar excepciones de lectura de registro
        }

        return false;
    }

    /// <summary>
    /// Intenta crear e instanciar un contenedor ActiveX In-Place para un control de diseño.
    /// Retorna null si el control no es ActiveX, es no-visual, no está registrado o el hospedaje está deshabilitado.
    /// </summary>
    public static ActiveXInPlaceHost? TryCreateHost(FormDesignerControl control, FrameworkElement? fallbackVisual = null)
    {
        if (!EnableInPlaceActiveXHosting) return null;

        // No instanciar controles intrínsecos de VB (VB.TextBox, VB.CommandButton, etc.) como ActiveX
        string normType = control.TypeName.ToUpperInvariant();
        if (normType.StartsWith("VB."))
        {
            return null;
        }

        // No instanciar como HWND visual los OCX que son invisibles en tiempo de ejecución (CommonDialog, ImageList, etc.)
        if (NonVisualOcxTypes.Contains(control.TypeName))
        {
            return null;
        }

        if (!TryResolveClsid(control.TypeName, out string clsid))
        {
            return null;
        }

        if (!IsControlRegistered(clsid, out string serverPath))
        {
            ActiveXDiagnosticsLogger.LogError("Discovery", clsid, unchecked((int)0x80040154), control.Name,
                $"El control '{control.TypeName}' ({clsid}) no está registrado en el subsistema COM x86.");
            return null;
        }

        // Si la biblioteca de tipos del archivo servidor OCX aún no estaba registrada en el catálogo, cargarla
        if (!string.IsNullOrWhiteSpace(serverPath) && File.Exists(serverPath))
        {
            try
            {
                if (!VbControlMemberCatalog.GetRegisteredOcxLibraries().Any(l => string.Equals(l.FilePath, serverPath, StringComparison.OrdinalIgnoreCase)))
                {
                    if (OcxTypeLibReader.TryReadTypeLib(serverPath, out var libInfo, out _) && libInfo != null)
                    {
                        VbControlMemberCatalog.RegisterOcxLibrary(libInfo);
                    }
                }
            }
            catch
            {
                // Ignorar errores en lectura de TypeLib secundaria
            }
        }

        try
        {
            int initW = Math.Max(24, (int)Math.Round(control.WidthPixels));
            int initH = Math.Max(20, (int)Math.Round(control.HeightPixels));

            var host = new ActiveXInPlaceHost(clsid, control.TypeName, initW, initH)
            {
                BoundControl = control,
                Width = control.WidthPixels,
                Height = control.HeightPixels,
                FallbackVisual = fallbackVisual
            };

            bool rendered = host.TryInitializeAndRender();
            if (!rendered && host.AxControl.HasInitializationError)
            {
                ActiveXDiagnosticsLogger.LogError("InPlaceHostInit", clsid, unchecked((int)0x80004005), control.Name,
                    $"Error al inicializar instancia ActiveX in-place para '{control.TypeName}'.");
                host.Dispose();
                return null;
            }

            if (!rendered && fallbackVisual == null)
            {
                host.Dispose();
                return null;
            }

            ActiveXDiagnosticsLogger.LogSuccess("InPlaceHostCreated", clsid, control.Name,
                $"Contenedor in-place creado exitosamente ({initW}x{initH} px, server: {serverPath})");
            return host;
        }
        catch (Exception ex)
        {
            ActiveXDiagnosticsLogger.LogError("InPlaceHostException", clsid, ex.HResult, control.Name, ex.Message);
            return null;
        }
    }

    /// <summary>
    /// Intenta crear e inicializar un contenedor interactivo de tiempo de ejecución (ActiveXRuntimeHost / HwndHost)
    /// con DISPID_AMBIENT_USERMODE = True para la ventana de ejecución F5 (LiveFormWindow).
    /// </summary>
    public static ActiveXRuntimeHost? TryCreateRuntimeHost(FormDesignerControl control, FrameworkElement? fallbackVisual = null)
    {
        if (!EnableInPlaceActiveXHosting) return null;

        string normType = control.TypeName.ToUpperInvariant();
        if (normType.StartsWith("VB."))
        {
            return null;
        }

        if (NonVisualOcxTypes.Contains(control.TypeName))
        {
            return null;
        }

        if (!TryResolveClsid(control.TypeName, out string clsid))
        {
            return null;
        }

        if (!IsControlRegistered(clsid, out string serverPath))
        {
            ActiveXDiagnosticsLogger.LogError("RuntimeDiscovery", clsid, unchecked((int)0x80040154), control.Name,
                $"El control runtime '{control.TypeName}' ({clsid}) no está registrado.");
            return null;
        }

        if (!string.IsNullOrWhiteSpace(serverPath) && File.Exists(serverPath))
        {
            try
            {
                if (!VbControlMemberCatalog.GetRegisteredOcxLibraries().Any(l => string.Equals(l.FilePath, serverPath, StringComparison.OrdinalIgnoreCase)))
                {
                    if (OcxTypeLibReader.TryReadTypeLib(serverPath, out var libInfo, out _) && libInfo != null)
                    {
                        VbControlMemberCatalog.RegisterOcxLibrary(libInfo);
                    }
                }
            }
            catch { }
        }

        try
        {
            int initW = Math.Max(24, (int)Math.Round(control.WidthPixels));
            int initH = Math.Max(20, (int)Math.Round(control.HeightPixels));

            var runtimeHost = new ActiveXRuntimeHost(clsid, control.TypeName, initW, initH)
            {
                BoundControl = control,
                Width = control.WidthPixels,
                Height = control.HeightPixels,
                FallbackVisual = fallbackVisual
            };

            if (!runtimeHost.TryInitializeRuntime())
            {
                ActiveXDiagnosticsLogger.LogError("RuntimeHostInit", clsid, unchecked((int)0x80004005), control.Name,
                    $"Error al inicializar contenedor interactivo runtime para '{control.TypeName}'.");
                runtimeHost.Dispose();
                return null;
            }

            ActiveXDiagnosticsLogger.LogSuccess("RuntimeHostCreated", clsid, control.Name,
                $"Contenedor interactivo runtime inicializado exitosamente ({initW}x{initH} px)");
            return runtimeHost;
        }
        catch (Exception ex)
        {
            ActiveXDiagnosticsLogger.LogError("RuntimeHostException", clsid, ex.HResult, control.Name, ex.Message);
            return null;
        }
    }
}
