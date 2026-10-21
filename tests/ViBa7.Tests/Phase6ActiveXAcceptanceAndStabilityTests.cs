using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using ViBa7.ActiveXHost;
using ViBa7.ActiveXHost.Diagnostics;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Execution;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

/// <summary>
/// Batería de Pruebas de Aceptación y Telemetría de Estabilidad (Fase 6).
/// Implementa las pruebas formales exigidas por la Sección 19 y Hito 8 de Manual_Tecnico_ActiveX_OCX_IDE_x86.txt:
/// - Descubrimiento y catálogo de controles ActiveX de referencia
/// - Telemetría diagnóstica estructurada ({Stage, Clsid, Iid, HResultHex, ErrorInfoDescription})
/// - Verificación de modelo STA de apartamentos COM
/// - Monitor de recursos y handles GDI/USER mediante GetGuiResources
/// - Prueba de estrés de 100 ciclos de creación y liberación sin fuga de handles
/// - Resiliencia ante controles ausentes / no registrados preservando metadatos (.frm/.frx)
/// - Despacho de eventos COM con parámetros ByRef y matrices de controles (Index As Integer)
/// - Fidelidad de transformaciones métricas HIMETRIC &lt;-&gt; Twips
/// </summary>
public class Phase6ActiveXAcceptanceAndStabilityTests
{
    private static void RunInStaThread(Action action)
    {
        Exception? exception = null;
        var thread = new Thread(() =>
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
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception).Throw();
        }
    }

    // =========================================================================
    // 1. Telemetría Diagnóstica Estructurada (Sección 17.1 y Hito 8)
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_DiagnosticsLogger_RecordsStructuredEntriesWithStaApartmentAndHResult()
    {
        RunInStaThread(() =>
        {
            ActiveXDiagnosticsLogger.Clear();

            string testClsid = "{6262D3A0-5316-11CE-BFC1-08002BE10318}"; // MSFlexGrid
            string testName = "MSFlexGrid1";

            // Registrar operaciones exitosas y fallidas simuladas
            ActiveXDiagnosticsLogger.LogSuccess("Discovery", testClsid, testName, "Control localizado en registro x86");
            ActiveXDiagnosticsLogger.LogSuccess("Created", testClsid, testName, "Instancia COM creada");
            ActiveXDiagnosticsLogger.LogError("QueryInterface", testClsid, unchecked((int)0x80004002), testName, "E_NOINTERFACE", iid: "{0000011B-0000-0000-C000-000000000046}");

            var recent = ActiveXDiagnosticsLogger.GetRecentEntries(10);
            Assert.True(recent.Count >= 3);

            var errList = ActiveXDiagnosticsLogger.GetErrors();
            Assert.Contains(errList, e => e.Stage == "QueryInterface" && e.HResultHex == "0x80004002");

            var ctrlEntries = ActiveXDiagnosticsLogger.GetEntriesForControl(testClsid);
            Assert.True(ctrlEntries.Count >= 3);

            // Validar que se registró el hilo como STA
            var lastEntry = ctrlEntries[^1];
            Assert.True(lastEntry.IsStaThread);
            Assert.Equal("QueryInterface", lastEntry.Stage);
            Assert.Equal(unchecked((int)0x80004002), lastEntry.HResult);
            Assert.Equal("0x80004002", lastEntry.HResultHex);
            Assert.Equal("{0000011B-0000-0000-C000-000000000046}", lastEntry.Iid);
        });
    }

    [Fact]
    public void Acceptance_Section19_DiagnosticsLogger_ExportToTextAndJson()
    {
        ActiveXDiagnosticsLogger.Clear();
        ActiveXDiagnosticsLogger.LogSuccess("TestStage", "{11111111-2222-3333-4444-555555555555}", "TestCtrl1", "Detalle de prueba");

        string textReport = ActiveXDiagnosticsLogger.ExportToText();
        Assert.Contains("=== ACTIVEX / OCX DIAGNOSTICS LOG", textReport);
        Assert.Contains("TestStage", textReport);
        Assert.Contains("TestCtrl1", textReport);

        string jsonReport = ActiveXDiagnosticsLogger.ExportToJson();
        Assert.Contains("\"Stage\": \"TestStage\"", jsonReport);
        Assert.Contains("\"Clsid\": \"{11111111-2222-3333-4444-555555555555}\"", jsonReport);
        Assert.Contains("\"HResultHex\": \"0x00000000\"", jsonReport);

        // Validar deserialización de JSON
        using var doc = JsonDocument.Parse(jsonReport);
        Assert.True(doc.RootElement.GetArrayLength() > 0);
    }

    // =========================================================================
    // 2. Monitor de Recursos GDI/USER y Prueba de Estrés de 100 Ciclos (Hito 8)
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_ResourceMonitor_SnapshotAndDeltaCalculations()
    {
        var baseline = ActiveXResourceMonitor.TakeSnapshot();
        Assert.True(baseline.TimestampUtc <= DateTime.UtcNow);

        // Simular creación y liberación de objeto COM en el monitor
        ActiveXResourceMonitor.TrackComObjectCreated();
        Assert.True(ActiveXResourceMonitor.ActiveComObjects >= 1);

        var intermediate = ActiveXResourceMonitor.TakeSnapshot();
        var deltaIntermediate = ActiveXResourceMonitor.Compare(baseline, intermediate);
        Assert.True(deltaIntermediate.ComObjectsDelta >= 1);

        ActiveXResourceMonitor.TrackComObjectReleased();
        var final = ActiveXResourceMonitor.TakeSnapshot();
        var deltaFinal = ActiveXResourceMonitor.Compare(baseline, final);
        Assert.Equal(0, deltaFinal.ComObjectsDelta);
    }

    [Fact]
    public void Acceptance_Section19_ResourceMonitor_100CycleStressTest_StableHandlesAndZeroLeakage()
    {
        RunInStaThread(() =>
        {
            // Ejecutar la prueba de estrés de 100 ciclos según el Hito 8 del Manual Técnico:
            // "Crear/destruir 100 veces no aumenta indefinidamente handles.
            // No quedan ventanas hijas tras cerrar. No quedan eventos conectados."
            var result = ActiveXResourceMonitor.RunStressTest(() =>
            {
                // En cada ciclo instanciamos un control ActiveX dinámico, su extensor VB y ejecutamos transformaciones métricas
                var dummyAx = new DynamicActiveXControl("{6262D3A0-5316-11CE-BFC1-08002BE10318}", 100, 100)
                {
                    Name = "MSFlexGrid1"
                };
                var extender = dummyAx.Extender;
                extender.SetBoundsDirect(150, 250, 3200, 2100);

                // Validar conversiones métricas
                int himetricX = VbControlExtender.TwipsToHimetric(extender.Width);
                float twipsX = VbControlExtender.HimetricToTwips(himetricX);
                Assert.True(Math.Abs(extender.Width - twipsX) <= 15);

                // Limpieza explícita
                extender.Dispose();
                dummyAx.Dispose();
            }, iterations: 100, warmupIterations: 5, maxHandleVariance: 5);

            Assert.Equal(100, result.Iterations);
            Assert.True(result.Success, $"Prueba de estrés fallida: {result.Summary}");
            Assert.True(result.Delta.IsStable(maxHandleVariance: 5));
            Assert.Contains("PASSED", result.Summary);
        });
    }

    // =========================================================================
    // 3. Catálogo y Descubrimiento de Controles ActiveX de Referencia (Sección 19)
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_ReferenceControls_CatalogAndDiscoveryCoverage()
    {
        // Matriz de controles de referencia exigida por Sección 19 y Hito 0 del Manual
        var referenceControls = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["MSFlexGridLib.MSFlexGrid"] = "{6262D3A0-5316-11CE-BFC1-08002BE10318}",
            ["MSComctlLib.TreeCtrl"] = "{C74190B6-8589-11D1-B16A-00C0F0283628}",
            ["MSComctlLib.ListViewCtrl"] = "{BDD1F04B-858B-11D1-B16A-00C0F0283628}",
            ["MSComctlLib.TabStrip"] = "{1EFB6596-857C-11D1-B16A-00C0F0283628}",
            ["MSComctlLib.StatusBar"] = "{8E3867A3-8586-11D1-B16A-00C0F0283628}",
            ["MSComctlLib.ProgBar"] = "{35053A22-8589-11D1-B16A-00C0F0283628}",
            ["MSComctlLib.Slider"] = "{F08DF954-8592-11D1-B16A-00C0F0283628}",
            ["RichTextLib.RichTextBox"] = "{3B7C8860-D78F-101B-B9B5-04021C009402}",
            ["TabDlg.SSTab"] = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}",
            ["MSWinsockLib.Winsock"] = "{248DD896-BB45-11CF-9ABC-0080C7E7B78D}",
            ["MSCommLib.MSComm"] = "{648A5600-2C6E-101B-82B6-000000000014}",
            ["SHDocVw.WebBrowser"] = "{8856F961-340A-11D0-A96B-00C04FD705A2}"
        };

        foreach (var kvp in referenceControls)
        {
            Assert.True(ActiveXHostFactory.TryResolveClsid(kvp.Key, out string resolvedClsid),
                $"No se pudo resolver el CLSID para el control de referencia: {kvp.Key}");
            Assert.Equal(kvp.Value, resolvedClsid);
        }
    }

    // =========================================================================
    // 4. Resiliencia ante Controles Ausentes y Preservación de Datos Desconocidos
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_UnresolvedMissingControl_PreservesUnknownDataWithoutCorruption()
    {
        // Sección 19: "Si el control falta, se preserva la información desconocida"
        string frmOriginal = @"VERSION 5.00
Object = ""{A1B2C3D4-E5F6-7890-ABCD-EF0123456789}#1.0#0""; ""MissingVendorOcx.ocx""
Begin VB.Form Form1
   Caption         =   ""Formulario con Control Ausente""
   ClientHeight    =   4800
   ClientWidth     =   6400
   Begin MissingVendor.SpecialWidget Special1
      Height          =   1500
      Left            =   300
      Top             =   450
      Width           =   2500
      _ExtentX        =   4410
      _ExtentY        =   2646
      CustomPropAlpha =   ""ValorAlpha""
      CustomPropBeta  =   12345
      BeginProperty NestedFont
         Name            =   ""Arial""
         Size            =   9.75
      EndProperty
   End
End
";

        var form = FormDesignParser.Parse(frmOriginal);
        Assert.NotNull(form);
        Assert.Single(form.Children);

        var ctrl = form.Children[0];
        Assert.Equal("Special1", ctrl.Name);
        Assert.Equal("MissingVendor.SpecialWidget", ctrl.TypeName);
        Assert.Equal("ValorAlpha", ctrl.GetCustomPropertyString("CustomPropAlpha"));
        Assert.Equal("12345", ctrl.GetCustomProperty("CustomPropBeta"));

        // Serializar de vuelta y verificar que se conserva intacta la cabecera Object = y los datos desconocidos
        string frmSerialized = FormDesignSerializer.Serialize(form);
        Assert.Contains("MissingVendorOcx.ocx", frmSerialized);
        Assert.Contains("Begin MissingVendor.SpecialWidget Special1", frmSerialized);
        Assert.Contains("CustomPropAlpha", frmSerialized);
        Assert.Contains("ValorAlpha", frmSerialized);
        Assert.Contains("CustomPropBeta", frmSerialized);
        Assert.Contains("12345", frmSerialized);
        Assert.Contains("BeginProperty NestedFont", frmSerialized);
    }

    // =========================================================================
    // 5. Eventos COM: Despacho, Parámetros ByRef y Matrices de Controles
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_EventSink_MultiParamByRef_And_ControlArrayIndex_Fidelity()
    {
        // Guid ficticia de interfaz de eventos Source
        var eventIid = Guid.NewGuid();
        var knownEvents = new List<OcxEventInfo>
        {
            new OcxEventInfo
            {
                Name = "Validate",
                DispId = 101,
                Parameters = new List<OcxParameterInfo>
                {
                    new OcxParameterInfo { Name = "Cancel", Type = "Boolean", IsByRef = true }
                }
            }
        };

        // Instanciar un sink para un control con índice de matriz (Index = 3)
        using var sink = new ActiveXEventSink(eventIid, "TextArray", "CustomOcx.TextInput", controlIndex: 3, knownEvents: knownEvents);

        ActiveXEventInvocation? receivedInvocation = null;
        sink.EventTriggered += inv =>
        {
            receivedInvocation = inv;
            // Modificar el parámetro ByRef Cancel (índice 1 en la llamada expuesta debido a que Index As Integer está en 0)
            inv.UpdateArgumentValue(1, VbValue.FromBoolean(true));
        };

        // Simular llamada DISPPARAMS desde COM: 1 parámetro ByRef (Cancel = False)
        IntPtr pCancel = Marshal.AllocHGlobal(2);
        Marshal.WriteInt16(pCancel, 0); // VARIANT_FALSE

        IntPtr pVariantArg = Marshal.AllocHGlobal(16);
        Marshal.WriteInt16(pVariantArg, 0, (short)(VarEnum.VT_BOOL | VarEnum.VT_BYREF)); // vt
        Marshal.WriteInt16(pVariantArg, 2, 0);
        Marshal.WriteInt16(pVariantArg, 4, 0);
        Marshal.WriteInt16(pVariantArg, 6, 0);
        Marshal.WriteIntPtr(pVariantArg, 8, pCancel);

        var dispParams = new DISPPARAMS
        {
            rgvarg = pVariantArg,
            rgdispidNamedArgs = IntPtr.Zero,
            cArgs = 1,
            cNamedArgs = 0
        };

        Guid dummyIid = Guid.Empty;
        int hr = sink.Invoke(101, ref dummyIid, 0, 1, ref dispParams, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

        Assert.Equal(0, hr);
        Assert.NotNull(receivedInvocation);
        Assert.Equal("Validate", receivedInvocation.EventName);
        Assert.Equal(2, receivedInvocation.Arguments.Length); // [Index, Cancel]
        Assert.Equal(3, receivedInvocation.Arguments[0].AsLong()); // Index inyectado
        Assert.False(receivedInvocation.Arguments[1].AsBoolean()); // Cancel original

        // Verificar que la mutación ByRef se reflejó en la memoria COM nativa
        short updatedCancel = Marshal.ReadInt16(pCancel);
        Assert.Equal(-1, updatedCancel); // VARIANT_TRUE = -1

        Marshal.FreeHGlobal(pVariantArg);
        Marshal.FreeHGlobal(pCancel);
    }

    // =========================================================================
    // 6. Fidelidad Bidireccional HIMETRIC <-> Twips
    // =========================================================================

    [Fact]
    public void Acceptance_Section19_CoordinateTransformations_HimetricTwipsFidelity()
    {
        // 1440 twips/in y 2540 HIMETRIC/in
        int[] testTwips = { 15, 60, 240, 1440, 2880, 5000, 10000, 20000 };

        foreach (int twips in testTwips)
        {
            int himetric = VbControlExtender.TwipsToHimetric(twips);
            float backTwips = VbControlExtender.HimetricToTwips(himetric);

            // La tolerancia máxima admisible por redondeo en enteros debe ser de 1 twip (0.015 mm)
            Assert.True(Math.Abs(twips - backTwips) <= 1.0f,
                $"Discrepancia en conversión Twips <-> HIMETRIC: Entrada={twips}, HIMETRIC={himetric}, Salida={backTwips}");
        }
    }
}
