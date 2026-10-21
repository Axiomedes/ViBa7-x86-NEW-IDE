using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using ViBa7.ActiveXHost.Native;

namespace ViBa7.ActiveXHost.Diagnostics;

/// <summary>
/// Registro inmutable de un evento o diagnóstico del subsistema OLE/ActiveX x86.
/// Cumple con las especificaciones de la Sección 17.1 y Hito 8 del Manual Técnico.
/// </summary>
public sealed record ActiveXDiagnosticRecord(
    DateTime TimestampUtc,
    int ThreadId,
    bool IsStaThread,
    string Stage,
    string Clsid,
    string? ProgId,
    string? ControlName,
    string? Iid,
    int HResult,
    string? ErrorInfoDescription,
    string? Detail)
{
    public string HResultHex => $"0x{HResult:X8}";
    public bool IsSuccess => HResult >= 0;

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append($"[{TimestampUtc:yyyy-MM-dd HH:mm:ss.fff}] [TID:{ThreadId}{(IsStaThread ? ":STA" : ":MTA/OTHER")}] ");
        sb.Append($"Stage: {Stage} | CLSID: {Clsid} | Name: {ControlName ?? "(anon)"} | HRESULT: {HResultHex}");
        if (!string.IsNullOrWhiteSpace(ProgId)) sb.Append($" | ProgID: {ProgId}");
        if (!string.IsNullOrWhiteSpace(Iid)) sb.Append($" | IID: {Iid}");
        if (!string.IsNullOrWhiteSpace(ErrorInfoDescription)) sb.Append($" | ErrorInfo: \"{ErrorInfoDescription}\"");
        if (!string.IsNullOrWhiteSpace(Detail)) sb.Append($" | Detail: {Detail}");
        return sb.ToString();
    }
}

/// <summary>
/// Registrador y agregador central de telemetría y diagnósticos de ciclo de vida ActiveX / COM.
/// Proporciona trazabilidad completa en tiempo real de etapas de instanciación, activación OLE,
/// persistencia, despacho de eventos, y captura automática de IErrorInfo.
/// </summary>
public static class ActiveXDiagnosticsLogger
{
    private static readonly ConcurrentQueue<ActiveXDiagnosticRecord> _buffer = new();
    private static readonly object _syncRoot = new();

    /// <summary>
    /// Capacidad máxima de retención del búfer de diagnósticos en memoria.
    /// </summary>
    public static int MaxHistoryCapacity { get; set; } = 2000;

    /// <summary>
    /// Evento disparado cada vez que se registra una nueva entrada de diagnóstico.
    /// </summary>
    public static event Action<ActiveXDiagnosticRecord>? DiagnosticLogged;

    /// <summary>
    /// Registra un evento de diagnóstico con inspección de apartamento STA y captura automática de IErrorInfo.
    /// </summary>
    public static ActiveXDiagnosticRecord Log(
        string stage,
        string clsid,
        string? controlName = null,
        int hResult = 0,
        string? iid = null,
        string? progId = null,
        string? detail = null,
        string? errorInfo = null)
    {
        int threadId = Environment.CurrentManagedThreadId;
        bool isSta = Thread.CurrentThread.GetApartmentState() == ApartmentState.STA;

        string? resolvedErrorInfo = errorInfo;
        if (hResult < 0 && string.IsNullOrWhiteSpace(resolvedErrorInfo))
        {
            if (Ole32Interop.TryGetComErrorDescription(out string? desc, out string? src))
            {
                resolvedErrorInfo = !string.IsNullOrWhiteSpace(src) ? $"[{src}] {desc}" : desc;
            }
        }

        var record = new ActiveXDiagnosticRecord(
            DateTime.UtcNow,
            threadId,
            isSta,
            stage ?? "General",
            clsid ?? "{00000000-0000-0000-0000-000000000000}",
            progId,
            controlName,
            iid,
            hResult,
            resolvedErrorInfo,
            detail);

        _buffer.Enqueue(record);

        // Controlar el tamaño del búfer para evitar acumulación indefinida
        while (_buffer.Count > MaxHistoryCapacity && _buffer.TryDequeue(out _))
        {
        }

        try
        {
            DiagnosticLogged?.Invoke(record);
        }
        catch
        {
            // Proteger la ejecución frente a excepciones en los suscriptores
        }

        return record;
    }

    /// <summary>
    /// Registra una operación exitosa (HRESULT = 0).
    /// </summary>
    public static ActiveXDiagnosticRecord LogSuccess(
        string stage,
        string clsid,
        string? controlName = null,
        string? detail = null,
        string? progId = null)
    {
        return Log(stage, clsid, controlName, hResult: 0, detail: detail, progId: progId);
    }

    /// <summary>
    /// Registra una operación fallida con HRESULT negativo.
    /// </summary>
    public static ActiveXDiagnosticRecord LogError(
        string stage,
        string clsid,
        int hResult,
        string? controlName = null,
        string? detail = null,
        string? iid = null,
        string? progId = null)
    {
        return Log(stage, clsid, controlName, hResult, iid, progId, detail);
    }

    /// <summary>
    /// Obtiene una instantánea de las entradas más recientes registradas.
    /// </summary>
    public static IReadOnlyList<ActiveXDiagnosticRecord> GetRecentEntries(int count = 100)
    {
        var items = _buffer.ToArray();
        if (items.Length <= count) return items;
        return items.Skip(items.Length - count).ToArray();
    }

    /// <summary>
    /// Obtiene todas las entradas registradas que representen fallos (HRESULT &lt; 0).
    /// </summary>
    public static IReadOnlyList<ActiveXDiagnosticRecord> GetErrors()
    {
        return _buffer.Where(e => !e.IsSuccess).ToArray();
    }

    /// <summary>
    /// Obtiene las entradas correspondientes a un CLSID o nombre de control específico.
    /// </summary>
    public static IReadOnlyList<ActiveXDiagnosticRecord> GetEntriesForControl(string clsidOrName)
    {
        if (string.IsNullOrWhiteSpace(clsidOrName)) return Array.Empty<ActiveXDiagnosticRecord>();

        return _buffer.Where(e =>
            string.Equals(e.Clsid, clsidOrName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(e.ControlName, clsidOrName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(e.ProgId, clsidOrName, StringComparison.OrdinalIgnoreCase)).ToArray();
    }

    /// <summary>
    /// Limpia el búfer en memoria de registros de diagnóstico.
    /// </summary>
    public static void Clear()
    {
        while (_buffer.TryDequeue(out _)) { }
    }

    /// <summary>
    /// Exporta el contenido actual del búfer como texto legible para informes o depuración.
    /// </summary>
    public static string ExportToText()
    {
        var records = _buffer.ToArray();
        var sb = new StringBuilder();
        sb.AppendLine($"=== ACTIVEX / OCX DIAGNOSTICS LOG ({records.Length} entries) ===");
        foreach (var r in records)
        {
            sb.AppendLine(r.ToString());
        }
        return sb.ToString();
    }

    /// <summary>
    /// Exporta el contenido actual como un documento JSON estructurado.
    /// </summary>
    public static string ExportToJson(bool indented = true)
    {
        var records = _buffer.ToArray();
        return JsonSerializer.Serialize(records, new JsonSerializerOptions
        {
            WriteIndented = indented
        });
    }
}
