using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using ViBa7.ActiveXHost.Native;

namespace ViBa7.ActiveXHost.Diagnostics;

/// <summary>
/// Instantánea de recursos GDI, USER, memoria y referencias COM en un instante determinado.
/// </summary>
public sealed record ActiveXResourceSnapshot(
    DateTime TimestampUtc,
    uint GdiObjects,
    uint UserObjects,
    uint GdiPeak,
    uint UserPeak,
    long PrivateMemoryBytes,
    long WorkingSetBytes,
    long ManagedMemoryBytes,
    int ActiveComObjects)
{
    public override string ToString()
    {
        return $"[GDI: {GdiObjects} (Peak: {GdiPeak}) | USER: {UserObjects} (Peak: {UserPeak}) | " +
               $"MemPriv: {PrivateMemoryBytes / 1024 / 1024} MB | Managed: {ManagedMemoryBytes / 1024 / 1024} MB | " +
               $"COM Active: {ActiveComObjects}]";
    }
}

/// <summary>
/// Diferencia neta de consumo de recursos entre dos instantáneas.
/// </summary>
public sealed record ActiveXResourceDelta(
    int GdiDelta,
    int UserDelta,
    long PrivateMemoryDeltaBytes,
    long ManagedMemoryDeltaBytes,
    int ComObjectsDelta)
{
    /// <summary>
    /// Determina si el consumo de handles es estable dentro de una tolerancia aceptable
    /// (por defecto 5 handles de margen para cachés de sistema de Windows/WPF).
    /// </summary>
    public bool IsStable(int maxHandleVariance = 5)
    {
        return Math.Abs(GdiDelta) <= maxHandleVariance &&
               Math.Abs(UserDelta) <= maxHandleVariance &&
               ComObjectsDelta <= 0;
    }

    public override string ToString()
    {
        return $"[Delta GDI: {GdiDelta:+0;-0;0} | Delta USER: {UserDelta:+0;-0;0} | " +
               $"Delta MemPriv: {PrivateMemoryDeltaBytes / 1024} KB | Delta COM: {ComObjectsDelta:+0;-0;0} | " +
               $"Stable: {IsStable()}]";
    }
}

/// <summary>
/// Resultado consolidado de una prueba de estrés de apertura y cierre repetido.
/// </summary>
public sealed record ActiveXStressTestResult(
    int Iterations,
    ActiveXResourceSnapshot Baseline,
    ActiveXResourceSnapshot Final,
    ActiveXResourceDelta Delta,
    bool Success)
{
    public string Summary
    {
        get
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== ACTIVEX STRESS TEST REPORT ({Iterations} cycles) ===");
            sb.AppendLine($"Baseline: {Baseline}");
            sb.AppendLine($"Final:    {Final}");
            sb.AppendLine($"Delta:    {Delta}");
            sb.AppendLine($"Status:   {(Success ? "PASSED (Zero / Bounded Leakage)" : "FAILED (Potential Leak Detected)")}");
            return sb.ToString();
        }
    }
}

/// <summary>
/// Monitor de recursos de bajo nivel para telemetría de estabilidad y detección de fugas (Leak Detection).
/// Supervisa handles GDI y USER mediante GetGuiResources, memoria del proceso y balance de objetos COM.
/// </summary>
public static class ActiveXResourceMonitor
{
    private static int _activeComObjectCount;

    /// <summary>
    /// Número de instancias COM administradas activas en el host.
    /// </summary>
    public static int ActiveComObjects => Volatile.Read(ref _activeComObjectCount);

    /// <summary>
    /// Incrementa el contador de objetos COM activos.
    /// </summary>
    public static void TrackComObjectCreated()
    {
        Interlocked.Increment(ref _activeComObjectCount);
    }

    /// <summary>
    /// Decrementa el contador de objetos COM activos al liberarse una instancia.
    /// </summary>
    public static void TrackComObjectReleased()
    {
        Interlocked.Decrement(ref _activeComObjectCount);
    }

    /// <summary>
    /// Reinicia el contador de objetos COM activos (utilizado principalmente para inicializar pruebas unitarias).
    /// </summary>
    public static void ResetComObjectCount()
    {
        Interlocked.Exchange(ref _activeComObjectCount, 0);
    }

    /// <summary>
    /// Captura una instantánea del estado actual de recursos del proceso.
    /// </summary>
    public static ActiveXResourceSnapshot TakeSnapshot()
    {
        uint gdi = 0;
        uint user = 0;
        uint gdiPeak = 0;
        uint userPeak = 0;

        try
        {
            IntPtr hProcess = Process.GetCurrentProcess().Handle;
            gdi = Ole32Interop.GetGuiResources(hProcess, Ole32Interop.GR_GDIOBJECTS);
            user = Ole32Interop.GetGuiResources(hProcess, Ole32Interop.GR_USEROBJECTS);
            gdiPeak = Ole32Interop.GetGuiResources(hProcess, Ole32Interop.GR_GDIOBJECTS_PEAK);
            userPeak = Ole32Interop.GetGuiResources(hProcess, Ole32Interop.GR_USEROBJECTS_PEAK);
        }
        catch
        {
            // Entornos de prueba donde GetGuiResources no esté disponible
        }

        long privMem = 0;
        long workingSet = 0;
        try
        {
            var proc = Process.GetCurrentProcess();
            proc.Refresh();
            privMem = proc.PrivateMemorySize64;
            workingSet = proc.WorkingSet64;
        }
        catch { }

        long managedMem = GC.GetTotalMemory(false);
        int comActive = ActiveComObjects;

        return new ActiveXResourceSnapshot(
            DateTime.UtcNow,
            gdi,
            user,
            gdiPeak,
            userPeak,
            privMem,
            workingSet,
            managedMem,
            comActive);
    }

    /// <summary>
    /// Compara dos instantáneas y calcula la diferencia neta de recursos.
    /// </summary>
    public static ActiveXResourceDelta Compare(ActiveXResourceSnapshot baseline, ActiveXResourceSnapshot current)
    {
        return new ActiveXResourceDelta(
            (int)current.GdiObjects - (int)baseline.GdiObjects,
            (int)current.UserObjects - (int)baseline.UserObjects,
            current.PrivateMemoryBytes - baseline.PrivateMemoryBytes,
            current.ManagedMemoryBytes - baseline.ManagedMemoryBytes,
            current.ActiveComObjects - baseline.ActiveComObjects);
    }

    /// <summary>
    /// Ejecuta una prueba de estrés repetitiva (por defecto 100 ciclos según el Hito 8 del Manual Técnico)
    /// ejecutando la acción de creación/interacción/cierre y evaluando la ausencia de fugas de handles.
    /// </summary>
    public static ActiveXStressTestResult RunStressTest(
        Action iterationAction,
        int iterations = 100,
        int warmupIterations = 5,
        int maxHandleVariance = 5)
    {
        // 1. Fase de calentamiento para estabilizar JIT, singletons estáticos y cachés
        for (int i = 0; i < warmupIterations; i++)
        {
            iterationAction();
        }

        // Forzar limpieza completa para establecer línea base
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var baseline = TakeSnapshot();

        // 2. Ejecución de la batería de ciclos repetidos
        for (int i = 0; i < iterations; i++)
        {
            iterationAction();
        }

        // Forzar limpieza posterior a los ciclos
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        var final = TakeSnapshot();
        var delta = Compare(baseline, final);
        bool success = delta.IsStable(maxHandleVariance);

        return new ActiveXStressTestResult(iterations, baseline, final, delta, success);
    }
}
