using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ViBa7.ProjectModel.Navigation;

/// <summary>
/// Gestor centralizado de puntos de interrupción (Breakpoints) del proyecto.
/// Mantiene los puntos de interrupción activos asociados a rutas de archivo y números de línea,
/// sincronizándose en tiempo real con el margen del editor y el motor de ejecución.
/// </summary>
public class BreakpointManager
{
    private static readonly Lazy<BreakpointManager> _instance = new(() => new BreakpointManager());
    public static BreakpointManager Instance => _instance.Value;

    private readonly object _lock = new();
    private readonly Dictionary<string, SortedSet<int>> _breakpoints = new(StringComparer.OrdinalIgnoreCase);

    public event Action? BreakpointsChanged;

    private string Normalize(string path)
    {
        try { return Path.GetFullPath(path); }
        catch { return path; }
    }

    /// <summary>
    /// Alterna el estado del punto de interrupción en una línea de un archivo.
    /// Retorna true si se agregó el breakpoint, false si se eliminó.
    /// </summary>
    public bool ToggleBreakpoint(string filePath, int line)
    {
        if (string.IsNullOrWhiteSpace(filePath) || line <= 0) return false;
        string key = Normalize(filePath);
        bool added;

        lock (_lock)
        {
            if (!_breakpoints.TryGetValue(key, out var set))
            {
                set = [];
                _breakpoints[key] = set;
            }

            if (set.Contains(line))
            {
                set.Remove(line);
                if (set.Count == 0) _breakpoints.Remove(key);
                added = false;
            }
            else
            {
                set.Add(line);
                added = true;
            }
        }

        BreakpointsChanged?.Invoke();
        return added;
    }

    /// <summary>
    /// Comprueba si existe un punto de interrupción en un archivo y línea especificados.
    /// </summary>
    public bool HasBreakpoint(string filePath, int line)
    {
        if (string.IsNullOrWhiteSpace(filePath) || line <= 0) return false;
        string key = Normalize(filePath);

        lock (_lock)
        {
            return _breakpoints.TryGetValue(key, out var set) && set.Contains(line);
        }
    }

    /// <summary>
    /// Obtiene todos los números de línea con punto de interrupción para un archivo.
    /// </summary>
    public IReadOnlyCollection<int> GetBreakpoints(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return Array.Empty<int>();
        string key = Normalize(filePath);

        lock (_lock)
        {
            if (_breakpoints.TryGetValue(key, out var set))
            {
                return set.ToArray();
            }
        }
        return Array.Empty<int>();
    }

    /// <summary>
    /// Obtiene todos los puntos de interrupción registrados en el proyecto como tuplas (FilePath, Line).
    /// </summary>
    public IReadOnlyCollection<(string FilePath, int Line)> GetAllBreakpoints()
    {
        lock (_lock)
        {
            var list = new List<(string FilePath, int Line)>();
            foreach (var kvp in _breakpoints)
            {
                foreach (int line in kvp.Value)
                {
                    list.Add((kvp.Key, line));
                }
            }
            return list;
        }
    }

    /// <summary>
    /// Total de puntos de interrupción activos en el proyecto.
    /// </summary>
    public int TotalBreakpointCount
    {
        get
        {
            lock (_lock)
            {
                return _breakpoints.Values.Sum(s => s.Count);
            }
        }
    }

    /// <summary>
    /// Elimina todos los puntos de interrupción del proyecto.
    /// </summary>
    public void ClearAll()
    {
        lock (_lock)
        {
            if (_breakpoints.Count == 0) return;
            _breakpoints.Clear();
        }

        BreakpointsChanged?.Invoke();
    }
}
