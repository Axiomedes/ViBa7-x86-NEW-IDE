using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.ProjectModel.IntelliSense;

/// <summary>
/// Catálogo exhaustivo de miembros (propiedades, métodos y eventos) para los controles
/// intrínsecos de Visual Basic 6.0, el Formulario (Form) y controles ActiveX / OCX cargados dinámicamente.
/// </summary>
public static class VbControlMemberCatalog
{
    private static readonly object _syncLock = new();
    private static readonly Dictionary<string, List<VbSymbolItem>> _controlMembers = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, List<VbSymbolItem>> _dynamicControlMembers = new(StringComparer.OrdinalIgnoreCase);
    private static readonly List<OcxLibraryInfo> _registeredOcxLibraries = [];
    private static readonly HashSet<string> _attemptedRegistryLookups = new(StringComparer.OrdinalIgnoreCase);
    private static readonly List<VbSymbolItem> _formMembers = [];

    static VbControlMemberCatalog()
    {
        InitializeCommonPropertiesAndMethods();
        InitializeFormMembers();
        InitializeCommandButton();
        InitializeTextBox();
        InitializeLabel();
        InitializeCheckBox();
        InitializeOptionButton();
        InitializeFrame();
        InitializeListBox();
        InitializeComboBox();
        InitializePictureBox();
        InitializeImage();
        InitializeTimer();
        InitializeHScrollBar();
        InitializeVScrollBar();
        InitializeShape();
        InitializeLine();
    }

    /// <summary>
    /// Retorna la lista de miembros de un control según su tipo (e.g. "CommandButton", "VB.CommandButton", "TextBox", "Form", "axJColCombo").
    /// </summary>
    public static IReadOnlyList<VbSymbolItem> GetMembersForControl(string typeName)
    {
        string normalized = NormalizeTypeName(typeName);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return Array.Empty<VbSymbolItem>();
        }

        if (normalized.Equals("Form", StringComparison.OrdinalIgnoreCase))
        {
            return _formMembers;
        }

        if (_controlMembers.TryGetValue(normalized, out var list))
        {
            return list;
        }

        lock (_syncLock)
        {
            // Buscar en controles dinámicos registrados (OCX)
            if (_dynamicControlMembers.TryGetValue(normalized, out var dynList))
            {
                return dynList;
            }

            // Si viene calificado ej. "AxJControls.axJColCombo", buscar por la parte no calificada
            int dot = normalized.LastIndexOf('.');
            if (dot >= 0 && dot < normalized.Length - 1)
            {
                string shortName = normalized[(dot + 1)..];
                if (_controlMembers.TryGetValue(shortName, out var shortList))
                {
                    return shortList;
                }
                if (_dynamicControlMembers.TryGetValue(shortName, out var shortDynList))
                {
                    return shortDynList;
                }
            }

            // Intentar auto-descubrir y registrar la TypeLib desde el Registro de Windows si aún no estaba cargada
            if (TryEnsureControlRegisteredFromRegistry(normalized))
            {
                if (_dynamicControlMembers.TryGetValue(normalized, out var resolvedDynList))
                {
                    return resolvedDynList;
                }
                if (dot >= 0 && dot < normalized.Length - 1)
                {
                    string shortName = normalized[(dot + 1)..];
                    if (_dynamicControlMembers.TryGetValue(shortName, out var resolvedShortList))
                    {
                        return resolvedShortList;
                    }
                }
            }

            return Array.Empty<VbSymbolItem>();
        }
    }

    /// <summary>
    /// Retorna los miembros del Formulario (Form / Me).
    /// </summary>
    public static IReadOnlyList<VbSymbolItem> GetFormMembers() => _formMembers;

    /// <summary>
    /// Retorna todos los nombres de tipos de controles intrínsecos de VB6.
    /// </summary>
    public static IEnumerable<string> GetKnownControlTypes()
    {
        yield return "Form";
        foreach (var key in _controlMembers.Keys)
        {
            yield return key;
        }
    }

    /// <summary>
    /// Busca si un tipo corresponde a una enumeración registrada en alguna TypeLib de OCX y retorna sus opciones para el PropertyGrid.
    /// </summary>
    public static List<DesignerPropertyChoice>? GetEnumChoices(string enumTypeName)
    {
        if (string.IsNullOrWhiteSpace(enumTypeName)) return null;
        string clean = enumTypeName.Trim();
        if (clean.StartsWith("OLE_", StringComparison.OrdinalIgnoreCase)) return null;

        lock (_syncLock)
        {
            foreach (var lib in _registeredOcxLibraries)
            {
                var found = lib.Enums.FirstOrDefault(e => e.Name.Equals(clean, StringComparison.OrdinalIgnoreCase));
                if (found != null && found.Items.Count > 0)
                {
                    return found.Items.Select(i => new DesignerPropertyChoice($"{i.Value} - {i.Name}", i.Value.ToString())).ToList();
                }
            }
        }
        return null;
    }

    /// <summary>
    /// Registra en el catálogo de IntelliSense una biblioteca OCX completa con todos sus controles, métodos y propiedades.
    /// </summary>
    public static void RegisterOcxLibrary(OcxLibraryInfo lib)
    {
        ArgumentNullException.ThrowIfNull(lib);

        lock (_syncLock)
        {
            // Si ya estaba registrada, desregistrar primero para refrescar
            UnregisterOcxLibrary(lib.LibraryName);

            _registeredOcxLibraries.Add(lib);
            string? fileStem = !string.IsNullOrWhiteSpace(lib.FilePath)
                ? Path.GetFileNameWithoutExtension(lib.FilePath)
                : null;

        foreach (var ctrl in lib.Controls)
        {
            var symbols = new List<VbSymbolItem>();

            // Propiedades y métodos del control (omitir miembros marcados con FUNCFLAG_FHIDDEN / VARFLAG_FHIDDEN)
            foreach (var m in ctrl.Members)
            {
                if (m.IsHidden) continue;

                symbols.Add(new VbSymbolItem
                {
                    Name = m.Name,
                    Kind = m.Kind,
                    ReturnType = m.ReturnType,
                    DispId = m.DispId,
                    IsHidden = m.IsHidden,
                    IsBrowsable = m.IsBrowsable,
                    IsReadOnly = m.Kind == VbSymbolKind.Property && m.IsPropertyGet && !m.IsPropertyPut,
                    Description = !string.IsNullOrEmpty(m.Description) ? m.Description : $"{m.Kind} {m.Name} de {ctrl.Name}",
                    Parameters = m.Parameters.Select(p => new VbParameterInfo
                    {
                        Name = p.Name,
                        Type = p.Type,
                        IsOptional = p.IsOptional,
                        IsByRef = p.IsByRef,
                        DefaultValue = p.DefaultValue,
                        Description = p.Description ?? string.Empty
                    }).ToList()
                });
            }

            // Eventos del control
            foreach (var ev in ctrl.Events)
            {
                symbols.Add(new VbSymbolItem
                {
                    Name = ev.Name,
                    Kind = VbSymbolKind.Event,
                    ReturnType = "Void",
                    DispId = ev.DispId,
                    Description = !string.IsNullOrEmpty(ev.Description) ? ev.Description : $"Evento {ev.Name} de {ctrl.Name}",
                    Parameters = ev.Parameters.Select(p => new VbParameterInfo
                    {
                        Name = p.Name,
                        Type = p.Type,
                        IsOptional = p.IsOptional,
                        IsByRef = p.IsByRef,
                        DefaultValue = p.DefaultValue,
                        Description = p.Description ?? string.Empty
                    }).ToList()
                });
            }

            // Registrar por nombre simple (ej. "axJColCombo")
            _dynamicControlMembers[ctrl.Name] = symbols;

            // Registrar por nombre calificado si tiene biblioteca (ej. "AxJControls.axJColCombo")
            if (!string.IsNullOrEmpty(lib.LibraryName))
            {
                _dynamicControlMembers[$"{lib.LibraryName}.{ctrl.Name}"] = symbols;
            }

            // Registrar también por nombre de archivo OCX (ej. "MSCOMCTL.ListViewCtrl")
            if (!string.IsNullOrEmpty(fileStem))
            {
                _dynamicControlMembers[$"{fileStem}.{ctrl.Name}"] = symbols;
            }
        }
    }
}

    /// <summary>
    /// Resuelve el nombre de una propiedad de un control ActiveX a partir de su DISPID.
    /// </summary>
    public static bool TryResolvePropertyNameByDispId(string controlTypeName, int dispId, out string propertyName)
    {
        propertyName = string.Empty;
        var members = GetMembersForControl(controlTypeName);
        var match = members.FirstOrDefault(m => m.Kind == VbSymbolKind.Property && m.DispId == dispId);
        if (match != null)
        {
            propertyName = match.Name;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Resuelve el símbolo de evento de un control ActiveX a partir de su DISPID.
    /// </summary>
    public static bool TryResolveEventByDispId(string controlTypeName, int dispId, out VbSymbolItem? eventSymbol)
    {
        eventSymbol = null;
        var members = GetMembersForControl(controlTypeName);
        var match = members.FirstOrDefault(m => m.Kind == VbSymbolKind.Event && m.DispId == dispId);
        if (match != null)
        {
            eventSymbol = match;
            return true;
        }
        return false;
    }

    /// <summary>
    /// Retorna todos los eventos expuestos por la interfaz Source de un control (intrínseco o ActiveX).
    /// </summary>
    public static IReadOnlyList<VbSymbolItem> GetEventsForControl(string controlTypeName)
    {
        var members = GetMembersForControl(controlTypeName);
        return members.Where(m => m.Kind == VbSymbolKind.Event).ToList();
    }

    /// <summary>
    /// Obtiene el evento predeterminado de un control (tanto intrínseco como ActiveX) y su lista de parámetros para generación de código.
    /// </summary>
    public static bool TryGetDefaultEvent(string controlTypeName, out string eventName, out IReadOnlyList<VbParameterInfo> parameters)
    {
        eventName = "Click";
        parameters = Array.Empty<VbParameterInfo>();

        if (string.IsNullOrWhiteSpace(controlTypeName)) return false;
        string norm = NormalizeTypeName(controlTypeName);

        // Mapeo curado para controles estándar y OCX populares de VB6
        if (norm.EndsWith("CommandButton", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("OptionButton", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("CheckBox", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("Menu", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("PictureBox", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("Image", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Click";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("TextBox", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Change";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("ListBox", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("ComboBox", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Click";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("Timer", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Timer";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("Form", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Load";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("MSFlexGrid", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("MSHFlexGrid", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "RowColChange";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("TreeCtrl", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("TreeView", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "NodeClick";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "Node", Type = "Node", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("ListViewCtrl", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("ListView", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "ItemClick";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "Item", Type = "ListItem", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("SSTab", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Click";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "PreviousTab", Type = "Integer", IsByRef = true }
            };
            return true;
        }

        if (norm.EndsWith("Toolbar", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "ButtonClick";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "Button", Type = "Button", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("StatusBar", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "PanelClick";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "Panel", Type = "Panel", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("TabStrip", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "TabClick";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "Tab", Type = "Tab", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("Slider", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("UpDown", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("DTPicker", StringComparison.OrdinalIgnoreCase) ||
            norm.EndsWith("FlatScrollBar", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Change";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("RichTextBox", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "Change";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        if (norm.EndsWith("Winsock", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "DataArrival";
            parameters = new List<VbParameterInfo>
            {
                new() { Name = "bytesTotal", Type = "Long", IsByRef = false }
            };
            return true;
        }

        if (norm.EndsWith("MSComm", StringComparison.OrdinalIgnoreCase))
        {
            eventName = "OnComm";
            parameters = Array.Empty<VbParameterInfo>();
            return true;
        }

        // Búsqueda dinámica en TypeLib
        var events = GetEventsForControl(controlTypeName);
        if (events.Count > 0)
        {
            // 1. Evento con DISPID 0 (DISPID_VALUE)
            var defDispId = events.FirstOrDefault(e => e.DispId == 0);
            if (defDispId != null)
            {
                eventName = defDispId.Name;
                parameters = defDispId.Parameters;
                return true;
            }

            // 2. Evento Click o Change
            var clickOrChange = events.FirstOrDefault(e => e.Name.Equals("Click", StringComparison.OrdinalIgnoreCase) ||
                                                           e.Name.Equals("Change", StringComparison.OrdinalIgnoreCase));
            if (clickOrChange != null)
            {
                eventName = clickOrChange.Name;
                parameters = clickOrChange.Parameters;
                return true;
            }

            // 3. Primer evento de la interfaz Source
            var firstEvent = events[0];
            eventName = firstEvent.Name;
            parameters = firstEvent.Parameters;
            return true;
        }

        return false;
    }

    /// <summary>
    /// Elimina el registro de una biblioteca OCX previamente cargada.
    /// </summary>
    public static void UnregisterOcxLibrary(string libraryNameOrPath)
    {
        if (string.IsNullOrWhiteSpace(libraryNameOrPath)) return;

        lock (_syncLock)
        {
            var matching = _registeredOcxLibraries.Where(l =>
                string.Equals(l.LibraryName, libraryNameOrPath, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(l.FilePath, libraryNameOrPath, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(l.Guid, libraryNameOrPath, StringComparison.OrdinalIgnoreCase)
            ).ToList();

            foreach (var lib in matching)
            {
                _registeredOcxLibraries.Remove(lib);
                string? fileStem = !string.IsNullOrWhiteSpace(lib.FilePath)
                    ? Path.GetFileNameWithoutExtension(lib.FilePath)
                    : null;

                foreach (var ctrl in lib.Controls)
                {
                    _dynamicControlMembers.Remove(ctrl.Name);
                    if (!string.IsNullOrEmpty(lib.LibraryName))
                    {
                        _dynamicControlMembers.Remove($"{lib.LibraryName}.{ctrl.Name}");
                    }
                    if (!string.IsNullOrEmpty(fileStem))
                    {
                        _dynamicControlMembers.Remove($"{fileStem}.{ctrl.Name}");
                    }
                }
            }
        }
    }

    /// <summary>
    /// Obtiene la declaración canónica de cabecera .frm (Object = "{GUID}#Major.Minor#0"; "Archivo.ocx")
    /// correspondiente a un tipo de control ActiveX registrado.
    /// </summary>
    public static bool TryGetOcxObjectHeaderForControl(string controlTypeName, out string objectHeaderLine)
    {
        objectHeaderLine = string.Empty;
        if (string.IsNullOrWhiteSpace(controlTypeName) || controlTypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string normalized = NormalizeTypeName(controlTypeName);
        int dot = normalized.LastIndexOf('.');
        string? prefix = dot > 0 ? normalized[..dot] : null;
        string shortName = dot >= 0 && dot < normalized.Length - 1 ? normalized[(dot + 1)..] : normalized;

        lock (_syncLock)
        {
            OcxLibraryInfo? matchedLib = FindLibraryForControl(prefix, shortName);
            if (matchedLib == null && TryEnsureControlRegisteredFromRegistry(normalized))
            {
                matchedLib = FindLibraryForControl(prefix, shortName);
            }

            if (matchedLib == null || string.IsNullOrWhiteSpace(matchedLib.Guid) || string.IsNullOrWhiteSpace(matchedLib.FilePath))
            {
                return false;
            }

            string guidStr = matchedLib.Guid.Trim();
            if (!guidStr.StartsWith('{')) guidStr = "{" + guidStr;
            if (!guidStr.EndsWith('}')) guidStr += "}";

            int major = matchedLib.MajorVersion <= 0 && matchedLib.MinorVersion <= 0 ? 1 : matchedLib.MajorVersion;
            int minor = Math.Max(0, matchedLib.MinorVersion);
            string fileName = Path.GetFileName(matchedLib.FilePath);

            objectHeaderLine = $"Object = \"{guidStr}#{major}.{minor}#0\"; \"{fileName}\"";
            return true;
        }
    }

    private static OcxLibraryInfo? FindLibraryForControl(string? prefix, string shortName)
    {
        foreach (var lib in _registeredOcxLibraries)
        {
            string? fileStem = !string.IsNullOrWhiteSpace(lib.FilePath)
                ? Path.GetFileNameWithoutExtension(lib.FilePath)
                : null;

            bool prefixMatches = string.IsNullOrWhiteSpace(prefix) ||
                                 string.Equals(lib.LibraryName, prefix, StringComparison.OrdinalIgnoreCase) ||
                                 string.Equals(fileStem, prefix, StringComparison.OrdinalIgnoreCase);

            if (!prefixMatches) continue;

            if (lib.Controls.Any(c => string.Equals(c.Name, shortName, StringComparison.OrdinalIgnoreCase)))
            {
                return lib;
            }
        }
        return null;
    }

    /// <summary>
    /// Retorna la lista de bibliotecas OCX registradas dinámicamente.
    /// </summary>
    public static IReadOnlyList<OcxLibraryInfo> GetRegisteredOcxLibraries()
    {
        lock (_syncLock)
        {
            return _registeredOcxLibraries.ToList();
        }
    }

    /// <summary>
    /// Intenta resolver el CLSID de un control OCX registrado dinámicamente por su nombre simple o calificado.
    /// </summary>
    public static bool TryResolveOcxClsid(string typeName, out string clsid)
    {
        clsid = string.Empty;
        if (string.IsNullOrWhiteSpace(typeName)) return false;

        lock (_syncLock)
        {
            if (TryResolveOcxClsidInternal(typeName, out clsid))
            {
                return true;
            }

            if (TryEnsureControlRegisteredFromRegistry(typeName))
            {
                return TryResolveOcxClsidInternal(typeName, out clsid);
            }

            return false;
        }
    }

    private static bool TryResolveOcxClsidInternal(string typeName, out string clsid)
    {
        clsid = string.Empty;
        string clean = typeName.Trim();
        int dot = clean.LastIndexOf('.');
        string? libPrefix = dot > 0 ? clean[..dot] : null;
        string shortName = dot >= 0 && dot < clean.Length - 1 ? clean[(dot + 1)..] : clean;

        foreach (var lib in _registeredOcxLibraries)
        {
            if (libPrefix != null &&
                !string.IsNullOrEmpty(lib.LibraryName) &&
                !lib.LibraryName.Equals(libPrefix, StringComparison.OrdinalIgnoreCase) &&
                !Path.GetFileNameWithoutExtension(lib.FilePath).Equals(libPrefix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var ctrl = lib.Controls.FirstOrDefault(c =>
                c.Name.Equals(shortName, StringComparison.OrdinalIgnoreCase) ||
                c.Name.Equals(clean, StringComparison.OrdinalIgnoreCase));

            if (ctrl != null && !string.IsNullOrWhiteSpace(ctrl.Clsid))
            {
                clsid = ctrl.Clsid;
                return true;
            }
        }

        // Segunda pasada sin restricción de prefijo de librería (por si el .frm usa el nombre del archivo OCX como prefijo en lugar del LibraryName de la TypeLib)
        if (libPrefix != null)
        {
            foreach (var lib in _registeredOcxLibraries)
            {
                var ctrl = lib.Controls.FirstOrDefault(c =>
                    c.Name.Equals(shortName, StringComparison.OrdinalIgnoreCase));

                if (ctrl != null && !string.IsNullOrWhiteSpace(ctrl.Clsid))
                {
                    clsid = ctrl.Clsid;
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Intenta localizar la TypeLib de un control OCX en el Registro de Windows (HKCR\ProgID\CLSID -> InprocServer32)
    /// y registrarla automáticamente en el catálogo si aún no estaba cargada.
    /// </summary>
    public static bool TryEnsureControlRegisteredFromRegistry(string typeName)
    {
        if (!OperatingSystem.IsWindows() || string.IsNullOrWhiteSpace(typeName))
            return false;

        string clean = NormalizeTypeName(typeName);
        if (string.IsNullOrWhiteSpace(clean) || _controlMembers.ContainsKey(clean))
            return false;

        lock (_syncLock)
        {
            if (!_attemptedRegistryLookups.Add(clean))
                return _dynamicControlMembers.ContainsKey(clean);

            try
            {
                string? resolvedClsid = null;
                string[] candidates =
                [
                    clean,
                    $"{clean}.1",
                    $"{clean}.2",
                    $"{clean}.3",
                    $"{clean}.4"
                ];

                foreach (var progId in candidates)
                {
                    using var clsidKey = Registry.ClassesRoot.OpenSubKey($@"{progId}\CLSID");
                    if (clsidKey?.GetValue(null) is string val && !string.IsNullOrWhiteSpace(val))
                    {
                        resolvedClsid = val.Trim();
                        break;
                    }
                }

                if (string.IsNullOrWhiteSpace(resolvedClsid))
                    return false;

                if (!resolvedClsid.StartsWith('{'))
                    resolvedClsid = "{" + resolvedClsid + "}";

                using var inprocKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{resolvedClsid}\InprocServer32");
                string? serverPath = inprocKey?.GetValue(null) as string;
                if (!string.IsNullOrWhiteSpace(serverPath))
                {
                    serverPath = Environment.ExpandEnvironmentVariables(serverPath.Trim('\"'));
                }

                if (string.IsNullOrWhiteSpace(serverPath) || !File.Exists(serverPath))
                {
                    using var tlibKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{resolvedClsid}\TypeLib");
                    string? tlibGuid = tlibKey?.GetValue(null) as string;
                    if (!string.IsNullOrWhiteSpace(tlibGuid))
                    {
                        serverPath = OcxTypeLibReader.ResolveOcxPath(clean, null, tlibGuid);
                    }
                }

                if (!string.IsNullOrWhiteSpace(serverPath) && File.Exists(serverPath))
                {
                    if (!_registeredOcxLibraries.Any(l => string.Equals(l.FilePath, serverPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        if (OcxTypeLibReader.TryReadTypeLib(serverPath, out var libInfo, out _) && libInfo != null)
                        {
                            RegisterOcxLibrary(libInfo);
                        }
                    }

                    // Si el ProgID usado tenía un prefijo distinto al LibraryName de la TypeLib, enlazar el alias directamente
                    var matchingLib = _registeredOcxLibraries.FirstOrDefault(l =>
                        string.Equals(l.FilePath, serverPath, StringComparison.OrdinalIgnoreCase));
                    if (matchingLib != null)
                    {
                        var matchingCtrl = matchingLib.Controls.FirstOrDefault(c =>
                            string.Equals(c.Clsid, resolvedClsid, StringComparison.OrdinalIgnoreCase));
                        if (matchingCtrl != null && _dynamicControlMembers.TryGetValue(matchingCtrl.Name, out var syms))
                        {
                            _dynamicControlMembers[clean] = syms;
                            return true;
                        }
                    }
                }
            }
            catch
            {
                // Ignorar errores de acceso al Registro
            }

            return _dynamicControlMembers.ContainsKey(clean);
        }
    }

    /// <summary>
    /// Limpia todos los controles dinámicos registrados.
    /// </summary>
    public static void ClearDynamicRegistrations()
    {
        lock (_syncLock)
        {
            _dynamicControlMembers.Clear();
            _registeredOcxLibraries.Clear();
            _attemptedRegistryLookups.Clear();
        }
    }

    /// <summary>
    /// Verifica si un tipo de control está registrado en el catálogo (intrínseco o dinámico).
    /// </summary>
    public static bool IsControlRegistered(string typeName)
    {
        string norm = NormalizeTypeName(typeName);
        if (_controlMembers.ContainsKey(norm))
        {
            return true;
        }

        lock (_syncLock)
        {
            if (_dynamicControlMembers.ContainsKey(norm))
            {
                return true;
            }
            int dot = norm.LastIndexOf('.');
            if (dot >= 0 && dot < norm.Length - 1)
            {
                string shortName = norm[(dot + 1)..];
                if (_controlMembers.ContainsKey(shortName) || _dynamicControlMembers.ContainsKey(shortName))
                {
                    return true;
                }
            }
            return TryEnsureControlRegisteredFromRegistry(norm);
        }
    }

    /// <summary>
    /// Busca un miembro específico en un tipo de control.
    /// </summary>
    public static VbSymbolItem? FindMember(string typeName, string memberName)
    {
        var members = GetMembersForControl(typeName);
        return members.FirstOrDefault(m => m.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase));
    }

    public static string NormalizeTypeName(string typeName)
    {
        if (string.IsNullOrWhiteSpace(typeName)) return string.Empty;
        string t = typeName.Trim();
        if (t.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        {
            t = t[3..];
        }
        return t;
    }

    #region Initializers

    private static VbSymbolItem Prop(string name, string type, string description) =>
        new()
        {
            Name = name,
            Kind = VbSymbolKind.Property,
            ReturnType = type,
            Description = description
        };

    private static VbSymbolItem Meth(string name, string? returnType, string description, params (string Name, string Type, bool Opt, string? Def, string Desc)[] parameters)
    {
        var pList = parameters.Select(p => new VbParameterInfo
        {
            Name = p.Name,
            Type = p.Type,
            IsOptional = p.Opt,
            DefaultValue = p.Def,
            Description = p.Desc
        }).ToList();

        return new VbSymbolItem
        {
            Name = name,
            Kind = VbSymbolKind.Method,
            ReturnType = returnType,
            Description = description,
            Parameters = pList
        };
    }

    private static VbSymbolItem Evt(string name, string description, params (string Name, string Type, bool Opt, string? Def, string Desc)[] parameters)
    {
        var pList = parameters.Select(p => new VbParameterInfo
        {
            Name = p.Name,
            Type = p.Type,
            IsOptional = p.Opt,
            DefaultValue = p.Def,
            Description = p.Desc
        }).ToList();

        return new VbSymbolItem
        {
            Name = name,
            Kind = VbSymbolKind.Event,
            Description = description,
            Parameters = pList
        };
    }

    private static List<VbSymbolItem> CreateBaseControlMembers()
    {
        return new List<VbSymbolItem>
        {
            Prop("Left", "Single", "Distancia entre el borde izquierdo del control y el contenedor."),
            Prop("Top", "Single", "Distancia entre el borde superior del control y el contenedor."),
            Prop("Width", "Single", "Ancho del control en Twips o unidades del contenedor."),
            Prop("Height", "Single", "Alto del control en Twips o unidades del contenedor."),
            Prop("Visible", "Boolean", "Determina si el control es visible o está oculto."),
            Prop("Enabled", "Boolean", "Determina si el control puede responder a eventos generados por el usuario."),
            Prop("Tag", "String", "Almacena datos adicionales asociados al control."),
            Prop("Name", "String", "Nombre que identifica al control en el código."),
            Prop("Parent", "Object", "Referencia al contenedor o formulario padre del control."),
            Prop("MousePointer", "Integer", "Indica el tipo de puntero de mouse que se muestra sobre el control."),
            Prop("ToolTipText", "String", "Texto emergente que aparece al pasar el cursor sobre el control."),
            Prop("TabIndex", "Integer", "Orden de tabulación del control dentro de su formulario."),
            Prop("TabStop", "Boolean", "Determina si el usuario puede tabular hacia este control."),
            Prop("HelpContextID", "Long", "Identificador de contexto de ayuda asociada."),
            Meth("Move", null, "Mueve el control y opcionalmente cambia su tamaño.",
                ("Left", "Single", false, null, "Nueva posición horizontal izquierda."),
                ("Top", "Single", true, null, "Nueva posición vertical superior."),
                ("Width", "Single", true, null, "Nuevo ancho."),
                ("Height", "Single", true, null, "Nuevo alto.")),
            Meth("Refresh", null, "Fuerza un redibujado completo del control."),
            Meth("ZOrder", null, "Coloca el control al frente o al fondo del orden Z.",
                ("Position", "Integer", true, "0", "0 para el frente, 1 para el fondo."))
        };
    }

    private static void AddFontProperties(List<VbSymbolItem> list)
    {
        list.Add(Prop("Font", "StdFont", "Objeto de fuente asociado al control."));
        list.Add(Prop("FontName", "String", "Nombre de la fuente tipográfica."));
        list.Add(Prop("FontSize", "Single", "Tamaño de la fuente en puntos."));
        list.Add(Prop("FontBold", "Boolean", "Indica si la fuente se dibuja en negrita."));
        list.Add(Prop("FontItalic", "Boolean", "Indica si la fuente se dibuja en cursiva."));
        list.Add(Prop("FontUnderline", "Boolean", "Indica si la fuente tiene subrayado."));
        list.Add(Prop("FontStrikethru", "Boolean", "Indica si la fuente tiene tachado."));
    }

    private static void InitializeCommonPropertiesAndMethods()
    {
        // Métodos y propiedades base ya encapsulados en CreateBaseControlMembers
    }

    private static void InitializeFormMembers()
    {
        _formMembers.AddRange(new[]
        {
            Prop("Caption", "String", "Texto que aparece en la barra de título del formulario."),
            Prop("Name", "String", "Nombre identificador del formulario."),
            Prop("BackColor", "Long", "Color de fondo del formulario."),
            Prop("ForeColor", "Long", "Color de primer plano para dibujo y texto."),
            Prop("Left", "Single", "Posición horizontal izquierda de la ventana."),
            Prop("Top", "Single", "Posición vertical superior de la ventana."),
            Prop("Width", "Single", "Ancho total del formulario."),
            Prop("Height", "Single", "Alto total del formulario."),
            Prop("ScaleWidth", "Single", "Ancho del área cliente interna."),
            Prop("ScaleHeight", "Single", "Alto del área cliente interna."),
            Prop("ScaleMode", "Integer", "Modo de escala de coordenadas (Twips, Pixels, etc.)."),
            Prop("BorderStyle", "Integer", "Estilo de borde de la ventana (None, Fixed Single, Sizable, etc.)."),
            Prop("WindowState", "Integer", "Estado de la ventana (0-Normal, 1-Minimizado, 2-Maximizado)."),
            Prop("StartupPosition", "Integer", "Posición de inicio al mostrar la ventana en pantalla."),
            Prop("AutoRedraw", "Boolean", "Determina si los gráficos del formulario persisten automáticamente en memoria."),
            Prop("ClipControls", "Boolean", "Determina si las regiones de dibujo recortan los controles contenidos."),
            Prop("Enabled", "Boolean", "Determina si el formulario responde a interacciones del usuario."),
            Prop("Visible", "Boolean", "Determina si la ventana es visible en pantalla."),
            Prop("KeyPreview", "Boolean", "Determina si los eventos de teclado del formulario se reciben antes que los controles."),
            Prop("MousePointer", "Integer", "Tipo de puntero de mouse en el área cliente del formulario."),
            Prop("Tag", "String", "Almacena datos adicionales asociados al formulario."),
            Prop("Icon", "StdPicture", "Ícono que se muestra en la barra de título y barra de tareas."),
            Prop("Picture", "StdPicture", "Imagen gráfica dibujada de fondo en el formulario."),
            Prop("MaxButton", "Boolean", "Indica si la ventana dispone de botón Maximizar."),
            Prop("MinButton", "Boolean", "Indica si la ventana dispone de botón Minimizar."),
            Prop("ControlBox", "Boolean", "Indica si la ventana dispone de menú de control del sistema."),
            Prop("Moveable", "Boolean", "Indica si el usuario puede mover la ventana con el mouse."),
            Prop("ShowInTaskbar", "Boolean", "Indica si la ventana aparece en la barra de tareas de Windows."),
            Prop("ActiveControl", "Object", "Retorna una referencia al control que tiene el foco de entrada."),
            Prop("Controls", "Object", "Colección de todos los controles contenidos en el formulario."),
            Prop("hDC", "Long", "Handle al Device Context de GDI del formulario."),
            Prop("hWnd", "Long", "Handle HWND de la ventana Win32 del formulario."),

            Meth("Show", null, "Muestra el formulario en pantalla de forma modal o no modal.",
                ("Modal", "Integer", true, "0", "0 = vbModeless (no modal), 1 = vbModal (modal)."),
                ("OwnerForm", "Form", true, null, "Formulario propietario.")),
            Meth("Hide", null, "Oculta el formulario sin descargarlo de la memoria."),
            Meth("SetFocus", null, "Asigna el foco de teclado al formulario."),
            Meth("Refresh", null, "Fuerza el redibujado inmediato del formulario y sus controles."),
            Meth("Cls", null, "Borra los gráficos y texto impresos en tiempo de ejecución."),
            Meth("Circle", null, "Dibuja un círculo, elipse o arco en el formulario.",
                ("Step", "Integer", true, null, "Indica coordenadas relativas."),
                ("X", "Single", false, null, "Coordenada X del centro."),
                ("Y", "Single", false, null, "Coordenada Y del centro."),
                ("Radius", "Single", false, null, "Radio del círculo."),
                ("Color", "Long", true, null, "Color del trazo."),
                ("Start", "Single", true, null, "Ángulo inicial."),
                ("End", "Single", true, null, "Ángulo final."),
                ("Aspect", "Single", true, null, "Relación de aspecto.")),
            Meth("Line", null, "Dibuja una línea o rectángulo en el formulario.",
                ("X1", "Single", false, null, "Coordenada X inicial."),
                ("Y1", "Single", false, null, "Coordenada Y inicial."),
                ("X2", "Single", false, null, "Coordenada X final."),
                ("Y2", "Single", false, null, "Coordenada Y final."),
                ("Color", "Long", true, null, "Color de la línea."),
                ("BF", "String", true, null, "B para caja, BF para caja rellena.")),
            Meth("PSet", null, "Dibuja un punto individual en las coordenadas especificadas.",
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y."),
                ("Color", "Long", true, null, "Color del punto.")),
            Meth("Point", "Long", "Retorna el valor de color RGB del píxel en las coordenadas indicadas.",
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y.")),
            Meth("Print", null, "Imprime texto directamente sobre el lienzo del formulario.",
                ("Output", "Variant", true, null, "Texto o expresión a imprimir.")),
            Meth("TextWidth", "Single", "Calcula el ancho que ocupará una cadena de texto con la fuente actual.",
                ("Str", "String", false, null, "Cadena de texto a medir.")),
            Meth("TextHeight", "Single", "Calcula el alto que ocupará una cadena de texto con la fuente actual.",
                ("Str", "String", false, null, "Cadena de texto a medir.")),
            Meth("Move", null, "Mueve y redimensiona la ventana del formulario.",
                ("Left", "Single", false, null, "Nueva posición izquierda."),
                ("Top", "Single", true, null, "Nueva posición superior."),
                ("Width", "Single", true, null, "Nuevo ancho."),
                ("Height", "Single", true, null, "Nuevo alto.")),
            Meth("PopupMenu", null, "Despliega un menú contextual emergente en la posición del mouse o coordenadas.",
                ("Menu", "Object", false, null, "Elemento de menú a desplegar."),
                ("Flags", "Integer", true, "0", "Banderas de alineación y clics."),
                ("X", "Single", true, null, "Coordenada X opcional."),
                ("Y", "Single", true, null, "Coordenada Y opcional.")),

            Evt("Load", "Se produce cuando el formulario se carga en la memoria."),
            Evt("Unload", "Se produce cuando el formulario se está descargando de la memoria.",
                ("Cancel", "Integer", false, null, "Establecer a distinto de 0 para cancelar la descarga.")),
            Evt("QueryUnload", "Se produce antes de que el formulario se descargue, indicando el motivo.",
                ("Cancel", "Integer", false, null, "Establecer a distinto de 0 para cancelar la descarga."),
                ("UnloadMode", "Integer", false, null, "Indica qué causó el cierre (0=Código, 1=Usuario, 2=Windows, etc.).")),
            Evt("Activate", "Se produce cuando el formulario se convierte en la ventana activa."),
            Evt("Deactivate", "Se produce cuando el formulario pierde el estado activo ante otra ventana del proyecto."),
            Evt("Resize", "Se produce cuando el tamaño del formulario ha cambiado."),
            Evt("Paint", "Se produce cuando es necesario repintar parte del área cliente del formulario."),
            Evt("Click", "Se produce cuando el usuario hace clic con el mouse en el formulario."),
            Evt("DblClick", "Se produce cuando el usuario hace doble clic con el mouse en el formulario."),
            Evt("KeyDown", "Se produce cuando se presiona una tecla mientras el formulario tiene el foco.",
                ("KeyCode", "Integer", false, null, "Código de la tecla."),
                ("Shift", "Integer", false, null, "Estado de Shift, Ctrl, Alt.")),
            Evt("KeyPress", "Se produce cuando se presiona y suelta una tecla de caracteres ANSI.",
                ("KeyAscii", "Integer", false, null, "Código ASCII de la tecla presionada.")),
            Evt("KeyUp", "Se produce cuando se libera una tecla previamente presionada.",
                ("KeyCode", "Integer", false, null, "Código de la tecla liberada."),
                ("Shift", "Integer", false, null, "Estado de Shift, Ctrl, Alt.")),
            Evt("MouseDown", "Se produce cuando se presiona un botón del mouse.",
                ("Button", "Integer", false, null, "Botón presionado (1=Izq, 2=Der, 4=Medio)."),
                ("Shift", "Integer", false, null, "Estado de Shift, Ctrl, Alt."),
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y.")),
            Evt("MouseMove", "Se produce cuando se mueve el mouse sobre el formulario.",
                ("Button", "Integer", false, null, "Botones presionados actualmente."),
                ("Shift", "Integer", false, null, "Estado de Shift, Ctrl, Alt."),
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y.")),
            Evt("MouseUp", "Se produce cuando se libera un botón del mouse.",
                ("Button", "Integer", false, null, "Botón liberado."),
                ("Shift", "Integer", false, null, "Estado de Shift, Ctrl, Alt."),
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y."))
        });
        AddFontProperties(_formMembers);
    }

    private static void InitializeCommandButton()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Caption", "String", "Texto mostrado en el botón de comando."),
            Prop("Default", "Boolean", "Determina si el botón responde a la tecla Enter por defecto."),
            Prop("Cancel", "Boolean", "Determina si el botón responde a la tecla Escape para cancelar."),
            Prop("Value", "Boolean", "Indica o simula la pulsación del botón de comando."),
            Prop("Style", "Integer", "Estilo visual del botón (0-Estándar, 1-Gráfico)."),
            Prop("Picture", "StdPicture", "Imagen que se muestra en el botón cuando Style es Gráfico."),
            Prop("DownPicture", "StdPicture", "Imagen cuando el botón está presionado."),
            Prop("DisabledPicture", "StdPicture", "Imagen cuando el botón está deshabilitado."),
            Prop("BackColor", "Long", "Color de fondo del botón (cuando Style es Gráfico)."),
            Meth("SetFocus", null, "Mueve el foco de entrada del teclado a este botón."),
            Evt("Click", "Se produce al hacer clic con el mouse o pulsar Enter sobre el botón."),
            Evt("GotFocus", "Se produce cuando el botón recibe el foco de entrada."),
            Evt("LostFocus", "Se produce cuando el botón pierde el foco de entrada.")
        });
        AddFontProperties(list);
        _controlMembers["CommandButton"] = list;
    }

    private static void InitializeTextBox()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Text", "String", "Texto contenido dentro de la caja de texto."),
            Prop("BackColor", "Long", "Color de fondo de la caja de texto."),
            Prop("ForeColor", "Long", "Color del texto contenido."),
            Prop("Locked", "Boolean", "Determina si el texto es de solo lectura impidiendo su edición."),
            Prop("MultiLine", "Boolean", "Determina si la caja de texto acepta y muestra múltiples líneas."),
            Prop("MaxLength", "Long", "Número máximo de caracteres que se pueden escribir (0 = sin límite)."),
            Prop("PasswordChar", "String", "Carácter que enmascara los caracteres ingresados (para contraseñas)."),
            Prop("ScrollBars", "Integer", "Barras de desplazamiento (0-Ninguna, 1-Horizontal, 2-Vertical, 3-Ambas)."),
            Prop("Alignment", "Integer", "Alineación del texto (0-Izquierda, 1-Derecha, 2-Centro)."),
            Prop("BorderStyle", "Integer", "Estilo de borde (0-Ninguno, 1-Fixed Single)."),
            Prop("HideSelection", "Boolean", "Determina si la selección de texto permanece visible al perder el foco."),
            Prop("SelStart", "Long", "Punto inicial del texto seleccionado o posición del cursor."),
            Prop("SelLength", "Long", "Número de caracteres seleccionados."),
            Prop("SelText", "String", "Cadena con el texto actualmente seleccionado."),
            Meth("SetFocus", null, "Mueve el foco de entrada a esta caja de texto."),
            Evt("Change", "Se produce cuando el contenido de la propiedad Text cambia."),
            Evt("Click", "Se produce al hacer clic sobre la caja de texto."),
            Evt("DblClick", "Se produce al hacer doble clic sobre la caja de texto."),
            Evt("GotFocus", "Se produce cuando el control recibe el foco de entrada."),
            Evt("LostFocus", "Se produce cuando el control pierde el foco de entrada."),
            Evt("KeyDown", "Se produce al presionar una tecla mientras el control tiene el foco.",
                ("KeyCode", "Integer", false, null, "Código de tecla."),
                ("Shift", "Integer", false, null, "Estado de teclas modificadoras.")),
            Evt("KeyPress", "Se produce al presionar una tecla de carácter ANSI.",
                ("KeyAscii", "Integer", false, null, "Código ASCII de la tecla.")),
            Evt("KeyUp", "Se produce al soltar una tecla previamente presionada.",
                ("KeyCode", "Integer", false, null, "Código de tecla."),
                ("Shift", "Integer", false, null, "Estado de teclas modificadoras."))
        });
        AddFontProperties(list);
        _controlMembers["TextBox"] = list;
    }

    private static void InitializeLabel()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Caption", "String", "Texto que muestra la etiqueta."),
            Prop("Alignment", "Integer", "Alineación del texto (0-Izquierda, 1-Derecha, 2-Centro)."),
            Prop("AutoSize", "Boolean", "Determina si la etiqueta ajusta su tamaño al texto automáticamente."),
            Prop("WordWrap", "Boolean", "Determina si el texto pasa a la siguiente línea si supera el ancho."),
            Prop("BackStyle", "Integer", "Estilo de fondo (0-Transparente, 1-Opaco)."),
            Prop("BorderStyle", "Integer", "Estilo de borde (0-Ninguno, 1-Fixed Single)."),
            Prop("BackColor", "Long", "Color de fondo de la etiqueta."),
            Prop("ForeColor", "Long", "Color del texto de la etiqueta."),
            Evt("Click", "Se produce al hacer clic sobre la etiqueta."),
            Evt("DblClick", "Se produce al hacer doble clic sobre la etiqueta.")
        });
        AddFontProperties(list);
        _controlMembers["Label"] = list;
    }

    private static void InitializeCheckBox()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Caption", "String", "Texto que acompaña a la casilla de verificación."),
            Prop("Value", "Integer", "Estado de la casilla (0-Desmarcado, 1-Marcado, 2-Indeterminado)."),
            Prop("Alignment", "Integer", "Ubicación del texto relativo a la casilla (0-Izquierda, 1-Derecha)."),
            Prop("Style", "Integer", "Estilo visual (0-Estándar, 1-Gráfico)."),
            Prop("BackColor", "Long", "Color de fondo de la casilla."),
            Prop("ForeColor", "Long", "Color del texto."),
            Meth("SetFocus", null, "Mueve el foco a esta casilla."),
            Evt("Click", "Se produce al alternar el estado de la casilla."),
            Evt("GotFocus", "Se produce cuando la casilla recibe el foco."),
            Evt("LostFocus", "Se produce cuando la casilla pierde el foco.")
        });
        AddFontProperties(list);
        _controlMembers["CheckBox"] = list;
    }

    private static void InitializeOptionButton()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Caption", "String", "Texto que acompaña al botón de opción."),
            Prop("Value", "Boolean", "Determina si el botón de opción está seleccionado."),
            Prop("Alignment", "Integer", "Ubicación del texto relativo al botón."),
            Prop("Style", "Integer", "Estilo visual (0-Estándar, 1-Gráfico)."),
            Prop("BackColor", "Long", "Color de fondo del botón."),
            Prop("ForeColor", "Long", "Color del texto."),
            Meth("SetFocus", null, "Mueve el foco a este botón de opción."),
            Evt("Click", "Se produce al seleccionar este botón de opción."),
            Evt("GotFocus", "Se produce cuando el botón de opción recibe el foco."),
            Evt("LostFocus", "Se produce cuando el botón de opción pierde el foco.")
        });
        AddFontProperties(list);
        _controlMembers["OptionButton"] = list;
    }

    private static void InitializeFrame()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Caption", "String", "Título mostrado en el marco contenedor."),
            Prop("BorderStyle", "Integer", "Estilo de borde (0-Ninguno, 1-Fixed Single)."),
            Prop("ClipControls", "Boolean", "Determina si se recortan las áreas de los controles secundarios."),
            Prop("BackColor", "Long", "Color de fondo del marco."),
            Prop("ForeColor", "Long", "Color del texto del título."),
            Evt("Click", "Se produce al hacer clic en el marco."),
            Evt("DblClick", "Se produce al hacer doble clic en el marco.")
        });
        AddFontProperties(list);
        _controlMembers["Frame"] = list;
    }

    private static void InitializeListBox()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("ListCount", "Integer", "Número total de elementos presentes en la lista."),
            Prop("ListIndex", "Integer", "Índice del elemento actualmente seleccionado (-1 si ninguno)."),
            Prop("MultiSelect", "Integer", "Modo de selección múltiple (0-Ninguno, 1-Simple, 2-Extendido)."),
            Prop("Sorted", "Boolean", "Indica si los elementos se ordenan alfabéticamente."),
            Prop("Columns", "Integer", "Número de columnas para mostrar los elementos."),
            Prop("Style", "Integer", "Estilo de la lista (0-Estándar, 1-Con casillas de verificación)."),
            Prop("IntegralHeight", "Boolean", "Indica si el alto de la lista se ajusta para mostrar líneas completas."),
            Prop("NewIndex", "Integer", "Índice del elemento añadido más recientemente."),
            Prop("BackColor", "Long", "Color de fondo de la lista."),
            Prop("ForeColor", "Long", "Color del texto de los elementos."),
            Meth("AddItem", null, "Añade un nuevo elemento a la lista.",
                ("Item", "String", false, null, "Texto del elemento."),
                ("Index", "Integer", true, null, "Posición opcional de inserción.")),
            Meth("RemoveItem", null, "Elimina un elemento de la lista por su índice.",
                ("Index", "Integer", false, null, "Índice del elemento a eliminar.")),
            Meth("Clear", null, "Elimina todos los elementos contenidos en la lista."),
            Meth("SetFocus", null, "Mueve el foco de entrada a la lista."),
            Evt("Click", "Se produce al seleccionar un elemento de la lista."),
            Evt("DblClick", "Se produce al hacer doble clic sobre un elemento de la lista."),
            Evt("Scroll", "Se produce al desplazar verticalmente la lista."),
            Evt("GotFocus", "Se produce cuando la lista recibe el foco."),
            Evt("LostFocus", "Se produce cuando la lista pierde el foco.")
        });
        AddFontProperties(list);
        _controlMembers["ListBox"] = list;
    }

    private static void InitializeComboBox()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Text", "String", "Texto actualmente seleccionado o escrito en el cuadro combinado."),
            Prop("ListCount", "Integer", "Número de elementos en la lista desplegable."),
            Prop("ListIndex", "Integer", "Índice del elemento seleccionado (-1 si ninguno)."),
            Prop("Style", "Integer", "Estilo del combo (0-Drop-down combo, 1-Simple, 2-Drop-down List)."),
            Prop("Locked", "Boolean", "Determina si el texto de la caja es de solo lectura."),
            Prop("Sorted", "Boolean", "Determina si los elementos se ordenan automáticamente."),
            Prop("IntegralHeight", "Boolean", "Determina si la lista desplegable se ajusta a líneas completas."),
            Prop("NewIndex", "Integer", "Índice del elemento más recientemente añadido."),
            Prop("SelStart", "Long", "Punto inicial del texto seleccionado."),
            Prop("SelLength", "Long", "Número de caracteres seleccionados."),
            Prop("SelText", "String", "Texto actualmente seleccionado en el cuadro."),
            Prop("BackColor", "Long", "Color de fondo."),
            Prop("ForeColor", "Long", "Color del texto."),
            Meth("AddItem", null, "Añade un nuevo elemento a la lista desplegable.",
                ("Item", "String", false, null, "Texto del elemento."),
                ("Index", "Integer", true, null, "Posición de inserción.")),
            Meth("RemoveItem", null, "Elimina un elemento por su índice.",
                ("Index", "Integer", false, null, "Índice del elemento a remover.")),
            Meth("Clear", null, "Elimina todos los elementos de la lista desplegable."),
            Meth("SetFocus", null, "Mueve el foco de entrada al cuadro combinado."),
            Evt("Change", "Se produce cuando se modifica el texto escrito."),
            Evt("Click", "Se produce al seleccionar un elemento de la lista."),
            Evt("DropDown", "Se produce cuando se despliega la lista hacia abajo."),
            Evt("Scroll", "Se produce al desplazar los elementos de la lista desplegable."),
            Evt("GotFocus", "Se produce cuando el cuadro combinado recibe el foco."),
            Evt("LostFocus", "Se produce cuando el cuadro combinado pierde el foco.")
        });
        AddFontProperties(list);
        _controlMembers["ComboBox"] = list;
    }

    private static void InitializePictureBox()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Picture", "StdPicture", "Imagen gráfica asignada al PictureBox."),
            Prop("Image", "StdPicture", "Retorna un handle a la imagen persistente en memoria."),
            Prop("AutoSize", "Boolean", "Determina si el control se redimensiona automáticamente para ajustarse a la imagen."),
            Prop("AutoRedraw", "Boolean", "Determina si las salidas gráficas son persistentes."),
            Prop("ClipControls", "Boolean", "Determina si se recortan los gráficos alrededor de los controles contenidos."),
            Prop("ScaleMode", "Integer", "Modo de escala de coordenadas de dibujo."),
            Prop("BorderStyle", "Integer", "Estilo de borde (0-Ninguno, 1-Fixed Single)."),
            Prop("BackColor", "Long", "Color de fondo del lienzo."),
            Prop("ForeColor", "Long", "Color de trazo para texto y primitivas de dibujo."),
            Meth("Cls", null, "Borra las operaciones de dibujo y texto en tiempo de ejecución."),
            Meth("Circle", null, "Dibuja un círculo, elipse o arco en el PictureBox.",
                ("X", "Single", false, null, "Coordenada X del centro."),
                ("Y", "Single", false, null, "Coordenada Y del centro."),
                ("Radius", "Single", false, null, "Radio."),
                ("Color", "Long", true, null, "Color del trazo.")),
            Meth("Line", null, "Dibuja una línea o rectángulo en el PictureBox.",
                ("X1", "Single", false, null, "Coordenada X inicial."),
                ("Y1", "Single", false, null, "Coordenada Y inicial."),
                ("X2", "Single", false, null, "Coordenada X final."),
                ("Y2", "Single", false, null, "Coordenada Y final."),
                ("Color", "Long", true, null, "Color de la línea."),
                ("BF", "String", true, null, "B para caja, BF para caja rellena.")),
            Meth("PSet", null, "Dibuja un punto individual.",
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y."),
                ("Color", "Long", true, null, "Color del punto.")),
            Meth("Point", "Long", "Retorna el color RGB del punto en las coordenadas indicadas.",
                ("X", "Single", false, null, "Coordenada X."),
                ("Y", "Single", false, null, "Coordenada Y.")),
            Meth("Print", null, "Imprime texto sobre el lienzo.",
                ("Output", "Variant", true, null, "Texto a imprimir.")),
            Meth("TextWidth", "Single", "Retorna el ancho del texto medido.",
                ("Str", "String", false, null, "Cadena a medir.")),
            Meth("TextHeight", "Single", "Retorna el alto del texto medido.",
                ("Str", "String", false, null, "Cadena a medir.")),
            Meth("SetFocus", null, "Mueve el foco de entrada al PictureBox."),
            Evt("Click", "Se produce al hacer clic en el PictureBox."),
            Evt("DblClick", "Se produce al hacer doble clic en el PictureBox."),
            Evt("Paint", "Se produce cuando el control necesita repintarse."),
            Evt("Resize", "Se produce cuando el tamaño del PictureBox cambia.")
        });
        AddFontProperties(list);
        _controlMembers["PictureBox"] = list;
    }

    private static void InitializeImage()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Picture", "StdPicture", "Imagen gráfica mostrada por el control Image."),
            Prop("Stretch", "Boolean", "Determina si la imagen se estira para llenar el control."),
            Prop("BorderStyle", "Integer", "Estilo de borde (0-Ninguno, 1-Fixed Single)."),
            Evt("Click", "Se produce al hacer clic sobre la imagen."),
            Evt("DblClick", "Se produce al hacer doble clic sobre la imagen.")
        });
        _controlMembers["Image"] = list;
    }

    private static void InitializeTimer()
    {
        var list = new List<VbSymbolItem>
        {
            Prop("Name", "String", "Nombre que identifica al Timer en el código."),
            Prop("Enabled", "Boolean", "Determina si el temporizador está activo y emite eventos periódicos."),
            Prop("Interval", "Long", "Intervalo en milisegundos entre eventos Timer (0 a 65535)."),
            Prop("Tag", "String", "Almacena datos adicionales asociados al temporizador."),
            Evt("Timer", "Se produce a intervalos periódicos especificados por la propiedad Interval.")
        };
        _controlMembers["Timer"] = list;
    }

    private static void InitializeHScrollBar()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Min", "Integer", "Valor numérico mínimo de la barra de desplazamiento."),
            Prop("Max", "Integer", "Valor numérico máximo de la barra de desplazamiento."),
            Prop("Value", "Integer", "Posición actual del indicador de la barra."),
            Prop("SmallChange", "Integer", "Magnitud del desplazamiento al hacer clic en las flechas."),
            Prop("LargeChange", "Integer", "Magnitud del desplazamiento al hacer clic en el riel de la barra."),
            Meth("SetFocus", null, "Mueve el foco de entrada a la barra."),
            Evt("Change", "Se produce cuando la posición Value de la barra cambia."),
            Evt("Scroll", "Se produce mientras el usuario arrastra la barra de desplazamiento.")
        });
        _controlMembers["HScrollBar"] = list;
    }

    private static void InitializeVScrollBar()
    {
        var list = CreateBaseControlMembers();
        list.AddRange(new[]
        {
            Prop("Min", "Integer", "Valor numérico mínimo de la barra de desplazamiento."),
            Prop("Max", "Integer", "Valor numérico máximo de la barra de desplazamiento."),
            Prop("Value", "Integer", "Posición actual del indicador de la barra."),
            Prop("SmallChange", "Integer", "Magnitud del desplazamiento al hacer clic en las flechas."),
            Prop("LargeChange", "Integer", "Magnitud del desplazamiento al hacer clic en el riel de la barra."),
            Meth("SetFocus", null, "Mueve el foco de entrada a la barra."),
            Evt("Change", "Se produce cuando la posición Value de la barra cambia."),
            Evt("Scroll", "Se produce mientras el usuario arrastra la barra de desplazamiento.")
        });
        _controlMembers["VScrollBar"] = list;
    }

    private static void InitializeShape()
    {
        var list = new List<VbSymbolItem>
        {
            Prop("Name", "String", "Nombre que identifica a la figura geométrica."),
            Prop("Left", "Single", "Posición horizontal izquierda."),
            Prop("Top", "Single", "Posición vertical superior."),
            Prop("Width", "Single", "Ancho de la figura."),
            Prop("Height", "Single", "Alto de la figura."),
            Prop("Visible", "Boolean", "Determina si la figura es visible."),
            Prop("Shape", "Integer", "Tipo de figura (0-Rectángulo, 1-Cuadrado, 2-Óvalo, 3-Círculo, 4-Rectángulo redondeado, 5-Cuadrado redondeado)."),
            Prop("BackColor", "Long", "Color de fondo interior de la figura."),
            Prop("BackStyle", "Integer", "Estilo de fondo (0-Transparente, 1-Opaco)."),
            Prop("BorderColor", "Long", "Color del borde exterior."),
            Prop("BorderStyle", "Integer", "Estilo del trazo del borde."),
            Prop("BorderWidth", "Integer", "Grosor en píxeles del borde."),
            Prop("FillColor", "Long", "Color del patrón de relleno."),
            Prop("FillStyle", "Integer", "Estilo del relleno (0-Sólido, 1-Transparente, etc.)."),
            Prop("Tag", "String", "Datos asociados."),
            Meth("Move", null, "Mueve y redimensiona la figura.",
                ("Left", "Single", false, null, "Nueva posición izquierda."),
                ("Top", "Single", true, null, "Nueva posición superior."),
                ("Width", "Single", true, null, "Nuevo ancho."),
                ("Height", "Single", true, null, "Nuevo alto.")),
            Meth("Refresh", null, "Fuerza el redibujado de la figura."),
            Meth("ZOrder", null, "Ajusta la posición en el orden Z.",
                ("Position", "Integer", true, "0", "0 frente, 1 fondo."))
        };
        _controlMembers["Shape"] = list;
    }

    private static void InitializeLine()
    {
        var list = new List<VbSymbolItem>
        {
            Prop("Name", "String", "Nombre identificador de la línea."),
            Prop("Visible", "Boolean", "Determina si la línea es visible."),
            Prop("X1", "Single", "Coordenada horizontal del punto inicial."),
            Prop("Y1", "Single", "Coordenada vertical del punto inicial."),
            Prop("X2", "Single", "Coordenada horizontal del punto final."),
            Prop("Y2", "Single", "Coordenada vertical del punto final."),
            Prop("BorderColor", "Long", "Color de la línea."),
            Prop("BorderStyle", "Integer", "Estilo de trazo de la línea."),
            Prop("BorderWidth", "Integer", "Grosor de la línea en píxeles."),
            Prop("Tag", "String", "Datos adicionales."),
            Meth("Refresh", null, "Fuerza el redibujado de la línea."),
            Meth("ZOrder", null, "Ajusta la posición en el orden Z.",
                ("Position", "Integer", true, "0", "0 frente, 1 fondo."))
        };
        _controlMembers["Line"] = list;
    }

    #endregion
}
