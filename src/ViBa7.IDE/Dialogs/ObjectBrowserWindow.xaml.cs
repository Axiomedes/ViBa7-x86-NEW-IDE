using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Documents;
using ViBa7.ProjectModel.IntelliSense;
using ViBa7.ProjectModel.Navigation;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.IDE.Dialogs;

public class ObMemberItem
{
    public string Name { get; set; } = string.Empty;
    public string Kind { get; set; } = "Method"; // Sub, Function, Property, Event, Const, Variable
    public string BadgeIcon { get; set; } = "⚡";
    public string Scope { get; set; } = "Public";
    public string Signature { get; set; } = string.Empty;
    public string ReturnType { get; set; } = string.Empty;
    public string FullSignatureText { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? FilePath { get; set; }
    public int LineNumber { get; set; }
    public ObClassItem? DeclaringClass { get; set; }

    public string TypeOrScopeDisplay
    {
        get
        {
            if (!string.IsNullOrEmpty(ReturnType)) return ReturnType;
            if (Kind == "Const") return "Constante";
            if (Kind == "Event") return "Evento";
            if (Kind == "Property") return "Propiedad";
            return Kind != "Sub" ? Kind : string.Empty;
        }
    }
}

public class ObClassItem
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName => !string.IsNullOrEmpty(LibraryName) && LibraryName != "Proyecto"
        ? $"{Name}"
        : Name;
    public string LibraryName { get; set; } = string.Empty;
    public string Kind { get; set; } = "Class"; // Class, Module, Form, UserControl, Control, Enum, Globals
    public string BadgeIcon { get; set; } = "📄";
    public string Description { get; set; } = string.Empty;
    public string? FilePath { get; set; }
    public int LineNumber { get; set; }
    public List<ObMemberItem> Members { get; } = [];
}

public class ObLibraryItem
{
    public string Name { get; set; } = string.Empty;
    public string DisplayTitle { get; set; } = string.Empty;
    public List<ObClassItem> Classes { get; } = [];

    public override string ToString() => DisplayTitle;
}

/// <summary>
/// Ventana clásica de inspección del Examinador de Objetos de VB6 (F2).
/// Permite explorar todas las bibliotecas de tipos, clases, métodos, propiedades,
/// eventos y constantes disponibles en el entorno y en el proyecto activo.
/// </summary>
public partial class ObjectBrowserWindow : Window
{
    private readonly VB6Project? _project;
    private readonly List<ObLibraryItem> _libraries = [];
    private readonly List<(string LibName, string ClassName, string? MemberName)> _history = [];
    private int _historyIndex = -1;
    private bool _isNavigatingHistory;

    public Action<string, int>? RequestNavigateToCode { get; set; }

    public ObjectBrowserWindow(VB6Project? project = null)
    {
        _project = project;
        InitializeComponent();

        BuildSymbolCatalog();
        PopulateLibrariesComboBox();

        KeyDown += ObjectBrowserWindow_KeyDown;
    }

    private void ObjectBrowserWindow_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            Close();
            e.Handled = true;
        }
        else if (e.Key is Key.Enter or Key.F7)
        {
            NavigateToSelectedMemberDefinition();
            e.Handled = true;
        }
    }

    private void BuildSymbolCatalog()
    {
        _libraries.Clear();

        // 1. Biblioteca Especial: <Todas las bibliotecas>
        var allLibs = new ObLibraryItem
        {
            Name = "<All>",
            DisplayTitle = "<Todas las bibliotecas>"
        };
        _libraries.Add(allLibs);

        // 2. Biblioteca del Proyecto Activo
        if (_project != null)
        {
            string projName = !string.IsNullOrWhiteSpace(_project.Name) ? _project.Name : "Proyecto1";
            var projLib = new ObLibraryItem
            {
                Name = projName,
                DisplayTitle = $"{projName} ({Path.GetFileName(_project.ProjectFilePath ?? projName)})"
            };

            LoadProjectSymbols(projLib, _project);
            _libraries.Add(projLib);
        }

        // 3. Biblioteca VB (Objetos intrínsecos y Controles de Visual Basic)
        var vbLib = new ObLibraryItem
        {
            Name = "VB",
            DisplayTitle = "VB (Visual Basic objects and procedures)"
        };
        LoadVbLibrary(vbLib);
        _libraries.Add(vbLib);

        // 4. Biblioteca VBA (Visual Basic For Applications - Runtime, Enums y Funciones)
        var vbaLib = new ObLibraryItem
        {
            Name = "VBA",
            DisplayTitle = "VBA (Visual Basic For Applications)"
        };
        LoadVbaLibrary(vbaLib);
        _libraries.Add(vbaLib);

        // 5. Bibliotecas Referenciadas y Controles OCX cargados
        LoadReferencedLibraries();

        // Poblar <Todas las bibliotecas> con la suma de clases
        foreach (var lib in _libraries.Skip(1))
        {
            allLibs.Classes.AddRange(lib.Classes);
        }
    }

    private void LoadProjectSymbols(ObLibraryItem projLib, VB6Project project)
    {
        foreach (var item in project.Items)
        {
            if (item.Type is not (ProjectItemType.Form or ProjectItemType.Module or ProjectItemType.Class or ProjectItemType.UserControl or ProjectItemType.PropertyPage))
                continue;

            string kindStr = item.Type switch
            {
                ProjectItemType.Form => "Form",
                ProjectItemType.Module => "Module",
                ProjectItemType.Class => "Class",
                ProjectItemType.UserControl => "UserControl",
                ProjectItemType.PropertyPage => "PropertyPage",
                _ => "Class"
            };

            string icon = item.Type switch
            {
                ProjectItemType.Form => "🗔",
                ProjectItemType.Module => "📦",
                ProjectItemType.Class => "📄",
                ProjectItemType.UserControl => "🧩",
                _ => "📄"
            };

            var cls = new ObClassItem
            {
                Name = item.Name,
                LibraryName = projLib.Name,
                Kind = kindStr,
                BadgeIcon = icon,
                Description = $"{kindStr} del proyecto '{projLib.Name}' ({item.RelativePath})",
                FilePath = item.AbsolutePath,
                LineNumber = 1
            };

            // Parsear miembros del archivo de código
            if (!string.IsNullOrEmpty(item.AbsolutePath) && File.Exists(item.AbsolutePath))
            {
                try
                {
                    var doc = SafeFileStorage.LoadDocument(item.AbsolutePath);
                    string codeText = !string.IsNullOrEmpty(doc.Code) ? doc.Code : doc.Text;

                    var symbols = CodeStructureParser.ParseMembers(codeText, item.AbsolutePath);
                    foreach (var s in symbols)
                    {
                        string memberIcon = s.Kind switch
                        {
                            CodeMemberKind.Sub => "⚡",
                            CodeMemberKind.Function => "⚡",
                            CodeMemberKind.PropertyGet or CodeMemberKind.PropertyLet or CodeMemberKind.PropertySet => "🔧",
                            CodeMemberKind.Event => "⚡",
                            CodeMemberKind.Constant => "🔒",
                            _ => "🔹"
                        };

                        string sigText = s.Kind switch
                        {
                            CodeMemberKind.Sub => $"Sub {s.Name}{s.Signature}",
                            CodeMemberKind.Function => $"Function {s.Name}{s.Signature}" + (!string.IsNullOrEmpty(s.TypeName) ? $" As {s.TypeName}" : ""),
                            CodeMemberKind.PropertyGet => $"Property Get {s.Name}{s.Signature}" + (!string.IsNullOrEmpty(s.TypeName) ? $" As {s.TypeName}" : ""),
                            CodeMemberKind.PropertyLet => $"Property Let {s.Name}{s.Signature}",
                            CodeMemberKind.PropertySet => $"Property Set {s.Name}{s.Signature}",
                            CodeMemberKind.Event => $"Event {s.Name}{s.Signature}",
                            CodeMemberKind.Constant => $"Const {s.Name}" + (!string.IsNullOrEmpty(s.TypeName) ? $" As {s.TypeName}" : ""),
                            _ => $"Dim {s.Name}" + (!string.IsNullOrEmpty(s.TypeName) ? $" As {s.TypeName}" : "")
                        };

                        cls.Members.Add(new ObMemberItem
                        {
                            Name = s.Name,
                            Kind = s.Kind.ToString(),
                            BadgeIcon = memberIcon,
                            Scope = s.Scope,
                            Signature = s.Signature,
                            ReturnType = s.TypeName,
                            FullSignatureText = $"{s.Scope} {sigText}",
                            Description = $"Miembro definido en '{item.Name}' ({Path.GetFileName(item.AbsolutePath)}:L{s.LineNumber})",
                            FilePath = item.AbsolutePath,
                            LineNumber = s.LineNumber,
                            DeclaringClass = cls
                        });
                    }
                }
                catch { }
            }

            projLib.Classes.Add(cls);
        }
    }

    private void LoadVbLibrary(ObLibraryItem vbLib)
    {
        // 1. Controles Intrínsecos
        foreach (var ctrlType in VbControlMemberCatalog.GetKnownControlTypes())
        {
            var cls = new ObClassItem
            {
                Name = ctrlType,
                LibraryName = "VB",
                Kind = "Control",
                BadgeIcon = ctrlType == "Form" ? "🗔" : "🔲",
                Description = $"Control intrínseco de Visual Basic 6.0 ({ctrlType})"
            };

            var members = VbControlMemberCatalog.GetMembersForControl(ctrlType);
            foreach (var m in members)
            {
                string icon = m.Kind switch
                {
                    VbSymbolKind.Method => "⚡",
                    VbSymbolKind.Event => "⚡",
                    VbSymbolKind.Property => "🔧",
                    _ => "🔹"
                };

                cls.Members.Add(new ObMemberItem
                {
                    Name = m.Name,
                    Kind = m.Kind.ToString(),
                    BadgeIcon = icon,
                    ReturnType = m.ReturnType ?? "Variant",
                    FullSignatureText = $"{m.Kind} {m.Name} As {m.ReturnType ?? "Variant"}",
                    Description = !string.IsNullOrWhiteSpace(m.Description) ? m.Description : $"Propiedad o método de VB.{ctrlType}",
                    DeclaringClass = cls
                });
            }

            vbLib.Classes.Add(cls);
        }

        // 2. Objetos Globales de VB: App, Clipboard, Screen, Printer, Err
        AddGlobalVbObject(vbLib, "App", "Objeto global App que provee información sobre la aplicación en ejecución (EXEName, Path, Title, PrevInstance, Major, Minor).",
            ("EXEName", "String", "Retorna el nombre del archivo ejecutable sin extensión."),
            ("Path", "String", "Retorna la ruta absoluta del directorio donde reside el proyecto o ejecutable."),
            ("Title", "String", "Retorna o establece el título de la aplicación en ejecución."),
            ("Major", "Integer", "Retorna el número de versión mayor."),
            ("Minor", "Integer", "Retorna el número de versión menor."),
            ("Revision", "Integer", "Retorna el número de revisión."),
            ("PrevInstance", "Boolean", "Indica si ya existe una instancia previa de la aplicación ejecutándose."));

        AddGlobalVbObject(vbLib, "Clipboard", "Objeto global Clipboard para interactuar con el portapapeles del sistema.",
            ("Clear", "Sub", "Vacía el contenido del portapapeles."),
            ("GetData", "Variant", "Retorna una imagen o gráfico del portapapeles."),
            ("GetText", "String", "Retorna una cadena de texto del portapapeles."),
            ("SetData", "Sub", "Coloca una imagen o gráfico en el portapapeles."),
            ("SetText", "Sub", "Coloca una cadena de texto en el portapapeles."));

        AddGlobalVbObject(vbLib, "Screen", "Objeto global Screen que provee información sobre la pantalla y el cursor activo.",
            ("Width", "Single", "Retorna el ancho de la pantalla en twips."),
            ("Height", "Single", "Retorna el alto de la pantalla en twips."),
            ("TwipsPerPixelX", "Single", "Retorna la cantidad de twips por píxel horizontal."),
            ("TwipsPerPixelY", "Single", "Retorna la cantidad de twips por píxel vertical."),
            ("MousePointer", "Integer", "Retorna o establece el puntero del mouse para toda la pantalla."));

        AddGlobalVbObject(vbLib, "Err", "Objeto de gestión de errores en tiempo de ejecución (ErrObject).",
            ("Number", "Long", "Retorna o establece un valor numérico que identifica un error."),
            ("Description", "String", "Retorna o establece una cadena descriptiva asociada al error."),
            ("Source", "String", "Retorna o establece el nombre del objeto o aplicación que generó el error."),
            ("Clear", "Sub", "Borra todas las propiedades del objeto Err."),
            ("Raise", "Sub", "Genera un error en tiempo de ejecución de manera explícita."));
    }

    private void AddGlobalVbObject(ObLibraryItem vbLib, string objName, string desc, params (string Name, string Type, string Desc)[] members)
    {
        var cls = new ObClassItem
        {
            Name = objName,
            LibraryName = "VB",
            Kind = "Class",
            BadgeIcon = "📄",
            Description = desc
        };

        foreach (var m in members)
        {
            cls.Members.Add(new ObMemberItem
            {
                Name = m.Name,
                Kind = m.Type == "Sub" ? "Sub" : "Property",
                BadgeIcon = m.Type == "Sub" ? "⚡" : "🔧",
                ReturnType = m.Type != "Sub" ? m.Type : string.Empty,
                FullSignatureText = m.Type == "Sub" ? $"Sub {m.Name}()" : $"Property {m.Name} As {m.Type}",
                Description = m.Desc,
                DeclaringClass = cls
            });
        }

        vbLib.Classes.Add(cls);
    }

    private void LoadVbaLibrary(ObLibraryItem vbaLib)
    {
        // 1. Módulo Global con Funciones Intrínsecas del Runtime
        var globals = new ObClassItem
        {
            Name = "<Globales>",
            LibraryName = "VBA",
            Kind = "Globals",
            BadgeIcon = "📦",
            Description = "Funciones intrínsecas globales de Visual Basic For Applications (VBA)"
        };

        foreach (var f in VbRuntimeCatalog.GetAllFunctions())
        {
            string sig = "(" + string.Join(", ", f.Parameters.Select(p =>
                (p.IsOptional ? "[" : "") + p.Name + (p.Type != "Variant" ? $" As {p.Type}" : "") +
                (!string.IsNullOrEmpty(p.DefaultValue) ? $" = {p.DefaultValue}" : "") + (p.IsOptional ? "]" : ""))) + ")";

            globals.Members.Add(new ObMemberItem
            {
                Name = f.Name,
                Kind = "Function",
                BadgeIcon = "⚡",
                ReturnType = f.ReturnType ?? "Variant",
                Signature = sig,
                FullSignatureText = $"Function {f.Name}{sig} As {f.ReturnType ?? "Variant"}",
                Description = !string.IsNullOrWhiteSpace(f.Description) ? f.Description : $"Función intrínseca VBA.{f.Name}",
                DeclaringClass = globals
            });
        }

        vbaLib.Classes.Add(globals);

        // 2. Agrupar Constantes en sus correspondientes Enums de VBA
        var enumGroups = VbRuntimeCatalog.GetAllConstants()
            .GroupBy(c => !string.IsNullOrEmpty(c.ReturnType) ? c.ReturnType : "VbConstants");

        foreach (var grp in enumGroups)
        {
            var enumClass = new ObClassItem
            {
                Name = grp.Key,
                LibraryName = "VBA",
                Kind = "Enum",
                BadgeIcon = "📋",
                Description = $"Enumeración de constantes estándar de Visual Basic For Applications ({grp.Key})"
            };

            foreach (var c in grp)
            {
                enumClass.Members.Add(new ObMemberItem
                {
                    Name = c.Name,
                    Kind = "Const",
                    BadgeIcon = "🔒",
                    ReturnType = grp.Key,
                    FullSignatureText = $"Const {c.Name} As {grp.Key}",
                    Description = !string.IsNullOrWhiteSpace(c.Description) ? c.Description : $"Constante de {grp.Key}",
                    DeclaringClass = enumClass
                });
            }

            vbaLib.Classes.Add(enumClass);
        }
    }

    private void LoadReferencedLibraries()
    {
        var registeredCatalog = VbControlMemberCatalog.GetRegisteredOcxLibraries();
        foreach (var lib in registeredCatalog)
        {
            string libName = !string.IsNullOrWhiteSpace(lib.LibraryName) ? lib.LibraryName : Path.GetFileNameWithoutExtension(lib.FilePath);
            if (string.IsNullOrWhiteSpace(libName) || _libraries.Any(l => l.Name.Equals(libName, StringComparison.OrdinalIgnoreCase)))
                continue;

            var obLib = new ObLibraryItem
            {
                Name = libName,
                DisplayTitle = $"{libName} ({Path.GetFileName(lib.FilePath)})"
            };

            // Controles y CoClasses
            foreach (var c in lib.Controls)
            {
                var cls = new ObClassItem
                {
                    Name = c.Name,
                    LibraryName = libName,
                    Kind = c.IsControl ? "Control" : "Class",
                    BadgeIcon = c.IsControl ? "🔲" : "📄",
                    Description = !string.IsNullOrWhiteSpace(c.Description) ? c.Description : $"Tipo COM {c.Name} en {libName}"
                };

                // Miembros (Métodos y Propiedades)
                foreach (var m in c.Members)
                {
                    cls.Members.Add(new ObMemberItem
                    {
                        Name = m.Name,
                        Kind = m.Kind == VbSymbolKind.Method ? "Method" : "Property",
                        BadgeIcon = m.Kind == VbSymbolKind.Method ? "⚡" : "🔧",
                        ReturnType = m.ReturnType,
                        FullSignatureText = $"{m.Kind} {m.Name} As {m.ReturnType}",
                        Description = !string.IsNullOrWhiteSpace(m.Description) ? m.Description : $"Miembro de {libName}.{c.Name}",
                        DeclaringClass = cls
                    });
                }

                // Eventos
                foreach (var ev in c.Events)
                {
                    cls.Members.Add(new ObMemberItem
                    {
                        Name = ev.Name,
                        Kind = "Event",
                        BadgeIcon = "⚡",
                        FullSignatureText = $"Event {ev.Name}()",
                        Description = !string.IsNullOrWhiteSpace(ev.Description) ? ev.Description : $"Evento de {libName}.{c.Name}",
                        DeclaringClass = cls
                    });
                }

                obLib.Classes.Add(cls);
            }

            // Enums
            foreach (var en in lib.Enums)
            {
                var enCls = new ObClassItem
                {
                    Name = en.Name,
                    LibraryName = libName,
                    Kind = "Enum",
                    BadgeIcon = "📋",
                    Description = !string.IsNullOrWhiteSpace(en.Description) ? en.Description : $"Enumeración COM {en.Name}"
                };

                foreach (var it in en.Items)
                {
                    enCls.Members.Add(new ObMemberItem
                    {
                        Name = it.Name,
                        Kind = "Const",
                        BadgeIcon = "🔒",
                        ReturnType = en.Name,
                        FullSignatureText = $"Const {it.Name} = {it.Value}",
                        Description = !string.IsNullOrWhiteSpace(it.Description) ? it.Description : $"Valor de {en.Name} = {it.Value}",
                        DeclaringClass = enCls
                    });
                }

                obLib.Classes.Add(enCls);
            }

            if (obLib.Classes.Count > 0)
            {
                _libraries.Add(obLib);
            }
        }
    }

    private void PopulateLibrariesComboBox()
    {
        CboLibraries.ItemsSource = _libraries;
        CboLibraries.SelectedIndex = _libraries.Count > 1 ? 1 : 0; // Seleccionar el proyecto o <Todas las bibliotecas>
    }

    private void CboLibraries_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ApplyFilter();
    }

    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        ApplyFilter();
    }

    private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
    {
        TxtSearch.Clear();
        TxtSearch.Focus();
    }

    private void ApplyFilter()
    {
        if (CboLibraries.SelectedItem is not ObLibraryItem selectedLib) return;

        string query = TxtSearch?.Text?.Trim() ?? string.Empty;

        var matchingClasses = new List<ObClassItem>();

        foreach (var cls in selectedLib.Classes)
        {
            if (string.IsNullOrEmpty(query))
            {
                matchingClasses.Add(cls);
            }
            else
            {
                // Coincide el nombre de la clase o alguno de sus miembros
                bool classMatches = cls.Name.Contains(query, StringComparison.OrdinalIgnoreCase);
                bool membersMatch = cls.Members.Any(m => m.Name.Contains(query, StringComparison.OrdinalIgnoreCase));

                if (classMatches || membersMatch)
                {
                    matchingClasses.Add(cls);
                }
            }
        }

        LstClasses.ItemsSource = matchingClasses.OrderBy(c => c.Name).ToList();

        if (matchingClasses.Count > 0)
        {
            LstClasses.SelectedIndex = 0;
        }
        else
        {
            LstMembers.ItemsSource = null;
            TxtSignature.Text = string.Empty;
            TxtMemberOf.Text = string.Empty;
            TxtDescription.Text = string.Empty;
            BtnGoToDefinition.IsEnabled = false;
        }
    }

    private void LstClasses_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LstClasses.SelectedItem is ObClassItem cls)
        {
            string query = TxtSearch?.Text?.Trim() ?? string.Empty;

            var membersToShow = string.IsNullOrEmpty(query)
                ? cls.Members
                : cls.Members.Where(m => m.Name.Contains(query, StringComparison.OrdinalIgnoreCase) || cls.Name.Contains(query, StringComparison.OrdinalIgnoreCase)).ToList();

            TxtMembersHeader.Text = $"Miembros de '{cls.Name}':";
            LstMembers.ItemsSource = membersToShow.OrderBy(m => m.Name).ToList();

            if (membersToShow.Count > 0)
            {
                LstMembers.SelectedIndex = 0;
            }
            else
            {
                ShowClassDetails(cls);
            }

            RecordHistory(cls.LibraryName, cls.Name, null);
        }
    }

    private void LstMembers_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LstMembers.SelectedItem is ObMemberItem member)
        {
            ShowMemberDetails(member);
            RecordHistory(member.DeclaringClass?.LibraryName ?? "", member.DeclaringClass?.Name ?? "", member.Name);
        }
        else if (LstClasses.SelectedItem is ObClassItem cls)
        {
            ShowClassDetails(cls);
        }
    }

    private void ShowMemberDetails(ObMemberItem member)
    {
        TxtSignature.Text = member.FullSignatureText;
        TxtMemberOf.Text = member.DeclaringClass != null
            ? $"Miembro de {member.DeclaringClass.LibraryName}.{member.DeclaringClass.Name}"
            : "Miembro";
        TxtDescription.Text = member.Description;

        bool canNavigate = !string.IsNullOrEmpty(member.FilePath) && File.Exists(member.FilePath);
        BtnGoToDefinition.IsEnabled = canNavigate;
    }

    private void ShowClassDetails(ObClassItem cls)
    {
        TxtSignature.Text = $"{cls.Kind} {cls.Name}";
        TxtMemberOf.Text = $"Clase / Módulo de la biblioteca '{cls.LibraryName}'";
        TxtDescription.Text = cls.Description;

        bool canNavigate = !string.IsNullOrEmpty(cls.FilePath) && File.Exists(cls.FilePath);
        BtnGoToDefinition.IsEnabled = canNavigate;
    }

    private void LstMembers_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        NavigateToSelectedMemberDefinition();
    }

    private void BtnGoToDefinition_Click(object sender, RoutedEventArgs e)
    {
        NavigateToSelectedMemberDefinition();
    }

    private void NavigateToSelectedMemberDefinition()
    {
        string? targetFile = null;
        int targetLine = 1;

        if (LstMembers.SelectedItem is ObMemberItem member && !string.IsNullOrEmpty(member.FilePath))
        {
            targetFile = member.FilePath;
            targetLine = Math.Max(1, member.LineNumber);
        }
        else if (LstClasses.SelectedItem is ObClassItem cls && !string.IsNullOrEmpty(cls.FilePath))
        {
            targetFile = cls.FilePath;
            targetLine = Math.Max(1, cls.LineNumber);
        }

        if (!string.IsNullOrEmpty(targetFile) && File.Exists(targetFile))
        {
            RequestNavigateToCode?.Invoke(targetFile, targetLine);
            Close();
        }
    }

    private void RecordHistory(string libName, string className, string? memberName)
    {
        if (_isNavigatingHistory) return;
        if (string.IsNullOrEmpty(className)) return;

        // Truncar historial si estábamos en medio
        if (_historyIndex >= 0 && _historyIndex < _history.Count - 1)
        {
            _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1));
        }

        _history.Add((libName, className, memberName));
        _historyIndex = _history.Count - 1;

        UpdateHistoryButtons();
    }

    private void UpdateHistoryButtons()
    {
        BtnBack.IsEnabled = _historyIndex > 0;
        BtnForward.IsEnabled = _historyIndex >= 0 && _historyIndex < _history.Count - 1;
    }

    private void BtnBack_Click(object sender, RoutedEventArgs e)
    {
        if (_historyIndex > 0)
        {
            _historyIndex--;
            NavigateHistoryItem(_history[_historyIndex]);
        }
    }

    private void BtnForward_Click(object sender, RoutedEventArgs e)
    {
        if (_historyIndex >= 0 && _historyIndex < _history.Count - 1)
        {
            _historyIndex++;
            NavigateHistoryItem(_history[_historyIndex]);
        }
    }

    private void NavigateHistoryItem((string LibName, string ClassName, string? MemberName) entry)
    {
        _isNavigatingHistory = true;
        try
        {
            var lib = _libraries.FirstOrDefault(l => l.Name.Equals(entry.LibName, StringComparison.OrdinalIgnoreCase)) ?? _libraries[0];
            CboLibraries.SelectedItem = lib;

            var cls = lib.Classes.FirstOrDefault(c => c.Name.Equals(entry.ClassName, StringComparison.OrdinalIgnoreCase));
            if (cls != null)
            {
                LstClasses.SelectedItem = cls;
                LstClasses.ScrollIntoView(cls);

                if (!string.IsNullOrEmpty(entry.MemberName))
                {
                    var m = cls.Members.FirstOrDefault(mb => mb.Name.Equals(entry.MemberName, StringComparison.OrdinalIgnoreCase));
                    if (m != null)
                    {
                        LstMembers.SelectedItem = m;
                        LstMembers.ScrollIntoView(m);
                    }
                }
            }
        }
        finally
        {
            _isNavigatingHistory = false;
            UpdateHistoryButtons();
        }
    }
}
