using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using ICSharpCode.AvalonEdit;
using ViBa7.IDE.CodeBuilders;
using ViBa7.IDE.Designer;
using ViBa7.IDE.Dialogs;
using ViBa7.IDE.Editor;
using ViBa7.IDE.Logging;
using ViBa7.IDE.Settings;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Analysis;
using ViBa7.ProjectModel.CodeBuilders;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Documents;
using ViBa7.ProjectModel.ErrorHandling;
using ViBa7.ProjectModel.Formatting;
using ViBa7.ProjectModel.IntelliSense;
using ViBa7.IDE.Execution;
using ViBa7.ProjectModel.Execution;
using ViBa7.ProjectModel.Navigation;
using ViBa7.ProjectModel.Compilation;
using ViBa7.ProjectModel.Settings;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.IDE;

public partial class MainWindow : Window
{
    private readonly AppSettings _settings;
    private VB6Project? _currentProject;
    private VB6ProjectGroup? _currentProjectGroup;
    private LiveExecutionHost? _liveExecutionHost;
    private LiveFormWindow? _liveFormWindow;

    public EditorTabItem? ActiveTab => DocumentsTabControl.SelectedItem as EditorTabItem;
    public TextEditor? ActiveEditor => ActiveTab?.Editor;

    public MainWindow()
    {
        InitializeComponent();
        _settings = AppSettings.Load();

        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
        DocumentsTabControl.SelectionChanged += DocumentsTabControl_SelectionChanged;

        // Suscribirse al stream de log en tiempo real
        AppLogger.LogMessageReceived += OnLogMessageReceived;

        Title = $"{AppVersion.AppTitle} v{AppVersion.VersionString} - Entorno de Desarrollo ViBa7 (Compatibilidad con VB6)";
        UpdateDocumentsVisibility();
    }

    private void OnLogMessageReceived(string logEntry)
    {
        Dispatcher.Invoke(() =>
        {
            if (TxtOutputLog.Text.Length > 100000)
            {
                TxtOutputLog.Text = TxtOutputLog.Text[^50000..];
            }
            TxtOutputLog.AppendText(logEntry + Environment.NewLine);
            TxtOutputLog.ScrollToEnd();
        });
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        // Restaurar posiciÃ³n y tamaÃ±o de ventana
        if (_settings.WindowWidth > 200) Width = _settings.WindowWidth;
        if (_settings.WindowHeight > 200) Height = _settings.WindowHeight;
        if (_settings.WindowLeft >= 0) Left = _settings.WindowLeft;
        if (_settings.WindowTop >= 0) Top = _settings.WindowTop;
        if (_settings.IsMaximized) WindowState = WindowState.Maximized;

        // Restaurar estado de paneles clÃ¡sicos VB6
        MenuToggleToolbox.IsChecked = _settings.ShowToolbox;
        ColToolbox.Width = _settings.ShowToolbox ? new GridLength(_settings.ToolboxWidth) : new GridLength(0);

        MenuToggleProjectExplorer.IsChecked = _settings.ShowProjectExplorer;
        ColProjectExplorer.Width = _settings.ShowProjectExplorer ? new GridLength(_settings.ProjectExplorerWidth) : new GridLength(0);

        MenuToggleProperties.IsChecked = _settings.ShowProperties;
        RowProperties.Height = _settings.ShowProperties ? new GridLength(_settings.PropertiesHeight) : new GridLength(0);
        RowPropertiesSplitter.Height = _settings.ShowProperties ? new GridLength(4) : new GridLength(0);

        MenuToggleOutput.IsChecked = _settings.ShowOutputPanel;
        SetOutputPanelVisibility(_settings.ShowOutputPanel);

        // Inicializar integraciÃ³n de paneles de diseÃ±o visual
        InitializeDesignerPanels();

        // Restaurar estado de Funciones Avanzadas VB7
        MenuToggleLineHighlight.IsChecked = _settings.EnableCurrentLineHighlight;
        BtnToolbarToggleLineHighlight.IsChecked = _settings.EnableCurrentLineHighlight;
        MenuToggleBranchLines.IsChecked = _settings.EnableBranchLines;
        BtnToolbarToggleBranchLines.IsChecked = _settings.EnableBranchLines;
        MenuToggleEndOfLineDetails.IsChecked = _settings.EnableEndOfLineDetails;
        BtnToolbarToggleEndOfLineDetails.IsChecked = _settings.EnableEndOfLineDetails;

        UpdateExplorerModeUI();

        PreviewKeyDown += MainWindow_PreviewKeyDown;

        UpdateRecentMenus();

        AppLogger.Info("Ventana principal cargada correctamente.");
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        // Verificar si hay documentos con cambios pendientes
        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab && tab.Document.IsDirty)
            {
                var response = MessageBox.Show(
                    $"Â¿Deseas guardar los cambios en '{tab.Document.DisplayName}' antes de salir?",
                    "Cambios pendientes",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                if (response == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }

                if (response == MessageBoxResult.Yes)
                {
                    var result = tab.Save(_settings.CreateBackupBeforeSave);
                    if (result.Status != SaveStatus.Success && result.Status != SaveStatus.NotModified)
                    {
                        MessageBox.Show(result.Message, "Error al guardar", MessageBoxButton.OK, MessageBoxImage.Error);
                        e.Cancel = true;
                        return;
                    }
                }
            }
        }

        // Guardar configuraciÃ³n de ventana y paneles
        _settings.IsMaximized = WindowState == WindowState.Maximized;
        if (WindowState == WindowState.Normal)
        {
            _settings.WindowWidth = Width;
            _settings.WindowHeight = Height;
            _settings.WindowLeft = Left;
            _settings.WindowTop = Top;
        }

        _settings.ShowToolbox = MenuToggleToolbox.IsChecked;
        if (ColToolbox.Width.Value > 30)
        {
            _settings.ToolboxWidth = ColToolbox.Width.Value;
        }

        _settings.ShowProjectExplorer = MenuToggleProjectExplorer.IsChecked;
        if (ColProjectExplorer.Width.Value > 50)
        {
            _settings.ProjectExplorerWidth = ColProjectExplorer.Width.Value;
        }

        _settings.ShowProperties = MenuToggleProperties.IsChecked;
        if (RowProperties.Height.Value > 50)
        {
            _settings.PropertiesHeight = RowProperties.Height.Value;
        }

        _settings.ShowOutputPanel = MenuToggleOutput.IsChecked;
        if (RowOutput.Height.Value > 30)
        {
            _settings.OutputPanelHeight = RowOutput.Height.Value;
        }

        _settings.Save();
        AppLogger.Info("ConfiguraciÃ³n guardada.");
    }

    private void UpdateRecentMenus()
    {
        // Actualizar menÃº de proyectos recientes
        MenuRecentProjects.Items.Clear();
        if (_settings.RecentProjects.Count == 0)
        {
            MenuRecentProjects.Items.Add(new MenuItem { Header = "(VacÃ­o)", IsEnabled = false });
        }
        else
        {
            foreach (var prj in _settings.RecentProjects)
            {
                var mi = new MenuItem { Header = Path.GetFileName(prj), ToolTip = prj };
                mi.Click += (_, _) => LoadProject(prj);
                MenuRecentProjects.Items.Add(mi);
            }
        }

        // Actualizar menÃº de archivos recientes
        MenuRecentFiles.Items.Clear();
        if (_settings.RecentFiles.Count == 0)
        {
            MenuRecentFiles.Items.Add(new MenuItem { Header = "(VacÃ­o)", IsEnabled = false });
        }
        else
        {
            foreach (var file in _settings.RecentFiles)
            {
                var mi = new MenuItem { Header = Path.GetFileName(file), ToolTip = file };
                mi.Click += (_, _) => OpenFile(file);
                MenuRecentFiles.Items.Add(mi);
            }
        }
    }

    private void DocumentsTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateStatusFromActiveTab();
        SyncDesignerPanelsWithActiveTab();
    }

    private void UpdateStatusFromActiveTab()
    {
        if (ActiveTab != null)
        {
            int line = ActiveTab.Editor.TextArea.Caret.Line;
            int col = ActiveTab.Editor.TextArea.Caret.Column;
            SetLineCol(line, col);

            string encName = ActiveTab.Document.Encoding.EncodingName;
            string bomInfo = ActiveTab.Document.HasByteOrderMark ? " (BOM)" : "";
            string nl = ActiveTab.Document.NewLineConvention.ToString().ToUpperInvariant();
            SetEncodingInfo($"{encName}{bomInfo}", nl);
            SetStatusMessage($"Documento activo: {ActiveTab.Document.DisplayName}");
        }
        else
        {
            SetLineCol(1, 1);
            SetEncodingInfo("Windows-1252 (ANSI)", "CRLF");
            SetStatusMessage("Listo");
        }
    }

    private void InitializeDesignerPanels()
    {
        // Resolver de modelos visuales para UserControls del proyecto actual o grupo (.vbg)
        ControlVisualRenderer.UserControlModelResolver = (typeName) => _currentProjectGroup?.GetUserControlDesignModel(typeName) ?? _currentProject?.GetUserControlDesignModel(typeName);
        FormDesignerControl.IsUserControlContainerResolver = (typeName) => (_currentProjectGroup?.GetUserControlDesignModel(typeName) ?? _currentProject?.GetUserControlDesignModel(typeName))?.ControlContainer ?? false;
        DesignerPropertyFactory.UserControlPropertiesResolver = (typeName) => (_currentProjectGroup?.GetUserControlDesignModel(typeName) ?? _currentProject?.GetUserControlDesignModel(typeName))?.ExposedProperties;

        // Enlazar eventos de Toolbox global al diseÃ±ador activo
        MainToolboxPanel.ToolSelected += typeName =>
        {
            if (ActiveTab?.DesignSurface != null && ActiveTab.IsDesignerActive)
            {
                ActiveTab.DesignSurface.ActiveCreationTool = typeName;
            }
        };

        MainToolboxPanel.ToolDoubleClicked += typeName =>
        {
            if (ActiveTab?.DesignSurface != null && ActiveTab.IsDesignerActive)
            {
                ActiveTab.DesignSurface.CreateControlInCenter(typeName);
            }
        };

        // Enlazar eventos de PropertyGrid global
        MainPropertyGridPanel.ObjectSelected += ctrl =>
        {
            if (ActiveTab?.DesignSurface != null && ActiveTab.IsDesignerActive)
            {
                ActiveTab.DesignSurface.SelectControl(ctrl);
            }
        };

        MainPropertyGridPanel.ControlRenamed += (ctrl, oldName, newName) =>
        {
            ActiveTab?.RenameControlInCode(oldName, newName);
        };

        MainPropertyGridPanel.ModelChanged += () =>
        {
            if (ActiveTab?.DesignSurface != null && ActiveTab.DesignModel != null)
            {
                ActiveTab.Document.Header = FormDesignSerializer.Serialize(ActiveTab.DesignModel);
                ActiveTab.Document.IsDirty = true;
                ActiveTab.DesignSurface.RefreshVisuals();
            }
        };

        MainPropertyGridPanel.VisualResolver = ctrl => ActiveTab?.DesignSurface?.GetVisualForControl(ctrl);
    }

    private void SyncDesignerPanelsWithActiveTab()
    {
        if (ActiveTab?.DesignSurface != null && ActiveTab.DesignModel != null)
        {
            var surface = ActiveTab.DesignSurface;
            var model = ActiveTab.DesignModel;

            MainPropertyGridPanel.BindModel(model, surface.SelectedControl, ActiveTab?.Document?.FilePath);

            // Suscribir eventos del diseÃ±ador de esta pestaÃ±a
            surface.SelectionChanged += selList =>
            {
                if (ActiveTab?.DesignSurface == surface)
                {
                    if (selList.Count == 0)
                        MainPropertyGridPanel.SelectForm();
                    else
                        MainPropertyGridPanel.SelectControl(selList.LastOrDefault());
                }
            };

            surface.ControlAdded += ctrl =>
            {
                if (ActiveTab?.DesignSurface == surface)
                {
                    MainPropertyGridPanel.RefreshObjectList();
                    MainPropertyGridPanel.SelectControl(ctrl);
                }
            };

            surface.ControlDeleted += ctrl =>
            {
                if (ActiveTab?.DesignSurface == surface)
                {
                    MainPropertyGridPanel.RefreshObjectList();
                    MainPropertyGridPanel.SelectForm();
                }
            };

            surface.ModelModified += () =>
            {
                if (ActiveTab?.DesignSurface == surface)
                {
                    MainPropertyGridPanel.RefreshCurrentProperties();
                }
            };

            surface.ToolCreationFinished += () =>
            {
                MainToolboxPanel.ResetToPointer();
            };
        }
        else
        {
            MainPropertyGridPanel.ClearModel();
            MainToolboxPanel.ResetToPointer();
        }

        // Sincronizar estado de habilitación de herramientas en la Caja de Herramientas
        MainToolboxPanel.EnableAllTools();
        if (ActiveTab?.Document?.FilePath?.EndsWith(".ctl", StringComparison.OrdinalIgnoreCase) == true)
        {
            string activeUcName = Path.GetFileNameWithoutExtension(ActiveTab.Document.FilePath);
            MainToolboxPanel.SetToolEnabled(activeUcName, false);
        }
    }

    private void UpdateDocumentsVisibility()
    {
        if (DocumentsTabControl.Items.Count == 0)
        {
            EmptyDocsNotice.Visibility = Visibility.Visible;
            DocumentsTabControl.Visibility = Visibility.Collapsed;
        }
        else
        {
            EmptyDocsNotice.Visibility = Visibility.Collapsed;
            DocumentsTabControl.Visibility = Visibility.Visible;
        }
    }

    public void SetStatusMessage(string message) => TxtStatusMessage.Text = message;
    public void SetLineCol(int line, int col) => TxtStatusLineCol.Text = $"Ln {line}, Col {col}";
    public void SetEncodingInfo(string encoding, string lineEnding) => TxtStatusEncoding.Text = $"{encoding} | {lineEnding}";

    public EditorTabItem? OpenFile(string filePath)
    {
        try
        {
            string fullPath = Path.GetFullPath(filePath);

            // Evitar abrir duplicados para la misma ruta canÃ³nica
            foreach (var item in DocumentsTabControl.Items)
            {
                if (item is EditorTabItem openTab &&
                    string.Equals(Path.GetFullPath(openTab.Document.FilePath), fullPath, StringComparison.OrdinalIgnoreCase))
                {
                    DocumentsTabControl.SelectedItem = openTab;
                    UpdateStatusFromActiveTab();
                    return openTab;
                }
            }

            var doc = SafeFileStorage.LoadDocument(fullPath);
            var tabItem = new EditorTabItem(doc, _settings);

            tabItem.CaretPositionChanged += (line, col) =>
            {
                if (DocumentsTabControl.SelectedItem == tabItem)
                {
                    SetLineCol(line, col);
                }
            };

            tabItem.ViewModeChanged += isDesigner =>
            {
                SyncDesignerPanelsWithActiveTab();
            };

            DocumentsTabControl.Items.Add(tabItem);
            DocumentsTabControl.SelectedItem = tabItem;

            _settings.AddRecentFile(fullPath);
            UpdateRecentMenus();

            UpdateDocumentsVisibility();
            UpdateStatusFromActiveTab();
            AppLogger.Info($"Archivo abierto: {fullPath}");
            return tabItem;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Error al abrir el archivo {filePath}", ex);
            MessageBox.Show($"No se pudo abrir el archivo:\n\n{ex.Message}", "Error al abrir archivo", MessageBoxButton.OK, MessageBoxImage.Error);
            return null;
        }
    }

    public bool CloseTab(EditorTabItem tab)
    {
        if (tab.Document.IsDirty)
        {
            var response = MessageBox.Show(
                $"Â¿Deseas guardar los cambios en '{tab.Document.DisplayName}' antes de cerrar?",
                "Cambios pendientes",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (response == MessageBoxResult.Cancel) return false;

            if (response == MessageBoxResult.Yes)
            {
                var result = tab.Save(_settings.CreateBackupBeforeSave);
                if (result.Status != SaveStatus.Success && result.Status != SaveStatus.NotModified)
                {
                    MessageBox.Show(result.Message, "Error al guardar", MessageBoxButton.OK, MessageBoxImage.Error);
                    return false;
                }
            }
        }

        int index = DocumentsTabControl.Items.IndexOf(tab);
        DocumentsTabControl.Items.Remove(tab);

        if (DocumentsTabControl.Items.Count > 0)
        {
            int nextIndex = Math.Min(index, DocumentsTabControl.Items.Count - 1);
            DocumentsTabControl.SelectedIndex = Math.Max(0, nextIndex);
        }
        else
        {
            DocumentsTabControl.SelectedItem = null;
        }

        UpdateDocumentsVisibility();
        UpdateStatusFromActiveTab();
        SyncDesignerPanelsWithActiveTab();
        return true;
    }

    public bool CloseAllDocuments()
    {
        var tabs = DocumentsTabControl.Items.OfType<EditorTabItem>().ToList();
        foreach (var tab in tabs)
        {
            if (!CloseTab(tab))
            {
                return false; // El usuario cancelÃ³
            }
        }
        return true;
    }

    public void CloseProjectAndAllDocuments()
    {
        if (!CloseAllDocuments()) return;

        _currentProject = null;
        _currentProjectGroup = null;

        if (MenuSaveProjectGroup != null) MenuSaveProjectGroup.Visibility = Visibility.Collapsed;
        if (MenuSetStartupProject != null) MenuSetStartupProject.Visibility = Visibility.Collapsed;
        if (SepGroupOptions != null) SepGroupOptions.Visibility = Visibility.Collapsed;

        ProjectTreeView.Items.Clear();

        MainPropertyGridPanel.ClearModel();
        ControlToolboxRegistry.UnregisterUserControlTools();
        MainToolboxPanel.ReloadTools();
        MainToolboxPanel.ResetToPointer();

        Title = $"{AppVersion.AppTitle} v{AppVersion.VersionString} - Entorno de Desarrollo ViBa7 (Compatibilidad con VB6)";
        SetStatusMessage("Proyecto y documentos cerrados.");
        AppLogger.Info("Proyecto y todos los documentos cerrados.");
    }

    #region Project Handling (Milestone 5 & VBG Groups)
    public void LoadProject(string filePath)
    {
        string ext = Path.GetExtension(filePath);
        if (string.Equals(ext, ".vbg", StringComparison.OrdinalIgnoreCase))
        {
            LoadProjectGroup(filePath);
        }
        else
        {
            LoadSingleProject(filePath);
        }
    }

    public void LoadProjectGroup(string vbgFilePath)
    {
        try
        {
            var group = VbgParser.Parse(vbgFilePath);
            _currentProjectGroup = group;
            _currentProject = group.PrimaryProject ?? group.Projects.FirstOrDefault();

            foreach (var p in group.Projects)
            {
                LoadProjectActiveXComponents(p);
            }
            ControlToolboxRegistry.RegisterUserControlTools(group);
            MainToolboxPanel.ReloadTools();

            if (MenuSaveProjectGroup != null) MenuSaveProjectGroup.Visibility = Visibility.Visible;
            if (MenuSetStartupProject != null) MenuSetStartupProject.Visibility = Visibility.Visible;
            if (SepGroupOptions != null) SepGroupOptions.Visibility = Visibility.Visible;

            RefreshProjectTree();
            string primaryName = group.PrimaryProject?.Name ?? "Sin definir";
            int secCount = group.SecondaryProjects.Count;
            SetStatusMessage($"Grupo de proyectos '{group.Name}' cargado: Principal '{primaryName}' (Inicio), {secCount} secundario(s) (subprocesos).");
            Title = $"{AppVersion.AppTitle} v{AppVersion.VersionString} - [Grupo: {group.Name}] Principal: {primaryName} ({Path.GetFileName(group.GroupFilePath)})";

            _settings.AddRecentProject(group.GroupFilePath);
            UpdateRecentMenus();

            AppLogger.Info($"Grupo de proyectos cargado: {group.GroupFilePath}");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Error al abrir el grupo de proyectos {vbgFilePath}", ex);
            MessageBox.Show($"Error al abrir el grupo de proyectos:\n\n{ex.Message}", "Error al abrir grupo de proyectos", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void LoadSingleProject(string vbpFilePath)
    {
        try
        {
            var project = VbpParser.Parse(vbpFilePath);
            _currentProject = project;
            _currentProjectGroup = null;

            if (MenuSaveProjectGroup != null) MenuSaveProjectGroup.Visibility = Visibility.Collapsed;
            if (MenuSetStartupProject != null) MenuSetStartupProject.Visibility = Visibility.Collapsed;
            if (SepGroupOptions != null) SepGroupOptions.Visibility = Visibility.Collapsed;

            LoadProjectActiveXComponents(project);
            ControlToolboxRegistry.RegisterUserControlTools(project);
            MainToolboxPanel.ReloadTools();

            RefreshProjectTree();
            SetStatusMessage($"Proyecto '{project.Name}' cargado exitosamente.");
            Title = $"{AppVersion.AppTitle} v{AppVersion.VersionString} - {project.Name} ({Path.GetFileName(project.ProjectFilePath)})";

            _settings.AddRecentProject(project.ProjectFilePath);
            UpdateRecentMenus();

            AppLogger.Info($"Proyecto cargado: {project.ProjectFilePath}");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Error al abrir el proyecto {vbpFilePath}", ex);
            MessageBox.Show($"Error al abrir el proyecto:\n\n{ex.Message}", "Error al abrir proyecto", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void RefreshProjectTree()
    {
        string filter = TxtProjectSearch?.Text.Trim() ?? "";

        if (_currentProjectGroup != null)
        {
            ProjectTreeView.Items.Clear();
            var groupRoot = new TreeViewItem
            {
                Header = $"?? Grupo de Proyectos: {_currentProjectGroup.Name} ({Path.GetFileName(_currentProjectGroup.GroupFilePath)})",
                IsExpanded = true,
                FontWeight = FontWeights.Bold
            };

            var groupMenu = new ContextMenu();
            var mnuSaveGroup = new MenuItem
            {
                Header = "Guardar grupo de proyectos (.vbg)"
            };
            mnuSaveGroup.Click += (s, e) => MenuSaveProjectGroup_Click(s, e);
            groupMenu.Items.Add(mnuSaveGroup);
            groupRoot.ContextMenu = groupMenu;

            foreach (var project in _currentProjectGroup.Projects)
            {
                bool isStartup = _currentProjectGroup.IsPrimaryProject(project);
                var prjNode = CreateProjectNode(project, isStartup, filter, _settings.ProjectExplorerMode);
                if (prjNode != null)
                {
                    groupRoot.Items.Add(prjNode);
                }
            }

            if (groupRoot.Items.Count > 0)
            {
                ProjectTreeView.Items.Add(groupRoot);
            }
            else
            {
                ProjectTreeView.Items.Add(new TreeViewItem
                {
                    Header = "No se encontraron coincidencias para la bÃºsqueda",
                    Foreground = Brushes.Gray,
                    FontStyle = FontStyles.Italic
                });
            }
        }
        else if (_currentProject != null)
        {
            ProjectTreeView.Items.Clear();
            var prjNode = CreateProjectNode(_currentProject, false, filter, _settings.ProjectExplorerMode);
            if (prjNode != null)
            {
                ProjectTreeView.Items.Add(prjNode);
            }
            else
            {
                ProjectTreeView.Items.Add(new TreeViewItem
                {
                    Header = "No se encontraron coincidencias para la bÃºsqueda",
                    Foreground = Brushes.Gray,
                    FontStyle = FontStyles.Italic
                });
            }
        }
        else
        {
            ProjectTreeView.Items.Clear();
            ProjectTreeView.Items.Add(new TreeViewItem
            {
                Header = "NingÃºn proyecto abierto",
                IsSelected = true,
                Foreground = Brushes.Gray,
                FontStyle = FontStyles.Italic
            });
        }

        UpdateMakeExeMenuState();
    }

    private TreeViewItem? CreateProjectNode(VB6Project project, bool isStartup, string filter, ProjectExplorerMode mode)
    {
        bool hasFilter = !string.IsNullOrEmpty(filter);
        string fileName = Path.GetFileName(project.ProjectFilePath);
        string title;
        string toolTipText;

        if (_currentProjectGroup != null)
        {
            if (isStartup)
            {
                title = $"?? [Principal] {project.Name} ({fileName}) [Inicio / Proceso]";
                toolTipText = $"Proyecto Principal (Inicio de ejecuciÃ³n)\nArchivo: {project.ProjectFilePath}\nTipo: {project.ProjectType}\nSe ejecuta como proceso principal del grupo.";
            }
            else
            {
                title = $"?? [Secundario] {project.Name} ({fileName}) [Subproceso - {project.ProjectType}]";
                toolTipText = $"Proyecto Secundario (Subproceso)\nArchivo: {project.ProjectFilePath}\nTipo: {project.ProjectType}\nSe ejecuta como subproceso / componente dependiente del proyecto principal.";
            }
        }
        else
        {
            title = $"?? {project.Name} ({fileName})";
            toolTipText = $"Proyecto: {project.Name}\nArchivo: {project.ProjectFilePath}\nTipo: {project.ProjectType}";
        }

        var rootItem = new TreeViewItem
        {
            Header = title,
            IsExpanded = true,
            FontWeight = isStartup ? FontWeights.Bold : FontWeights.SemiBold,
            ToolTip = toolTipText,
            Tag = project
        };

        var contextMenu = new ContextMenu();
        if (_currentProjectGroup != null)
        {
            var mnuSetStartup = new MenuItem
            {
                Header = isStartup ? "? Proyecto Principal (Inicio)" : "Establecer como proyecto principal (Inicio)",
                IsEnabled = !isStartup,
                FontWeight = isStartup ? FontWeights.Bold : FontWeights.Normal
            };
            mnuSetStartup.Click += (s, e) =>
            {
                _currentProjectGroup.SetPrimaryProject(project);
                _currentProject = project;
                RefreshProjectTree();
                UpdateMakeExeMenuState();
                SetStatusMessage($"'{project.Name}' establecido como Proyecto Principal (Inicio).");
            };
            contextMenu.Items.Add(mnuSetStartup);
            contextMenu.Items.Add(new Separator());
        }

        var mnuExplore = new MenuItem
        {
            Header = "Abrir carpeta en el explorador"
        };
        mnuExplore.Click += (s, e) =>
        {
            string dir = project.ProjectDirectory;
            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = dir,
                        UseShellExecute = true
                    });
                }
                catch { }
            }
        };
        contextMenu.Items.Add(mnuExplore);
        rootItem.ContextMenu = contextMenu;

        int totalMatches = 0;

        void AddCategory(string headerTitle, IEnumerable<ProjectItem> items)
        {
            var itemList = items.ToList();
            if (itemList.Count == 0) return;

            var folderItem = new TreeViewItem
            {
                Header = $"{headerTitle} ({itemList.Count})",
                IsExpanded = hasFilter,
                FontWeight = FontWeights.SemiBold
            };

            int folderMatches = 0;
            foreach (var item in itemList)
            {
                var itemNode = CreateProjectItemNode(item, filter, mode);
                if (itemNode != null)
                {
                    folderItem.Items.Add(itemNode);
                    folderMatches++;
                }
            }

            if (folderMatches > 0)
            {
                totalMatches += folderMatches;
                rootItem.Items.Add(folderItem);
            }
        }

        AddCategory("?? Formularios", project.Forms);
        AddCategory("?? MÃ³dulos", project.Modules);
        AddCategory("?? MÃ³dulos de Clase", project.Classes);
        AddCategory("??? Controles de Usuario", project.UserControls);

        // Referencias
        if (project.References.Any())
        {
            var matchingRefs = project.References
                .Where(r => !hasFilter || r.DisplayTitle.Contains(filter, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (matchingRefs.Count > 0)
            {
                var refFolder = new TreeViewItem
                {
                    Header = $"?? Referencias ({project.References.Count})",
                    IsExpanded = false,
                    FontWeight = FontWeights.Normal
                };

                foreach (var r in matchingRefs)
                {
                    refFolder.Items.Add(new TreeViewItem
                    {
                        Header = $"â€¢ {r.DisplayTitle}",
                        Foreground = Brushes.Gray,
                        FontWeight = FontWeights.Normal
                    });
                }
                rootItem.Items.Add(refFolder);
                totalMatches += matchingRefs.Count;
            }
        }

        if (hasFilter && totalMatches == 0 && !project.Name.Contains(filter, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return rootItem;
    }

    private TreeViewItem? CreateProjectItemNode(ProjectItem item, string filter, ProjectExplorerMode mode)
    {
        bool hasFilter = !string.IsNullOrEmpty(filter);
        bool fileMatches = !hasFilter || item.DisplayTitle.Contains(filter, StringComparison.OrdinalIgnoreCase);

        var node = new TreeViewItem
        {
            Header = $"{item.DisplayIcon} {item.DisplayTitle}" + (!item.FileExists ? " (No encontrado)" : ""),
            Tag = item,
            FontWeight = FontWeights.Normal,
            IsExpanded = hasFilter
        };

        if (!item.FileExists)
        {
            node.Foreground = Brushes.Crimson;
            node.ToolTip = $"El archivo no fue encontrado en: {item.AbsolutePath}";
            return fileMatches ? node : null;
        }

        node.ToolTip = item.AbsolutePath;

        // Si estamos en modo Avanzado y es un archivo de cÃ³digo editable, parsear miembros
        if (mode == ProjectExplorerMode.Advanced && item.IsEditable)
        {
            List<CodeMemberSymbol> members = [];
            try
            {
                string text = "";
                // Si estÃ¡ abierto en un tab, usar el texto en vivo
                foreach (var tabObj in DocumentsTabControl.Items)
                {
                    if (tabObj is EditorTabItem openTab &&
                        string.Equals(Path.GetFullPath(openTab.Document.FilePath), Path.GetFullPath(item.AbsolutePath), StringComparison.OrdinalIgnoreCase))
                    {
                        text = openTab.Editor.Text;
                        break;
                    }
                }

                if (string.IsNullOrEmpty(text))
                {
                    text = SafeFileStorage.LoadDocument(item.AbsolutePath).Text;
                }

                members = CodeStructureParser.ParseMembers(text, item.AbsolutePath);
            }
            catch
            {
                members = [];
            }

            int matchingMembers = 0;
            foreach (var member in members)
            {
                bool memberMatches = !hasFilter ||
                    member.Name.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                    member.DisplayText.Contains(filter, StringComparison.OrdinalIgnoreCase);

                if (memberMatches)
                {
                    matchingMembers++;
                    var memberNode = new TreeViewItem
                    {
                        Header = member.DisplayText,
                        Tag = member,
                        FontWeight = FontWeights.Normal,
                        ToolTip = $"{member.DisplayText}\nLÃ­nea {member.LineNumber}\n{member.SourceFilePath}"
                    };
                    node.Items.Add(memberNode);
                }
            }

            if (hasFilter)
            {
                if (!fileMatches && matchingMembers == 0)
                {
                    return null;
                }
                node.IsExpanded = true;
            }
        }
        else
        {
            if (hasFilter && !fileMatches)
            {
                return null;
            }
        }

        return node;
    }

    private void ProjectTreeView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (ProjectTreeView.SelectedItem is TreeViewItem selItem)
        {
            if (selItem.Tag is CodeMemberSymbol sym)
            {
                var tab = OpenFile(sym.SourceFilePath);
                if (tab != null && sym.LineNumber > 0 && sym.LineNumber <= tab.Editor.Document.LineCount)
                {
                    var line = tab.Editor.Document.GetLineByNumber(sym.LineNumber);
                    tab.Editor.CaretOffset = line.Offset;
                    tab.Editor.ScrollToLine(sym.LineNumber);
                    tab.Editor.Focus();
                }
                e.Handled = true;
                return;
            }

            if (selItem.Tag is ProjectItem item)
            {
                if (!item.FileExists)
                {
                    MessageBox.Show($"El archivo '{item.DisplayTitle}' no existe en la ruta esperada:\n\n{item.AbsolutePath}",
                        "Archivo no encontrado", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (item.IsEditable)
                {
                    OpenFile(item.AbsolutePath);
                }
                e.Handled = true;
            }
        }
    }

    private void ProjectTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (e.NewValue is TreeViewItem item)
        {
            if (item.Tag is VB6Project prj)
            {
                _currentProject = prj;
                UpdateMakeExeMenuState();
                if (_currentProjectGroup != null)
                {
                    SetStatusMessage($"{prj.RoleDescription}: '{prj.Name}' ({Path.GetFileName(prj.ProjectFilePath)})");
                }
                else
                {
                    SetStatusMessage($"Proyecto: '{prj.Name}'");
                }
            }
            else if (item.Tag is ProjectItem pItem && pItem.ContainingProject != null)
            {
                _currentProject = pItem.ContainingProject;
                UpdateMakeExeMenuState();
            }
        }
    }

    private void UpdateExplorerModeUI()
    {
        bool isAdvanced = _settings.ProjectExplorerMode == ProjectExplorerMode.Advanced;
        MenuExplorerMode.IsChecked = isAdvanced;
        MenuExplorerMode.Header = isAdvanced ? "Modo de Explorador: _Avanzado (activo)" : "Modo de Explorador: _EstÃ¡ndar (activo)";
        BtnToggleExplorerMode.Opacity = isAdvanced ? 1.0 : 0.6;
        BtnToggleExplorerMode.ToolTip = isAdvanced
            ? "Modo de Explorador: Avanzado (clic para cambiar a EstÃ¡ndar)"
            : "Modo de Explorador: EstÃ¡ndar (clic para cambiar a Avanzado)";
    }

    private void BtnToggleExplorerMode_Click(object sender, RoutedEventArgs e)
    {
        _settings.ProjectExplorerMode = _settings.ProjectExplorerMode == ProjectExplorerMode.Advanced
            ? ProjectExplorerMode.Standard
            : ProjectExplorerMode.Advanced;
        _settings.Save();
        UpdateExplorerModeUI();
        RefreshProjectTree();
        SetStatusMessage($"Modo de Explorador cambiado a: {(_settings.ProjectExplorerMode == ProjectExplorerMode.Advanced ? "Avanzado" : "EstÃ¡ndar")}.");
    }

    private void MenuExplorerMode_Click(object sender, RoutedEventArgs e)
    {
        BtnToggleExplorerMode_Click(sender, e);
    }

    private void TxtProjectSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        string filter = TxtProjectSearch.Text.Trim();
        BtnClearSearch.Visibility = string.IsNullOrEmpty(filter) ? Visibility.Collapsed : Visibility.Visible;
        RefreshProjectTree();
    }

    private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
    {
        TxtProjectSearch.Text = string.Empty;
    }
    #endregion

    #region Menu File Actions
    private void MenuNewProject_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dlg = new Dialogs.NewProjectDialog();
            dlg.Owner = this;
            if (dlg.ShowDialog() != true) return;

            string projectDir = dlg.ProjectLocation;
            string projectName = dlg.ProjectName;
            string vbpPath = dlg.ProjectFilePath;
            var dialect = dlg.SelectedDialect;
            var projType = dlg.SelectedProjectType;

            if (!Directory.Exists(projectDir))
            {
                Directory.CreateDirectory(projectDir);
            }

            string startupObject = "Form1";
            string initialItemPath = Path.Combine(projectDir, "Form1.frm");
            string itemEntry = "Form=Form1.frm";
            string typeStr = "Exe";

            if (projType == Dialogs.NewProjectType.ActiveXControl)
            {
                typeStr = "Control";
                startupObject = "(None)";
                initialItemPath = Path.Combine(projectDir, "UserControl1.ctl");
                itemEntry = "UserControl=UserControl1.ctl";

                string ctlContent =
                    "VERSION 5.00\r\n" +
                    "Begin VB.UserControl UserControl1 \r\n" +
                    "   ClientHeight    =   3600\r\n" +
                    "   ClientLeft      =   0\r\n" +
                    "   ClientTop       =   0\r\n" +
                    "   ClientWidth     =   4800\r\n" +
                    "   ScaleHeight     =   3600\r\n" +
                    "   ScaleWidth      =   4800\r\n" +
                    "End\r\n" +
                    "Attribute VB_Name = \"UserControl1\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = True\r\n" +
                    "Attribute VB_PredeclaredId = False\r\n" +
                    "Attribute VB_Exposed = True\r\n" +
                    "\r\n" +
                    "Private Sub UserControl_Initialize()\r\n" +
                    "    \r\n" +
                    "End Sub\r\n";
                File.WriteAllText(initialItemPath, ctlContent, Encoding.GetEncoding(1252));
            }
            else if (projType == Dialogs.NewProjectType.ActiveXDll)
            {
                typeStr = "OleDll";
                startupObject = "(None)";
                initialItemPath = Path.Combine(projectDir, "Class1.cls");
                itemEntry = "Class=Class1; Class1.cls";

                string clsContent =
                    "VERSION 1.0 CLASS\r\n" +
                    "BEGIN\r\n" +
                    "  MultiUse = -1  'True\r\n" +
                    "  Persistable = 0  'NotPersistable\r\n" +
                    "  DataBindingBehavior = 0  'vbNone\r\n" +
                    "  DataSourceBehavior  = 0  'vbNone\r\n" +
                    "  MTSTransactionMode  = 0  'NotAnMTSObject\r\n" +
                    "END\r\n" +
                    "Attribute VB_Name = \"Class1\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = True\r\n" +
                    "Attribute VB_PredeclaredId = False\r\n" +
                    "Attribute VB_Exposed = True\r\n" +
                    "Option Explicit\r\n" +
                    "\r\n" +
                    "Private Sub Class_Initialize()\r\n" +
                    "    \r\n" +
                    "End Sub\r\n";
                File.WriteAllText(initialItemPath, clsContent, Encoding.GetEncoding(1252));
            }
            else
            {
                // Standard EXE
                string formContent =
                    "VERSION 5.00\r\n" +
                    "Begin VB.Form Form1 \r\n" +
                    "   Caption         =   \"Form1\"\r\n" +
                    "   ClientHeight    =   4500\r\n" +
                    "   ClientLeft      =   60\r\n" +
                    "   ClientTop       =   450\r\n" +
                    "   ClientWidth     =   6800\r\n" +
                    "   LinkTopic       =   \"Form1\"\r\n" +
                    "   ScaleHeight     =   4500\r\n" +
                    "   ScaleWidth      =   6800\r\n" +
                    "   StartUpPosition =   3  'Windows Default\r\n" +
                    "End\r\n" +
                    "Attribute VB_Name = \"Form1\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = False\r\n" +
                    "Attribute VB_PredeclaredId = True\r\n" +
                    "Attribute VB_Exposed = False\r\n" +
                    "\r\n" +
                    "Private Sub Form_Load()\r\n" +
                    "    \r\n" +
                    "End Sub\r\n";
                File.WriteAllText(initialItemPath, formContent, Encoding.GetEncoding(1252));
            }

            string dialectStr = dialect switch
            {
                VbLanguageDialect.Vb7Office => "Vb7Office",
                VbLanguageDialect.ViBaModern => "ViBaModern",
                _ => "VB6"
            };

            string vbpContent =
                $"Type={typeStr}\r\n" +
                $"{itemEntry}\r\n" +
                $"Startup=\"{startupObject}\"\r\n" +
                $"Name=\"{projectName}\"\r\n" +
                $"Title=\"{projectName}\"\r\n" +
                $"ExeName32=\"{projectName}.exe\"\r\n" +
                $"LanguageDialect=\"{dialectStr}\"\r\n" +
                "Command32=\"\"\r\n" +
                "CompatibleMode=\"0\"\r\n" +
                "MajorVer=1\r\n" +
                "MinorVer=0\r\n" +
                "RevisionVer=0\r\n" +
                "AutoIncrementVer=0\r\n" +
                "ServerSupportFiles=0\r\n" +
                "CompilationType=0\r\n" +
                "OptimizationType=0\r\n" +
                "FavorPentiumPro(tm)=0\r\n" +
                "CodeViewDebugInfo=0\r\n" +
                "NoAliasing=0\r\n" +
                "BoundsCheck=0\r\n" +
                "OverflowCheck=0\r\n" +
                "FlPointCheck=0\r\n" +
                "FDIVCheck=0\r\n" +
                "UnroundedFP=0\r\n" +
                "StartMode=0\r\n" +
                "Unattended=0\r\n" +
                "Retained=0\r\n" +
                "ThreadPerObject=0\r\n" +
                "MaxNumberOfThreads=1\r\n";

            File.WriteAllText(vbpPath, vbpContent, Encoding.GetEncoding(1252));

            LoadProject(vbpPath);
            OpenFile(initialItemPath);
            SetStatusMessage($"Nuevo proyecto creado: {projectName} [{dialect}] ({vbpPath})");
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al crear nuevo proyecto", ex);
            MessageBox.Show($"Error al crear el nuevo proyecto:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuNewForm_Click(object? sender, RoutedEventArgs? e)
    {
        try
        {
            if (_currentProject != null)
            {
                int idx = 1;
                string formName = $"Form{idx}";
                while (_currentProject.Forms.Any(f => string.Equals(Path.GetFileNameWithoutExtension(f.AbsolutePath), formName, StringComparison.OrdinalIgnoreCase)) ||
                       File.Exists(Path.Combine(_currentProject.ProjectDirectory, $"{formName}.frm")))
                {
                    idx++;
                    formName = $"Form{idx}";
                }

                string formPath = Path.Combine(_currentProject.ProjectDirectory, $"{formName}.frm");
                string formContent =
                    "VERSION 5.00\r\n" +
                    $"Begin VB.Form {formName} \r\n" +
                    $"   Caption         =   \"{formName}\"\r\n" +
                    "   ClientHeight    =   4500\r\n" +
                    "   ClientLeft      =   60\r\n" +
                    "   ClientTop       =   450\r\n" +
                    "   ClientWidth     =   6800\r\n" +
                    $"   LinkTopic       =   \"{formName}\"\r\n" +
                    "   ScaleHeight     =   4500\r\n" +
                    "   ScaleWidth      =   6800\r\n" +
                    "   StartUpPosition =   3  'Windows Default\r\n" +
                    "End\r\n" +
                    $"Attribute VB_Name = \"{formName}\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = False\r\n" +
                    "Attribute VB_PredeclaredId = True\r\n" +
                    "Attribute VB_Exposed = False\r\n" +
                    "\r\n" +
                    "Private Sub Form_Load()\r\n" +
                    "    \r\n" +
                    "End Sub\r\n";

                File.WriteAllText(formPath, formContent, Encoding.GetEncoding(1252));

                string vbpText = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
                string newEntry = $"Form={formName}.frm\r\n";
                int lastFormIdx = vbpText.LastIndexOf("Form=", StringComparison.OrdinalIgnoreCase);
                if (lastFormIdx >= 0)
                {
                    int endOfLine = vbpText.IndexOf("\r\n", lastFormIdx);
                    if (endOfLine >= 0)
                        vbpText = vbpText.Insert(endOfLine + 2, newEntry);
                    else
                        vbpText += "\r\n" + newEntry;
                }
                else
                {
                    vbpText = newEntry + vbpText;
                }
                File.WriteAllText(_currentProject.ProjectFilePath, vbpText, Encoding.GetEncoding(1252));

                LoadProject(_currentProject.ProjectFilePath);
                OpenFile(formPath);
                SetStatusMessage($"Formulario '{formName}.frm' agregado al proyecto.");
            }
            else
            {
                var saveDlg = new SaveFileDialog
                {
                    Title = "Guardar Nuevo Formulario",
                    Filter = "Formulario ViBa7 (Compatibilidad con VB6) (*.frm)|*.frm",
                    FileName = "Form1.frm",
                    DefaultExt = ".frm"
                };

                if (saveDlg.ShowDialog() != true) return;

                string formPath = saveDlg.FileName;
                string formName = Path.GetFileNameWithoutExtension(formPath);
                string formContent =
                    "VERSION 5.00\r\n" +
                    $"Begin VB.Form {formName} \r\n" +
                    $"   Caption         =   \"{formName}\"\r\n" +
                    "   ClientHeight    =   4500\r\n" +
                    "   ClientLeft      =   60\r\n" +
                    "   ClientTop       =   450\r\n" +
                    "   ClientWidth     =   6800\r\n" +
                    $"   LinkTopic       =   \"{formName}\"\r\n" +
                    "   ScaleHeight     =   4500\r\n" +
                    "   ScaleWidth      =   6800\r\n" +
                    "   StartUpPosition =   3  'Windows Default\r\n" +
                    "End\r\n" +
                    $"Attribute VB_Name = \"{formName}\"\r\n" +
                    "Attribute VB_GlobalNameSpace = False\r\n" +
                    "Attribute VB_Creatable = False\r\n" +
                    "Attribute VB_PredeclaredId = True\r\n" +
                    "Attribute VB_Exposed = False\r\n" +
                    "\r\n" +
                    "Private Sub Form_Load()\r\n" +
                    "    \r\n" +
                    "End Sub\r\n";

                File.WriteAllText(formPath, formContent, Encoding.GetEncoding(1252));
                OpenFile(formPath);
                SetStatusMessage($"Nuevo formulario creado: {formName}.frm");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al agregar nuevo formulario", ex);
            MessageBox.Show($"Error al crear el formulario:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuAddModule_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentProject == null)
            {
                MessageBox.Show("Debes tener un proyecto abierto para agregar un mÃ³dulo.", "InformaciÃ³n", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int idx = 1;
            string modName = $"Module{idx}";
            while (_currentProject.Modules.Any(m => string.Equals(Path.GetFileNameWithoutExtension(m.AbsolutePath), modName, StringComparison.OrdinalIgnoreCase)) ||
                   File.Exists(Path.Combine(_currentProject.ProjectDirectory, $"{modName}.bas")))
            {
                idx++;
                modName = $"Module{idx}";
            }

            string modPath = Path.Combine(_currentProject.ProjectDirectory, $"{modName}.bas");
            string modContent =
                $"Attribute VB_Name = \"{modName}\"\r\n" +
                "Option Explicit\r\n\r\n";

            File.WriteAllText(modPath, modContent, Encoding.GetEncoding(1252));

            string vbpText = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
            string newEntry = $"Module={modName}; {modName}.bas\r\n";
            vbpText += "\r\n" + newEntry;
            File.WriteAllText(_currentProject.ProjectFilePath, vbpText, Encoding.GetEncoding(1252));

            LoadProject(_currentProject.ProjectFilePath);
            OpenFile(modPath);
            SetStatusMessage($"MÃ³dulo '{modName}.bas' agregado al proyecto.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al agregar mÃ³dulo", ex);
            MessageBox.Show($"Error al crear el mÃ³dulo:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuAddClass_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentProject == null)
            {
                MessageBox.Show("Debes tener un proyecto abierto para agregar un mÃ³dulo de clase.", "InformaciÃ³n", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int idx = 1;
            string clsName = $"Class{idx}";
            while (_currentProject.Classes.Any(c => string.Equals(Path.GetFileNameWithoutExtension(c.AbsolutePath), clsName, StringComparison.OrdinalIgnoreCase)) ||
                   File.Exists(Path.Combine(_currentProject.ProjectDirectory, $"{clsName}.cls")))
            {
                idx++;
                clsName = $"Class{idx}";
            }

            string clsPath = Path.Combine(_currentProject.ProjectDirectory, $"{clsName}.cls");
            string clsContent =
                "VERSION 1.0 CLASS\r\n" +
                "BEGIN\r\n" +
                "  MultiUse = -1  'True\r\n" +
                "  Persistable = 0  'NotPersistable\r\n" +
                "  DataBindingBehavior = 0  'vbNone\r\n" +
                "  DataSourceBehavior  = 0  'vbNone\r\n" +
                "  MTSTransactionMode  = 0  'NotAnMTSObject\r\n" +
                "END\r\n" +
                $"Attribute VB_Name = \"{clsName}\"\r\n" +
                "Attribute VB_GlobalNameSpace = False\r\n" +
                "Attribute VB_Creatable = True\r\n" +
                "Attribute VB_PredeclaredId = False\r\n" +
                "Attribute VB_Exposed = False\r\n" +
                "Option Explicit\r\n\r\n";

            File.WriteAllText(clsPath, clsContent, Encoding.GetEncoding(1252));

            string vbpText = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
            string newEntry = $"Class={clsName}; {clsName}.cls\r\n";
            vbpText += "\r\n" + newEntry;
            File.WriteAllText(_currentProject.ProjectFilePath, vbpText, Encoding.GetEncoding(1252));

            LoadProject(_currentProject.ProjectFilePath);
            OpenFile(clsPath);
            SetStatusMessage($"Clase '{clsName}.cls' agregada al proyecto.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al agregar clase", ex);
            MessageBox.Show($"Error al crear la clase:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuAddUserControl_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentProject == null)
            {
                MessageBox.Show("Debes tener un proyecto abierto para agregar un control de usuario.", "Información", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            int idx = 1;
            string ucName = $"UserControl{idx}";
            while (_currentProject.UserControls.Any(u => string.Equals(Path.GetFileNameWithoutExtension(u.AbsolutePath), ucName, StringComparison.OrdinalIgnoreCase)) ||
                   File.Exists(Path.Combine(_currentProject.ProjectDirectory, $"{ucName}.ctl")))
            {
                idx++;
                ucName = $"UserControl{idx}";
            }

            string ucPath = Path.Combine(_currentProject.ProjectDirectory, $"{ucName}.ctl");
            string ucContent =
                "VERSION 5.00\r\n" +
                $"Begin VB.UserControl {ucName}\r\n" +
                "   ClientHeight    =   3600\r\n" +
                "   ClientLeft      =   0\r\n" +
                "   ClientTop       =   0\r\n" +
                "   ClientWidth     =   4800\r\n" +
                "   ScaleHeight     =   3600\r\n" +
                "   ScaleWidth      =   4800\r\n" +
                "End\r\n" +
                $"Attribute VB_Name = \"{ucName}\"\r\n" +
                "Attribute VB_GlobalNameSpace = False\r\n" +
                "Attribute VB_Creatable = True\r\n" +
                "Attribute VB_PredeclaredId = False\r\n" +
                "Attribute VB_Exposed = True\r\n" +
                "Option Explicit\r\n\r\n";

            File.WriteAllText(ucPath, ucContent, Encoding.GetEncoding(1252));

            string vbpText = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
            string newEntry = $"UserControl={ucName}.ctl\r\n";
            vbpText += "\r\n" + newEntry;
            File.WriteAllText(_currentProject.ProjectFilePath, vbpText, Encoding.GetEncoding(1252));

            LoadProject(_currentProject.ProjectFilePath);
            OpenFile(ucPath);
            SetStatusMessage($"Control de usuario '{ucName}.ctl' agregado y proyectado en la Caja de Herramientas.");
            AppLogger.Info($"Control de usuario agregado al proyecto: {ucPath}");
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al agregar control de usuario", ex);
            MessageBox.Show($"Error al crear el control de usuario:\n\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuReferences_Click(object sender, RoutedEventArgs e)
    {
        ShowReferencesDialog();
    }

    private void ShowReferencesDialog()
    {
        if (_currentProject == null)
        {
            MessageBox.Show("Debes abrir o crear un proyecto para gestionar sus referencias de bibliotecas de tipos.", "Referencias del Proyecto", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new Dialogs.ReferencesWindow(_currentProject)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true)
        {
            // 1. Registrar TypeLibs seleccionadas en IntelliSense para autocompletado y catálogo de miembros
            foreach (var r in _currentProject.References)
            {
                if (!string.IsNullOrEmpty(r.FilePath) && File.Exists(r.FilePath))
                {
                    if (OcxTypeLibReader.TryReadTypeLib(r.FilePath, out var lib, out _) && lib != null)
                    {
                        VbControlMemberCatalog.RegisterOcxLibrary(lib);
                    }
                }
            }

            // 2. Persistir lista completa de referencias en el archivo .vbp
            UpdateProjectReferencesInVbpFile(_currentProject);

            SetStatusMessage($"Referencias actualizadas: {_currentProject.References.Count} bibliotecas enlazadas.");
            AppLogger.Info($"Referencias del proyecto '{_currentProject.Name}' actualizadas. Total: {_currentProject.References.Count}");
        }
    }

    private void UpdateProjectReferencesInVbpFile(VB6Project project)
    {
        if (project == null || string.IsNullOrEmpty(project.ProjectFilePath) || !File.Exists(project.ProjectFilePath))
            return;

        try
        {
            var lines = File.ReadAllLines(project.ProjectFilePath, Encoding.GetEncoding(1252)).ToList();

            // Encontrar el primer índice donde estaban las referencias o insertar antes de Form=/Object=
            int insertIndex = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].TrimStart().StartsWith("Reference=", StringComparison.OrdinalIgnoreCase))
                {
                    insertIndex = i;
                    break;
                }
            }

            // Si no había referencias previas, buscar Form= o Object= o Startup=
            if (insertIndex < 0)
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    string trimmed = lines[i].TrimStart();
                    if (trimmed.StartsWith("Object=", StringComparison.OrdinalIgnoreCase) ||
                        trimmed.StartsWith("Form=", StringComparison.OrdinalIgnoreCase) ||
                        trimmed.StartsWith("Module=", StringComparison.OrdinalIgnoreCase) ||
                        trimmed.StartsWith("Class=", StringComparison.OrdinalIgnoreCase) ||
                        trimmed.StartsWith("Startup=", StringComparison.OrdinalIgnoreCase))
                    {
                        insertIndex = i;
                        break;
                    }
                }
            }

            if (insertIndex < 0) insertIndex = lines.Count;

            // Remover todas las líneas Reference= existentes
            lines.RemoveAll(l => l.TrimStart().StartsWith("Reference=", StringComparison.OrdinalIgnoreCase));

            if (insertIndex > lines.Count) insertIndex = lines.Count;

            // Insertar las referencias actualizadas en orden de prioridad
            var refLines = project.References
                .Select(r => !string.IsNullOrEmpty(r.RawLine)
                    ? r.RawLine
                    : $"Reference=*\\G{(r.Guid.StartsWith('{') ? r.Guid : "{" + r.Guid + "}")}#{r.MajorVersion}.{r.MinorVersion}#{r.Lcid}#{r.FilePath}#{r.Description}")
                .ToList();

            lines.InsertRange(insertIndex, refLines);

            File.WriteAllLines(project.ProjectFilePath, lines, Encoding.GetEncoding(1252));
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Error al actualizar referencias en {project.ProjectFilePath}: {ex.Message}", ex);
        }
    }

    private void MenuComponents_Click(object sender, RoutedEventArgs e)
    {
        ShowComponentsDialog();
    }

    private void ShowComponentsDialog()
    {
        var dlg = new Dialogs.ComponentsWindow(_currentProject)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true)
        {
            int addedCount = 0;
            int totalTools = 0;

            // 1. Procesar componentes seleccionados
            foreach (var item in dlg.SelectedItems)
            {
                var lib = item.LibraryInfo;
                if (lib == null && !string.IsNullOrEmpty(item.FilePath) && File.Exists(item.FilePath))
                {
                    if (OcxTypeLibReader.TryReadTypeLib(item.FilePath, out var readLib, out _))
                    {
                        lib = readLib;
                        item.LibraryInfo = readLib;
                    }
                }

                if (lib != null)
                {
                    // Registrar en el catÃ¡logo de miembros de IntelliSense
                    VbControlMemberCatalog.RegisterOcxLibrary(lib);

                    // Si es control visual o contiene controles, registrar en la caja de herramientas (Toolbox)
                    if (item.IsControl || lib.Controls.Any(c => c.IsControl))
                    {
                        ControlToolboxRegistry.RegisterOcxTools(lib);
                        totalTools += lib.Controls.Count(c => c.IsControl);
                    }

                    // Agregar al proyecto actual si estÃ¡ abierto
                    if (_currentProject != null)
                    {
                        string relPath = !string.IsNullOrEmpty(_currentProject.ProjectDirectory) && lib.FilePath.StartsWith(_currentProject.ProjectDirectory, StringComparison.OrdinalIgnoreCase)
                            ? Path.GetRelativePath(_currentProject.ProjectDirectory, lib.FilePath)
                            : Path.GetFileName(lib.FilePath);

                        _currentProject.AddActiveXControl(lib, relPath);

                        // Persistir directiva Object= o Reference= en el archivo .vbp en disco
                        try
                        {
                            if (File.Exists(_currentProject.ProjectFilePath))
                            {
                                string vbpContent = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
                                string guidFormatted = lib.Guid.StartsWith('{') ? lib.Guid : "{" + lib.Guid + "}";

                                if (item.IsControl || lib.Controls.Any(c => c.IsControl))
                                {
                                    string objLine = $"Object={guidFormatted}#{lib.MajorVersion}.{lib.MinorVersion}#0; {Path.GetFileName(lib.FilePath)}";
                                    if (!vbpContent.Contains(guidFormatted, StringComparison.OrdinalIgnoreCase))
                                    {
                                        int firstForm = vbpContent.IndexOf("Form=", StringComparison.OrdinalIgnoreCase);
                                        if (firstForm >= 0)
                                        {
                                            vbpContent = vbpContent.Insert(firstForm, objLine + "\r\n");
                                        }
                                        else
                                        {
                                            vbpContent += "\r\n" + objLine;
                                        }
                                        File.WriteAllText(_currentProject.ProjectFilePath, vbpContent, Encoding.GetEncoding(1252));
                                    }
                                }
                                else
                                {
                                    // Referencia COM / DLL ActiveX
                                    string refLine = $"Reference=*{guidFormatted}#{lib.MajorVersion}.{lib.MinorVersion}#0#{lib.FilePath}#{lib.Description}";
                                    if (!vbpContent.Contains(guidFormatted, StringComparison.OrdinalIgnoreCase))
                                    {
                                        int firstForm = vbpContent.IndexOf("Form=", StringComparison.OrdinalIgnoreCase);
                                        if (firstForm >= 0)
                                        {
                                            vbpContent = vbpContent.Insert(firstForm, refLine + "\r\n");
                                        }
                                        else
                                        {
                                            vbpContent += "\r\n" + refLine;
                                        }
                                        File.WriteAllText(_currentProject.ProjectFilePath, vbpContent, Encoding.GetEncoding(1252));
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            AppLogger.Warn($"No se pudo persistir la lÃ­nea del componente en el archivo VBP: {ex.Message}");
                        }
                    }

                    addedCount++;
                }
            }

            // 2. Procesar componentes deseleccionados
            foreach (var item in dlg.DeselectedItems)
            {
                string libName = !string.IsNullOrEmpty(item.LibraryName) ? item.LibraryName : Path.GetFileNameWithoutExtension(item.FilePath);

                // Desregistrar de VbControlMemberCatalog por nombre, ruta y GUID
                VbControlMemberCatalog.UnregisterOcxLibrary(libName);
                if (!string.IsNullOrEmpty(item.FilePath)) VbControlMemberCatalog.UnregisterOcxLibrary(item.FilePath);
                if (!string.IsNullOrEmpty(item.Guid)) VbControlMemberCatalog.UnregisterOcxLibrary(item.Guid);

                // Desregistrar de ControlToolboxRegistry por nombre y por prefijo
                ControlToolboxRegistry.UnregisterOcxTools(libName);
                if (item.LibraryInfo != null && !string.IsNullOrEmpty(item.LibraryInfo.LibraryName))
                {
                    ControlToolboxRegistry.UnregisterOcxTools(item.LibraryInfo.LibraryName);
                }

                // Remover del proyecto en memoria
                if (_currentProject != null)
                {
                    _currentProject.RemoveActiveXControl(libName);
                    if (!string.IsNullOrEmpty(item.FilePath)) _currentProject.RemoveActiveXControl(item.FilePath);
                    if (!string.IsNullOrEmpty(item.Guid)) _currentProject.RemoveActiveXControl(item.Guid);
                    _currentProject.RemoveReference(item.Guid);
                    if (!string.IsNullOrEmpty(item.FilePath)) _currentProject.RemoveReference(item.FilePath);

                    if (File.Exists(_currentProject.ProjectFilePath))
                    {
                        try
                        {
                            string vbpContent = File.ReadAllText(_currentProject.ProjectFilePath, Encoding.GetEncoding(1252));
                            string cleanGuid = item.Guid.Trim('{', '}');
                            string fileName = Path.GetFileName(item.FilePath);
                            string[] lines = vbpContent.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);

                            var filteredLines = lines.Where(l =>
                            {
                                if (!string.IsNullOrEmpty(cleanGuid) && l.Contains(cleanGuid, StringComparison.OrdinalIgnoreCase))
                                    return false;
                                if (!string.IsNullOrEmpty(fileName) && l.Contains(fileName, StringComparison.OrdinalIgnoreCase) && (l.StartsWith("Object=", StringComparison.OrdinalIgnoreCase) || l.StartsWith("Reference=", StringComparison.OrdinalIgnoreCase)))
                                    return false;
                                return true;
                            });

                            string newContent = string.Join("\r\n", filteredLines);
                            if (newContent != vbpContent)
                            {
                                File.WriteAllText(_currentProject.ProjectFilePath, newContent, Encoding.GetEncoding(1252));
                            }
                        }
                        catch (Exception ex)
                        {
                            AppLogger.Warn($"No se pudo remover la directiva del componente deseleccionado en VBP: {ex.Message}");
                        }
                    }
                }
            }

            // 3. Refrescar caja de herramientas en modo DiseÃ±o
            MainToolboxPanel.ReloadTools();

            // 3b. Refrescar superficies de diseÃ±o abiertas para re-resolver controles OCX que estaban pendientes
            foreach (var item in DocumentsTabControl.Items)
            {
                if (item is EditorTabItem tab && tab.DesignSurface != null && tab.DesignModel != null)
                {
                    foreach (var ctrl in tab.DesignModel.GetAllControls())
                    {
                        if (ctrl.IsUnresolved && VbControlMemberCatalog.IsControlRegistered(ctrl.TypeName))
                        {
                            ctrl.IsUnresolved = false;
                            ctrl.UnresolvedErrorMessage = null;
                        }
                    }
                    tab.DesignSurface.RefreshCanvas();
                }
            }

            SetStatusMessage($"Componentes actualizados: {addedCount} componentes activos ({totalTools} controles disponibles en Toolbox e IntelliSense).");
        }
    }

    private void LoadProjectActiveXComponents(VB6Project project)
    {
        if (project == null) return;

        int loadedCount = 0;
        foreach (var item in project.ActiveXControls)
        {
            string? resolvedPath = OcxTypeLibReader.ResolveOcxPath(item.RelativePath, project.ProjectDirectory, item.Guid);
            if (string.IsNullOrEmpty(resolvedPath) && !string.IsNullOrEmpty(item.AbsolutePath) && File.Exists(item.AbsolutePath))
            {
                resolvedPath = item.AbsolutePath;
            }

            if (!string.IsNullOrEmpty(resolvedPath) && File.Exists(resolvedPath))
            {
                if (OcxTypeLibReader.TryReadTypeLib(resolvedPath, out var lib, out _))
                {
                    VbControlMemberCatalog.RegisterOcxLibrary(lib!);
                    ControlToolboxRegistry.RegisterOcxTools(lib!);
                    if (!project.LoadedOcxLibraries.Contains(lib!))
                    {
                        project.LoadedOcxLibraries.Add(lib!);
                    }
                    loadedCount++;
                }
            }
        }

        if (loadedCount > 0)
        {
            MainToolboxPanel.ReloadTools();
            AppLogger.Info($"Se cargaron {loadedCount} componentes ActiveX del proyecto '{project.Name}'.");
        }
    }

    private void MenuOpenProject_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Abrir Proyecto o Grupo de Proyectos ViBa7 (Compatibilidad con VB6)",
            Filter = "Proyectos y Grupos ViBa7 (*.vbp;*.vbg)|*.vbp;*.vbg|Grupos de Proyectos (*.vbg)|*.vbg|Proyectos ViBa7 (*.vbp)|*.vbp|Todos los archivos (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            LoadProject(dialog.FileName);
        }
    }

    private void MenuOpenFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Abrir Archivo de CÃ³digo ViBa7 (Compatibilidad con VB6)",
            Filter = "Archivos de CÃ³digo ViBa7 (*.bas;*.cls;*.frm;*.ctl;*.pag;*.dob;*.dsr)|*.bas;*.cls;*.frm;*.ctl;*.pag;*.dob;*.dsr|MÃ³dulos (*.bas)|*.bas|Clases (*.cls)|*.cls|Formularios (*.frm)|*.frm|User Controls (*.ctl)|*.ctl|PÃ¡ginas de Propiedades (*.pag)|*.pag|Documentos de Usuario (*.dob)|*.dob|DiseÃ±adores (*.dsr)|*.dsr|Todos los archivos (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            OpenFile(dialog.FileName);
        }
    }

    private void MenuSave_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab == null) return;

        var result = ActiveTab.Save(_settings.CreateBackupBeforeSave);
        switch (result.Status)
        {
            case SaveStatus.Success:
                if (ActiveTab.Document.FilePath?.EndsWith(".ctl", StringComparison.OrdinalIgnoreCase) == true)
                {
                    _currentProject?.InvalidateUserControlCache();
                    _currentProjectGroup?.InvalidateUserControlCache();

                    foreach (var item in DocumentsTabControl.Items)
                    {
                        if (item is EditorTabItem tab && tab.DesignSurface != null)
                        {
                            tab.DesignSurface.RefreshCanvas();
                        }
                    }

                    if (_currentProjectGroup != null)
                        ControlToolboxRegistry.RegisterUserControlTools(_currentProjectGroup);
                    else if (_currentProject != null)
                        ControlToolboxRegistry.RegisterUserControlTools(_currentProject);

                    MainToolboxPanel.ReloadTools();
                }
                SetStatusMessage($"'{ActiveTab.Document.DisplayName}' guardado exitosamente.");
                break;
            case SaveStatus.NotModified:
                SetStatusMessage($"'{ActiveTab.Document.DisplayName}' no tiene cambios.");
                break;
            case SaveStatus.ReadOnlyFile:
            case SaveStatus.ExternalConflict:
            case SaveStatus.Error:
                MessageBox.Show(result.Message, "Error al guardar", MessageBoxButton.OK, MessageBoxImage.Warning);
                SetStatusMessage(result.Message);
                break;
        }
    }

    private void MenuSaveAs_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab == null) return;

        var dialog = new SaveFileDialog
        {
            Title = "Guardar como",
            FileName = ActiveTab.Document.DisplayName,
            Filter = "Archivos de CÃ³digo VB6 (*.bas;*.cls;*.frm)|*.bas;*.cls;*.frm|Todos los archivos (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            ActiveTab.Document.FilePath = dialog.FileName;
            ActiveTab.Document.DisplayName = Path.GetFileName(dialog.FileName);
            ActiveTab.Document.IsDirty = true;
            MenuSave_Click(sender, e);
        }
    }

    private void MenuSaveAll_Click(object sender, RoutedEventArgs e)
    {
        int count = 0;
        bool userControlSaved = false;
        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab && tab.Document.IsDirty)
            {
                var result = tab.Save(_settings.CreateBackupBeforeSave);
                if (result.Status == SaveStatus.Success)
                {
                    count++;
                    if (tab.Document.FilePath?.EndsWith(".ctl", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        userControlSaved = true;
                    }
                }
            }
        }

        if (userControlSaved)
        {
            _currentProject?.InvalidateUserControlCache();
            _currentProjectGroup?.InvalidateUserControlCache();

            foreach (var item in DocumentsTabControl.Items)
            {
                if (item is EditorTabItem tab && tab.DesignSurface != null)
                {
                    tab.DesignSurface.RefreshCanvas();
                }
            }

            if (_currentProjectGroup != null)
                ControlToolboxRegistry.RegisterUserControlTools(_currentProjectGroup);
            else if (_currentProject != null)
                ControlToolboxRegistry.RegisterUserControlTools(_currentProject);

            MainToolboxPanel.ReloadTools();
        }

        SetStatusMessage(count > 0 ? $"Se guardaron {count} documentos modificados." : "Todos los documentos estÃ¡n actualizados.");
    }

    private void MenuCloseDoc_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null)
        {
            CloseTab(ActiveTab);
        }
    }

    private void MenuCloseAll_Click(object sender, RoutedEventArgs e)
    {
        CloseProjectAndAllDocuments();
    }

    private void MenuSaveProjectGroup_Click(object sender, RoutedEventArgs e)
    {
        if (_currentProjectGroup == null) return;
        try
        {
            _currentProjectGroup.Save();
            SetStatusMessage($"Grupo de proyectos '{_currentProjectGroup.Name}' guardado exitosamente.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error al guardar grupo de proyectos:\n\n{ex.Message}", "Guardar Grupo", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuSetStartupProject_Click(object sender, RoutedEventArgs e)
    {
        if (_currentProjectGroup == null) return;

        var targetProj = _currentProject;
        if (targetProj == null || _currentProjectGroup.IsPrimaryProject(targetProj))
        {
            targetProj = _currentProjectGroup.SecondaryProjects.FirstOrDefault();
        }

        if (targetProj != null)
        {
            _currentProjectGroup.SetPrimaryProject(targetProj);
            _currentProject = targetProj;
            RefreshProjectTree();
            UpdateMakeExeMenuState();
            SetStatusMessage($"'{targetProj.Name}' establecido como Proyecto Principal (Inicio).");
        }
    }

    private void UpdateMakeExeMenuState()
    {
        if (MenuMakeExe == null) return;

        var projToBuild = _currentProject ?? _currentProjectGroup?.PrimaryProject;
        if (projToBuild != null)
        {
            string name = !string.IsNullOrWhiteSpace(projToBuild.Name) ? projToBuild.Name : "Proyecto1";
            string roleSuffix = _currentProjectGroup != null && _currentProjectGroup.IsPrimaryProject(projToBuild) ? " (Principal)" : "";
            MenuMakeExe.Header = $"Generar _{name}.exe{roleSuffix}...";
            MenuMakeExe.IsEnabled = true;
            if (BtnToolbarMakeExe != null) BtnToolbarMakeExe.IsEnabled = true;
        }
        else
        {
            MenuMakeExe.Header = "Generar _Ejecutable...";
            MenuMakeExe.IsEnabled = false;
            if (BtnToolbarMakeExe != null) BtnToolbarMakeExe.IsEnabled = false;
        }
    }

    private void MenuMakeExe_Click(object sender, RoutedEventArgs e)
    {
        if (_currentProject == null)
        {
            MessageBox.Show("No hay ningÃºn proyecto activo cargado para generar el ejecutable.\n\nPor favor abra o cree un proyecto primero.", "Generar Ejecutable", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new MakeProjectExeDialog(_currentProject)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && dlg.CompileResult != null)
        {
            var res = dlg.CompileResult;
            SetStatusMessage($"Ejecutable '{Path.GetFileName(res.OutputExePath)}' generado exitosamente ({res.OutputSizeBytes:N0} bytes).");
            AppLogger.Info($"CompilaciÃ³n nativa x32 exitosa: {res.OutputExePath} ({res.OutputSizeBytes} bytes)");

            var ask = MessageBox.Show(
                $"El ejecutable nativo PE32 fue generado exitosamente:\n\n{res.OutputExePath}\nTamaÃ±o: {res.OutputSizeBytes:N0} bytes\nEnlace: msvbvm60.dll!ThunRTMain\n\nÂ¿Desea abrir la carpeta contenedora?",
                "CompilaciÃ³n Exitosa",
                MessageBoxButton.YesNo,
                MessageBoxImage.Information);

            if (ask == MessageBoxResult.Yes)
            {
                try
                {
                    string? dir = Path.GetDirectoryName(res.OutputExePath);
                    if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = $"/select,\"{res.OutputExePath}\"",
                            UseShellExecute = true
                        });
                    }
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"No se pudo abrir el explorador para {res.OutputExePath}: {ex.Message}");
                }
            }
        }
    }

    private void MenuExit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
    #endregion

    #region Menu Edit Actions
    private void MenuUndo_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null && ActiveEditor.CanUndo) ActiveEditor.Undo();
    }

    private void MenuRedo_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null && ActiveEditor.CanRedo) ActiveEditor.Redo();
    }

    private void MenuCut_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab?.IsDesignerActive == true && ActiveTab.DesignSurface != null)
        {
            ActiveTab.DesignSurface.CutSelectedControls();
            return;
        }

        if (ActiveEditor != null)
        {
            string sel = ActiveEditor.SelectedText;
            if (!string.IsNullOrWhiteSpace(sel))
            {
                ClipboardRing.Instance.Add(sel);
            }
            ActiveEditor.Cut();
        }
    }

    private void MenuCopy_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab?.IsDesignerActive == true && ActiveTab.DesignSurface != null)
        {
            ActiveTab.DesignSurface.CopySelectedControls();
            return;
        }

        if (ActiveEditor != null)
        {
            string sel = ActiveEditor.SelectedText;
            if (!string.IsNullOrWhiteSpace(sel))
            {
                ClipboardRing.Instance.Add(sel);
            }
            ActiveEditor.Copy();
        }
    }

    private void MenuPaste_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab?.IsDesignerActive == true && ActiveTab.DesignSurface != null)
        {
            ActiveTab.DesignSurface.PasteControls();
            return;
        }

        ActiveEditor?.Paste();
    }

    private void MenuFind_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null) return;
        var dlg = new FindReplaceDialog(ActiveEditor) { Owner = this };
        dlg.Show();
    }

    private void MenuReplace_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null) return;
        var dlg = new FindReplaceDialog(ActiveEditor) { Owner = this };
        dlg.Show();
    }

    private void MenuFindInProject_Click(object sender, RoutedEventArgs e)
    {
        IReadOnlyList<VB6Project> projectsToSearch;
        if (_currentProjectGroup != null && _currentProjectGroup.Projects.Count > 0)
        {
            projectsToSearch = _currentProjectGroup.Projects;
        }
        else if (_currentProject != null)
        {
            projectsToSearch = [_currentProject];
        }
        else
        {
            MessageBox.Show("Debes abrir un proyecto .vbp o grupo .vbg para buscar.", "Buscar en Proyecto", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new FindInProjectDialog(projectsToSearch, (filePath, line) =>
        {
            var tab = OpenFile(filePath);
            if (tab != null)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(Math.Min(line, tab.Editor.Document.LineCount));
                tab.Editor.CaretOffset = lineObj.Offset;
                tab.Editor.ScrollToLine(line);
                tab.Editor.Focus();
            }
        })
        {
            Owner = this
        };
        dlg.Show();
    }

    private void MenuGoToLine_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null) return;

        int currentLine = ActiveEditor.TextArea.Caret.Line;
        int lineCount = ActiveEditor.Document.LineCount;

        var dlg = new GoToLineDialog(currentLine, lineCount) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            int targetLine = dlg.SelectedLine;
            var line = ActiveEditor.Document.GetLineByNumber(targetLine);
            ActiveEditor.CaretOffset = line.Offset;
            ActiveEditor.ScrollToLine(targetLine);
        }
    }
    #endregion

    #region Menu View Actions
    private void BtnViewCode_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null && ActiveTab.DesignSurface != null)
        {
            ActiveTab.SwitchToCodeView();
            return;
        }

        if (ProjectTreeView.SelectedItem is TreeViewItem selItem && selItem.Tag is ProjectItem item)
        {
            var tab = OpenFile(item.AbsolutePath);
            tab?.SwitchToCodeView();
        }
    }

    private void BtnViewObject_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null && ActiveTab.DesignSurface != null)
        {
            ActiveTab.SwitchToDesignerView();
            return;
        }

        if (ProjectTreeView.SelectedItem is TreeViewItem selItem && selItem.Tag is ProjectItem item)
        {
            var tab = OpenFile(item.AbsolutePath);
            tab?.SwitchToDesignerView();
        }
    }

    private void MenuObjectBrowser_Click(object sender, RoutedEventArgs e)
    {
        OpenObjectBrowser();
    }

    public void OpenObjectBrowser()
    {
        var ob = new Dialogs.ObjectBrowserWindow(_currentProject)
        {
            Owner = this
        };

        ob.RequestNavigateToCode = (filePath, lineNumber) =>
        {
            var tab = OpenFile(filePath);
            if (tab != null)
            {
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
                {
                    tab.SwitchToCodeView();
                    tab.Editor.TextArea.Caret.Line = Math.Max(1, lineNumber);
                    tab.Editor.TextArea.Caret.Column = 1;
                    tab.Editor.ScrollTo(lineNumber, 1);
                    tab.Editor.Focus();
                }));
            }
        };

        ob.ShowDialog();
    }

    private void MenuToggleToolbox_Click(object sender, RoutedEventArgs e)
    {
        ColToolbox.Width = MenuToggleToolbox.IsChecked ? new GridLength(_settings.ToolboxWidth) : new GridLength(0);
    }

    private void MenuToggleProjectExplorer_Click(object sender, RoutedEventArgs e)
    {
        ColProjectExplorer.Width = MenuToggleProjectExplorer.IsChecked ? new GridLength(_settings.ProjectExplorerWidth) : new GridLength(0);
    }

    private void MenuToggleProperties_Click(object sender, RoutedEventArgs e)
    {
        RowProperties.Height = MenuToggleProperties.IsChecked ? new GridLength(_settings.PropertiesHeight) : new GridLength(0);
        RowPropertiesSplitter.Height = MenuToggleProperties.IsChecked ? new GridLength(4) : new GridLength(0);
    }

    private void MenuToggleStatusBar_Click(object sender, RoutedEventArgs e)
    {
        MainStatusBar.Visibility = MenuToggleStatusBar.IsChecked ? Visibility.Visible : Visibility.Collapsed;
    }

    private void MenuToggleOutput_Click(object sender, RoutedEventArgs e)
    {
        SetOutputPanelVisibility(MenuToggleOutput.IsChecked);
        if (MenuToggleOutput.IsChecked)
        {
            ShowBottomTab("Output");
        }
    }

    private void SetOutputPanelVisibility(bool isVisible)
    {
        if (RowOutput == null || RowOutputSplitter == null) return;
        double h = _settings.OutputPanelHeight > 50 ? _settings.OutputPanelHeight : 160;
        RowOutput.Height = isVisible ? new GridLength(h) : new GridLength(0);
        RowOutputSplitter.Height = isVisible ? new GridLength(4) : new GridLength(0);
        if (MenuToggleOutput != null) MenuToggleOutput.IsChecked = isVisible;
    }

    private void BottomTabRadio_Click(object sender, RoutedEventArgs e)
    {
        if (TabRadioOutput?.IsChecked == true) ShowBottomTab("Output");
        else if (TabRadioReferences?.IsChecked == true) ShowBottomTab("References");
        else if (TabRadioTasks?.IsChecked == true) ShowBottomTab("Tasks");
        else if (TabRadioAnalysis?.IsChecked == true) ShowBottomTab("Analysis");
        else if (TabRadioImmediate?.IsChecked == true) ShowBottomTab("Immediate");
        else if (TabRadioLocals?.IsChecked == true) ShowBottomTab("Locals");
        else if (TabRadioCallStack?.IsChecked == true) ShowBottomTab("CallStack");
    }

    public void ShowBottomTab(string tabName)
    {
        if (ViewOutput == null || ViewReferences == null || ViewTasks == null || ViewAnalysis == null || ViewImmediate == null || ViewLocals == null || ViewCallStack == null) return;
        if (TabRadioOutput == null || TabRadioReferences == null || TabRadioTasks == null || TabRadioAnalysis == null || TabRadioImmediate == null || TabRadioLocals == null || TabRadioCallStack == null) return;

        ViewOutput.Visibility = (tabName == "Output") ? Visibility.Visible : Visibility.Collapsed;
        ViewReferences.Visibility = (tabName == "References") ? Visibility.Visible : Visibility.Collapsed;
        ViewTasks.Visibility = (tabName == "Tasks") ? Visibility.Visible : Visibility.Collapsed;
        ViewAnalysis.Visibility = (tabName == "Analysis") ? Visibility.Visible : Visibility.Collapsed;
        ViewImmediate.Visibility = (tabName == "Immediate") ? Visibility.Visible : Visibility.Collapsed;
        ViewLocals.Visibility = (tabName == "Locals") ? Visibility.Visible : Visibility.Collapsed;
        ViewCallStack.Visibility = (tabName == "CallStack") ? Visibility.Visible : Visibility.Collapsed;

        TabRadioOutput.IsChecked = (tabName == "Output");
        TabRadioReferences.IsChecked = (tabName == "References");
        TabRadioTasks.IsChecked = (tabName == "Tasks");
        TabRadioAnalysis.IsChecked = (tabName == "Analysis");
        TabRadioImmediate.IsChecked = (tabName == "Immediate");
        TabRadioLocals.IsChecked = (tabName == "Locals");
        TabRadioCallStack.IsChecked = (tabName == "CallStack");

        SetOutputPanelVisibility(true);
    }

    public class LocalVariableDisplayItem
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Scope { get; set; } = string.Empty;
    }

    public class CallStackDisplayItem
    {
        public string ProcedureName { get; set; } = string.Empty;
        public string ContainerName { get; set; } = string.Empty;
        public int LineNumber { get; set; }
    }

    public void RefreshLocalsWindow()
    {
        if (_liveExecutionHost == null || LstLocals == null) return;

        var items = new List<LocalVariableDisplayItem>();
        var frame = _liveExecutionHost.Context.CurrentFrame;
        if (frame != null)
        {
            foreach (var kvp in frame.LocalVariables.OrderBy(k => k.Key))
            {
                items.Add(new LocalVariableDisplayItem
                {
                    Name = kvp.Key,
                    Value = kvp.Value.AsString(),
                    Type = kvp.Value.VarType.ToString(),
                    Scope = "Local"
                });
            }

            foreach (var kvp in _liveExecutionHost.Context.StaticVariables.OrderBy(k => k.Key))
            {
                if (kvp.Key.StartsWith(frame.ProcedureName + ".", StringComparison.OrdinalIgnoreCase))
                {
                    string varName = kvp.Key[(frame.ProcedureName.Length + 1)..];
                    items.Add(new LocalVariableDisplayItem
                    {
                        Name = varName,
                        Value = kvp.Value.AsString(),
                        Type = kvp.Value.VarType.ToString(),
                        Scope = "Static"
                    });
                }
            }
        }

        foreach (var kvp in _liveExecutionHost.Context.ModuleVariables.OrderBy(k => k.Key))
        {
            if (!kvp.Key.Contains('.'))
            {
                items.Add(new LocalVariableDisplayItem
                {
                    Name = kvp.Key,
                    Value = kvp.Value.AsString(),
                    Type = kvp.Value.VarType.ToString(),
                    Scope = "Module"
                });
            }
        }

        LstLocals.ItemsSource = items;
    }

    public void RefreshCallStackWindow()
    {
        if (_liveExecutionHost == null || LstCallStack == null) return;

        var stackItems = new List<CallStackDisplayItem>();
        foreach (var frame in _liveExecutionHost.Context.CallStack)
        {
            stackItems.Add(new CallStackDisplayItem
            {
                ProcedureName = frame.ProcedureName,
                ContainerName = frame.ModuleOrFormName ?? "(General)",
                LineNumber = frame.CurrentLineNumber
            });
        }

        LstCallStack.ItemsSource = stackItems;
    }

    private void LstCallStack_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (LstCallStack.SelectedItem is CallStackDisplayItem item)
        {
            if (item.LineNumber > 0 && ActiveEditor != null && item.LineNumber <= ActiveEditor.Document.LineCount)
            {
                ActiveEditor.ScrollToLine(item.LineNumber);
                ActiveEditor.TextArea.Caret.Line = item.LineNumber;
                ActiveEditor.Focus();
            }
        }
    }

    private void MenuImmediateWindow_Click(object sender, RoutedEventArgs e)
    {
        ShowBottomTab("Immediate");
        MainImmediateWindowControl?.TxtInput.Focus();
    }

    private void MenuToggleReferences_Click(object sender, RoutedEventArgs e)
    {
        ShowBottomTab("References");
    }

    private void MenuToggleTasks_Click(object sender, RoutedEventArgs e)
    {
        ShowBottomTab("Tasks");
        RefreshTaskList();
    }

    private void BtnClearOutput_Click(object sender, RoutedEventArgs e)
    {
        TxtOutputLog.Clear();
    }

    private void BtnCloseOutput_Click(object sender, RoutedEventArgs e)
    {
        MenuToggleOutput.IsChecked = false;
        SetOutputPanelVisibility(false);
    }

    private void MenuOpenLogsFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string dir = Path.GetDirectoryName(AppLogger.LogFilePath)!;
            Process.Start(new ProcessStartInfo
            {
                FileName = dir,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error("No se pudo abrir la carpeta de logs", ex);
        }
    }
    #endregion

    #region Tools & Options
    private void MenuMenuEditor_Click(object sender, RoutedEventArgs e)
    {
        OpenMenuEditor();
    }

    public void OpenMenuEditor()
    {
        if (ActiveTab == null)
        {
            MessageBox.Show("No hay ningÃºn documento activo para editar menÃºs.", "Editor de MenÃºs", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (ActiveTab.DesignModel == null)
        {
            MessageBox.Show("El Editor de MenÃºs solo estÃ¡ disponible para formularios y controles de usuario ViBa7 (Compatibilidad con VB6) (.frm / .ctl).", "Editor de MenÃºs", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            var editorWnd = new ViBa7.IDE.Designer.MenuEditor.MenuEditorWindow(ActiveTab.DesignModel)
            {
                Owner = this
            };

            if (editorWnd.ShowDialog() == true)
            {
                // El modelo fue actualizado con la nueva jerarquÃ­a de menÃºs
                ActiveTab.Document.Header = FormDesignSerializer.Serialize(ActiveTab.DesignModel);
                ActiveTab.Document.IsDirty = true;
                ActiveTab.DesignSurface?.RefreshVisuals();
                MainPropertyGridPanel.RefreshObjectList();
                SetStatusMessage("Estructura de menÃºs actualizada.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("Error al abrir el Editor de MenÃºs", ex);
            MessageBox.Show($"OcurriÃ³ un error al abrir el Editor de MenÃºs:\n\n{ex.Message}", "Editor de MenÃºs", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void MenuOptions_Click(object sender, RoutedEventArgs e)
    {
        var opt = new OptionsWindow(_settings) { Owner = this };
        if (opt.ShowDialog() == true)
        {
            // Aplicar cambios a las pestaÃ±as abiertas
            foreach (var item in DocumentsTabControl.Items)
            {
                if (item is EditorTabItem tab)
                {
                    tab.ApplyThemeSettings(_settings);
                }
            }

            UpdateExplorerModeUI();
            RefreshProjectTree();
        }
    }
    #endregion

    #region Menu Help Actions
    private void MenuAbout_Click(object sender, RoutedEventArgs e)
    {
        var about = new AboutWindow { Owner = this };
        about.ShowDialog();
    }
    #endregion

    #region Funciones Avanzadas
    private bool _isChordCtrlK;

    private void MainWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        // F5: Iniciar EjecuciÃ³n en Caliente / Shift+F5: Detener
        if (e.Key == Key.F5)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                MenuRunStop_Click(this, new RoutedEventArgs());
            }
            else
            {
                MenuRunStart_Click(this, new RoutedEventArgs());
            }
            e.Handled = true;
            return;
        }

        // Ctrl+Shift+B: Generar Ejecutable (Build / Make EXE)
        if (e.Key == Key.B && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
        {
            MenuMakeExe_Click(this, new RoutedEventArgs());
            e.Handled = true;
            return;
        }

        // Ctrl+G: Ventana Inmediata (Immediate Window)
        if (e.Key == Key.G && Keyboard.Modifiers == ModifierKeys.Control)
        {
            MenuImmediateWindow_Click(this, new RoutedEventArgs());
            e.Handled = true;
            return;
        }

        // F2: Examinador de Objetos (VB6 Clásico)
        if (e.Key == Key.F2 && Keyboard.Modifiers == ModifierKeys.None)
        {
            OpenObjectBrowser();
            e.Handled = true;
            return;
        }

        // Ctrl+Break / Pause: Interrumpir / Pausa
        if (e.Key is Key.Pause or Key.Cancel && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            MenuRunPause_Click(this, new RoutedEventArgs());
            e.Handled = true;
            return;
        }

        // F8: Paso a paso por instrucciones / Shift+F8: Paso a paso por procedimientos
        if (e.Key == Key.F8)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                MenuStepOver_Click(this, new RoutedEventArgs());
            }
            else
            {
                MenuStepInto_Click(this, new RoutedEventArgs());
            }
            e.Handled = true;
            return;
        }

        // F9: Alternar punto de interrupciÃ³n
        if (e.Key == Key.F9)
        {
            if ((Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                MenuClearAllBreakpoints_Click(this, new RoutedEventArgs());
                e.Handled = true;
                return;
            }
            if (Keyboard.Modifiers == ModifierKeys.None)
            {
                MenuToggleBreakpoint_Click(this, new RoutedEventArgs());
                e.Handled = true;
                return;
            }
        }

        // Ctrl+Shift+N: Nuevo Proyecto / Ctrl+N: Nuevo Formulario
        if (e.Key == Key.N && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                MenuNewProject_Click(this, new RoutedEventArgs());
            }
            else
            {
                MenuNewForm_Click(this, new RoutedEventArgs());
            }
            e.Handled = true;
            return;
        }

        // Ctrl+O: Abrir Proyecto o Grupo
        if (e.Key == Key.O && Keyboard.Modifiers == ModifierKeys.Control)
        {
            MenuOpenProject_Click(this, new RoutedEventArgs());
            e.Handled = true;
            return;
        }

        // Ctrl+S / Ctrl+Shift+S: Guardar / Guardar Todo
        if (e.Key == Key.S && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                MenuSaveAll_Click(this, new RoutedEventArgs());
            }
            else
            {
                MenuSave_Click(this, new RoutedEventArgs());
            }
            e.Handled = true;
            return;
        }

        // Ctrl+W: Cerrar Documento / Ctrl+Shift+W: Cerrar Todo
        if (e.Key == Key.W && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                MenuCloseAll_Click(this, new RoutedEventArgs());
            }
            else
            {
                MenuCloseDoc_Click(this, new RoutedEventArgs());
            }
            e.Handled = true;
            return;
        }

        // Ctrl+E: Editor de MenÃºs (VB6 ClÃ¡sico)
        if (e.Key == Key.E && Keyboard.Modifiers == ModifierKeys.Control)
        {
            OpenMenuEditor();
            e.Handled = true;
            return;
        }

        // F7: Ver CÃ³digo / Shift+F7: Ver DiseÃ±ador (Flujo Dual VB6)
        if (e.Key == Key.F7)
        {
            if (ActiveTab?.DesignSurface != null)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                {
                    ActiveTab.SwitchToDesignerView();
                }
                else
                {
                    ActiveTab.SwitchToCodeView();
                }
                e.Handled = true;
                return;
            }
        }

        // IntelliSense: Ctrl+J (List Members), Ctrl+I (Parameter Info), Ctrl+Espacio (Complete Word)
        if (ActiveTab != null && !ActiveTab.IsDesignerActive)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.J)
            {
                ActiveTab.Editor.Focus();
                ActiveTab.IntelliSenseController.ShowCompletionWindow(explicitTrigger: true);
                e.Handled = true;
                return;
            }
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.I)
            {
                ActiveTab.Editor.Focus();
                ActiveTab.IntelliSenseController.ShowParameterInsight(explicitTrigger: true);
                e.Handled = true;
                return;
            }
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Space)
            {
                ActiveTab.Editor.Focus();
                ActiveTab.IntelliSenseController.ShowCompletionWindow(explicitTrigger: true);
                e.Handled = true;
                return;
            }
        }

        // F4: Mostrar / Enfocar Ventana de Propiedades
        if (e.Key == Key.F4 && Keyboard.Modifiers == ModifierKeys.None)
        {
            if (RowProperties.Height.Value <= 30)
            {
                RowProperties.Height = new GridLength(_settings.PropertiesHeight);
                RowPropertiesSplitter.Height = new GridLength(4);
                MenuToggleProperties.IsChecked = true;
            }
            MainPropertyGridPanel.Focus();
            e.Handled = true;
            return;
        }

        // Ctrl+Alt+X: Alternar Caja de Herramientas (Toolbox)
        if (e.Key == Key.X && (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Alt)) == (ModifierKeys.Control | ModifierKeys.Alt))
        {
            MenuToggleToolbox.IsChecked = !MenuToggleToolbox.IsChecked;
            ColToolbox.Width = MenuToggleToolbox.IsChecked ? new GridLength(_settings.ToolboxWidth) : new GridLength(0);
            e.Handled = true;
            return;
        }

        // Ctrl+T: Componentes ActiveX / OCX
        if (e.Key == Key.T && Keyboard.Modifiers == ModifierKeys.Control)
        {
            ShowComponentsDialog();
            e.Handled = true;
            return;
        }

        // 0. NavegaciÃ³n AtrÃ¡s / Adelante: Alt+Left o Ctrl+- / Alt+Right o Ctrl+Shift+-
        if ((e.Key == Key.Left && Keyboard.Modifiers == ModifierKeys.Alt) ||
            (e.Key == Key.OemMinus && Keyboard.Modifiers == ModifierKeys.Control))
        {
            NavigateBackward();
            e.Handled = true;
            return;
        }

        if ((e.Key == Key.Right && Keyboard.Modifiers == ModifierKeys.Alt) ||
            (e.Key == Key.OemMinus && (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift)))
        {
            NavigateForward();
            e.Handled = true;
            return;
        }

        // Buscar todas las referencias: Shift+F12
        if (e.Key == Key.F12 && (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
        {
            FindReferencesForCurrentSymbol();
            e.Handled = true;
            return;
        }

        // Marcadores directos: Ctrl+F2 (alternar), F2 (siguiente), Shift+F2 (anterior), Ctrl+Shift+F2 (limpiar)
        if (e.Key == Key.F2)
        {
            if ((Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                ClearAllBookmarks();
                e.Handled = true;
                return;
            }
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                ToggleBookmarkAtCurrentLine();
                e.Handled = true;
                return;
            }
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                GoToPreviousBookmark();
                e.Handled = true;
                return;
            }
            if (Keyboard.Modifiers == ModifierKeys.None)
            {
                GoToNextBookmark();
                e.Handled = true;
                return;
            }
        }

        // 1. Invertir Asignaciones: Ctrl+Shift+R
        if (e.Key == Key.R && (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
        {
            if (ActiveEditor != null)
            {
                EditorOperationsHelper.SwitchAssignment(ActiveEditor);
                SetStatusMessage("Asignaciones invertidas (A = B ? B = A).");
                e.Handled = true;
                return;
            }
        }

        // 2. Secuencia acorde Ctrl+K
        if (_isChordCtrlK)
        {
            _isChordCtrlK = false;
            if (e.Key == Key.D || e.Key == Key.F)
            {
                FormatActiveDocument();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.C)
            {
                if (ActiveEditor != null)
                {
                    EditorOperationsHelper.CommentSelection(ActiveEditor);
                    SetStatusMessage("LÃ­neas comentadas.");
                }
                e.Handled = true;
                return;
            }
            if (e.Key == Key.U)
            {
                if (ActiveEditor != null)
                {
                    EditorOperationsHelper.UncommentSelection(ActiveEditor);
                    SetStatusMessage("LÃ­neas descomentadas.");
                }
                e.Handled = true;
                return;
            }
            if (e.Key == Key.K)
            {
                ToggleBookmarkAtCurrentLine();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.N)
            {
                GoToNextBookmark();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.P)
            {
                GoToPreviousBookmark();
                e.Handled = true;
                return;
            }
            if (e.Key == Key.L)
            {
                ClearAllBookmarks();
                e.Handled = true;
                return;
            }
            SetStatusMessage("CombinaciÃ³n Ctrl+K cancelada.");
        }
        else if (e.Key == Key.K && Keyboard.Modifiers == ModifierKeys.Control)
        {
            _isChordCtrlK = true;
            SetStatusMessage("Ctrl+K presionado. Esperando segunda tecla (D: Formatear, C: Comentar, U: Descomentar, K: Marcador, N: Siguiente, P: Anterior, L: Limpiar)...");
            e.Handled = true;
            return;
        }

        // 3. Portapapeles mÃºltiple: Ctrl+Shift+V
        if (e.Key == Key.V && (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
        {
            OpenClipboardRingDialog();
            e.Handled = true;
            return;
        }
    }

    private void MenuToggleLineHighlight_Click(object sender, RoutedEventArgs e)
    {
        SetLineHighlightState(MenuToggleLineHighlight.IsChecked);
    }

    private void BtnToolbarToggleLineHighlight_Click(object sender, RoutedEventArgs e)
    {
        SetLineHighlightState(BtnToolbarToggleLineHighlight.IsChecked == true);
    }

    private void SetLineHighlightState(bool enabled)
    {
        _settings.EnableCurrentLineHighlight = enabled;
        MenuToggleLineHighlight.IsChecked = enabled;
        BtnToolbarToggleLineHighlight.IsChecked = enabled;

        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab)
            {
                tab.SetCurrentLineHighlightEnabled(enabled);
            }
        }

        SetStatusMessage(enabled ? "Resaltado de lÃ­nea activa activado." : "Resaltado de lÃ­nea activa desactivado.");
    }

    private void MenuToggleBranchLines_Click(object sender, RoutedEventArgs e)
    {
        SetBranchLinesState(MenuToggleBranchLines.IsChecked);
    }

    private void BtnToolbarToggleBranchLines_Click(object sender, RoutedEventArgs e)
    {
        SetBranchLinesState(BtnToolbarToggleBranchLines.IsChecked == true);
    }

    private void SetBranchLinesState(bool enabled)
    {
        _settings.EnableBranchLines = enabled;
        MenuToggleBranchLines.IsChecked = enabled;
        BtnToolbarToggleBranchLines.IsChecked = enabled;

        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab)
            {
                tab.SetBranchLinesEnabled(enabled);
            }
        }

        SetStatusMessage(enabled ? "GuÃ­as de bloques (Branch Lines) activadas." : "GuÃ­as de bloques desactivadas.");
    }

    private void MenuToggleEndOfLineDetails_Click(object sender, RoutedEventArgs e)
    {
        SetEndOfLineDetailsState(MenuToggleEndOfLineDetails.IsChecked);
    }

    private void BtnToolbarToggleEndOfLineDetails_Click(object sender, RoutedEventArgs e)
    {
        SetEndOfLineDetailsState(BtnToolbarToggleEndOfLineDetails.IsChecked == true);
    }

    private void SetEndOfLineDetailsState(bool enabled)
    {
        _settings.EnableEndOfLineDetails = enabled;
        MenuToggleEndOfLineDetails.IsChecked = enabled;
        BtnToolbarToggleEndOfLineDetails.IsChecked = enabled;

        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab)
            {
                tab.SetEndOfLineDetailsEnabled(enabled);
            }
        }

        SetStatusMessage(enabled ? "Detalles de fin de bloque (Inlay Hints) activados." : "Detalles de fin de bloque desactivados.");
    }

    private void MenuCommentBlock_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.CommentSelection(ActiveEditor);
            SetStatusMessage("Bloque comentado exitosamente.");
        }
    }

    private void MenuUncommentBlock_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.UncommentSelection(ActiveEditor);
            SetStatusMessage("Bloque descomentado exitosamente.");
        }
    }

    private void MenuSwitchAssignment_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.SwitchAssignment(ActiveEditor);
            SetStatusMessage("Asignaciones invertidas exitosamente (A = B ? B = A).");
        }
    }

    private void MenuCaseUpper_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.ConvertCase(ActiveEditor, ProjectModel.Advanced.TextCaseMode.Upper);
            SetStatusMessage("Texto convertido a MAYÃšSCULAS.");
        }
    }

    private void MenuCaseLower_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.ConvertCase(ActiveEditor, ProjectModel.Advanced.TextCaseMode.Lower);
            SetStatusMessage("Texto convertido a minÃºsculas.");
        }
    }

    private void MenuCaseTitle_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.ConvertCase(ActiveEditor, ProjectModel.Advanced.TextCaseMode.Title);
            SetStatusMessage("Texto convertido a Tipo TÃ­tulo.");
        }
    }

    private void MenuPropertyBuilder_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null)
        {
            MessageBox.Show("Debes abrir un archivo de cÃ³digo para usar el Asistente de Propiedades.",
                "Asistente de Propiedades", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string initialProp = ActiveEditor.SelectedText.Trim();
        var dlg = new PropertyBuilderDialog(initialProp) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            InsertCodeAtCaret(dlg.GeneratedCode, "Propiedad");
        }
    }

    private void MenuProcedureBuilder_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null)
        {
            MessageBox.Show("Debes abrir un archivo de cÃ³digo para usar el Asistente de Procedimientos.",
                "Asistente de Procedimientos", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string initialName = ActiveEditor.SelectedText.Trim();
        var dlg = new ProcedureBuilderDialog(initialName) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            InsertCodeAtCaret(dlg.GeneratedCode, "Procedimiento");
        }
    }

    private void MenuInsertSnippet_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor == null)
        {
            MessageBox.Show("Debes abrir un archivo de cÃ³digo para insertar snippets.",
                "Insertar Snippet", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var cm = new ContextMenu();
        foreach (var snip in SnippetManager.GetAllSnippets())
        {
            var mi = new MenuItem
            {
                Header = $"{snip.Title}  ({snip.Shortcut})",
                ToolTip = snip.Description,
                Tag = snip
            };
            mi.Click += (_, _) => InsertSnippetAtCaret(snip);
            cm.Items.Add(mi);
        }

        if (sender is FrameworkElement fe)
        {
            cm.PlacementTarget = fe;
            cm.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
        }
        cm.IsOpen = true;
    }

    private void InsertSnippetAtCaret(SnippetDefinition snippet)
    {
        if (ActiveEditor == null || ActiveEditor.Document == null) return;

        int caretOffset = ActiveEditor.CaretOffset;
        var line = ActiveEditor.Document.GetLineByOffset(caretOffset);
        string lineTextBeforeCaret = ActiveEditor.Document.GetText(line.Offset, caretOffset - line.Offset);

        int indentLen = 0;
        while (indentLen < lineTextBeforeCaret.Length && (lineTextBeforeCaret[indentLen] == ' ' || lineTextBeforeCaret[indentLen] == '\t'))
        {
            indentLen++;
        }
        string indent = lineTextBeforeCaret[..indentLen];

        if (SnippetManager.TryExpandSnippet(snippet.Shortcut, indent, out string expanded, out int relativeCaret))
        {
            ActiveEditor.Document.Insert(caretOffset, expanded);
            ActiveEditor.CaretOffset = caretOffset + relativeCaret;
            ActiveEditor.Focus();
            SetStatusMessage($"Snippet '{snippet.Title}' insertado.");
        }
    }

    private void MenuClipboardRing_Click(object sender, RoutedEventArgs e)
    {
        OpenClipboardRingDialog();
    }

    private void OpenClipboardRingDialog()
    {
        if (ActiveEditor == null)
        {
            MessageBox.Show("Debes abrir un archivo de cÃ³digo para pegar desde el portapapeles.",
                "Portapapeles MÃºltiple", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new ClipboardRingDialog { Owner = this };
        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.SelectedText))
        {
            ActiveEditor.Document.Insert(ActiveEditor.CaretOffset, dlg.SelectedText);
            ActiveEditor.CaretOffset += dlg.SelectedText.Length;
            ActiveEditor.Focus();
            SetStatusMessage("Fragmento del portapapeles pegado.");
        }
    }

    private void InsertCodeAtCaret(string codeToInsert, string operationDesc)
    {
        if (ActiveEditor == null || ActiveEditor.Document == null || string.IsNullOrWhiteSpace(codeToInsert)) return;

        int caretOffset = ActiveEditor.CaretOffset;
        var line = ActiveEditor.Document.GetLineByOffset(caretOffset);
        string lineText = ActiveEditor.Document.GetText(line.Offset, line.Length);
        string prefix = (!string.IsNullOrWhiteSpace(lineText)) ? "\r\n\r\n" : "";
        string text = prefix + codeToInsert + "\r\n";

        ActiveEditor.Document.Insert(caretOffset, text);
        ActiveEditor.CaretOffset = caretOffset + text.Length;
        ActiveEditor.Focus();
        SetStatusMessage($"{operationDesc} insertado exitosamente.");
    }

    #region IntelliSense & Auto List Members
    private void MenuListMembers_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null && !ActiveTab.IsDesignerActive)
        {
            ActiveTab.Editor.Focus();
            ActiveTab.IntelliSenseController.ShowCompletionWindow(explicitTrigger: true);
        }
    }

    private void MenuParameterInfo_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null && !ActiveTab.IsDesignerActive)
        {
            ActiveTab.Editor.Focus();
            ActiveTab.IntelliSenseController.ShowParameterInsight(explicitTrigger: true);
        }
    }

    private void MenuCompleteWord_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveTab != null && !ActiveTab.IsDesignerActive)
        {
            ActiveTab.Editor.Focus();
            ActiveTab.IntelliSenseController.ShowCompletionWindow(explicitTrigger: true);
        }
    }
    #endregion

    #region Fase 4: Formato, OrdenaciÃ³n, Diff y EstadÃ­sticas
    private void MenuFormatDocument_Click(object sender, RoutedEventArgs e)
    {
        FormatActiveDocument();
    }

    private void MenuAlignDeclarations_Click(object sender, RoutedEventArgs e)
    {
        AlignDeclarationsInActiveDocument();
    }

    private void MenuSortProcedures_Click(object sender, RoutedEventArgs e)
    {
        SortProceduresInActiveDocument();
    }

    private void MenuShowDiff_Click(object sender, RoutedEventArgs e)
    {
        ShowFileDiffForActiveDocument();
    }

    private void MenuCodeStatistics_Click(object sender, RoutedEventArgs e)
    {
        ShowCodeStatistics();
    }

    public void FormatActiveDocument()
    {
        if (ActiveEditor == null || ActiveEditor.Document == null)
        {
            SetStatusMessage("No hay ningÃºn documento activo para formatear.");
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        var options = new CodeFormatterOptions
        {
            IndentSize = _settings.TabSize > 0 ? _settings.TabSize : 4,
            UseTabs = false,
            MaxConsecutiveBlankLines = 1,
            EnsureBlankLineBetweenProcedures = true,
            AlignDeclarations = false
        };

        string formatted = CodeFormatter.FormatCode(currentText, options);
        if (!string.Equals(currentText, formatted, StringComparison.Ordinal))
        {
            int caretOffset = Math.Min(ActiveEditor.CaretOffset, formatted.Length);
            ActiveEditor.Document.BeginUpdate();
            try
            {
                ActiveEditor.Document.Text = formatted;
            }
            finally
            {
                ActiveEditor.Document.EndUpdate();
            }
            ActiveEditor.CaretOffset = caretOffset;
            SetStatusMessage("Documento formateado correctamente.");
        }
        else
        {
            SetStatusMessage("El documento ya se encuentra correctamente formateado.");
        }
    }

    public void AlignDeclarationsInActiveDocument()
    {
        if (ActiveEditor == null || ActiveEditor.Document == null)
        {
            SetStatusMessage("No hay ningÃºn documento activo.");
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        string aligned = CodeFormatter.AlignAsInDeclarations(currentText);
        if (!string.Equals(currentText, aligned, StringComparison.Ordinal))
        {
            ActiveEditor.Document.BeginUpdate();
            try
            {
                ActiveEditor.Document.Text = aligned;
            }
            finally
            {
                ActiveEditor.Document.EndUpdate();
            }
            SetStatusMessage("Declaraciones 'As' alineadas correctamente.");
        }
        else
        {
            SetStatusMessage("No se encontraron declaraciones contiguas para alinear.");
        }
    }

    public void SortProceduresInActiveDocument()
    {
        if (ActiveEditor == null || ActiveEditor.Document == null)
        {
            MessageBox.Show(this, "Debes abrir un archivo de cÃ³digo para ordenar sus procedimientos.", "Ordenar Procedimientos", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        var dlg = new SortProceduresDialog(currentText)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.ResultingCode))
        {
            if (!string.Equals(currentText, dlg.ResultingCode, StringComparison.Ordinal))
            {
                ActiveEditor.Document.BeginUpdate();
                try
                {
                    ActiveEditor.Document.Text = dlg.ResultingCode;
                }
                finally
                {
                    ActiveEditor.Document.EndUpdate();
                }
                SetStatusMessage("Procedimientos del mÃ³dulo reordenados exitosamente.");
            }
        }
    }

    public void ShowFileDiffForActiveDocument()
    {
        if (ActiveEditor == null || ActiveTab == null)
        {
            MessageBox.Show(this, "No hay ningÃºn documento activo para comparar diferencias.", "Visual Diff", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string filePath = ActiveTab.Document?.FilePath ?? string.Empty;
        string editorText = ActiveEditor.Text;
        string diskText = string.Empty;

        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        {
            try
            {
                var doc = SafeFileStorage.LoadDocument(filePath);
                diskText = doc.Text;
            }
            catch
            {
                diskText = File.ReadAllText(filePath);
            }
        }

        var dlg = new FileDiffDialog(filePath, diskText, editorText, onSaveAction: () =>
        {
            MenuSave_Click(this, new RoutedEventArgs());
        })
        {
            Owner = this
        };

        dlg.ShowDialog();
    }

    public void ShowCodeStatistics()
    {
        ProjectStatisticsSummary summary;

        if (_currentProjectGroup != null && _currentProjectGroup.Projects.Count > 0)
        {
            summary = new ProjectStatisticsSummary
            {
                ProjectName = $"Grupo: {_currentProjectGroup.Name}"
            };
            foreach (var prj in _currentProjectGroup.Projects)
            {
                var prjSummary = ProjectStatistics.AnalyzeProject(prj);
                summary.Files.AddRange(prjSummary.Files);
            }
        }
        else if (_currentProject != null)
        {
            summary = ProjectStatistics.AnalyzeProject(_currentProject);
        }
        else if (ActiveEditor != null)
        {
            string filePath = ActiveTab?.Document?.FilePath ?? "Documento Activo";
            var fileStat = ProjectStatistics.AnalyzeText(ActiveEditor.Text, filePath);
            summary = new ProjectStatisticsSummary
            {
                ProjectName = Path.GetFileName(filePath),
                Files = new List<FileCodeStatistics> { fileStat }
            };
        }
        else
        {
            MessageBox.Show(this, "No hay ningÃºn proyecto ni documento abierto para calcular estadÃ­sticas.", "EstadÃ­sticas de CÃ³digo", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new CodeStatisticsDialog(summary)
        {
            Owner = this
        };

        dlg.ShowDialog();
    }
    #endregion

    #region Fase 5: Manejo de Errores y Trazas (Tracing)
    private void MenuErrorHandler_Click(object sender, RoutedEventArgs e)
    {
        OpenErrorHandlerDialog();
    }

    private void MenuTracing_Click(object sender, RoutedEventArgs e)
    {
        OpenTracingDialog();
    }

    private void MenuRemoveTraces_Click(object sender, RoutedEventArgs e)
    {
        RemoveAllTracesFromActiveDocument();
    }

    public void OpenErrorHandlerDialog()
    {
        if (ActiveEditor == null || ActiveTab == null)
        {
            MessageBox.Show(this, "Debes abrir un archivo de cÃ³digo para gestionar el manejo de errores.", "Manejo de Errores", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        string moduleName = Path.GetFileNameWithoutExtension(ActiveTab.Document?.FilePath ?? "") ?? "Module1";
        int caretLine = ActiveEditor.TextArea.Caret.Line;

        var dlg = new ErrorHandlerDialog(currentText, moduleName, caretLine)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.ResultingCode))
        {
            if (!string.Equals(currentText, dlg.ResultingCode, StringComparison.Ordinal))
            {
                int caretOffset = Math.Min(ActiveEditor.CaretOffset, dlg.ResultingCode.Length);
                ActiveEditor.Document.BeginUpdate();
                try
                {
                    ActiveEditor.Document.Text = dlg.ResultingCode;
                }
                finally
                {
                    ActiveEditor.Document.EndUpdate();
                }
                ActiveEditor.CaretOffset = caretOffset;

                string actionDesc = dlg.IsRemoval ? "retirado de" : "aplicado a";
                SetStatusMessage($"Manejo de errores {actionDesc} {dlg.AffectedCount} procedimiento(s).");
            }
        }
    }

    public void OpenTracingDialog()
    {
        if (ActiveEditor == null || ActiveTab == null)
        {
            MessageBox.Show(this, "Debes abrir un archivo de cÃ³digo para inyectar o retirar trazas.", "InyecciÃ³n de Trazas", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        string moduleName = Path.GetFileNameWithoutExtension(ActiveTab.Document?.FilePath ?? "") ?? "Module1";
        int caretLine = ActiveEditor.TextArea.Caret.Line;

        var dlg = new TracingDialog(currentText, moduleName, caretLine)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.ResultingCode))
        {
            if (!string.Equals(currentText, dlg.ResultingCode, StringComparison.Ordinal))
            {
                int caretOffset = Math.Min(ActiveEditor.CaretOffset, dlg.ResultingCode.Length);
                ActiveEditor.Document.BeginUpdate();
                try
                {
                    ActiveEditor.Document.Text = dlg.ResultingCode;
                }
                finally
                {
                    ActiveEditor.Document.EndUpdate();
                }
                ActiveEditor.CaretOffset = caretOffset;

                string actionDesc = dlg.IsRemoval ? "retiradas de" : "inyectadas en";
                SetStatusMessage($"Trazas {actionDesc} {dlg.AffectedCount} procedimiento(s).");
            }
        }
    }

    public void RemoveAllTracesFromActiveDocument()
    {
        if (ActiveEditor == null || ActiveTab == null)
        {
            SetStatusMessage("No hay ningÃºn documento activo.");
            return;
        }

        string currentText = ActiveEditor.Text;
        if (string.IsNullOrEmpty(currentText)) return;

        string cleaned = TracingEngine.RemoveTracesFromModule(currentText, TracingOptions.DefaultTag, out int count);
        if (!string.Equals(currentText, cleaned, StringComparison.Ordinal))
        {
            int caretOffset = Math.Min(ActiveEditor.CaretOffset, cleaned.Length);
            ActiveEditor.Document.BeginUpdate();
            try
            {
                ActiveEditor.Document.Text = cleaned;
            }
            finally
            {
                ActiveEditor.Document.EndUpdate();
            }
            ActiveEditor.CaretOffset = caretOffset;
            SetStatusMessage($"Se retiraron todas las trazas de {count} procedimiento(s).");
        }
        else
        {
            SetStatusMessage("No se detectaron trazas inyectadas en el mÃ³dulo activo.");
        }
    }
    #endregion

    #region Fase 6: DiseÃ±adores y Analizador de Calidad
    private void MenuTabOrder_Click(object sender, RoutedEventArgs e)
    {
        OpenTabOrderDialog();
    }

    public void OpenTabOrderDialog()
    {
        if (ActiveTab == null || ActiveTab.Document == null)
        {
            MessageBox.Show(this, "Abre un formulario (.frm, .ctl, .pag) para gestionar el orden de tabulaciÃ³n.", "Orden de TabulaciÃ³n", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string header = ActiveTab.Document.Header ?? string.Empty;
        if (string.IsNullOrWhiteSpace(header))
        {
            MessageBox.Show(this, "El archivo actual no posee una definiciÃ³n visual de formulario/controles o es un mÃ³dulo estÃ¡ndar (.bas / .cls).", "Orden de TabulaciÃ³n", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string formName = Path.GetFileNameWithoutExtension(ActiveTab.Document.FilePath);
        var dlg = new TabOrderDialog(header, formName)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.ResultingHeader))
        {
            if (!string.Equals(header, dlg.ResultingHeader, StringComparison.Ordinal))
            {
                ActiveTab.Document.Header = dlg.ResultingHeader;
                ActiveTab.Document.IsDirty = true;
                ActiveTab.UpdateHeaderDisplay();
                SetStatusMessage($"Orden de tabulaciÃ³n actualizado en '{formName}'.");
            }
        }
    }

    private void MenuHotkeys_Click(object sender, RoutedEventArgs e)
    {
        OpenHotkeyDialog();
    }

    public void OpenHotkeyDialog()
    {
        if (ActiveTab == null || ActiveTab.Document == null)
        {
            MessageBox.Show(this, "Abre un formulario (.frm, .ctl, .pag) para verificar las teclas rÃ¡pidas de acceso.", "Teclas RÃ¡pidas (Hotkeys)", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string header = ActiveTab.Document.Header ?? string.Empty;
        if (string.IsNullOrWhiteSpace(header))
        {
            MessageBox.Show(this, "El archivo actual no posee una definiciÃ³n visual de formulario/controles o es un mÃ³dulo estÃ¡ndar (.bas / .cls).", "Teclas RÃ¡pidas (Hotkeys)", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        string formName = Path.GetFileNameWithoutExtension(ActiveTab.Document.FilePath);
        var dlg = new HotkeyDialog(header, formName)
        {
            Owner = this
        };

        if (dlg.ShowDialog() == true && !string.IsNullOrEmpty(dlg.ResultingHeader))
        {
            if (!string.Equals(header, dlg.ResultingHeader, StringComparison.Ordinal))
            {
                ActiveTab.Document.Header = dlg.ResultingHeader;
                ActiveTab.Document.IsDirty = true;
                ActiveTab.UpdateHeaderDisplay();
                SetStatusMessage($"Teclas rÃ¡pidas actualizadas en '{formName}'.");
            }
        }
    }

    private void MenuRunAnalysis_Click(object sender, RoutedEventArgs e)
    {
        RunCodeAnalysis();
    }

    private void BtnRunAnalysis_Click(object sender, RoutedEventArgs e)
    {
        RunCodeAnalysis();
    }

    private void BtnClearAnalysis_Click(object sender, RoutedEventArgs e)
    {
        LstAnalysis.ItemsSource = null;
        TxtAnalysisStatus.Text = "0 advertencias encontradas";
        SetStatusMessage("Lista de anÃ¡lisis de calidad limpiada.");
    }

    public void RunCodeAnalysis()
    {
        var issues = new List<AnalysisIssue>();

        if (_currentProject != null)
        {
            var projectFiles = GetSearchableFiles();
            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var file in projectFiles)
            {
                if (visited.Add(file))
                {
                    string? openText = GetOpenDocumentText(file);
                    if (openText != null)
                    {
                        issues.AddRange(CodeAnalyzer.AnalyzeText(openText, file));
                    }
                    else
                    {
                        issues.AddRange(CodeAnalyzer.AnalyzeFile(file));
                    }
                }
            }
        }
        else if (ActiveTab != null && ActiveEditor != null)
        {
            issues.AddRange(CodeAnalyzer.AnalyzeText(ActiveEditor.Text, ActiveTab.Document.FilePath));
        }
        else
        {
            MessageBox.Show(this, "No hay ningÃºn proyecto ni archivo abierto para analizar.", "Analizador de CÃ³digo", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var sortedIssues = issues.OrderBy(i => i.Severity).ThenBy(i => i.FileName).ThenBy(i => i.LineNumber).ToList();
        LstAnalysis.ItemsSource = sortedIssues;
        TxtAnalysisStatus.Text = $"{sortedIssues.Count} advertencia(s) / sugerencia(s) encontradas";
        ShowBottomTab("Analysis");
        SetStatusMessage($"AnÃ¡lisis de calidad completado: {sortedIssues.Count} observaciÃ³n(es).");
    }

    private void LstAnalysis_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (LstAnalysis.SelectedItem is AnalysisIssue issue)
        {
            RecordCurrentNavigation();
            var tab = OpenFile(issue.FilePath);
            if (tab != null && issue.LineNumber > 0 && issue.LineNumber <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(issue.LineNumber);
                tab.Editor.CaretOffset = lineObj.Offset;
                tab.Editor.ScrollToLine(issue.LineNumber);
                tab.Editor.Focus();
                SetStatusMessage($"Navegando a advertencia: {issue.RuleId} en {Path.GetFileName(issue.FilePath)} (LÃ­nea {issue.LineNumber})");
            }
        }
    }
    #endregion

    #region NavegaciÃ³n e Historial (Alt+? / Alt+?)
    private void MenuNavBackward_Click(object sender, RoutedEventArgs e)
    {
        NavigateBackward();
    }

    private void MenuNavForward_Click(object sender, RoutedEventArgs e)
    {
        NavigateForward();
    }

    public void NavigateBackward()
    {
        string curFile = ActiveTab?.Document.FilePath ?? "";
        int curLine = ActiveEditor?.TextArea.Caret.Line ?? 1;
        int curCol = ActiveEditor?.TextArea.Caret.Column ?? 1;

        var target = NavigationHistory.Instance.NavigateBackward(curFile, curLine, curCol);
        if (target != null)
        {
            var tab = OpenFile(target.FilePath);
            if (tab != null && target.LineNumber > 0 && target.LineNumber <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(target.LineNumber);
                int targetOffset = Math.Min(lineObj.Offset + Math.Max(0, target.ColumnNumber - 1), lineObj.EndOffset);
                tab.Editor.CaretOffset = targetOffset;
                tab.Editor.ScrollToLine(target.LineNumber);
                tab.Editor.Focus();
                SetStatusMessage($"NavegaciÃ³n atrÃ¡s: {Path.GetFileName(target.FilePath)} (LÃ­nea {target.LineNumber})");
            }
        }
        else
        {
            SetStatusMessage("No hay mÃ¡s ubicaciones anteriores en el historial de navegaciÃ³n.");
        }
    }

    public void NavigateForward()
    {
        string curFile = ActiveTab?.Document.FilePath ?? "";
        int curLine = ActiveEditor?.TextArea.Caret.Line ?? 1;
        int curCol = ActiveEditor?.TextArea.Caret.Column ?? 1;

        var target = NavigationHistory.Instance.NavigateForward(curFile, curLine, curCol);
        if (target != null)
        {
            var tab = OpenFile(target.FilePath);
            if (tab != null && target.LineNumber > 0 && target.LineNumber <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(target.LineNumber);
                int targetOffset = Math.Min(lineObj.Offset + Math.Max(0, target.ColumnNumber - 1), lineObj.EndOffset);
                tab.Editor.CaretOffset = targetOffset;
                tab.Editor.ScrollToLine(target.LineNumber);
                tab.Editor.Focus();
                SetStatusMessage($"NavegaciÃ³n adelante: {Path.GetFileName(target.FilePath)} (LÃ­nea {target.LineNumber})");
            }
        }
        else
        {
            SetStatusMessage("No hay mÃ¡s ubicaciones posteriores en el historial de navegaciÃ³n.");
        }
    }
    #endregion

    #region Marcadores (Bookmarks)
    private void MenuToggleBookmark_Click(object sender, RoutedEventArgs e)
    {
        ToggleBookmarkAtCurrentLine();
    }

    private void MenuNextBookmark_Click(object sender, RoutedEventArgs e)
    {
        GoToNextBookmark();
    }

    private void MenuPrevBookmark_Click(object sender, RoutedEventArgs e)
    {
        GoToPreviousBookmark();
    }

    private void MenuClearBookmarks_Click(object sender, RoutedEventArgs e)
    {
        ClearAllBookmarks();
    }

    public void ToggleBookmarkAtCurrentLine()
    {
        if (ActiveTab == null) return;
        int line = ActiveTab.Editor.TextArea.Caret.Line;
        bool added = BookmarkManager.Instance.ToggleBookmark(ActiveTab.Document.FilePath, line);
        ActiveTab.BookmarkMargin?.InvalidateVisual();
        SetStatusMessage(added ? $"Marcador agregado en lÃ­nea {line}." : $"Marcador eliminado de lÃ­nea {line}.");
    }

    public void GoToNextBookmark()
    {
        string curFile = ActiveTab?.Document.FilePath ?? "";
        int curLine = ActiveEditor?.TextArea.Caret.Line ?? 0;

        var next = BookmarkManager.Instance.GetNextBookmark(curFile, curLine);
        if (next != null)
        {
            RecordCurrentNavigation();
            var tab = OpenFile(next.Value.FilePath);
            if (tab != null && next.Value.Line <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(next.Value.Line);
                tab.Editor.CaretOffset = lineObj.Offset;
                tab.Editor.ScrollToLine(next.Value.Line);
                tab.Editor.Focus();
                SetStatusMessage($"Marcador: {Path.GetFileName(next.Value.FilePath)} (LÃ­nea {next.Value.Line})");
            }
        }
        else
        {
            SetStatusMessage("No hay marcadores activos en el proyecto.");
        }
    }

    public void GoToPreviousBookmark()
    {
        string curFile = ActiveTab?.Document.FilePath ?? "";
        int curLine = ActiveEditor?.TextArea.Caret.Line ?? 0;

        var prev = BookmarkManager.Instance.GetPreviousBookmark(curFile, curLine);
        if (prev != null)
        {
            RecordCurrentNavigation();
            var tab = OpenFile(prev.Value.FilePath);
            if (tab != null && prev.Value.Line <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(prev.Value.Line);
                tab.Editor.CaretOffset = lineObj.Offset;
                tab.Editor.ScrollToLine(prev.Value.Line);
                tab.Editor.Focus();
                SetStatusMessage($"Marcador: {Path.GetFileName(prev.Value.FilePath)} (LÃ­nea {prev.Value.Line})");
            }
        }
        else
        {
            SetStatusMessage("No hay marcadores activos en el proyecto.");
        }
    }

    public void ClearAllBookmarks()
    {
        BookmarkManager.Instance.ClearAllBookmarks();
        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab)
            {
                tab.BookmarkMargin?.InvalidateVisual();
            }
        }
        SetStatusMessage("Todos los marcadores han sido borrados.");
    }
    #endregion

    #region BÃºsqueda de Referencias (Find References)
    private void MenuFindReferences_Click(object sender, RoutedEventArgs e)
    {
        FindReferencesForCurrentSymbol();
    }

    private void BtnSearchReferencesPrompt_Click(object sender, RoutedEventArgs e)
    {
        FindReferencesForCurrentSymbol();
    }

    private void BtnClearReferences_Click(object sender, RoutedEventArgs e)
    {
        LstReferences.ItemsSource = null;
        TxtReferencesStatus.Text = "BÃºsqueda limpiada.";
    }

    public void FindReferencesForCurrentSymbol()
    {
        string symbol = ActiveTab?.GetWordUnderCaret() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(symbol))
        {
            var dlg = new InputPromptDialog("Buscar todas las referencias", "Ingresa el nombre del sÃ­mbolo, rutina o variable a buscar:")
            {
                Owner = this
            };
            if (dlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(dlg.InputValue))
            {
                symbol = dlg.InputValue;
            }
            else
            {
                return;
            }
        }

        var projectFiles = GetSearchableFiles();
        if (projectFiles.Count == 0)
        {
            MessageBox.Show("Debes abrir un proyecto o archivo de cÃ³digo para buscar referencias.", "Buscar Referencias", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SetStatusMessage($"Buscando referencias de '{symbol}'...");

        var results = ReferenceFinder.FindReferences(symbol, projectFiles, path => GetOpenDocumentText(path));
        LstReferences.ItemsSource = results;
        int fileCount = results.Select(r => r.FilePath).Distinct().Count();
        TxtReferencesStatus.Text = $"Referencias de '{symbol}': {results.Count} coincidencia(s) encontrada(s) en {fileCount} archivo(s).";

        ShowBottomTab("References");
        SetStatusMessage($"Referencias de '{symbol}': {results.Count} encontrada(s).");
    }

    private void LstReferences_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (LstReferences.SelectedItem is ReferenceMatch match)
        {
            RecordCurrentNavigation();
            var tab = OpenFile(match.FilePath);
            if (tab != null && match.LineNumber > 0 && match.LineNumber <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(match.LineNumber);
                int tokenOffset = lineObj.Offset + Math.Max(0, match.ColumnNumber - 1);
                tab.Editor.CaretOffset = tokenOffset;
                tab.Editor.ScrollToLine(match.LineNumber);
                tab.Editor.Focus();
            }
        }
    }
    #endregion

    #region Explorador de Tareas (TODO Explorer)
    private List<TaskItem> _allTasks = [];

    private void BtnRefreshTasks_Click(object sender, RoutedEventArgs e)
    {
        RefreshTaskList();
    }

    private void CmbTaskFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        ApplyTaskFilter();
    }

    public void RefreshTaskList()
    {
        var projectFiles = GetSearchableFiles();
        if (projectFiles.Count == 0)
        {
            LstTasks.ItemsSource = null;
            TxtTasksStatus.Text = "No hay archivos de proyecto cargados.";
            return;
        }

        SetStatusMessage("Escaneando tareas en el proyecto...");
        _allTasks = TaskScanner.ScanTasks(projectFiles, path => GetOpenDocumentText(path));
        ApplyTaskFilter();
        SetStatusMessage($"Lista de tareas actualizada: {_allTasks.Count} tarea(s) encontrada(s).");
    }

    private void ApplyTaskFilter()
    {
        if (LstTasks == null || TxtTasksStatus == null) return;

        string filter = (CmbTaskFilter?.SelectedItem is ComboBoxItem cbi) ? (cbi.Content?.ToString() ?? "Todas") : "Todas";

        IEnumerable<TaskItem> filtered = _allTasks;
        if (filter != "Todas")
        {
            filtered = _allTasks.Where(t => string.Equals(t.Tag, filter, StringComparison.OrdinalIgnoreCase));
        }

        var list = filtered.ToList();
        LstTasks.ItemsSource = list;
        TxtTasksStatus.Text = $"{list.Count} tarea(s) visible(s) ({_allTasks.Count} en total).";
    }

    private void LstTasks_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (LstTasks.SelectedItem is TaskItem task)
        {
            RecordCurrentNavigation();
            var tab = OpenFile(task.FilePath);
            if (tab != null && task.LineNumber > 0 && task.LineNumber <= tab.Editor.Document.LineCount)
            {
                var lineObj = tab.Editor.Document.GetLineByNumber(task.LineNumber);
                tab.Editor.CaretOffset = lineObj.Offset;
                tab.Editor.ScrollToLine(task.LineNumber);
                tab.Editor.Focus();
            }
        }
    }
    #endregion

    #region Utilidades de NavegaciÃ³n y Documentos
    private List<string> GetSearchableFiles()
    {
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (_currentProjectGroup != null)
        {
            foreach (var proj in _currentProjectGroup.Projects)
            {
                foreach (var item in proj.Items)
                {
                    if (item.IsEditable && !string.IsNullOrWhiteSpace(item.AbsolutePath))
                    {
                        files.Add(item.AbsolutePath);
                    }
                }
            }
        }
        else if (_currentProject != null)
        {
            foreach (var item in _currentProject.Items)
            {
                if (item.IsEditable && !string.IsNullOrWhiteSpace(item.AbsolutePath))
                {
                    files.Add(item.AbsolutePath);
                }
            }
        }

        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab && !string.IsNullOrWhiteSpace(tab.Document.FilePath))
            {
                files.Add(tab.Document.FilePath);
            }
        }

        return files.ToList();
    }

    private string? GetOpenDocumentText(string filePath)
    {
        foreach (var item in DocumentsTabControl.Items)
        {
            if (item is EditorTabItem tab && string.Equals(Path.GetFullPath(tab.Document.FilePath), Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase))
            {
                return tab.Editor.Text;
            }
        }
        return null;
    }

    private void RecordCurrentNavigation()
    {
        if (ActiveTab != null)
        {
            int line = ActiveEditor?.TextArea.Caret.Line ?? 1;
            int col = ActiveEditor?.TextArea.Caret.Column ?? 1;
            NavigationHistory.Instance.RecordNavigation(ActiveTab.Document.FilePath, line, col);
        }
    }

    public void CommentSelection()
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.CommentSelection(ActiveEditor);
            SetStatusMessage("LÃ­neas comentadas.");
        }
    }

    public void UncommentSelection()
    {
        if (ActiveEditor != null)
        {
            EditorOperationsHelper.UncommentSelection(ActiveEditor);
            SetStatusMessage("LÃ­neas descomentadas.");
        }
    }
    #endregion
    #endregion

    private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is EditorTabItem tabItem)
        {
            CloseTab(tabItem);
        }
    }

    #region EjecuciÃ³n en Caliente (Hot-Run / F5)

    private System.Windows.Threading.DispatcherFrame? _activeDebugFrame;

    private bool HandleDebuggerPause()
    {
        string? file = _liveExecutionHost?.Context.PausedFilePath;
        int line = _liveExecutionHost?.Context.PausedLineNumber ?? 0;

        Dispatcher.Invoke(() =>
        {
            // 1. Mostrar puntero de ejecuci?n (Flecha amarilla en margen + l?nea amarilla)
            BookmarkMargin.SetExecutionPointer(file, line);

            // 2. Abrir o activar pesta?a correspondiente al archivo en pausa
            if (!string.IsNullOrEmpty(file) && File.Exists(file))
            {
                OpenFile(file);
                if (ActiveEditor != null && line > 0 && line <= ActiveEditor.Document.LineCount)
                {
                    ActiveEditor.ScrollToLine(line);
                    ActiveEditor.TextArea.Caret.Line = line;
                }
            }

            // 3. Actualizar paneles de inspecci?n (Variables Locales y Pila de Llamadas)
            RefreshLocalsWindow();
            RefreshCallStackWindow();
            ShowBottomTab("Locals");

            // 4. Actualizar botones de depuraci?n
            MenuRunStart.IsEnabled = true;
            BtnToolbarRunStart.IsEnabled = true;
            MenuRunPause.IsEnabled = false;
            BtnToolbarRunPause.IsEnabled = false;
            MenuRunStop.IsEnabled = true;
            BtnToolbarRunStop.IsEnabled = true;

            SetStatusMessage($"Depuraci?n en pausa: L?nea {line} ({Path.GetFileName(file)})");
        });

        // 5. Entrar en bucle modal para bombear mensajes de la UI mientras est? en pausa
        _activeDebugFrame = new System.Windows.Threading.DispatcherFrame();
        System.Windows.Threading.Dispatcher.PushFrame(_activeDebugFrame);

        return _liveExecutionHost?.Context.State != ExecutionState.Stopped;
    }

    private void MenuRunStart_Click(object sender, RoutedEventArgs e)
    {
        if (_activeDebugFrame != null && _liveExecutionHost != null)
        {
            // Reanudar ejecuci?n continua
            _liveExecutionHost.Context.CurrentStepMode = StepMode.None;
            _liveExecutionHost.Context.State = ExecutionState.Running;
            BookmarkMargin.ClearExecutionPointer();
            SetStatusMessage("Ejecutando...");
            MenuRunStart.IsEnabled = false;
            BtnToolbarRunStart.IsEnabled = false;
            MenuRunPause.IsEnabled = true;
            BtnToolbarRunPause.IsEnabled = true;
            _activeDebugFrame.Continue = false;
            _activeDebugFrame = null;
            return;
        }

        StartHotRun();
    }

    private void MenuRunPause_Click(object sender, RoutedEventArgs e)
    {
        if (_liveExecutionHost != null)
        {
            _liveExecutionHost.Context.State = ExecutionState.Paused;
            SetStatusMessage("Ejecuci?n en pausa (Interrupci?n)");
            MenuRunStart.IsEnabled = true;
            BtnToolbarRunStart.IsEnabled = true;
            MenuRunPause.IsEnabled = false;
            BtnToolbarRunPause.IsEnabled = false;
        }
    }

    private void MenuRunStop_Click(object sender, RoutedEventArgs e)
    {
        if (_activeDebugFrame != null && _liveExecutionHost != null)
        {
            _liveExecutionHost.Context.State = ExecutionState.Stopped;
            BookmarkMargin.ClearExecutionPointer();
            _activeDebugFrame.Continue = false;
            _activeDebugFrame = null;
        }

        StopHotRun();
    }

    private void MenuStepInto_Click(object sender, RoutedEventArgs e)
    {
        if (_activeDebugFrame != null && _liveExecutionHost != null)
        {
            _liveExecutionHost.Context.CurrentStepMode = StepMode.StepInto;
            _liveExecutionHost.Context.State = ExecutionState.Running;
            BookmarkMargin.ClearExecutionPointer();
            SetStatusMessage("Paso a paso por instrucciones (Step Into)");
            _activeDebugFrame.Continue = false;
            _activeDebugFrame = null;
            return;
        }

        if (_liveExecutionHost == null)
        {
            StartHotRun(StepMode.StepInto);
        }
        else
        {
            _liveExecutionHost.Context.CurrentStepMode = StepMode.StepInto;
            _liveExecutionHost.Context.State = ExecutionState.Running;
        }
    }

    private void MenuStepOver_Click(object sender, RoutedEventArgs e)
    {
        if (_activeDebugFrame != null && _liveExecutionHost != null)
        {
            _liveExecutionHost.Context.CurrentStepMode = StepMode.StepOver;
            _liveExecutionHost.Context.StepOverFrameDepth = _liveExecutionHost.Context.CallStack.Count;
            _liveExecutionHost.Context.State = ExecutionState.Running;
            BookmarkMargin.ClearExecutionPointer();
            SetStatusMessage("Paso a paso por procedimientos (Step Over)");
            _activeDebugFrame.Continue = false;
            _activeDebugFrame = null;
            return;
        }

        if (_liveExecutionHost == null)
        {
            StartHotRun(StepMode.StepOver);
        }
        else
        {
            _liveExecutionHost.Context.CurrentStepMode = StepMode.StepOver;
            _liveExecutionHost.Context.StepOverFrameDepth = _liveExecutionHost.Context.CallStack.Count;
            _liveExecutionHost.Context.State = ExecutionState.Running;
        }
    }

    private void MenuToggleBreakpoint_Click(object sender, RoutedEventArgs e)
    {
        if (ActiveEditor != null && ActiveTab?.Document?.FilePath != null)
        {
            int line = ActiveEditor.TextArea.Caret.Line;
            string file = ActiveTab.Document.FilePath;
            if (_liveExecutionHost == null)
            {
                _liveExecutionHost = new LiveExecutionHost();
                MainImmediateWindowControl?.AttachHost(_liveExecutionHost);
            }

            bool added = BreakpointManager.Instance.ToggleBreakpoint(file, line);
            if (added)
            {
                _liveExecutionHost.Context.Breakpoints.Add((file, line));
                SetStatusMessage($"Punto de interrupci?n establecido en l?nea {line} ({Path.GetFileName(file)})");
            }
            else
            {
                _liveExecutionHost.Context.Breakpoints.Remove((file, line));
                SetStatusMessage($"Punto de interrupci?n eliminado en l?nea {line} ({Path.GetFileName(file)})");
            }
        }
    }

    private void MenuClearAllBreakpoints_Click(object sender, RoutedEventArgs e)
    {
        BreakpointManager.Instance.ClearAll();
        _liveExecutionHost?.Context.Breakpoints.Clear();
        SetStatusMessage("Todos los puntos de interrupci?n eliminados.");
    }
    public void StartHotRun(StepMode initialStepMode = StepMode.None)
    {
        try
        {
            var primaryProject = _currentProjectGroup?.PrimaryProject ?? _currentProject;

            // 1. Determinar el formulario objetivo para ejecutar
            FormDesignModel? targetFormModel = null;
            string? targetFormCode = null;

            // OpciÃ³n A: La pestaÃ±a activa tiene un diseÃ±ador de formulario abierto Y es un Formulario (.frm)
            // (Si es un UserControl .ctl, no es ejecutable de forma independiente; debe ejecutarse el formulario de inicio del proyecto principal)
            bool isActiveTabRunnableForm = ActiveTab?.DesignModel != null &&
                                          ActiveTab.Document.FilePath?.EndsWith(".frm", StringComparison.OrdinalIgnoreCase) == true;

            if (isActiveTabRunnableForm && ActiveTab != null)
            {
                bool belongsToPrimary = _currentProjectGroup == null ||
                                        primaryProject == null ||
                                        primaryProject.Forms.Any(f => string.Equals(f.AbsolutePath, ActiveTab.Document.FilePath, StringComparison.OrdinalIgnoreCase));

                if (belongsToPrimary)
                {
                    targetFormModel = ActiveTab.DesignModel;
                    targetFormCode = ActiveTab.Document.Text;
                }
            }

            // OpciÃ³n B: Si no, buscar en el proyecto principal su formulario de inicio o el primer formulario
            if (targetFormModel == null && primaryProject != null)
            {
                var startupFormRef = primaryProject.Startup;
                var formItem = !string.IsNullOrEmpty(startupFormRef)
                    ? primaryProject.Forms.FirstOrDefault(f => f.Name.Equals(startupFormRef, StringComparison.OrdinalIgnoreCase) || Path.GetFileNameWithoutExtension(f.RelativePath).Equals(startupFormRef, StringComparison.OrdinalIgnoreCase))
                    : primaryProject.Forms.FirstOrDefault();

                if (formItem != null && !string.IsNullOrEmpty(formItem.AbsolutePath) && File.Exists(formItem.AbsolutePath))
                {
                    var doc = SafeFileStorage.LoadDocument(formItem.AbsolutePath);
                    if (!string.IsNullOrEmpty(doc.Header))
                    {
                        targetFormModel = FormDesignParser.Parse(doc.Header);
                        string? basePath = Path.GetDirectoryName(formItem.AbsolutePath);
                        if (!string.IsNullOrEmpty(basePath))
                        {
                            FormDesignParser.ResolveControlPictures(targetFormModel, basePath);
                        }
                    }
                    targetFormCode = doc.Text;
                }
            }

            // OpciÃ³n C: Buscar cualquier pestaÃ±a abierta que tenga diseÃ±ador de formulario .frm
            if (targetFormModel == null)
            {
                foreach (var item in DocumentsTabControl.Items)
                {
                    if (item is EditorTabItem tab && tab.DesignModel != null && tab.Document.FilePath?.EndsWith(".frm", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        targetFormModel = tab.DesignModel;
                        targetFormCode = tab.Document.Text;
                        string? tabDir = Path.GetDirectoryName(tab.Document.FilePath);
                        if (!string.IsNullOrEmpty(tabDir))
                        {
                            FormDesignParser.ResolveControlPictures(targetFormModel, tabDir);
                        }
                        break;
                    }
                }
            }

            if (targetFormModel == null)
            {
                MessageBox.Show(
                    "No se encontrÃ³ ningÃºn formulario para iniciar la ejecuciÃ³n en caliente.\nAbra un formulario o agregue uno al proyecto principal.",
                    "EjecuciÃ³n en Caliente",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            // 2. Preparar el LiveExecutionHost
            if (_liveExecutionHost == null)
            {
                _liveExecutionHost = new LiveExecutionHost();
            }
            else
            {
                _liveExecutionHost.Reset();
            }

            MainImmediateWindowControl?.AttachHost(_liveExecutionHost);

            // Configurar depurador interactivo
            _liveExecutionHost.Context.PauseHook = HandleDebuggerPause;
            _liveExecutionHost.Context.CurrentStepMode = initialStepMode;
            _liveExecutionHost.Context.Breakpoints.Clear();
            foreach (var bp in BreakpointManager.Instance.GetAllBreakpoints())
            {
                _liveExecutionHost.Context.Breakpoints.Add(bp);
            }
            if (primaryProject != null)
            {
                _liveExecutionHost.Context.Dialect = primaryProject.Dialect;
            }

            // 3. Si estamos en un grupo (.vbg), cargar primero los proyectos secundarios como subprocesos / componentes
            if (_currentProjectGroup != null)
            {
                foreach (var secProj in _currentProjectGroup.SecondaryProjects)
                {
                    // MÃ³dulos (.bas) del subproceso
                    foreach (var mod in secProj.Modules)
                    {
                        if (!string.IsNullOrEmpty(mod.AbsolutePath) && File.Exists(mod.AbsolutePath))
                        {
                            string modCode = File.ReadAllText(mod.AbsolutePath);
                            _liveExecutionHost.LoadSourceCode(modCode, mod.Name);
                        }
                    }

                    // Clases (.cls) del subproceso
                    foreach (var cls in secProj.Classes)
                    {
                        if (!string.IsNullOrEmpty(cls.AbsolutePath) && File.Exists(cls.AbsolutePath))
                        {
                            var doc = SafeFileStorage.LoadDocument(cls.AbsolutePath);
                            _liveExecutionHost.LoadSourceCode(doc.Text, cls.Name);
                        }
                    }

                    // UserControls (.ctl) del subproceso
                    foreach (var uc in secProj.UserControls)
                    {
                        if (!string.IsNullOrEmpty(uc.AbsolutePath) && File.Exists(uc.AbsolutePath))
                        {
                            var doc = SafeFileStorage.LoadDocument(uc.AbsolutePath);
                            _liveExecutionHost.LoadSourceCode(doc.Text, uc.Name);
                        }
                    }
                }
            }

            // MÃ³dulos, clases y UserControls del proyecto principal
            if (primaryProject != null)
            {
                foreach (var mod in primaryProject.Modules)
                {
                    if (!string.IsNullOrEmpty(mod.AbsolutePath) && File.Exists(mod.AbsolutePath))
                    {
                        string modCode = File.ReadAllText(mod.AbsolutePath);
                        _liveExecutionHost.LoadSourceCode(modCode, mod.Name);
                    }
                }

                foreach (var cls in primaryProject.Classes)
                {
                    if (!string.IsNullOrEmpty(cls.AbsolutePath) && File.Exists(cls.AbsolutePath))
                    {
                        var doc = SafeFileStorage.LoadDocument(cls.AbsolutePath);
                        _liveExecutionHost.LoadSourceCode(doc.Text, cls.Name);
                    }
                }

                foreach (var uc in primaryProject.UserControls)
                {
                    if (!string.IsNullOrEmpty(uc.AbsolutePath) && File.Exists(uc.AbsolutePath))
                    {
                        var doc = SafeFileStorage.LoadDocument(uc.AbsolutePath);
                        _liveExecutionHost.LoadSourceCode(doc.Text, uc.Name);
                    }
                }
            }

            // 4. Cargar el cÃ³digo del formulario en el host
            if (!string.IsNullOrEmpty(targetFormCode))
            {
                _liveExecutionHost.LoadSourceCode(targetFormCode, targetFormModel.Name);
            }

            // 5. Configurar manejador de MsgBox y DoEvents del host para integrarlo con WPF
            _liveExecutionHost.Context.MsgBoxHandler = (prompt, buttons, title) =>
            {
                int buttonStyle = buttons & 0x07;
                int iconStyle = buttons & 0x70;

                MessageBoxImage icon = iconStyle switch
                {
                    16 => MessageBoxImage.Error,       // vbCritical
                    32 => MessageBoxImage.Question,    // vbQuestion
                    48 => MessageBoxImage.Warning,     // vbExclamation
                    64 => MessageBoxImage.Information, // vbInformation
                    _ => MessageBoxImage.None          // Sin icono en VB6 por defecto
                };

                MessageBoxButton btn = buttonStyle switch
                {
                    0 => MessageBoxButton.OK,          // vbOKOnly
                    1 => MessageBoxButton.OKCancel,    // vbOKCancel
                    2 => MessageBoxButton.OKCancel,    // vbAbortRetryIgnore (fallback estÃ¡ndar en WPF)
                    3 => MessageBoxButton.YesNoCancel, // vbYesNoCancel
                    4 => MessageBoxButton.YesNo,       // vbYesNo
                    5 => MessageBoxButton.OKCancel,    // vbRetryCancel (fallback estÃ¡ndar en WPF)
                    _ => MessageBoxButton.OK
                };

                string dialogTitle = !string.IsNullOrEmpty(title) ? title : targetFormModel.Caption;

                var mbr = MessageBox.Show(prompt, dialogTitle, btn, icon);
                return mbr switch
                {
                    MessageBoxResult.OK => 1,     // vbOK
                    MessageBoxResult.Cancel => 2, // vbCancel
                    MessageBoxResult.Yes => 6,    // vbYes
                    MessageBoxResult.No => 7,     // vbNo
                    _ => 1
                };
            };

            _liveExecutionHost.Context.DoEventsHandler = () =>
            {
                Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() => { }));
            };

            // 6. Instanciar la ventana en vivo (LiveFormWindow)
            Func<string, FormDesignModel?>? ucResolver = _currentProjectGroup != null
                ? typeName => _currentProjectGroup.GetUserControlDesignModel(typeName)
                : (primaryProject != null ? typeName => primaryProject.GetUserControlDesignModel(typeName) : null);

            _liveFormWindow = new LiveFormWindow(targetFormModel, _liveExecutionHost, targetFormCode, ucResolver)
            {
                Owner = this
            };

            _liveFormWindow.Closed += (s, ev) =>
            {
                OnLiveExecutionStopped();
            };

            // 7. Actualizar estado de la interfaz
            MenuRunStart.IsEnabled = false;
            BtnToolbarRunStart.IsEnabled = false;
            MenuRunPause.IsEnabled = true;
            BtnToolbarRunPause.IsEnabled = true;
            MenuRunStop.IsEnabled = true;
            BtnToolbarRunStop.IsEnabled = true;

            if (_currentProjectGroup != null)
            {
                string secNames = string.Join(", ", _currentProjectGroup.SecondaryProjects.Select(p => $"'{p.Name}'"));
                SetStatusMessage($"Ejecutando en caliente: Proyecto Principal '{primaryProject?.Name}' (Inicio) con subproceso {secNames}");
            }
            else
            {
                SetStatusMessage($"Ejecutando en caliente: {targetFormModel.Name}");
            }
            ShowBottomTab("Immediate");

            // 8. Mostrar la ventana y arrancar
            _liveFormWindow.Show();
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Error al iniciar ejecuciÃ³n en caliente: {ex.Message}", ex);
            MessageBox.Show($"Error al iniciar la ejecuciÃ³n en caliente:\n{ex.Message}", "Error de EjecuciÃ³n", MessageBoxButton.OK, MessageBoxImage.Error);
            OnLiveExecutionStopped();
        }
    }

    public void StopHotRun()
    {
        if (_liveFormWindow != null)
        {
            var win = _liveFormWindow;
            _liveFormWindow = null;
            try { win.Close(); } catch { }
        }

        OnLiveExecutionStopped();
    }

    private void OnLiveExecutionStopped()
    {
        BookmarkMargin.ClearExecutionPointer();
        if (_activeDebugFrame != null)
        {
            _activeDebugFrame.Continue = false;
            _activeDebugFrame = null;
        }
        if (_liveExecutionHost != null)
        {
            _liveExecutionHost.Context.State = ExecutionState.Stopped;
        }

        _liveFormWindow = null;

        Dispatcher.Invoke(() =>
        {
            MenuRunStart.IsEnabled = true;
            BtnToolbarRunStart.IsEnabled = true;
            MenuRunPause.IsEnabled = false;
            BtnToolbarRunPause.IsEnabled = false;
            MenuRunStop.IsEnabled = false;
            BtnToolbarRunStop.IsEnabled = false;

            SetStatusMessage("Listo");
        });
    }

    #endregion
}
