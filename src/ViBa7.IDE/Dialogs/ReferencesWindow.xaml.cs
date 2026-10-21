using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.IDE.Dialogs;

public class ReferenceItemViewModel : INotifyPropertyChanged
{
    private bool _isSelected;
    private string _displayTitle = string.Empty;
    private string _subtitle = string.Empty;
    private string _filePath = string.Empty;
    private bool _fileExists = true;

    public event PropertyChangedEventHandler? PropertyChanged;
    public Action? SelectionChangedCallback { get; set; }

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected != value)
            {
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                SelectionChangedCallback?.Invoke();
            }
        }
    }

    public string DisplayTitle
    {
        get => _displayTitle;
        set
        {
            if (_displayTitle != value)
            {
                _displayTitle = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayTitle)));
            }
        }
    }

    public string Subtitle
    {
        get => _subtitle;
        set
        {
            if (_subtitle != value)
            {
                _subtitle = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Subtitle)));
            }
        }
    }

    public string FilePath
    {
        get => _filePath;
        set
        {
            if (_filePath != value)
            {
                _filePath = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FilePath)));
            }
        }
    }

    public bool FileExists
    {
        get => _fileExists;
        set
        {
            if (_fileExists != value)
            {
                _fileExists = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(FileExists)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(BadgeIcon)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TitleForeground)));
            }
        }
    }

    public string Guid { get; set; } = string.Empty;
    public string MajorVersion { get; set; } = "1";
    public string MinorVersion { get; set; } = "0";
    public string Lcid { get; set; } = "0";
    public string Description { get; set; } = string.Empty;
    public string RawLine { get; set; } = string.Empty;

    public string BadgeIcon => FileExists ? "📚" : "⚠️";

    public Brush TitleForeground => FileExists
        ? new SolidColorBrush(Color.FromRgb(17, 24, 39))
        : new SolidColorBrush(Color.FromRgb(220, 38, 38));
}

/// <summary>
/// Diálogo clásico de gestión de referencias COM / TypeLib del proyecto activo (Proyecto -> Referencias...).
/// Permite seleccionar, ordenar por prioridad y buscar bibliotecas de tipos en disco y en el Registro de Windows.
/// </summary>
public partial class ReferencesWindow : Window
{
    private readonly VB6Project? _project;
    private readonly List<ReferenceItemViewModel> _allReferences = [];
    public ObservableCollection<ReferenceItemViewModel> FilteredReferences { get; } = [];

    public ReferencesWindow(VB6Project? project = null)
    {
        _project = project;
        InitializeComponent();

        if (_project != null && !string.IsNullOrWhiteSpace(_project.Name))
        {
            Title = $"Referencias - {_project.Name}";
        }

        LstReferences.ItemsSource = FilteredReferences;
        Loaded += async (s, e) => await LoadAllReferencesAsync();
    }

    private async Task LoadAllReferencesAsync()
    {
        _allReferences.Clear();

        // 1. Cargar referencias declaradas en el proyecto actual (.vbp)
        if (_project != null)
        {
            foreach (var r in _project.References)
            {
                bool exists = !string.IsNullOrEmpty(r.FilePath) && File.Exists(r.FilePath);
                string title = !string.IsNullOrWhiteSpace(r.Description)
                    ? r.Description
                    : (!string.IsNullOrWhiteSpace(r.DisplayTitle) ? r.DisplayTitle : Path.GetFileName(r.FilePath));

                string display = exists ? title : $"[FALTA] {title}";
                string sub = exists ? r.FilePath : $"(No encontrado en disco: {r.FilePath})";

                var vm = new ReferenceItemViewModel
                {
                    IsSelected = true,
                    DisplayTitle = display,
                    Subtitle = sub,
                    FilePath = r.FilePath,
                    FileExists = exists,
                    Guid = r.Guid,
                    MajorVersion = r.MajorVersion,
                    MinorVersion = r.MinorVersion,
                    Lcid = r.Lcid,
                    Description = r.Description,
                    RawLine = r.RawLine,
                    SelectionChangedCallback = UpdateSelectedCount
                };

                _allReferences.Add(vm);
            }
        }

        ApplyFilter();
        UpdateSelectedCount();

        // 2. Descubrir todas las TypeLibs y DLLs registradas en Windows en segundo plano
        if (PnlLoading != null)
        {
            PnlLoading.Visibility = Visibility.Visible;
        }

        try
        {
            var systemComponents = await Task.Run(() => SystemActiveXRegistryScanner.GetRegisteredComponents());

            foreach (var sys in systemComponents)
            {
                string normPath = sys.FilePath;
                string cleanGuid = sys.Guid.Trim('{', '}');

                // Evitar duplicados con las ya cargadas del proyecto
                bool alreadyExists = _allReferences.Any(r =>
                    (!string.IsNullOrEmpty(normPath) && string.Equals(r.FilePath, normPath, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(cleanGuid) && string.Equals(r.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase)));

                if (alreadyExists) continue;

                var vm = new ReferenceItemViewModel
                {
                    IsSelected = false,
                    DisplayTitle = sys.DisplayTitle,
                    Subtitle = $"{sys.FilePath} (v{sys.MajorVersion}.{sys.MinorVersion})",
                    FilePath = sys.FilePath,
                    FileExists = sys.FileExists,
                    Guid = sys.Guid,
                    MajorVersion = sys.MajorVersion.ToString(),
                    MinorVersion = sys.MinorVersion.ToString(),
                    Lcid = sys.Lcid,
                    Description = sys.Description,
                    SelectionChangedCallback = UpdateSelectedCount
                };

                _allReferences.Add(vm);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error al escanear TypeLibs del sistema: {ex.Message}");
        }
        finally
        {
            if (PnlLoading != null)
            {
                PnlLoading.Visibility = Visibility.Collapsed;
            }

            ApplyFilter();
            UpdateSelectedCount();

            if (FilteredReferences.Count > 0 && LstReferences.SelectedIndex < 0)
            {
                LstReferences.SelectedIndex = 0;
            }
        }
    }

    private void ApplyFilter()
    {
        string query = TxtSearch?.Text?.Trim() ?? string.Empty;

        // Separar referencias seleccionadas (manteniendo su orden de prioridad) y no seleccionadas (ordenadas alfabéticamente)
        var selectedList = _allReferences.Where(r => r.IsSelected).ToList();
        var unselectedList = _allReferences.Where(r => !r.IsSelected)
            .OrderBy(r => r.DisplayTitle, StringComparer.CurrentCultureIgnoreCase)
            .ToList();

        var combined = new List<ReferenceItemViewModel>(selectedList);
        combined.AddRange(unselectedList);

        FilteredReferences.Clear();
        foreach (var item in combined)
        {
            if (string.IsNullOrEmpty(query) ||
                item.DisplayTitle.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                item.Subtitle.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                item.FilePath.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                item.Guid.Contains(query, StringComparison.OrdinalIgnoreCase))
            {
                FilteredReferences.Add(item);
            }
        }

        UpdatePriorityButtonsState();
    }

    private void UpdateSelectedCount()
    {
        int selected = _allReferences.Count(r => r.IsSelected);
        int total = _allReferences.Count;

        if (TxtStatusSelected != null)
        {
            TxtStatusSelected.Text = $"Seleccionadas: {selected}";
        }
        if (TxtStatusTotal != null)
        {
            TxtStatusTotal.Text = $"Total: {total}";
        }

        UpdatePriorityButtonsState();
    }

    private void UpdatePriorityButtonsState()
    {
        if (BtnPriorityUp == null || BtnPriorityDown == null) return;

        if (LstReferences.SelectedItem is ReferenceItemViewModel selectedItem && selectedItem.IsSelected)
        {
            // Obtener el índice dentro de las seleccionadas
            var selectedItems = _allReferences.Where(r => r.IsSelected).ToList();
            int idx = selectedItems.IndexOf(selectedItem);

            BtnPriorityUp.IsEnabled = idx > 0;
            BtnPriorityDown.IsEnabled = idx >= 0 && idx < selectedItems.Count - 1;
        }
        else
        {
            BtnPriorityUp.IsEnabled = false;
            BtnPriorityDown.IsEnabled = false;
        }
    }

    private void LstReferences_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LstReferences.SelectedItem is ReferenceItemViewModel item)
        {
            TxtFilePath.Text = !string.IsNullOrEmpty(item.FilePath)
                ? (item.FileExists ? item.FilePath : $"[FALTA] {item.FilePath}")
                : "(Ninguna)";
            TxtFilePath.Foreground = item.FileExists
                ? new SolidColorBrush(Color.FromRgb(17, 24, 39))
                : new SolidColorBrush(Color.FromRgb(220, 38, 38));

            TxtLanguage.Text = !string.IsNullOrEmpty(item.Lcid) && item.Lcid != "0"
                ? $"LCID: {item.Lcid}"
                : "0 (Estándar)";

            TxtGuid.Text = !string.IsNullOrEmpty(item.Guid) ? item.Guid : "(Sin GUID registrado)";
            TxtVersion.Text = $"{item.MajorVersion}.{item.MinorVersion}";
        }
        else
        {
            TxtFilePath.Text = "(Ninguna)";
            TxtLanguage.Text = "-";
            TxtGuid.Text = "-";
            TxtVersion.Text = "-";
        }

        UpdatePriorityButtonsState();
    }

    private void BtnPriorityUp_Click(object sender, RoutedEventArgs e)
    {
        if (LstReferences.SelectedItem is not ReferenceItemViewModel item || !item.IsSelected) return;

        var selectedItems = _allReferences.Where(r => r.IsSelected).ToList();
        int idx = selectedItems.IndexOf(item);
        if (idx > 0)
        {
            // Intercambiar en la lista maestra _allReferences
            int masterIdxA = _allReferences.IndexOf(item);
            int masterIdxB = _allReferences.IndexOf(selectedItems[idx - 1]);

            _allReferences[masterIdxA] = selectedItems[idx - 1];
            _allReferences[masterIdxB] = item;

            ApplyFilter();
            LstReferences.SelectedItem = item;
            LstReferences.ScrollIntoView(item);
        }
    }

    private void BtnPriorityDown_Click(object sender, RoutedEventArgs e)
    {
        if (LstReferences.SelectedItem is not ReferenceItemViewModel item || !item.IsSelected) return;

        var selectedItems = _allReferences.Where(r => r.IsSelected).ToList();
        int idx = selectedItems.IndexOf(item);
        if (idx >= 0 && idx < selectedItems.Count - 1)
        {
            // Intercambiar en la lista maestra _allReferences
            int masterIdxA = _allReferences.IndexOf(item);
            int masterIdxB = _allReferences.IndexOf(selectedItems[idx + 1]);

            _allReferences[masterIdxA] = selectedItems[idx + 1];
            _allReferences[masterIdxB] = item;

            ApplyFilter();
            LstReferences.SelectedItem = item;
            LstReferences.ScrollIntoView(item);
        }
    }

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog
        {
            Title = "Examinar Biblioteca de Tipos",
            Filter = "Bibliotecas de tipos (*.olb;*.tlb;*.dll;*.ocx;*.exe)|*.olb;*.tlb;*.dll;*.ocx;*.exe|Todos los archivos (*.*)|*.*",
            CheckFileExists = true
        };

        if (ofd.ShowDialog(this) == true)
        {
            string path = ofd.FileName;

            // Intentar leer la TypeLib para obtener GUID, versión y descripción
            string guid = string.Empty;
            string major = "1";
            string minor = "0";
            string lcid = "0";
            string desc = Path.GetFileNameWithoutExtension(path);

            if (OcxTypeLibReader.TryReadTypeLib(path, out var lib, out _) && lib != null)
            {
                guid = lib.Guid;
                major = lib.MajorVersion.ToString();
                minor = lib.MinorVersion.ToString();
                lcid = "0";
                desc = !string.IsNullOrWhiteSpace(lib.Description) ? lib.Description : lib.LibraryName;
            }

            // Buscar si ya existe en la lista
            var existing = _allReferences.FirstOrDefault(r =>
                string.Equals(r.FilePath, path, StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrEmpty(guid) && string.Equals(r.Guid.Trim('{', '}'), guid.Trim('{', '}'), StringComparison.OrdinalIgnoreCase)));

            if (existing != null)
            {
                existing.IsSelected = true;
                existing.FileExists = true;
                existing.FilePath = path;
                LstReferences.SelectedItem = existing;
                LstReferences.ScrollIntoView(existing);
            }
            else
            {
                var newRef = new ReferenceItemViewModel
                {
                    IsSelected = true,
                    DisplayTitle = desc,
                    Subtitle = path,
                    FilePath = path,
                    FileExists = true,
                    Guid = guid,
                    MajorVersion = major,
                    MinorVersion = minor,
                    Lcid = lcid,
                    Description = desc,
                    SelectionChangedCallback = UpdateSelectedCount
                };

                // Insertar al inicio de las seleccionadas
                int firstUnselected = _allReferences.FindIndex(r => !r.IsSelected);
                if (firstUnselected >= 0)
                {
                    _allReferences.Insert(firstUnselected, newRef);
                }
                else
                {
                    _allReferences.Add(newRef);
                }

                ApplyFilter();
                LstReferences.SelectedItem = newRef;
                LstReferences.ScrollIntoView(newRef);
            }

            UpdateSelectedCount();
        }
    }

    private void BtnClearSearch_Click(object sender, RoutedEventArgs e)
    {
        TxtSearch.Clear();
        TxtSearch.Focus();
    }

    private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        ApplyFilter();
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        if (_project != null)
        {
            // Obtener todas las referencias marcadas en su orden de prioridad actual
            var selectedVms = _allReferences.Where(r => r.IsSelected).ToList();
            var newRefs = new List<ProjectReference>();

            foreach (var vm in selectedVms)
            {
                string guidFormatted = vm.Guid;
                if (!string.IsNullOrEmpty(guidFormatted) && !guidFormatted.StartsWith('{'))
                {
                    guidFormatted = "{" + guidFormatted + "}";
                }

                string rawLine = !string.IsNullOrEmpty(vm.RawLine)
                    ? vm.RawLine
                    : $"Reference=*\\G{guidFormatted}#{vm.MajorVersion}.{vm.MinorVersion}#{vm.Lcid}#{vm.FilePath}#{vm.Description}";

                newRefs.Add(new ProjectReference
                {
                    Guid = guidFormatted,
                    MajorVersion = vm.MajorVersion,
                    MinorVersion = vm.MinorVersion,
                    Lcid = vm.Lcid,
                    FilePath = vm.FilePath,
                    Description = vm.Description,
                    RawLine = rawLine
                });
            }

            _project.SetReferences(newRefs);
        }

        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
