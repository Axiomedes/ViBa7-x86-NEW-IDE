using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;
using System.Windows.Media;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Resources;

namespace ViBa7.IDE.Designer;

public class PropertyTemplateSelector : DataTemplateSelector
{
    public DataTemplate? TextTemplate { get; set; }
    public DataTemplate? ChoiceTemplate { get; set; }
    public DataTemplate? ColorTemplate { get; set; }
    public DataTemplate? PictureTemplate { get; set; }
    public DataTemplate? CustomDialogTemplate { get; set; }
    public DataTemplate? FontTemplate { get; set; }

    public override DataTemplate? SelectTemplate(object item, DependencyObject container)
    {
        if (item is PropertyItemViewModel prop)
        {
            if (prop.Type == DesignerPropertyType.FontDescription)
                return FontTemplate;
            if (prop.Type == DesignerPropertyType.CustomDialog)
                return CustomDialogTemplate;
            if (prop.Type == DesignerPropertyType.Picture)
                return PictureTemplate;
            if (prop.Type == DesignerPropertyType.ColorHex)
                return ColorTemplate;
            if (prop.Choices.Count > 0)
                return ChoiceTemplate;
            return TextTemplate;
        }
        return base.SelectTemplate(item, container);
    }
}

public class PropertyItemViewModel : INotifyPropertyChanged
{
    private readonly DesignerPropertyItem _model;
    private readonly Action _onModelChanged;

    public DesignerPropertyItem Model => _model;
    public string Name => _model.Name;
    public string DisplayName => _model.DisplayName;
    public string Category => _model.Category;
    public string Description => _model.Description;
    public DesignerPropertyType Type => _model.Type;
    public List<DesignerPropertyChoice> Choices => _model.Choices;
    public bool IsReadOnly => _model.IsReadOnly;

    public string CurrentValue
    {
        get => _model.CurrentValue;
        set
        {
            if (!string.Equals(_model.CurrentValue, value, StringComparison.Ordinal))
            {
                if (_model.SetValue(value))
                {
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentValue)));
                    if (Type == DesignerPropertyType.ColorHex)
                    {
                        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ColorBrush)));
                    }
                }
            }
        }
    }

    public Brush ColorBrush
    {
        get
        {
            if (Type == DesignerPropertyType.ColorHex)
            {
                return VbColorHelper.GetBrushFromHex(CurrentValue, VbColorModel.DefaultControlColor);
            }
            return Brushes.Transparent;
        }
    }

    public PropertyItemViewModel(DesignerPropertyItem model, Action onModelChanged)
    {
        _model = model;
        _onModelChanged = onModelChanged;
    }

    public void Refresh()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentValue)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ColorBrush)));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
}

public class ObjectSelectorItem
{
    public string DisplayName { get; init; } = string.Empty;
    public object Target { get; init; } = null!;
}

public partial class PropertyGridPanel : UserControl
{
    private FormDesignModel? _model;
    private object? _currentSelectedTarget;
    private readonly List<PropertyItemViewModel> _allProperties = new();
    private bool _isUpdatingSelectionInternal = false;

    public event Action<FormDesignerControl?>? ObjectSelected;
    public event Action? ModelChanged;
    public event Action<FormDesignerControl, string, string>? ControlRenamed;

    public string? CurrentFormFilePath { get; set; }
    public Func<FormDesignerControl, FrameworkElement?>? VisualResolver { get; set; }

    public PropertyGridPanel()
    {
        InitializeComponent();
    }

    public void BindModel(FormDesignModel model, FormDesignerControl? selectedControl = null, string? formFilePath = null)
    {
        _model = model;
        CurrentFormFilePath = formFilePath;
        RefreshObjectList();

        if (selectedControl != null)
        {
            SelectControl(selectedControl);
        }
        else
        {
            SelectForm();
        }
    }

    public void ClearModel()
    {
        _model = null;
        _currentSelectedTarget = null;
        _allProperties.Clear();

        _isUpdatingSelectionInternal = true;
        if (CmbObjects != null)
        {
            CmbObjects.ItemsSource = null;
        }
        _isUpdatingSelectionInternal = false;

        if (PropertiesListView != null)
        {
            PropertiesListView.ItemsSource = null;
        }

        if (TxtHeaderTitle != null)
        {
            TxtHeaderTitle.Text = "Propiedades";
        }

        if (TxtHelpPropName != null)
        {
            TxtHelpPropName.Text = "Propiedad";
        }

        if (TxtHelpDescription != null)
        {
            TxtHelpDescription.Text = "No hay ningún objeto o formulario activo.";
        }
    }

    public void RefreshObjectList()
    {
        if (_model == null || CmbObjects == null) return;

        var items = new List<ObjectSelectorItem>
        {
            new ObjectSelectorItem
            {
                DisplayName = $"{_model.Name} (Form)",
                Target = _model
            }
        };

        foreach (var ctrl in _model.GetAllControls().OrderBy(c => c.Name).ThenBy(c => c.Index ?? 0))
        {
            items.Add(new ObjectSelectorItem
            {
                DisplayName = $"{ctrl.DisplayNameWithIndex} ({ctrl.DisplayType})",
                Target = ctrl
            });
        }

        _isUpdatingSelectionInternal = true;
        CmbObjects.ItemsSource = items;

        if (_currentSelectedTarget is FormDesignerControl ctrlTarget)
        {
            CmbObjects.SelectedItem = items.FirstOrDefault(i => i.Target == ctrlTarget) ?? items.FirstOrDefault();
        }
        else
        {
            CmbObjects.SelectedItem = items.FirstOrDefault();
        }
        _isUpdatingSelectionInternal = false;
    }

    public void SelectForm()
    {
        if (_model == null || CmbObjects == null) return;
        _currentSelectedTarget = _model;

        _isUpdatingSelectionInternal = true;
        if (CmbObjects.ItemsSource is IEnumerable<ObjectSelectorItem> items)
        {
            CmbObjects.SelectedItem = items.FirstOrDefault(i => i.Target == _model);
        }
        _isUpdatingSelectionInternal = false;

        if (TxtHeaderTitle != null)
        {
            TxtHeaderTitle.Text = $"Propiedades - {_model.Name}";
        }
        LoadPropertiesForTarget(_model);
    }

    public void SelectControl(FormDesignerControl? control)
    {
        if (control == null)
        {
            SelectForm();
            return;
        }

        if (CmbObjects == null) return;

        _currentSelectedTarget = control;

        _isUpdatingSelectionInternal = true;
        if (CmbObjects.ItemsSource is IEnumerable<ObjectSelectorItem> items)
        {
            CmbObjects.SelectedItem = items.FirstOrDefault(i => i.Target == control);
        }
        _isUpdatingSelectionInternal = false;

        if (TxtHeaderTitle != null)
        {
            TxtHeaderTitle.Text = $"Propiedades - {control.Name}";
        }
        LoadPropertiesForTarget(control);
    }

    public void RefreshCurrentProperties()
    {
        foreach (var prop in _allProperties)
        {
            prop.Refresh();
        }
    }

    private void LoadPropertiesForTarget(object target)
    {
        if (_model == null) return;

        _allProperties.Clear();

        if (target is FormDesignModel formModel)
        {
            var raw = DesignerPropertyFactory.GetPropertiesForForm(formModel, OnInnerModelChanged);
            foreach (var p in raw)
            {
                _allProperties.Add(new PropertyItemViewModel(p, OnInnerModelChanged));
            }
        }
        else if (target is FormDesignerControl ctrl)
        {
            var raw = DesignerPropertyFactory.GetPropertiesForControl(
                ctrl,
                _model,
                OnInnerModelChanged,
                (c, oldName, newName) =>
                {
                    RefreshObjectList();
                    ControlRenamed?.Invoke(c, oldName, newName);
                },
                name =>
                {
                    var res = MessageBox.Show(
                        $"Ya tiene un control llamado '{name}'. ¿Desea crear una matriz de controles?",
                        "Microsoft Visual Basic",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    return res == MessageBoxResult.Yes;
                },
                msg =>
                {
                    MessageBox.Show(msg, "Microsoft Visual Basic", MessageBoxButton.OK, MessageBoxImage.Warning);
                });

            foreach (var p in raw)
            {
                _allProperties.Add(new PropertyItemViewModel(p, OnInnerModelChanged));
            }
        }

        ApplyFilterAndSorting();
    }

    private void OnInnerModelChanged()
    {
        if (_currentSelectedTarget is FormDesignerControl ctrl)
        {
            var visual = VisualResolver?.Invoke(ctrl);
            ViBa7.ActiveXHost.ActiveXInPlaceHost? axHost = visual as ViBa7.ActiveXHost.ActiveXInPlaceHost
                ?? (visual is Grid g && g.Children.Count > 0 ? g.Children[0] as ViBa7.ActiveXHost.ActiveXInPlaceHost : null);

            axHost?.SyncFromDesignerModel();
        }

        ModelChanged?.Invoke();
    }

    private void ApplyFilterAndSorting()
    {
        if (TxtFilter == null || RadioCategorized == null || PropertiesListView == null)
        {
            return;
        }

        IEnumerable<PropertyItemViewModel> query = _allProperties;

        string filter = TxtFilter.Text.Trim();
        if (!string.IsNullOrEmpty(filter))
        {
            query = query.Where(p =>
                p.DisplayName.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                p.Category.Contains(filter, StringComparison.OrdinalIgnoreCase)
            );
        }

        if (RadioCategorized.IsChecked == true)
        {
            query = query.OrderBy(p => p.Category).ThenBy(p => p.DisplayName);
        }
        else
        {
            query = query.OrderBy(p => p.DisplayName.StartsWith("(") ? 0 : 1).ThenBy(p => p.DisplayName);
        }

        PropertiesListView.ItemsSource = query.ToList();
    }

    private void CmbObjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isUpdatingSelectionInternal) return;
        if (CmbObjects == null || TxtHeaderTitle == null) return;

        if (CmbObjects.SelectedItem is ObjectSelectorItem item)
        {
            _currentSelectedTarget = item.Target;
            if (item.Target is FormDesignModel)
            {
                TxtHeaderTitle.Text = $"Propiedades - {_model?.Name}";
                LoadPropertiesForTarget(item.Target);
                ObjectSelected?.Invoke(null);
            }
            else if (item.Target is FormDesignerControl ctrl)
            {
                TxtHeaderTitle.Text = $"Propiedades - {ctrl.Name}";
                LoadPropertiesForTarget(ctrl);
                ObjectSelected?.Invoke(ctrl);
            }
        }
    }

    private void RadioViewMode_Checked(object sender, RoutedEventArgs e)
    {
        ApplyFilterAndSorting();
    }

    private void TxtFilter_TextChanged(object sender, TextChangedEventArgs e)
    {
        ApplyFilterAndSorting();
    }

    private void PropertiesListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (PropertiesListView == null || TxtHelpPropName == null || TxtHelpDescription == null) return;

        if (PropertiesListView.SelectedItem is PropertyItemViewModel prop)
        {
            TxtHelpPropName.Text = prop.DisplayName;
            TxtHelpDescription.Text = prop.Description;
        }
        else
        {
            TxtHelpPropName.Text = "Propiedad";
            TxtHelpDescription.Text = "Selecciona una propiedad para ver su descripción.";
        }
    }

    private void PropertyTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && sender is TextBox tb)
        {
            var binding = tb.GetBindingExpression(TextBox.TextProperty);
            binding?.UpdateSource();
            e.Handled = true;
        }
    }

    private void PropertyTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is TextBox tb)
        {
            var binding = tb.GetBindingExpression(TextBox.TextProperty);
            binding?.UpdateSource();
        }
    }

    private void PropertyChoice_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox cb && cb.IsDropDownOpen)
        {
            var binding = cb.GetBindingExpression(ComboBox.SelectedValueProperty);
            binding?.UpdateSource();
        }
    }

    private void BtnEyedropper_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is PropertyItemViewModel prop)
        {
            ScreenColorDropper.StartColorPick(color =>
            {
                string vbHex = VbColorHelper.ToVbColorHex(color);
                prop.CurrentValue = vbHex;
            });
        }
    }

    private void PropertyName_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe)
        {
            if (fe.DataContext is PropertyItemViewModel prop && PropertiesListView != null)
            {
                PropertiesListView.SelectedItem = prop;
            }

            DependencyObject? parent = VisualTreeHelper.GetParent(fe);
            while (parent != null && parent is not Grid)
            {
                parent = VisualTreeHelper.GetParent(parent);
            }

            if (parent is Grid grid)
            {
                var tb = FindVisualChild<TextBox>(grid);
                if (tb != null && !tb.IsReadOnly)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        tb.Focus();
                        tb.SelectAll();
                    }), System.Windows.Threading.DispatcherPriority.Input);
                    e.Handled = true;
                    return;
                }

                var cb = FindVisualChild<ComboBox>(grid);
                if (cb != null)
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        cb.Focus();
                    }), System.Windows.Threading.DispatcherPriority.Input);
                    e.Handled = true;
                    return;
                }
            }
        }
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T typedChild)
            {
                return typedChild;
            }
            var result = FindVisualChild<T>(child);
            if (result != null)
            {
                return result;
            }
        }
        return null;
    }

    private void BtnPickColor_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is PropertyItemViewModel prop)
        {
            ShowColorPaletteMenu(btn, prop);
        }
    }

    private void ShowColorPaletteMenu(Button anchor, PropertyItemViewModel prop)
    {
        var menu = new ContextMenu();

        var colors = new (string Name, string Hex)[]
        {
            // Colores del sistema
            ("ButtonFace (&H8000000F&)", "&H8000000F&"),
            ("Window Blanco (&H80000005&)", "&H80000005&"),
            ("WindowText Negro (&H80000008&)", "&H80000008&"),
            ("ButtonHighlight (&H80000014&)", "&H80000014&"),
            ("ButtonShadow (&H80000010&)", "&H80000010&"),
            ("Highlight Azul (&H8000000D&)", "&H8000000D&"),
            // Colores RGB Clásicos VB6
            ("Negro (vbBlack)", "&H00000000&"),
            ("Blanco (vbWhite)", "&H00FFFFFF&"),
            ("Rojo (vbRed)", "&H000000FF&"),
            ("Verde (vbGreen)", "&H0000FF00&"),
            ("Azul (vbBlue)", "&H00FF0000&"),
            ("Amarillo (vbYellow)", "&H0000FFFF&"),
            ("Cian (vbCyan)", "&H00FFFF00&"),
            ("Magenta (vbMagenta)", "&H00FF00FF&"),
            ("Gris (vbGray)", "&H00808080&")
        };

        foreach (var (name, hex) in colors)
        {
            var mi = new MenuItem
            {
                Header = name,
                Icon = new Border
                {
                    Width = 14,
                    Height = 14,
                    Background = VbColorHelper.GetBrushFromHex(hex, VbColorModel.DefaultControlColor),
                    BorderBrush = Brushes.Gray,
                    BorderThickness = new Thickness(1)
                }
            };
            mi.Click += (_, _) =>
            {
                prop.CurrentValue = hex;
            };
            menu.Items.Add(mi);
        }

        menu.PlacementTarget = anchor;
        menu.IsOpen = true;
    }

    private void BtnBrowsePicture_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is PropertyItemViewModel prop)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Cargar Imagen (PictureBox / CommandButton / Form)",
                Filter = "Todos los formatos de imagen (*.png;*.webp;*.jpg;*.jpeg;*.bmp;*.ico;*.gif)|*.png;*.webp;*.jpg;*.jpeg;*.bmp;*.ico;*.gif|PNG (*.png)|*.png|WebP (*.webp)|*.webp|JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Mapas de bits (*.bmp)|*.bmp|Iconos (*.ico)|*.ico|Todos los archivos (*.*)|*.*"
            };

            if (ofd.ShowDialog() == true)
            {
                try
                {
                    byte[] imageBytes = File.ReadAllBytes(ofd.FileName);

                    // Determinar archivo .frx correspondiente
                    string frxPath;
                    string? formPath = CurrentFormFilePath;
                    if (string.IsNullOrWhiteSpace(formPath) && Application.Current?.MainWindow is MainWindow mw)
                    {
                        formPath = mw.ActiveTab?.Document?.FilePath;
                    }

                    if (!string.IsNullOrEmpty(formPath))
                    {
                        string fullFormPath = Path.GetFullPath(formPath);
                        frxPath = Path.ChangeExtension(fullFormPath, ".frx");
                    }
                    else
                    {
                        string dir = Directory.GetCurrentDirectory();
                        frxPath = Path.Combine(dir, $"{_model?.Name ?? "Form1"}.frx");
                    }

                    int offset = FrxResourceFile.AppendResource(frxPath, imageBytes);
                    string frxFileName = Path.GetFileName(frxPath);
                    string reference = FrxResourceFile.FormatReference(frxFileName, offset);

                    if (_currentSelectedTarget is FormDesignerControl ctrl)
                    {
                        ctrl.Picture = reference;
                        ctrl.PictureData = imageBytes;

                        // Si es CommandButton, activar estilo Graphical
                        if (ctrl.TypeName.EndsWith("Button", StringComparison.OrdinalIgnoreCase))
                        {
                            ctrl.Style = 1;
                        }

                        // Si es PictureBox con AutoSize, ajustar tamaño
                        if (ctrl.AutoSize)
                        {
                            var bmp = VbImageHelper.LoadBitmapFromBytes(imageBytes);
                            if (bmp != null)
                            {
                                ctrl.Width = (int)(bmp.PixelWidth * FormDesignerControl.TwipsPerPixel);
                                ctrl.Height = (int)(bmp.PixelHeight * FormDesignerControl.TwipsPerPixel);
                            }
                        }
                    }
                    else if (_currentSelectedTarget is FormDesignModel formModel)
                    {
                        formModel.Picture = reference;
                        formModel.PictureData = imageBytes;
                    }

                    prop.Model.CurrentValue = FrxResourceFile.GetFormatDescription(imageBytes, reference);
                    OnInnerModelChanged();
                    prop.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al cargar la imagen: {ex.Message}", "Error de Imagen", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    private void BtnClearPicture_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is PropertyItemViewModel prop)
        {
            ClearPictureProperty(prop);
        }
    }

    private void PropertyPicture_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Delete && sender is TextBox tb && tb.DataContext is PropertyItemViewModel prop)
        {
            ClearPictureProperty(prop);
            e.Handled = true;
        }
    }

    private void ClearPictureProperty(PropertyItemViewModel prop)
    {
        if (_currentSelectedTarget is FormDesignerControl ctrl)
        {
            ctrl.Picture = null;
            ctrl.PictureData = null;
        }
        else if (_currentSelectedTarget is FormDesignModel formModel)
        {
            formModel.Picture = null;
            formModel.PictureData = null;
        }

        prop.Model.CurrentValue = "(None)";
        OnInnerModelChanged();
        prop.Refresh();
    }

    private void BtnOpenPropertyPages_Click(object sender, RoutedEventArgs e)
    {
        OpenPropertyPagesForCurrentSelection();
    }

    private void CustomPropertyText_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        OpenPropertyPagesForCurrentSelection();
    }

    public void OpenPropertyPagesForCurrentSelection()
    {
        if (_currentSelectedTarget is FormDesignerControl ctrl)
        {
            OpenPropertyPages(ctrl);
        }
    }

    public void OpenPropertyPages(FormDesignerControl ctrl)
    {
        if (ctrl == null) return;

        // Si el control está hospedado In-Place como ActiveX, intentar invocar sus hojas de propiedades OLE nativas
        ViBa7.ActiveXHost.ActiveXInPlaceHost? axHost = null;
        var vis = VisualResolver?.Invoke(ctrl);
        if (vis is ViBa7.ActiveXHost.ActiveXInPlaceHost directHost)
        {
            axHost = directHost;
        }
        else if (vis is Grid g && g.Children.Count > 0 && g.Children[0] is ViBa7.ActiveXHost.ActiveXInPlaceHost gridHost)
        {
            axHost = gridHost;
        }

        if (axHost != null)
        {
            if (axHost.ShowPropertyPages(Window.GetWindow(this)))
            {
                axHost.SyncToDesignerModel();
                LoadPropertiesForTarget(ctrl);
                OnInnerModelChanged();
                return;
            }
        }

        var dlg = new Dialogs.OcxPropertyPagesWindow(ctrl, _model, () =>
        {
            OnInnerModelChanged();
            RefreshCurrentProperties();
        });
        dlg.Owner = Window.GetWindow(this);
        dlg.ShowDialog();
    }

    private void BtnBrowseFont_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.DataContext is PropertyItemViewModel prop)
        {
            OpenFontPickerDialog(prop);
        }
    }

    private void PropertyFont_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is PropertyItemViewModel prop)
        {
            OpenFontPickerDialog(prop);
        }
    }

    private void OpenFontPickerDialog(PropertyItemViewModel prop)
    {
        string propName = prop.Model.Name;
        FormFontDescription? currentFont = null;
        if (_currentSelectedTarget is FormDesignerControl ctrl)
        {
            currentFont = ctrl.GetFontProperty(propName);
            if (currentFont == null)
            {
                if (propName.Equals("Font", StringComparison.OrdinalIgnoreCase))
                {
                    currentFont = ctrl.Font ?? _model?.Font ?? new FormFontDescription();
                }
                else
                {
                    string curVal = prop.Model.CurrentValue ?? "";
                    if (!string.IsNullOrWhiteSpace(curVal))
                    {
                        var parts = curVal.Split(',');
                        string fontName = parts[0].Trim();
                        double fontSize = 8.25;
                        if (parts.Length > 1 && double.TryParse(parts[1].Replace("pt", "").Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double sz))
                        {
                            fontSize = sz;
                        }
                        currentFont = new FormFontDescription { Name = fontName, Size = fontSize };
                    }
                    else
                    {
                        currentFont = new FormFontDescription();
                    }
                }
            }
        }
        else if (_currentSelectedTarget is FormDesignModel formModel)
        {
            currentFont = formModel.Font ?? new FormFontDescription();
        }
        else
        {
            currentFont = new FormFontDescription();
        }

        var parentWindow = Window.GetWindow(this);
        if (Dialogs.FontPickerDialog.ShowDialog(parentWindow, currentFont, out var selectedFont))
        {
            string fontDesc = $"{selectedFont.Name}, {selectedFont.Size.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt";

            if (_currentSelectedTarget is FormDesignerControl selectedCtrl)
            {
                selectedCtrl.SetFontProperty(propName, selectedFont);
                if (propName.Equals("Font", StringComparison.OrdinalIgnoreCase))
                {
                    selectedCtrl.Font = selectedFont;
                }
                selectedCtrl.SetCustomProperty(propName, fontDesc);
                DesignerPropertyFactory.SynchronizeStandardProperties(selectedCtrl, propName, selectedFont.Name);

                // Si es un control ActiveX alojado in-place, propagar al control nativo COM
                if (VisualResolver != null)
                {
                    var visual = VisualResolver(selectedCtrl);
                    ViBa7.ActiveXHost.ActiveXInPlaceHost? axHost = visual as ViBa7.ActiveXHost.ActiveXInPlaceHost
                        ?? (visual is Grid g && g.Children.Count > 0 ? g.Children[0] as ViBa7.ActiveXHost.ActiveXInPlaceHost : null);

                    if (axHost != null)
                    {
                        if (propName.Equals("Font", StringComparison.OrdinalIgnoreCase))
                        {
                            axHost.AxControl.ConfigureAmbientFont(
                                selectedFont.Name,
                                (float)selectedFont.Size,
                                selectedFont.IsBold,
                                selectedFont.Italic,
                                selectedFont.Underline,
                                selectedFont.Strikethrough,
                                (short)selectedFont.Charset);
                        }

                        axHost.AxControl.SetFontProperty(
                            propName,
                            selectedFont.Name,
                            (float)selectedFont.Size,
                            selectedFont.IsBold,
                            selectedFont.Italic,
                            selectedFont.Underline,
                            selectedFont.Strikethrough,
                            (short)selectedFont.Charset);
                    }
                }
            }
            else if (_currentSelectedTarget is FormDesignModel formModel)
            {
                formModel.Font = selectedFont;
            }

            prop.Model.SetValue(fontDesc);
            OnInnerModelChanged();
            prop.Refresh();
        }
    }
}
