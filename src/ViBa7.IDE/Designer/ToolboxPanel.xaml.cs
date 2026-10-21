using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.IDE.Designer;

public partial class ToolboxPanel : UserControl, INotifyPropertyChanged
{
    public class ToolboxItemViewModel : INotifyPropertyChanged
    {
        private bool _isSelected;
        private bool _isEnabled = true;

        public string TypeName { get; init; } = string.Empty;
        public string DisplayName { get; init; } = string.Empty;
        public string Glyph { get; init; } = string.Empty;
        public ImageSource? IconSource { get; init; }
        public string ToolTip { get; init; } = string.Empty;

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsEnabled)));
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public ObservableCollection<ToolboxItemViewModel> Tools { get; } = new();

    public event Action<string?>? ToolSelected;
    public event Action<string>? ToolDoubleClicked;
    public event PropertyChangedEventHandler? PropertyChanged;

    private bool _isCompactMode = false;
    public bool IsCompactMode
    {
        get => _isCompactMode;
        set
        {
            if (_isCompactMode != value)
            {
                _isCompactMode = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsCompactMode)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ItemWidth)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TextVisibility)));
            }
        }
    }

    public double ItemWidth => IsCompactMode ? 28 : 124;
    public Visibility TextVisibility => IsCompactMode ? Visibility.Collapsed : Visibility.Visible;

    public string? ActiveToolTypeName => Tools.FirstOrDefault(t => t.IsSelected && t.TypeName != ControlToolboxRegistry.PointerTool)?.TypeName;

    public ToolboxPanel()
    {
        InitializeComponent();
        DataContext = this;
        LoadTools();
    }

    public void ReloadTools() => LoadTools();

    public void LoadTools()
    {
        Tools.Clear();
        foreach (var tool in ControlToolboxRegistry.GetAllTools())
        {
            ImageSource? iconSrc = null;
            if (tool.IconData != null && tool.IconData.Length > 0)
            {
                iconSrc = VbImageHelper.LoadBitmapFromBytes(tool.IconData);
            }
            iconSrc ??= GetToolIcon(tool.TypeName);

            Tools.Add(new ToolboxItemViewModel
            {
                TypeName = tool.TypeName,
                DisplayName = tool.DisplayName,
                Glyph = tool.Glyph,
                IconSource = iconSrc,
                ToolTip = tool.ToolTip,
                IsSelected = tool.TypeName == ControlToolboxRegistry.PointerTool
            });
        }
        ToolsItemsControl.ItemsSource = Tools;
    }

    private static ImageSource? GetToolIcon(string typeName)
    {
        string resourceKey = typeName switch
        {
            ControlToolboxRegistry.PointerTool => "Icon_Tool_Pointer",
            "VB.PictureBox" => "Icon_Tool_PictureBox",
            "VB.Label" => "Icon_Tool_Label",
            "VB.TextBox" => "Icon_Tool_TextBox",
            "VB.Frame" => "Icon_Tool_Frame",
            "VB.CommandButton" => "Icon_Tool_CommandButton",
            "VB.CheckBox" => "Icon_Tool_CheckBox",
            "VB.OptionButton" => "Icon_Tool_OptionButton",
            "VB.ComboBox" => "Icon_Tool_ComboBox",
            "VB.ListBox" => "Icon_Tool_ListBox",
            "VB.HScrollBar" => "Icon_Tool_HScrollBar",
            "VB.VScrollBar" => "Icon_Tool_VScrollBar",
            "VB.Timer" => "Icon_Tool_Timer",
            "VB.DriveListBox" => "Icon_Tool_DriveListBox",
            "VB.DirListBox" => "Icon_Tool_DirListBox",
            "VB.FileListBox" => "Icon_Tool_FileListBox",
            "VB.Shape" => "Icon_Tool_Shape",
            "VB.Line" => "Icon_Tool_Line",
            "VB.Image" => "Icon_Tool_Image",
            "VB.Data" => "Icon_Tool_Data",
            "VB.OLE" => "Icon_Tool_OLE",
            _ => (ControlToolboxRegistry.Find(typeName)?.Glyph == "🎛️" || ControlToolboxRegistry.Find(typeName)?.ToolTip.StartsWith("Control de Usuario", StringComparison.OrdinalIgnoreCase) == true)
                ? "Icon_FileUserControl"
                : "Icon_Tool_CommandButton"
        };

        return Application.Current?.TryFindResource(resourceKey) as ImageSource;
    }

    public void SetToolEnabled(string typeName, bool isEnabled)
    {
        var tool = Tools.FirstOrDefault(t => t.TypeName.Equals(typeName, StringComparison.OrdinalIgnoreCase) ||
                                             t.DisplayName.Equals(typeName, StringComparison.OrdinalIgnoreCase) ||
                                             t.TypeName.EndsWith($".{typeName}", StringComparison.OrdinalIgnoreCase));
        if (tool != null)
        {
            tool.IsEnabled = isEnabled;
            if (!isEnabled && tool.IsSelected)
            {
                ResetToPointer();
            }
        }
    }

    public void EnableAllTools()
    {
        foreach (var item in Tools)
        {
            item.IsEnabled = true;
        }
    }

    public void ResetToPointer()
    {
        foreach (var item in Tools)
        {
            item.IsSelected = (item.TypeName == ControlToolboxRegistry.PointerTool);
        }
        ToolSelected?.Invoke(null);
    }

    private void ToolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is ToggleButton btn && btn.DataContext is ToolboxItemViewModel clickedItem)
        {
            // Radio button behavior
            foreach (var item in Tools)
            {
                item.IsSelected = (item == clickedItem);
            }

            if (clickedItem.TypeName == ControlToolboxRegistry.PointerTool)
            {
                ToolSelected?.Invoke(null);
            }
            else
            {
                ToolSelected?.Invoke(clickedItem.TypeName);
            }
        }
    }

    private void ToolButton_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is ToggleButton btn && btn.DataContext is ToolboxItemViewModel clickedItem)
        {
            if (clickedItem.TypeName != ControlToolboxRegistry.PointerTool)
            {
                ToolDoubleClicked?.Invoke(clickedItem.TypeName);
                ResetToPointer();
            }
        }
    }

    private void BtnToggleViewMode_Click(object sender, RoutedEventArgs e)
    {
        IsCompactMode = !IsCompactMode;
    }
}
