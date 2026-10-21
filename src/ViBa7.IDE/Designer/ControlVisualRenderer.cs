using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Media.Imaging;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Designers;
using ViBa7.ActiveXHost;

namespace ViBa7.IDE.Designer;

/// <summary>
/// Fábrica y renderizador visual de controles de VB6 en WPF con alta fidelidad gráfica histórica.
/// </summary>
public static class ControlVisualRenderer
{
    private static readonly FontFamily DefaultFontFamily = new("Tahoma, MS Sans Serif, Segoe UI, Arial");

    /// <summary>
    /// Delegado global para resolver la definición de modelo de diseño de un UserControl por su TypeName.
    /// </summary>
    public static Func<string, FormDesignModel?>? UserControlModelResolver { get; set; }

    static ControlVisualRenderer()
    {
        ActiveXInPlaceHost.FallbackVisualFactory = ctrl => CreateWpfSimulatedVisual(ctrl);
    }

    /// <summary>
    /// Crea el elemento visual de WPF para un control del diseñador (Modo Diseño: UserMode = False).
    /// </summary>
    public static FrameworkElement CreateVisual(FormDesignerControl control, Action<FormDesignerControl, FrameworkElement>? onChildCreated = null)
        => CreateVisualInternal(control, onChildCreated, isRuntime: false);

    /// <summary>
    /// Crea el elemento visual de WPF/HwndHost interactivo para un control en Modo Ejecución F5 (UserMode = True).
    /// </summary>
    public static FrameworkElement CreateRuntimeVisual(FormDesignerControl control, Action<FormDesignerControl, FrameworkElement>? onChildCreated = null)
        => CreateVisualInternal(control, onChildCreated, isRuntime: true);

    /// <summary>
    /// Crea el elemento visual simulado en WPF para un control (usado internamente o como fallback de ActiveX).
    /// </summary>
    public static FrameworkElement CreateWpfSimulatedVisual(FormDesignerControl control, Action<FormDesignerControl, FrameworkElement>? onChildCreated = null)
    {
        bool hasVbPrefix = control.TypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase);
        string normType = control.TypeName.ToUpperInvariant();
        if (hasVbPrefix)
        {
            normType = normType[3..];
        }

        string shortName = control.TypeName.Contains('.')
            ? control.TypeName.Substring(control.TypeName.LastIndexOf('.') + 1)
            : control.TypeName;

        // Si es un UserControl registrado en el proyecto, componerlo visualmente
        if (UserControlModelResolver != null)
        {
            var ucModel = UserControlModelResolver(control.TypeName)
                       ?? UserControlModelResolver(shortName)
                       ?? UserControlModelResolver(normType);
            if (ucModel != null)
            {
                return CreateUserControlVisual(control, ucModel, onChildCreated);
            }
        }

        return normType switch
        {
            "COMMANDBUTTON" => CreateCommandButton(control),
            "TEXTBOX" => CreateTextBox(control),
            "LABEL" => CreateLabel(control),
            "CHECKBOX" => CreateCheckBox(control),
            "OPTIONBUTTON" => CreateOptionButton(control),
            "FRAME" => CreateFrame(control, onChildCreated),
            "LISTBOX" => CreateListBox(control),
            "COMBOBOX" => CreateComboBox(control),
            "PICTUREBOX" => CreatePictureBox(control, onChildCreated),
            "IMAGE" => CreateImage(control),
            "TIMER" => CreateTimer(control),
            "HSCROLLBAR" => CreateHScrollBar(control),
            "VSCROLLBAR" => CreateVScrollBar(control),
            "SHAPE" => CreateShape(control),
            "LINE" => CreateLine(control),
            _ => DispatchSpecializedOrGenericControl(control, normType, onChildCreated)
        };
    }

    private static FrameworkElement CreateVisualInternal(
        FormDesignerControl control,
        Action<FormDesignerControl, FrameworkElement>? onChildCreated,
        bool isRuntime)
    {
        bool hasVbPrefix = control.TypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase);
        string normType = control.TypeName.ToUpperInvariant();
        if (hasVbPrefix)
        {
            normType = normType[3..];
        }

        string shortName = control.TypeName.Contains('.')
            ? control.TypeName.Substring(control.TypeName.LastIndexOf('.') + 1)
            : control.TypeName;

        bool isIntrinsicVb = hasVbPrefix || IsIntrinsicVbControlType(normType);
        bool isNonVisual = control.IsNonResizable || IsNonVisualControl(normType, control);

        FrameworkElement? visual = null;

        // Si el control está marcado como no resuelto (OCX faltante o no registrado), renderizar visual distintivo preservando propiedades
        if (control.IsUnresolved)
        {
            visual = CreateUnresolvedControlVisual(control);
        }

        // Si es un UserControl registrado en el proyecto, componerlo visualmente a partir de sus controles constitutivos
        if (visual == null && UserControlModelResolver != null)
        {
            var ucModel = UserControlModelResolver(control.TypeName)
                       ?? UserControlModelResolver(shortName)
                       ?? UserControlModelResolver(normType);
            if (ucModel != null)
            {
                visual = CreateUserControlVisual(control, ucModel, onChildCreated);
            }
        }

        if (visual == null)
        {
            var wpfVisual = CreateWpfSimulatedVisual(control, onChildCreated);

            // Si es un control OCX externo y es visual, intentar hospedaje ActiveX (ActiveXRuntimeHost en F5 o ActiveXInPlaceHost en Diseño)
            if (!isIntrinsicVb && !isNonVisual && ActiveXHostFactory.EnableInPlaceActiveXHosting)
            {
                FrameworkElement? axVisual = isRuntime
                    ? ActiveXHostFactory.TryCreateRuntimeHost(control, wpfVisual)
                    : ActiveXHostFactory.TryCreateHost(control, wpfVisual);

                if (axVisual != null)
                {
                    if (control.Children.Count > 0)
                    {
                        var containerGrid = new Grid();
                        containerGrid.Children.Add(axVisual);
                        var overlayCanvas = new Canvas { ClipToBounds = true, Background = Brushes.Transparent };
                        foreach (var child in control.Children)
                        {
                            var childVisual = CreateVisualInternal(child, onChildCreated, isRuntime);
                            overlayCanvas.Children.Add(childVisual);
                            onChildCreated?.Invoke(child, childVisual);
                        }
                        containerGrid.Children.Add(overlayCanvas);
                        visual = containerGrid;
                    }
                    else
                    {
                        visual = axVisual;
                    }
                }
                else
                {
                    visual = wpfVisual;
                }
            }
            else
            {
                visual = wpfVisual;
            }
        }

        if (isNonVisual)
        {
            visual.Width = 32;
            visual.Height = 32;
            control.Width = 480;
            control.Height = 480;
        }
        else
        {
            visual.Width = control.WidthPixels;
            visual.Height = control.HeightPixels;
        }
        visual.Tag = control;
        visual.SnapsToDevicePixels = true;
        visual.UseLayoutRounding = true;

        Canvas.SetLeft(visual, control.LeftPixels);
        Canvas.SetTop(visual, control.TopPixels);

        return visual;
    }

    private static bool IsIntrinsicVbControlType(string normType)
    {
        return normType is "COMMANDBUTTON" or "TEXTBOX" or "LABEL" or "CHECKBOX" or "OPTIONBUTTON"
            or "FRAME" or "LISTBOX" or "COMBOBOX" or "PICTUREBOX" or "IMAGE" or "TIMER"
            or "HSCROLLBAR" or "VSCROLLBAR" or "SHAPE" or "LINE" or "DATA" or "OLE"
            or "DRIVELISTBOX" or "DIRLISTBOX" or "FILELISTBOX" or "MENU";
    }

    private static FrameworkElement CreateCommandButton(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, new SolidColorBrush(Color.FromRgb(236, 233, 216))),
            BorderBrush = new SolidColorBrush(Color.FromRgb(115, 115, 115)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(1)
        };

        // Bisel interior 3D clásico
        var innerBorder = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            Padding = new Thickness(4, 2, 4, 2)
        };

        var textBlock = BuildTextBlockWithHotkey(
            string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
            ctrl.Font,
            ctrl.ForeColorHex,
            TextAlignment.Center
        );
        textBlock.VerticalAlignment = VerticalAlignment.Center;
        textBlock.HorizontalAlignment = HorizontalAlignment.Center;

        var bmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
        if (bmp != null && (ctrl.Style == 1 || !string.IsNullOrEmpty(ctrl.Picture) || ctrl.PictureData != null))
        {
            var image = new System.Windows.Controls.Image
            {
                Source = bmp,
                Stretch = Stretch.None,
                Width = bmp.PixelWidth,
                Height = bmp.PixelHeight,
                SnapsToDevicePixels = true,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            bool hasCaption = !string.IsNullOrEmpty(ctrl.Caption);
            if (!hasCaption)
            {
                innerBorder.Child = image;
            }
            else
            {
                switch (ctrl.PictureAlignment)
                {
                    case PictureAlignment.Bottom:
                        var spBottom = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        spBottom.Children.Add(textBlock);
                        image.Margin = new Thickness(0, 2, 0, 0);
                        spBottom.Children.Add(image);
                        innerBorder.Child = spBottom;
                        break;

                    case PictureAlignment.Left:
                        var spLeft = new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        image.Margin = new Thickness(0, 0, 4, 0);
                        spLeft.Children.Add(image);
                        spLeft.Children.Add(textBlock);
                        innerBorder.Child = spLeft;
                        break;

                    case PictureAlignment.Right:
                        var spRight = new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        spRight.Children.Add(textBlock);
                        image.Margin = new Thickness(4, 0, 0, 0);
                        spRight.Children.Add(image);
                        innerBorder.Child = spRight;
                        break;

                    case PictureAlignment.Center:
                        var gridCenter = new Grid
                        {
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        gridCenter.Children.Add(image);
                        gridCenter.Children.Add(textBlock);
                        innerBorder.Child = gridCenter;
                        break;

                    case PictureAlignment.Top:
                    default:
                        var spTop = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        image.Margin = new Thickness(0, 0, 0, 2);
                        spTop.Children.Add(image);
                        spTop.Children.Add(textBlock);
                        innerBorder.Child = spTop;
                        break;
                }
            }
        }
        else
        {
            innerBorder.Child = textBlock;
        }

        border.Child = innerBorder;
        return border;
    }

    private static FrameworkElement CreateTextBox(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
            BorderBrush = new SolidColorBrush(Color.FromRgb(127, 157, 185)), // Bisel Windows clásico
            BorderThickness = new Thickness(1),
            Padding = new Thickness(3, 1, 3, 1)
        };

        var textBlock = new TextBlock
        {
            Text = ctrl.Text,
            FontFamily = GetFontFamily(ctrl.Font),
            FontSize = GetFontSize(ctrl.Font),
            FontWeight = GetFontWeight(ctrl.Font),
            FontStyle = GetFontStyle(ctrl.Font),
            Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, Brushes.Black),
            TextWrapping = (ctrl.MultiLine == true) ? TextWrapping.Wrap : TextWrapping.NoWrap,
            VerticalAlignment = VerticalAlignment.Top
        };

        if (ctrl.Alignment.HasValue)
        {
            textBlock.TextAlignment = ctrl.Alignment.Value switch
            {
                1 => TextAlignment.Right,
                2 => TextAlignment.Center,
                _ => TextAlignment.Left
            };
        }

        border.Child = textBlock;
        return border;
    }

    private static FrameworkElement CreateLabel(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.Transparent),
            Padding = new Thickness(1)
        };

        if (ctrl.BorderStyle == 1)
        {
            border.BorderBrush = new SolidColorBrush(Color.FromRgb(128, 128, 128));
            border.BorderThickness = new Thickness(1);
        }

        var align = ctrl.Alignment.HasValue ? (ctrl.Alignment.Value switch
        {
            1 => TextAlignment.Right,
            2 => TextAlignment.Center,
            _ => TextAlignment.Left
        }) : TextAlignment.Left;

        var textBlock = BuildTextBlockWithHotkey(
            string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
            ctrl.Font,
            ctrl.ForeColorHex,
            align
        );
        textBlock.TextWrapping = TextWrapping.Wrap;
        textBlock.VerticalAlignment = VerticalAlignment.Top;

        border.Child = textBlock;
        return border;
    }

    private static FrameworkElement CreateCheckBox(FormDesignerControl ctrl)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // Cajita de verificación 3D hundida (13x13)
        var box = new Border
        {
            Width = 13,
            Height = 13,
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(115, 115, 115)),
            BorderThickness = new Thickness(1),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Left
        };

        var textBlock = BuildTextBlockWithHotkey(
            ctrl.Caption,
            ctrl.Font,
            ctrl.ForeColorHex,
            TextAlignment.Left
        );
        textBlock.VerticalAlignment = VerticalAlignment.Center;
        textBlock.Margin = new Thickness(4, 0, 0, 0);

        Grid.SetColumn(box, 0);
        Grid.SetColumn(textBlock, 1);

        grid.Children.Add(box);
        grid.Children.Add(textBlock);

        return grid;
    }

    private static FrameworkElement CreateOptionButton(FormDesignerControl ctrl)
    {
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        // Círculo de opción (13x13)
        var circle = new Ellipse
        {
            Width = 13,
            Height = 13,
            Fill = Brushes.White,
            Stroke = new SolidColorBrush(Color.FromRgb(115, 115, 115)),
            StrokeThickness = 1,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Left
        };

        var textBlock = BuildTextBlockWithHotkey(
            ctrl.Caption,
            ctrl.Font,
            ctrl.ForeColorHex,
            TextAlignment.Left
        );
        textBlock.VerticalAlignment = VerticalAlignment.Center;
        textBlock.Margin = new Thickness(4, 0, 0, 0);

        Grid.SetColumn(circle, 0);
        Grid.SetColumn(textBlock, 1);

        grid.Children.Add(circle);
        grid.Children.Add(textBlock);

        return grid;
    }

    private static FrameworkElement CreateFrame(FormDesignerControl ctrl, Action<FormDesignerControl, FrameworkElement>? onChildCreated)
    {
        var groupBox = new GroupBox
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(208, 206, 191)),
            BorderThickness = new Thickness(1),
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.Transparent),
            Padding = new Thickness(2)
        };

        var headerText = BuildTextBlockWithHotkey(
            string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
            ctrl.Font,
            ctrl.ForeColorHex,
            TextAlignment.Left
        );
        groupBox.Header = headerText;

        // Lienzo interno para alojar los controles anidados del Frame
        var innerCanvas = new Canvas
        {
            ClipToBounds = true,
            Background = Brushes.Transparent
        };

        foreach (var child in ctrl.Children)
        {
            var childVisual = CreateVisual(child, onChildCreated);
            innerCanvas.Children.Add(childVisual);
            onChildCreated?.Invoke(child, childVisual);
        }

        groupBox.Content = innerCanvas;
        return groupBox;
    }

    private static FrameworkElement CreateListBox(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
            BorderBrush = new SolidColorBrush(Color.FromRgb(127, 157, 185)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });

        var itemsPanel = new StackPanel { Margin = new Thickness(3) };
        itemsPanel.Children.Add(new TextBlock { Text = $"{ctrl.Name} Item 1", FontSize = 11, Foreground = Brushes.Gray });
        itemsPanel.Children.Add(new TextBlock { Text = $"{ctrl.Name} Item 2", FontSize = 11, Foreground = Brushes.Gray });

        // Scrollbar falso visual a la derecha
        var scrollbar = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(215, 215, 215)),
            BorderThickness = new Thickness(1, 0, 0, 0)
        };

        Grid.SetColumn(itemsPanel, 0);
        Grid.SetColumn(scrollbar, 1);

        grid.Children.Add(itemsPanel);
        grid.Children.Add(scrollbar);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateComboBox(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
            BorderBrush = new SolidColorBrush(Color.FromRgb(127, 157, 185)),
            BorderThickness = new Thickness(1),
            Height = Math.Max(21, ctrl.HeightPixels)
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });

        var text = new TextBlock
        {
            Text = string.IsNullOrEmpty(ctrl.Text) ? ctrl.Name : ctrl.Text,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 2, 0),
            FontSize = 11.5
        };

        var arrowBtn = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1, 0, 0, 0),
            Child = new TextBlock
            {
                Text = "▼",
                FontSize = 8,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = Brushes.Black
            }
        };

        Grid.SetColumn(text, 0);
        Grid.SetColumn(arrowBtn, 1);

        grid.Children.Add(text);
        grid.Children.Add(arrowBtn);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreatePictureBox(FormDesignerControl ctrl, Action<FormDesignerControl, FrameworkElement>? onChildCreated)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
            BorderBrush = new SolidColorBrush(Color.FromRgb(128, 128, 128)),
            BorderThickness = new Thickness(ctrl.BorderStyle == 0 ? 0 : 1)
        };

        var canvas = new Canvas { ClipToBounds = true };

        var bmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
        if (bmp != null)
        {
            if (ctrl.AutoSize)
            {
                ctrl.Width = (int)(bmp.PixelWidth * FormDesignerControl.TwipsPerPixel);
                ctrl.Height = (int)(bmp.PixelHeight * FormDesignerControl.TwipsPerPixel);
            }

            var image = new System.Windows.Controls.Image
            {
                Source = bmp,
                Stretch = Stretch.None,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top
            };
            Canvas.SetLeft(image, 0);
            Canvas.SetTop(image, 0);
            canvas.Children.Add(image);
        }

        foreach (var child in ctrl.Children)
        {
            var childVisual = CreateVisual(child, onChildCreated);
            canvas.Children.Add(childVisual);
            onChildCreated?.Invoke(child, childVisual);
        }

        border.Child = canvas;
        return border;
    }

    private static FrameworkElement CreateImage(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1),
            Background = Brushes.Transparent
        };

        var text = new TextBlock
        {
            Text = $"🖼 {ctrl.Name}",
            FontSize = 10,
            Foreground = Brushes.Gray,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        border.Child = text;
        return border;
    }

    private static FrameworkElement CreateTimer(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Width = 32,
            Height = 32,
            Background = new SolidColorBrush(Color.FromRgb(255, 255, 230)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(2),
            ToolTip = $"{ctrl.Name} (Timer)"
        };

        var text = new TextBlock
        {
            Text = "⏱",
            FontSize = 14,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        border.Child = text;
        return border;
    }

    private static FrameworkElement CreateHScrollBar(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });

        var leftArrow = new TextBlock { Text = "◄", FontSize = 9, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        var rightArrow = new TextBlock { Text = "►", FontSize = 9, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        var thumb = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(4, 1, 4, 1)
        };

        Grid.SetColumn(leftArrow, 0);
        Grid.SetColumn(thumb, 1);
        Grid.SetColumn(rightArrow, 2);

        grid.Children.Add(leftArrow);
        grid.Children.Add(thumb);
        grid.Children.Add(rightArrow);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateVScrollBar(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(230, 230, 230)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(16) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(16) });

        var upArrow = new TextBlock { Text = "▲", FontSize = 9, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        var downArrow = new TextBlock { Text = "▼", FontSize = 9, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        var thumb = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1),
            Margin = new Thickness(1, 4, 1, 4)
        };

        Grid.SetRow(upArrow, 0);
        Grid.SetRow(thumb, 1);
        Grid.SetRow(downArrow, 2);

        grid.Children.Add(upArrow);
        grid.Children.Add(thumb);
        grid.Children.Add(downArrow);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateShape(FormDesignerControl ctrl)
    {
        var rect = new Rectangle
        {
            Stroke = VbColorHelper.ToBrush(ctrl.ForeColorHex, Brushes.Black),
            StrokeThickness = 1,
            Fill = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.Transparent)
        };
        return rect;
    }

    private static FrameworkElement CreateLine(FormDesignerControl ctrl)
    {
        var line = new Line
        {
            X1 = 0,
            Y1 = 0,
            X2 = Math.Max(10, ctrl.WidthPixels),
            Y2 = Math.Max(10, ctrl.HeightPixels),
            Stroke = VbColorHelper.ToBrush(ctrl.ForeColorHex, Brushes.Black),
            StrokeThickness = 1
        };
        return line;
    }

    #region Motores de Resolución de Propiedades Universales (Fase 4)
    public static string GetEffectiveCaption(FormDesignerControl ctrl)
    {
        // 1. Si ctrl.Caption no está vacío y no es referencia cruda .frx
        if (!string.IsNullOrWhiteSpace(ctrl.Caption) && !IsFrxReference(ctrl.Caption))
        {
            return ctrl.Caption;
        }

        // 2. Caption1 (usado por AxFramework, AxLabelPlus, etc.)
        string? cap1 = ctrl.GetCustomProperty("Caption1");
        if (!string.IsNullOrWhiteSpace(cap1) && !IsFrxReference(cap1))
        {
            return cap1;
        }

        // 3. Caption en CustomProperties
        string? capCustom = ctrl.GetCustomProperty("Caption");
        if (!string.IsNullOrWhiteSpace(capCustom) && !IsFrxReference(capCustom))
        {
            return capCustom;
        }

        // 4. Text estándar o personalizado
        if (!string.IsNullOrWhiteSpace(ctrl.Text) && !IsFrxReference(ctrl.Text))
        {
            return ctrl.Text;
        }
        string? textCustom = ctrl.GetCustomProperty("Text");
        if (!string.IsNullOrWhiteSpace(textCustom) && !IsFrxReference(textCustom))
        {
            return textCustom;
        }

        // 5. Title / Label / Prompt / Value / SimpleText / Header
        string? title = ctrl.GetCustomProperty("Title")
                     ?? ctrl.GetCustomProperty("Label")
                     ?? ctrl.GetCustomProperty("Prompt")
                     ?? ctrl.GetCustomProperty("SimpleText")
                     ?? ctrl.GetCustomProperty("Header");
        if (!string.IsNullOrWhiteSpace(title) && !IsFrxReference(title))
        {
            return title;
        }

        // 6. Fallback legible: nombre del control
        return !string.IsNullOrWhiteSpace(ctrl.Name) ? ctrl.Name : ctrl.DisplayType;
    }

    public static string? GetEffectiveSubCaption(FormDesignerControl ctrl)
    {
        string? cap2 = ctrl.GetCustomProperty("Caption2")
                    ?? ctrl.GetCustomProperty("SubTitle")
                    ?? ctrl.GetCustomProperty("Description");
        if (!string.IsNullOrWhiteSpace(cap2) && !IsFrxReference(cap2))
        {
            return cap2;
        }
        return null;
    }

    public static bool IsFrxReference(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        string t = val.Trim();
        return t.Contains(".frx\":", StringComparison.OrdinalIgnoreCase) ||
               t.Contains(".ctx\":", StringComparison.OrdinalIgnoreCase) ||
               t.Contains(".pgx\":", StringComparison.OrdinalIgnoreCase);
    }

    public static Brush GetEffectiveBackground(FormDesignerControl ctrl, Brush fallback)
    {
        // 1. Imagen o Skin asignado
        if (ctrl.PictureData != null && ctrl.PictureData.Length > 0)
        {
            var bmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
            if (bmp != null)
            {
                var brush = new ImageBrush(bmp) { Stretch = Stretch.Fill };
                brush.Freeze();
                return brush;
            }
        }

        // 2. Transparencia explícita
        string? trans = ctrl.GetCustomProperty("Transparent");
        string? bStyle = ctrl.GetCustomProperty("BackStyle");
        if ((trans != null && (trans.Equals("-1") || trans.Equals("1") || trans.Equals("True", StringComparison.OrdinalIgnoreCase))) ||
            (bStyle != null && bStyle.Equals("0")))
        {
            return Brushes.Transparent;
        }

        // 3. Chequeo de gradiente / degradado
        string? col1Str = ctrl.GetCustomProperty("GradientColor1")
                       ?? ctrl.GetCustomProperty("BackColor1")
                       ?? ctrl.BackColorHex
                       ?? ctrl.GetCustomProperty("BackGroundColor")
                       ?? ctrl.GetCustomProperty("BackColor");

        string? col2Str = ctrl.GetCustomProperty("GradientColor2")
                       ?? ctrl.GetCustomProperty("BackColor2")
                       ?? ctrl.GetCustomProperty("BackColorActive2")
                       ?? ctrl.GetCustomProperty("BackColorBkg");

        bool hasGradient = (ctrl.GetCustomProperty("Gradient")?.Equals("-1", StringComparison.OrdinalIgnoreCase) == true) ||
                           (ctrl.GetCustomProperty("Gradient")?.Equals("True", StringComparison.OrdinalIgnoreCase) == true) ||
                           (!string.IsNullOrWhiteSpace(ctrl.GetCustomProperty("BackColor1")) && !string.IsNullOrWhiteSpace(ctrl.GetCustomProperty("BackColor2")));

        if (hasGradient && !string.IsNullOrWhiteSpace(col1Str) && !string.IsNullOrWhiteSpace(col2Str))
        {
            Color c1 = VbColorHelper.ToWpfColor(col1Str, Colors.White);
            Color c2 = VbColorHelper.ToWpfColor(col2Str, Colors.LightGray);

            int angle = ctrl.GetCustomPropertyInt("BackAngle") ?? ctrl.GetCustomPropertyInt("GradientAngle") ?? 90;
            double rad = angle * Math.PI / 180.0;
            double cos = Math.Cos(rad);
            double sin = Math.Sin(rad);

            Point start = new(0.5 - cos * 0.5, 0.5 - sin * 0.5);
            Point end = new(0.5 + cos * 0.5, 0.5 + sin * 0.5);

            var grad = new LinearGradientBrush(c1, c2, start, end);
            grad.Freeze();
            return grad;
        }

        // 4. Color plano
        if (!string.IsNullOrWhiteSpace(col1Str))
        {
            return VbColorHelper.ToBrush(col1Str, fallback as SolidColorBrush ?? new SolidColorBrush(Color.FromRgb(240, 240, 240)));
        }

        return fallback;
    }

    public static (Brush BorderBrush, Thickness BorderThickness, CornerRadius CornerRadius) GetEffectiveBorder(
        FormDesignerControl ctrl,
        Brush defaultBorderBrush,
        int defaultThickness = 1,
        int defaultCorner = 0)
    {
        string? borderColStr = ctrl.GetCustomProperty("BorderColor")
                            ?? ctrl.GetCustomProperty("LineColor")
                            ?? ctrl.GetCustomProperty("HotLineColor")
                            ?? ctrl.GetCustomProperty("BorderColorActive");

        Brush bBrush = !string.IsNullOrWhiteSpace(borderColStr)
            ? VbColorHelper.ToBrush(borderColStr, defaultBorderBrush as SolidColorBrush ?? new SolidColorBrush(Color.FromRgb(140, 140, 140)))
            : defaultBorderBrush;

        int bWidth = ctrl.GetCustomPropertyInt("BorderWidth") ?? defaultThickness;
        if (ctrl.BorderStyle == 0 && string.IsNullOrWhiteSpace(borderColStr))
        {
            bWidth = 0;
        }

        int curve = ctrl.GetCustomPropertyInt("CornerCurve")
                 ?? ctrl.GetCustomPropertyInt("RoundCorners")
                 ?? defaultCorner;

        return (bBrush, new Thickness(Math.Max(0, bWidth)), new CornerRadius(Math.Clamp(curve, 0, 30)));
    }

    public static Brush GetEffectiveForeColor(FormDesignerControl ctrl, Brush fallback)
    {
        string? foreStr = ctrl.GetCustomProperty("ForeColor1")
                       ?? ctrl.ForeColorHex
                       ?? ctrl.GetCustomProperty("ForeColor")
                       ?? ctrl.GetCustomProperty("ForeColorList")
                       ?? ctrl.GetCustomProperty("TextColor");

        if (!string.IsNullOrWhiteSpace(foreStr))
        {
            return VbColorHelper.ToBrush(foreStr, fallback as SolidColorBrush ?? Brushes.Black);
        }
        return fallback;
    }

    public static SolidColorBrush? GetColorProperty(FormDesignerControl ctrl, params string[] propertyNames)
    {
        foreach (var prop in propertyNames)
        {
            string? val = ctrl.GetCustomProperty(prop);
            if (!string.IsNullOrWhiteSpace(val))
            {
                var fallback = new SolidColorBrush(Colors.Transparent);
                var brush = VbColorHelper.ToBrush(val, fallback);
                if (brush != fallback) return brush;
            }
        }
        return null;
    }

    public static FormFontDescription? GetEffectiveFont(FormDesignerControl ctrl)
    {
        if (ctrl.Font != null) return ctrl.Font;

        foreach (var prop in ctrl.CustomProperties)
        {
            if (prop.IsBeginProperty && (prop.Name.EndsWith("Font", StringComparison.OrdinalIgnoreCase) || prop.Name.StartsWith("Font", StringComparison.OrdinalIgnoreCase)))
            {
                return FormDesignParser.ParseFontBlock(prop.RawPropertyLines);
            }
        }
        return null;
    }

    private static List<(string Title, TextAlignment Alignment)> ParseFormatString(string? formatStr)
    {
        var list = new List<(string Title, TextAlignment Alignment)>();
        if (string.IsNullOrWhiteSpace(formatStr)) return list;

        string[] parts = formatStr.Split('|');
        foreach (var p in parts)
        {
            string trimmed = p.Trim();
            if (trimmed.Length == 0) continue;

            TextAlignment align = TextAlignment.Left;
            if (trimmed.StartsWith('^'))
            {
                align = TextAlignment.Center;
                trimmed = trimmed[1..].Trim();
            }
            else if (trimmed.StartsWith('>'))
            {
                align = TextAlignment.Right;
                trimmed = trimmed[1..].Trim();
            }
            else if (trimmed.StartsWith('<'))
            {
                align = TextAlignment.Left;
                trimmed = trimmed[1..].Trim();
            }

            list.Add((trimmed, align));
        }
        return list;
    }
    #endregion

    private static FrameworkElement CreateUnresolvedControlVisual(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromArgb(235, 254, 243, 199)), // Amber-100
            BorderBrush = new SolidColorBrush(Color.FromRgb(217, 119, 6)), // Amber-600
            BorderThickness = new Thickness(1.5),
            CornerRadius = new CornerRadius(2),
            SnapsToDevicePixels = true,
            ToolTip = string.IsNullOrWhiteSpace(ctrl.UnresolvedErrorMessage)
                ? $"Componente no registrado o faltante: {ctrl.TypeName}\nLas propiedades y datos binarios se preservarán intactos al guardar."
                : $"Componente no registrado: {ctrl.TypeName}\n{ctrl.UnresolvedErrorMessage}\nLas propiedades y datos binarios se preservarán intactos al guardar."
        };

        var grid = new Grid
        {
            ClipToBounds = true,
            Margin = new Thickness(2)
        };

        var sp = new StackPanel
        {
            Orientation = Orientation.Vertical,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var headerPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 0, 0, 2)
        };

        var warnIcon = new TextBlock
        {
            Text = "⚠️",
            FontSize = 12,
            Margin = new Thickness(0, 0, 4, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

        var badgeBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(180, 83, 9)),
            CornerRadius = new CornerRadius(2),
            Padding = new Thickness(4, 1, 4, 1),
            VerticalAlignment = VerticalAlignment.Center
        };

        var badgeText = new TextBlock
        {
            Text = "NO REGISTRADO",
            FontSize = 8.5,
            FontWeight = FontWeights.Bold,
            Foreground = Brushes.White
        };
        badgeBorder.Child = badgeText;

        headerPanel.Children.Add(warnIcon);
        headerPanel.Children.Add(badgeBorder);
        sp.Children.Add(headerPanel);

        var nameText = new TextBlock
        {
            Text = ctrl.Name,
            FontSize = 10.5,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(120, 53, 15)),
            TextAlignment = TextAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxWidth = Math.Max(40, ctrl.WidthPixels - 8)
        };
        sp.Children.Add(nameText);

        var typeText = new TextBlock
        {
            Text = ctrl.TypeName,
            FontSize = 9,
            Foreground = new SolidColorBrush(Color.FromRgb(180, 83, 9)),
            TextAlignment = TextAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            MaxWidth = Math.Max(40, ctrl.WidthPixels - 8)
        };
        sp.Children.Add(typeText);

        grid.Children.Add(sp);
        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateUserControlVisual(
        FormDesignerControl ctrl,
        FormDesignModel ucModel,
        Action<FormDesignerControl, FrameworkElement>? onChildCreated)
    {
        // Fondo: textura/skin del UserControl (.ctx o .frx), degradado o color de fondo
        Brush bgBrush = GetEffectiveBackground(ctrl, VbColorHelper.ToBrush(ctrl.BackColorHex ?? ucModel.BackColorHex, new SolidColorBrush(Color.FromRgb(240, 240, 240))));

        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(140, 140, 140)),
            defaultThickness: ctrl.BorderStyle == 0 ? 0 : 1,
            defaultCorner: 0);

        // Si el UserControl es transparente y sin borde, dibujar un contorno tenue de tiempo de diseño
        if (bgBrush == Brushes.Transparent && borderThickness == new Thickness(0))
        {
            borderBrush = new SolidColorBrush(Color.FromArgb(140, 140, 140, 140));
            borderThickness = new Thickness(1);
        }

        var border = new Border
        {
            Background = bgBrush,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            ClipToBounds = true,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };

        var rootGrid = new Grid
        {
            ClipToBounds = true,
            Background = Brushes.Transparent
        };

        var ucCanvas = new Canvas
        {
            ClipToBounds = true,
            Background = Brushes.Transparent
        };
        rootGrid.Children.Add(ucCanvas);
        border.Child = rootGrid;

        double scaleX = (ucModel.ClientWidth > 0 && ctrl.Width > 0) ? (double)ctrl.Width / ucModel.ClientWidth : 1.0;
        double scaleY = (ucModel.ClientHeight > 0 && ctrl.Height > 0) ? (double)ctrl.Height / ucModel.ClientHeight : 1.0;

        if (ucModel.Children.Count > 0)
        {
            // Propagar propiedades de la instancia sobre los controles constitutivos
            string? instanceCaption = !string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Caption : GetEffectiveCaption(ctrl);
            string? instanceText = !string.IsNullOrEmpty(ctrl.Text) ? ctrl.Text : ctrl.GetCustomPropertyString("Text");
            string? instanceForeColor = ctrl.ForeColorHex ?? ctrl.GetCustomProperty("ForeColor");
            string? instanceBackColor = ctrl.BackColorHex ?? ctrl.GetCustomProperty("BackColor");
            var instanceFont = ctrl.Font ?? GetEffectiveFont(ctrl);
            string? caption1 = ctrl.GetCustomPropertyString("Caption1");
            string? caption2 = ctrl.GetCustomPropertyString("Caption2");

            int labelIndex = 0;
            foreach (var child in ucModel.Children)
            {
                if (child.IsMenu) continue;

                var composedChild = child.Clone();

                // Propagar CustomProperties del contenedor sobre el hijo si no las tiene (delegación universal)
                foreach (var cp in ctrl.CustomProperties)
                {
                    if (!cp.IsBeginProperty && !composedChild.CustomProperties.Any(p => p.Name.Equals(cp.Name, StringComparison.OrdinalIgnoreCase)))
                    {
                        composedChild.CustomProperties.Add(new FormPropertyItem(cp.Name, cp.Value));
                    }
                }

                // Escalado proporcional de coordenadas y tamaños constitutivos
                if (Math.Abs(scaleX - 1.0) > 0.02 || Math.Abs(scaleY - 1.0) > 0.02)
                {
                    composedChild.Left = (int)Math.Round(child.Left * scaleX);
                    composedChild.Top = (int)Math.Round(child.Top * scaleY);
                    composedChild.Width = (int)Math.Round(child.Width * scaleX);
                    composedChild.Height = (int)Math.Round(child.Height * scaleY);
                }

                // Mapeo específico por nombre de control constitutivo (e.g. Label1.Caption, Text1.Text)
                string? specificCaption = ctrl.GetCustomPropertyString($"{composedChild.Name}.Caption");
                string? specificText = ctrl.GetCustomPropertyString($"{composedChild.Name}.Text");

                // Mapear Caption/Text según tipo de control constitutivo
                if (composedChild.TypeName.EndsWith("Label", StringComparison.OrdinalIgnoreCase))
                {
                    string? charCodeStr = ctrl.GetCustomProperty("IconCharCode") ?? ctrl.GetCustomProperty("IconCode") ?? ctrl.GetCustomProperty("CharCode");
                    int? parsedCharCode = DesignerPropertyFactory.ParseCharCode(charCodeStr);
                    if (parsedCharCode.HasValue && parsedCharCode.Value > 0 &&
                        (composedChild.Name.Contains("Icon", StringComparison.OrdinalIgnoreCase) ||
                         (child.Font?.Name != null && (child.Font.Name.Contains("Font", StringComparison.OrdinalIgnoreCase) || child.Font.Name.Contains("Ico", StringComparison.OrdinalIgnoreCase)))))
                    {
                        composedChild.Caption = SafeConvertFromUtf32(parsedCharCode.Value);
                        var iconFont = ctrl.GetFontProperty("IconFont") ?? ctrl.GetFontProperty("Font");
                        if (iconFont != null) composedChild.Font = iconFont;
                        string? iconCol = ctrl.GetCustomProperty("IconForeColor") ?? ctrl.GetCustomProperty("IconColor");
                        if (iconCol != null) composedChild.ForeColorHex = iconCol;
                    }
                    else if (specificCaption != null)
                    {
                        composedChild.Caption = specificCaption;
                    }
                    else if (labelIndex == 0 && caption1 != null)
                    {
                        composedChild.Caption = caption1;
                    }
                    else if (labelIndex == 1 && caption2 != null)
                    {
                        composedChild.Caption = caption2;
                    }
                    else if (instanceCaption != null)
                    {
                        composedChild.Caption = instanceCaption;
                    }
                    labelIndex++;
                }
                else if (composedChild.TypeName.EndsWith("CommandButton", StringComparison.OrdinalIgnoreCase) ||
                         composedChild.TypeName.Contains("Button", StringComparison.OrdinalIgnoreCase))
                {
                    composedChild.Caption = specificCaption ?? instanceCaption ?? composedChild.Caption;
                }

                if (composedChild.TypeName.EndsWith("TextBox", StringComparison.OrdinalIgnoreCase) ||
                    composedChild.TypeName.Contains("Text", StringComparison.OrdinalIgnoreCase))
                {
                    if (specificText != null)
                    {
                        composedChild.Text = specificText;
                    }
                    else if (instanceText != null)
                    {
                        composedChild.Text = instanceText;
                    }
                }

                if (instanceForeColor != null && string.IsNullOrEmpty(child.ForeColorHex))
                {
                    composedChild.ForeColorHex = instanceForeColor;
                }

                if (instanceFont != null && composedChild.Font == null)
                {
                    composedChild.Font = instanceFont;
                }

                if (instanceBackColor != null && (string.IsNullOrEmpty(child.BackColorHex) || child.BackColorHex.Equals("&H8000000F&", StringComparison.OrdinalIgnoreCase)))
                {
                    composedChild.BackColorHex = instanceBackColor;
                }

                var childVisual = CreateVisual(composedChild, onChildCreated: null);
                childVisual.IsHitTestVisible = false;
                ucCanvas.Children.Add(childVisual);
            }
        }
        else
        {
            // UserControl sin controles constitutivos (dibujado por código / GDI)
            string? charCodeStr = ctrl.GetCustomProperty("IconCharCode") ?? ctrl.GetCustomProperty("IconCode") ?? ctrl.GetCustomProperty("CharCode");
            int? charCode = DesignerPropertyFactory.ParseCharCode(charCodeStr);

            string labelText = GetEffectiveCaption(ctrl);
            var effFont = ctrl.Font ?? ucModel.Font ?? GetEffectiveFont(ctrl);
            var foreBrush = GetEffectiveForeColor(ctrl, VbColorHelper.ToBrush(ctrl.ForeColorHex ?? ucModel.ForeColorHex, Brushes.DarkSlateGray));

            if (charCode.HasValue && charCode.Value > 0)
            {
                string glyph = SafeConvertFromUtf32(charCode.Value);
                var iconFont = ctrl.GetFontProperty("IconFont") ?? ctrl.GetFontProperty("Font") ?? effFont;
                string? iconColorHex = ctrl.GetCustomProperty("IconForeColor") ?? ctrl.GetCustomProperty("IconColor") ?? ctrl.ForeColorHex;
                var iconBrush = VbColorHelper.ToBrush(iconColorHex, foreBrush as SolidColorBrush ?? Brushes.DarkSlateGray);

                var iconBlock = new TextBlock
                {
                    Text = glyph,
                    FontFamily = GetFontFamily(iconFont),
                    FontSize = iconFont?.Size > 0 ? GetFontSize(iconFont) : GetFontSize(effFont) + 2,
                    FontWeight = GetFontWeight(iconFont),
                    FontStyle = GetFontStyle(iconFont),
                    Foreground = iconBrush,
                    VerticalAlignment = VerticalAlignment.Center
                };

                if (!string.IsNullOrWhiteSpace(labelText))
                {
                    iconBlock.Margin = new Thickness(0, 0, 6, 0);
                    var hPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        IsHitTestVisible = false
                    };
                    hPanel.Children.Add(iconBlock);
                    hPanel.Children.Add(new TextBlock
                    {
                        Text = labelText,
                        FontFamily = GetFontFamily(effFont),
                        FontSize = GetFontSize(effFont),
                        FontWeight = GetFontWeight(effFont),
                        FontStyle = GetFontStyle(effFont),
                        Foreground = foreBrush,
                        VerticalAlignment = VerticalAlignment.Center
                    });
                    rootGrid.Children.Add(hPanel);
                }
                else
                {
                    rootGrid.Children.Add(iconBlock);
                }
            }
            else
            {
                var sp = new StackPanel
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    IsHitTestVisible = false
                };

                var iconSource = ResolveControlIcon(ctrl);
                if (iconSource != null)
                {
                    var img = new Image
                    {
                        Source = iconSource,
                        Width = 16,
                        Height = 16,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 0, 0, 2)
                    };
                    RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.NearestNeighbor);
                    sp.Children.Add(img);
                }

                sp.Children.Add(new TextBlock
                {
                    Text = labelText,
                    FontFamily = GetFontFamily(effFont),
                    FontSize = GetFontSize(effFont),
                    FontWeight = GetFontWeight(effFont),
                    FontStyle = GetFontStyle(effFont),
                    Foreground = foreBrush,
                    HorizontalAlignment = HorizontalAlignment.Center
                });

                string? subText = GetEffectiveSubCaption(ctrl);
                if (!string.IsNullOrWhiteSpace(subText))
                {
                    sp.Children.Add(new TextBlock
                    {
                        Text = subText,
                        FontFamily = GetFontFamily(effFont),
                        FontSize = Math.Max(7, GetFontSize(effFont) - 1.5),
                        Foreground = Brushes.Gray,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 2, 0, 0)
                    });
                }

                rootGrid.Children.Add(sp);
            }
        }

        // Renderizar controles hijos alojados dentro de esta instancia (si es un UserControl contenedor)
        if (ctrl.Children.Count > 0)
        {
            foreach (var hostedChild in ctrl.Children)
            {
                var hostedVisual = CreateVisual(hostedChild, onChildCreated);
                ucCanvas.Children.Add(hostedVisual);
                onChildCreated?.Invoke(hostedChild, hostedVisual);
            }
        }

        return border;
    }

    private static FrameworkElement DispatchSpecializedOrGenericControl(
        FormDesignerControl ctrl,
        string normType,
        Action<FormDesignerControl, FrameworkElement>? onChildCreated = null)
    {
        // Controles de Windows Common Controls y ActiveX comunes
        if (normType.Contains("PROGRESSBAR") || normType.Contains("PROGBAR") || normType.Contains("PROGRESSCIRCULAR"))
        {
            return CreateProgressBar(ctrl);
        }
        if (normType.Contains("SLIDER") || normType.Contains("TRACKBAR"))
        {
            return CreateSlider(ctrl);
        }
        if (normType.Contains("STATUSBAR"))
        {
            return CreateStatusBar(ctrl);
        }
        if (normType.Contains("TOOLBAR"))
        {
            return CreateToolbar(ctrl);
        }
        if (normType.Contains("SSTAB") || normType.Contains("TABSTRIP") || normType.Contains("TABCONTROL"))
        {
            return CreateTabStrip(ctrl);
        }
        if (normType.Contains("LISTVIEW"))
        {
            return CreateListView(ctrl);
        }
        if (normType.Contains("TREEVIEW"))
        {
            return CreateTreeView(ctrl);
        }
        if (normType.Contains("RICHTEXT") || normType.Contains("RTF"))
        {
            return CreateRichTextBox(ctrl);
        }
        if (normType.Contains("FLEXGRID") || normType.Contains("SGRID") || normType.Contains("LYNXGRID") ||
            normType.Contains("GRID") || normType.Contains("TABLE") || normType.Contains("SHEET") || normType.Contains("LIST"))
        {
            return CreateEnhancedGrid(ctrl, normType);
        }
        if (normType.Contains("UPDOWN") || normType.Contains("SPINBUTTON"))
        {
            return CreateUpDown(ctrl);
        }
        if (normType.Contains("MONTHVIEW"))
        {
            return CreateMonthView(ctrl);
        }
        if (normType.Contains("DTPICKER") || normType.Contains("DATETIME"))
        {
            return CreateDatePicker(ctrl);
        }
        if (normType.Contains("MASK") || normType.Contains("MSMASK"))
        {
            return CreateMaskEdBox(ctrl);
        }
        if (normType.Contains("WEBBROWSER") || normType.Contains("IEFRAME") || normType.Contains("INTERNETEXPLORER"))
        {
            return CreateWebBrowser(ctrl);
        }
        if (normType.Contains("ALPHAIMG") || normType.Contains("ALPHAIMAGE"))
        {
            return CreateAlphaImage(ctrl);
        }

        // Semántica de Controles de Terceros (Fase 4: Botones, Textos, Checks, Opciones, Marcos)
        if (normType.Contains("BUTTON") || normType.Contains("BTN") || normType.Contains("COMMAND") || normType.Contains("CMD"))
        {
            return CreateThirdPartyButton(ctrl);
        }
        if (normType.Contains("LABEL") || normType.Contains("BANNER") || normType.Contains("TITLE"))
        {
            return CreateEnhancedLabel(ctrl);
        }
        if (normType.Contains("TEXT") || normType.Contains("EDIT") || normType.Contains("INPUT"))
        {
            return CreateThirdPartyTextBox(ctrl);
        }
        if (normType.Contains("CHECK"))
        {
            return CreateThirdPartyCheckOrOption(ctrl, isOption: false);
        }
        if (normType.Contains("OPTION") || normType.Contains("RADIO"))
        {
            return CreateThirdPartyCheckOrOption(ctrl, isOption: true);
        }
        if (normType.Contains("PANEL") || normType.Contains("FRAME") || normType.Contains("GROUP") || normType.Contains("BOX"))
        {
            return CreateThirdPartyPanel(ctrl, onChildCreated);
        }
        if (normType.Contains("COMBO") || normType.Contains("DROPDOWN") || normType.Contains("CHOICE"))
        {
            return CreateMultiColumnCombo(ctrl);
        }

        return CreateGenericControl(ctrl);
    }

    private static FrameworkElement CreateThirdPartyButton(FormDesignerControl ctrl)
    {
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(140, 140, 140)),
            defaultThickness: 1,
            defaultCorner: 2);

        var border = new Border
        {
            Background = GetEffectiveBackground(ctrl, new SolidColorBrush(Color.FromRgb(236, 233, 216))),
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            Padding = new Thickness(3)
        };

        var sp = new StackPanel
        {
            Orientation = Orientation.Vertical,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var iconSource = ResolveControlIcon(ctrl);
        if (iconSource != null)
        {
            var img = new Image
            {
                Source = iconSource,
                Width = 16,
                Height = 16,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 2)
            };
            RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.NearestNeighbor);
            sp.Children.Add(img);
        }

        string caption = GetEffectiveCaption(ctrl);
        var font = GetEffectiveFont(ctrl);
        Brush foreBrush = GetEffectiveForeColor(ctrl, Brushes.Black);

        TextAlignment align = TextAlignment.Center;
        string? alignH = ctrl.GetCustomProperty("CaptionAlignH");
        if (alignH != null)
        {
            if (alignH == "0" || alignH.Equals("Left", StringComparison.OrdinalIgnoreCase)) align = TextAlignment.Left;
            else if (alignH == "2" || alignH.Equals("Right", StringComparison.OrdinalIgnoreCase)) align = TextAlignment.Right;
            else align = TextAlignment.Center;
        }

        var textBlock = new TextBlock
        {
            Text = caption,
            FontFamily = GetFontFamily(font),
            FontSize = GetFontSize(font),
            FontWeight = GetFontWeight(font) == FontWeights.Normal ? FontWeights.Medium : GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = foreBrush,
            TextAlignment = align,
            HorizontalAlignment = align == TextAlignment.Left ? HorizontalAlignment.Left : (align == TextAlignment.Right ? HorizontalAlignment.Right : HorizontalAlignment.Center)
        };
        sp.Children.Add(textBlock);

        string? subCaption = GetEffectiveSubCaption(ctrl);
        if (!string.IsNullOrWhiteSpace(subCaption))
        {
            var subText = new TextBlock
            {
                Text = subCaption,
                FontFamily = GetFontFamily(font),
                FontSize = Math.Max(7, GetFontSize(font) - 1.5),
                Foreground = Brushes.Gray,
                TextAlignment = align,
                HorizontalAlignment = sp.HorizontalAlignment,
                Margin = new Thickness(0, 1, 0, 0)
            };
            sp.Children.Add(subText);
        }

        border.Child = sp;
        return border;
    }

    private static FrameworkElement CreateThirdPartyTextBox(FormDesignerControl ctrl)
    {
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(127, 157, 185)),
            defaultThickness: 1,
            defaultCorner: 0);

        var border = new Border
        {
            Background = GetEffectiveBackground(ctrl, Brushes.White),
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            Padding = new Thickness(3, 1, 3, 1)
        };

        string displayText = !string.IsNullOrWhiteSpace(ctrl.Text)
            ? ctrl.Text
            : GetEffectiveCaption(ctrl);

        var font = GetEffectiveFont(ctrl);
        var textBlock = new TextBlock
        {
            Text = displayText,
            FontFamily = GetFontFamily(font),
            FontSize = GetFontSize(font),
            FontWeight = GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = GetEffectiveForeColor(ctrl, Brushes.Black),
            TextWrapping = ctrl.MultiLine == true ? TextWrapping.Wrap : TextWrapping.NoWrap,
            VerticalAlignment = VerticalAlignment.Top
        };

        if (ctrl.Alignment.HasValue)
        {
            textBlock.TextAlignment = ctrl.Alignment.Value switch
            {
                1 => TextAlignment.Right,
                2 => TextAlignment.Center,
                _ => TextAlignment.Left
            };
        }

        border.Child = textBlock;
        return border;
    }

    private static FrameworkElement CreateThirdPartyCheckOrOption(FormDesignerControl ctrl, bool isOption)
    {
        var grid = new Grid
        {
            Background = GetEffectiveBackground(ctrl, Brushes.Transparent),
            VerticalAlignment = VerticalAlignment.Center
        };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        int val = ctrl.GetCustomPropertyInt("Value") ?? (ctrl.Checked ? 1 : 0);
        bool isChecked = val != 0 || ctrl.Checked;

        var box = new Border
        {
            Width = 13,
            Height = 13,
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            CornerRadius = isOption ? new CornerRadius(6.5) : new CornerRadius(1),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        if (isChecked)
        {
            box.Child = new TextBlock
            {
                Text = isOption ? "●" : "✓",
                FontSize = isOption ? 8 : 9,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.Black,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        string caption = GetEffectiveCaption(ctrl);
        var font = GetEffectiveFont(ctrl);
        var lbl = new TextBlock
        {
            Text = caption,
            FontFamily = GetFontFamily(font),
            FontSize = GetFontSize(font),
            FontWeight = GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = GetEffectiveForeColor(ctrl, Brushes.Black),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(5, 0, 0, 0)
        };

        Grid.SetColumn(box, 0);
        Grid.SetColumn(lbl, 1);
        grid.Children.Add(box);
        grid.Children.Add(lbl);

        return grid;
    }

    private static FrameworkElement CreateThirdPartyPanel(FormDesignerControl ctrl, Action<FormDesignerControl, FrameworkElement>? onChildCreated)
    {
        var bg = GetEffectiveBackground(ctrl, new SolidColorBrush(Color.FromRgb(240, 240, 240)));
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            defaultThickness: 1,
            defaultCorner: 2);

        if (bg == Brushes.Transparent && borderThickness == new Thickness(0))
        {
            borderBrush = new SolidColorBrush(Color.FromArgb(150, 140, 140, 140));
            borderThickness = new Thickness(1);
        }

        var border = new Border
        {
            Background = bg,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            ClipToBounds = true
        };

        var panelGrid = new Grid();
        string caption = GetEffectiveCaption(ctrl);
        bool hasHeader = !string.IsNullOrWhiteSpace(caption) && !caption.Equals(ctrl.Name, StringComparison.OrdinalIgnoreCase);
        if (hasHeader)
        {
            panelGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var font = GetEffectiveFont(ctrl);
            var titleBlock = new TextBlock
            {
                Text = caption,
                FontFamily = GetFontFamily(font),
                FontSize = GetFontSize(font),
                FontWeight = FontWeights.SemiBold,
                Foreground = GetEffectiveForeColor(ctrl, Brushes.DarkSlateGray),
                Margin = new Thickness(6, 4, 6, 2)
            };
            Grid.SetRow(titleBlock, 0);
            panelGrid.Children.Add(titleBlock);
        }

        panelGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        var canvas = new Canvas
        {
            ClipToBounds = true,
            Background = Brushes.Transparent
        };
        Grid.SetRow(canvas, panelGrid.RowDefinitions.Count - 1);
        panelGrid.Children.Add(canvas);

        if (ctrl.Children.Count > 0)
        {
            foreach (var child in ctrl.Children)
            {
                var childVisual = CreateVisual(child, onChildCreated);
                canvas.Children.Add(childVisual);
                onChildCreated?.Invoke(child, childVisual);
            }
        }
        else if (!hasHeader)
        {
            var watermark = new TextBlock
            {
                Text = !string.IsNullOrWhiteSpace(ctrl.Name) ? ctrl.Name : ctrl.DisplayType,
                FontSize = 10,
                Foreground = GetEffectiveForeColor(ctrl, Brushes.DimGray),
                Opacity = 0.75,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false
            };
            Grid.SetRow(watermark, panelGrid.RowDefinitions.Count - 1);
            panelGrid.Children.Add(watermark);
        }

        border.Child = panelGrid;
        return border;
    }

    private static bool IsNonVisualControl(string normType, FormDesignerControl ctrl)
    {
        if (normType is "TIMER" or "COMMONDIALOG" or "COMMDLG" or "WINSOCK" or "MSCOMM" or "SYSINFO"
            or "CRYSTALREPORT" or "MAPI" or "MAPISESSION" or "MAPIMESSAGES"
            or "VBSINI" or "INET" or "INTERNETTRANSFER")
        {
            return true;
        }

        string? invisible = ctrl.GetCustomProperty("InvisibleAtRuntime");
        if (invisible != null && (invisible.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                  invisible.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                  invisible.Equals("True", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        return false;
    }

    private static ImageSource? ResolveControlIcon(FormDesignerControl ctrl)
    {
        // 1. Imagen asignada directamente (p.ej. PictureData desde .frx o .ctx)
        if (ctrl.PictureData != null && ctrl.PictureData.Length > 0)
        {
            var bmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
            if (bmp != null) return bmp;
        }

        // 2. Icono registrado en el Toolbox (OCX o UserControl)
        var tool = ControlToolboxRegistry.Find(ctrl.TypeName);
        if (tool?.IconData != null && tool.IconData.Length > 0)
        {
            var bmp = VbImageHelper.LoadBitmapFromBytes(tool.IconData);
            if (bmp != null) return bmp;
        }

        // 3. UserControl del proyecto: ToolboxBitmapData del modelo
        if (UserControlModelResolver != null)
        {
            var ucModel = UserControlModelResolver(ctrl.TypeName);
            if (ucModel?.ToolboxBitmapData != null && ucModel.ToolboxBitmapData.Length > 0)
            {
                var bmp = VbImageHelper.LoadBitmapFromBytes(ucModel.ToolboxBitmapData);
                if (bmp != null) return bmp;
            }
        }

        return null;
    }

    private static FrameworkElement CreateGenericControl(FormDesignerControl ctrl)
    {
        var bg = GetEffectiveBackground(ctrl, new SolidColorBrush(Color.FromRgb(245, 245, 245)));
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            defaultThickness: 1,
            defaultCorner: 0);

        if (bg == Brushes.Transparent && borderThickness == new Thickness(0))
        {
            borderBrush = new SolidColorBrush(Color.FromArgb(150, 140, 140, 140));
            borderThickness = new Thickness(1);
        }

        var border = new Border
        {
            Background = bg,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            Padding = new Thickness(4)
        };

        var sp = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        var iconSource = ResolveControlIcon(ctrl);
        if (iconSource != null)
        {
            var img = new Image
            {
                Source = iconSource,
                Width = 16,
                Height = 16,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 2)
            };
            RenderOptions.SetBitmapScalingMode(img, BitmapScalingMode.NearestNeighbor);
            sp.Children.Add(img);
        }

        string caption = GetEffectiveCaption(ctrl);
        string? subCaption = GetEffectiveSubCaption(ctrl);
        var effectiveFont = GetEffectiveFont(ctrl);
        var foreBrush = GetEffectiveForeColor(ctrl, Brushes.DarkSlateGray);

        if (!string.IsNullOrWhiteSpace(caption))
        {
            var capBlock = new TextBlock
            {
                Text = caption,
                FontFamily = GetFontFamily(effectiveFont),
                FontSize = GetFontSize(effectiveFont),
                FontWeight = GetFontWeight(effectiveFont),
                FontStyle = GetFontStyle(effectiveFont),
                Foreground = foreBrush,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            sp.Children.Add(capBlock);

            if (!string.IsNullOrWhiteSpace(subCaption))
            {
                sp.Children.Add(new TextBlock
                {
                    Text = subCaption,
                    FontSize = Math.Max(8, GetFontSize(effectiveFont) - 1.5),
                    Foreground = Brushes.DimGray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis
                });
            }

            if (!string.Equals(caption, ctrl.DisplayType, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(caption, ctrl.Name, StringComparison.OrdinalIgnoreCase))
            {
                sp.Children.Add(new TextBlock
                {
                    Text = ctrl.DisplayType,
                    FontSize = 8.5,
                    Foreground = Brushes.Gray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Opacity = 0.7
                });
            }
        }
        else
        {
            var titleBlock = new TextBlock
            {
                Text = ctrl.DisplayType,
                FontSize = 10,
                FontWeight = FontWeights.SemiBold,
                Foreground = foreBrush,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            sp.Children.Add(titleBlock);

            if (!string.Equals(ctrl.Name, ctrl.DisplayType, StringComparison.OrdinalIgnoreCase))
            {
                sp.Children.Add(new TextBlock
                {
                    Text = ctrl.Name,
                    FontSize = 9.5,
                    Foreground = Brushes.DimGray,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
            }
        }

        border.Child = sp;
        return border;
    }

    private static FrameworkElement CreateProgressBar(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(235, 235, 235)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(2)
        };

        var grid = new Grid();
        var bar = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(0, 120, 215)),
            HorizontalAlignment = HorizontalAlignment.Left,
            Width = Math.Max(20, ctrl.WidthPixels * 0.45)
        };
        grid.Children.Add(bar);
        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateSlider(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.Transparent)
        };

        var grid = new Grid { VerticalAlignment = VerticalAlignment.Center };
        var track = new Border
        {
            Height = 4,
            Background = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
            BorderThickness = new Thickness(0, 0, 0, 1)
        };
        var thumb = new Border
        {
            Width = 11,
            Height = 18,
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(115, 115, 115)),
            BorderThickness = new Thickness(1),
            HorizontalAlignment = HorizontalAlignment.Center
        };
        grid.Children.Add(track);
        grid.Children.Add(thumb);
        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateStatusBar(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(0, 1, 0, 0)
        };

        var grid = new Grid { Margin = new Thickness(2) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(2, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var panel1 = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(140, 140, 140)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4, 1, 4, 1),
            Margin = new Thickness(1)
        };
        panel1.Child = new TextBlock
        {
            Text = !string.IsNullOrEmpty(ctrl.Text) ? ctrl.Text : (!string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Caption : ctrl.Name),
            FontSize = 9.5,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetColumn(panel1, 0);
        grid.Children.Add(panel1);

        var panel2 = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(140, 140, 140)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4, 1, 4, 1),
            Margin = new Thickness(1)
        };
        panel2.Child = new TextBlock { Text = "Listo", FontSize = 9.5, VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(panel2, 1);
        grid.Children.Add(panel2);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateToolbar(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(2)
        };

        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        for (int i = 0; i < 4; i++)
        {
            var btn = new Border
            {
                Width = 22,
                Height = 22,
                Margin = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1)
            };
            sp.Children.Add(btn);
        }
        border.Child = sp;
        return border;
    }

    private static FrameworkElement CreateTabStrip(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(24) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var tabs = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(2, 2, 2, 0) };
        var activeTab = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1, 1, 1, 0),
            Padding = new Thickness(8, 2, 8, 2),
            Child = new TextBlock { Text = !string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Caption : "Tab 1", FontSize = 9.5 }
        };
        var inactiveTab = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1, 1, 1, 0),
            Padding = new Thickness(8, 2, 8, 2),
            Margin = new Thickness(2, 0, 0, 0),
            Child = new TextBlock { Text = "Tab 2", FontSize = 9.5, Foreground = Brushes.Gray }
        };
        tabs.Children.Add(activeTab);
        tabs.Children.Add(inactiveTab);
        Grid.SetRow(tabs, 0);
        grid.Children.Add(tabs);

        var body = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1)
        };
        Grid.SetRow(body, 1);
        grid.Children.Add(body);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateListView(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(20) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var header = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(235, 235, 235)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 0, 0, 1)
        };
        var headerCols = new StackPanel { Orientation = Orientation.Horizontal };
        headerCols.Children.Add(new Border
        {
            Width = 90,
            BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
            BorderThickness = new Thickness(0, 0, 1, 0),
            Padding = new Thickness(4, 1, 4, 1),
            Child = new TextBlock { Text = "Columna 1", FontSize = 9, FontWeight = FontWeights.SemiBold }
        });
        headerCols.Children.Add(new Border
        {
            Width = 90,
            Padding = new Thickness(4, 1, 4, 1),
            Child = new TextBlock { Text = "Columna 2", FontSize = 9, FontWeight = FontWeights.SemiBold }
        });
        header.Child = headerCols;
        Grid.SetRow(header, 0);
        grid.Children.Add(header);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateTreeView(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4)
        };

        var sp = new StackPanel();
        sp.Children.Add(new TextBlock { Text = "[-] Nodo Principal", FontSize = 9.5, Foreground = Brushes.Black });
        sp.Children.Add(new TextBlock { Text = "   [+] Subnodo 1", FontSize = 9.5, Foreground = Brushes.DimGray, Margin = new Thickness(12, 2, 0, 0) });
        sp.Children.Add(new TextBlock { Text = "   [+] Subnodo 2", FontSize = 9.5, Foreground = Brushes.DimGray, Margin = new Thickness(12, 2, 0, 0) });
        border.Child = sp;
        return border;
    }

    private static FrameworkElement CreateRichTextBox(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4)
        };

        string text = !string.IsNullOrEmpty(ctrl.Text) ? ctrl.Text : ctrl.Name;
        border.Child = new TextBlock
        {
            Text = text,
            FontSize = 10,
            Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, Brushes.Black)
        };
        return border;
    }

    private static FrameworkElement CreateFlexGrid(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        for (int r = 0; r < 4; r++) grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(18) });
        for (int c = 0; c < 4; c++) grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });

        for (int r = 0; r < 4; r++)
        {
            for (int c = 0; c < 4; c++)
            {
                bool isHeader = r == 0 || c == 0;
                var cell = new Border
                {
                    Background = isHeader ? new SolidColorBrush(Color.FromRgb(230, 230, 230)) : Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)),
                    BorderThickness = new Thickness(0, 0, 1, 1)
                };
                Grid.SetRow(cell, r);
                Grid.SetColumn(cell, c);
                grid.Children.Add(cell);
            }
        }

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateUpDown(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition());
        grid.RowDefinitions.Add(new RowDefinition());

        var up = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = new TextBlock { Text = "▲", FontSize = 7, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
        };
        var down = new Border
        {
            Child = new TextBlock { Text = "▼", FontSize = 7, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
        };
        Grid.SetRow(up, 0);
        Grid.SetRow(down, 1);
        grid.Children.Add(up);
        grid.Children.Add(down);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateDatePicker(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(2, 0, 0, 0)
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });

        var dateText = new TextBlock
        {
            Text = DateTime.Now.ToString("dd/MM/yyyy"),
            FontSize = 9.5,
            VerticalAlignment = VerticalAlignment.Center
        };
        var calBtn = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1, 0, 0, 0),
            Child = new TextBlock { Text = "📅", FontSize = 9, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
        };
        Grid.SetColumn(dateText, 0);
        Grid.SetColumn(calBtn, 1);
        grid.Children.Add(dateText);
        grid.Children.Add(calBtn);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateEnhancedGrid(FormDesignerControl ctrl, string flavor)
    {
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            defaultThickness: 1);

        var bg = GetEffectiveBackground(ctrl, Brushes.White);
        var border = new Border
        {
            Background = bg,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius
        };

        var outerGrid = new Grid();
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(20) });
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        outerGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(14) });
        outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        outerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });

        // Colores configurables de grilla VB6
        var headerBg = GetColorProperty(ctrl, "BackColorFixed", "BackColorHeader") ?? new SolidColorBrush(Color.FromRgb(236, 233, 216));
        var headerFg = GetColorProperty(ctrl, "ForeColorFixed", "ForeColorHeader") ?? Brushes.Black;
        var gridLineBrush = GetColorProperty(ctrl, "GridColor", "LineColor") ?? new SolidColorBrush(Color.FromRgb(215, 220, 225));
        var altRowBg = GetColorProperty(ctrl, "StripedColor", "BackColorAlternate", "BackColorAlt") ?? new SolidColorBrush(Color.FromRgb(246, 248, 250));

        // FormatString parsing (p.ej. | ^Código | <Descripción | >Cantidad |)
        string? formatStr = ctrl.GetCustomProperty("FormatString") ?? ctrl.GetCustomProperty("_FormatString");
        var parsedCols = ParseFormatString(formatStr);

        var headerBar = new Grid { Background = headerBg };
        var rowsGrid = new Grid();

        int colCount = parsedCols.Count > 0 ? parsedCols.Count : 4;
        // Columna 0: selector de fila fija
        headerBar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(22) });
        rowsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(22) });

        var cornerCell = new Border
        {
            Background = headerBg,
            BorderBrush = gridLineBrush,
            BorderThickness = new Thickness(0, 0, 1, 1)
        };
        Grid.SetColumn(cornerCell, 0);
        headerBar.Children.Add(cornerCell);

        for (int c = 0; c < colCount; c++)
        {
            GridLength widthDef;
            string title;
            TextAlignment align;

            if (parsedCols.Count > 0)
            {
                var colInfo = parsedCols[c];
                title = colInfo.Title;
                align = colInfo.Alignment;
                widthDef = c == colCount - 1 ? new GridLength(1, GridUnitType.Star) : new GridLength(70);
            }
            else
            {
                title = $"Col {c + 1}";
                align = TextAlignment.Left;
                widthDef = c == colCount - 1 ? new GridLength(1, GridUnitType.Star) : new GridLength(70);
            }

            headerBar.ColumnDefinitions.Add(new ColumnDefinition { Width = widthDef });
            rowsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = widthDef });

            var cellHdr = new Border
            {
                Background = headerBg,
                BorderBrush = gridLineBrush,
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(4, 1, 4, 1),
                Child = new TextBlock
                {
                    Text = title,
                    FontSize = 9,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = headerFg,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextAlignment = align,
                    TextTrimming = TextTrimming.CharacterEllipsis
                }
            };
            Grid.SetColumn(cellHdr, c + 1);
            headerBar.Children.Add(cellHdr);
        }

        Grid.SetRow(headerBar, 0);
        Grid.SetColumn(headerBar, 0);
        outerGrid.Children.Add(headerBar);

        // Filas de ejemplo
        for (int r = 0; r < 5; r++) rowsGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(18) });

        for (int r = 0; r < 5; r++)
        {
            bool isAlt = (r % 2 == 1);
            Brush rowBg = isAlt ? altRowBg : (bg is SolidColorBrush scb && scb.Color.A > 0 ? bg : Brushes.White);

            for (int c = 0; c <= colCount; c++)
            {
                var rowCell = new Border
                {
                    Background = c == 0 ? headerBg : rowBg,
                    BorderBrush = gridLineBrush,
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Padding = new Thickness(2, 0, 2, 0)
                };
                if (c == 0 && r == 0)
                {
                    rowCell.Child = new TextBlock
                    {
                        Text = "►",
                        FontSize = 7,
                        Foreground = headerFg,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                }
                Grid.SetRow(rowCell, r);
                Grid.SetColumn(rowCell, c);
                rowsGrid.Children.Add(rowCell);
            }
        }

        Grid.SetRow(rowsGrid, 1);
        Grid.SetColumn(rowsGrid, 0);
        outerGrid.Children.Add(rowsGrid);

        // Barras de scroll
        var vScroll = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1, 0, 0, 0)
        };
        Grid.SetRow(vScroll, 0);
        Grid.SetRowSpan(vScroll, 2);
        Grid.SetColumn(vScroll, 1);
        outerGrid.Children.Add(vScroll);

        var hScroll = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 1, 0, 0)
        };
        Grid.SetRow(hScroll, 2);
        Grid.SetColumn(hScroll, 0);
        outerGrid.Children.Add(hScroll);

        border.Child = outerGrid;
        return border;
    }

    private static FrameworkElement CreateMaskEdBox(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
            BorderBrush = new SolidColorBrush(Color.FromRgb(127, 157, 185)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(3, 1, 3, 1)
        };

        string mask = ctrl.GetCustomProperty("Mask") ?? ctrl.GetCustomProperty("Format") ?? string.Empty;
        string display = !string.IsNullOrEmpty(ctrl.Text) ? ctrl.Text : (!string.IsNullOrEmpty(mask) ? mask : "___-____");

        var tb = new TextBlock
        {
            Text = display,
            FontFamily = GetFontFamily(ctrl.Font),
            FontSize = GetFontSize(ctrl.Font),
            FontWeight = GetFontWeight(ctrl.Font),
            FontStyle = GetFontStyle(ctrl.Font),
            Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, Brushes.Black),
            VerticalAlignment = VerticalAlignment.Center
        };
        border.Child = tb;
        return border;
    }

    private static FrameworkElement CreateWebBrowser(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(22) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var topBar = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(4, 1, 4, 1)
        };

        var barPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        barPanel.Children.Add(new TextBlock { Text = "🌐", FontSize = 10, Margin = new Thickness(0, 0, 4, 0), VerticalAlignment = VerticalAlignment.Center });

        var addressBox = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1),
            Padding = new Thickness(4, 0, 4, 0),
            Width = Math.Max(100, ctrl.WidthPixels - 50),
            Child = new TextBlock { Text = "about:blank", FontSize = 9, Foreground = Brushes.DimGray, VerticalAlignment = VerticalAlignment.Center }
        };
        barPanel.Children.Add(addressBox);
        topBar.Child = barPanel;
        Grid.SetRow(topBar, 0);
        grid.Children.Add(topBar);

        var contentArea = new Grid { Background = Brushes.White };
        contentArea.Children.Add(new TextBlock
        {
            Text = $"{ctrl.Name} (WebBrowser)",
            FontSize = 10,
            Foreground = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        });
        Grid.SetRow(contentArea, 1);
        grid.Children.Add(contentArea);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateAlphaImage(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(160, 160, 160)),
            BorderThickness = new Thickness(1),
            Background = Brushes.Transparent
        };

        var bmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
        if (bmp != null)
        {
            var img = new System.Windows.Controls.Image
            {
                Source = bmp,
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            border.Child = img;
            return border;
        }

        var sp = new StackPanel
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        sp.Children.Add(new TextBlock
        {
            Text = "🎨 AlphaImage",
            FontSize = 10,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.DarkSlateGray,
            HorizontalAlignment = HorizontalAlignment.Center
        });
        sp.Children.Add(new TextBlock
        {
            Text = ctrl.Name,
            FontSize = 9,
            Foreground = Brushes.Gray,
            HorizontalAlignment = HorizontalAlignment.Center
        });
        border.Child = sp;
        return border;
    }

    private static FrameworkElement CreateEnhancedLabel(FormDesignerControl ctrl)
    {
        string typeUpper = ctrl.TypeName.ToUpperInvariant();
        bool isDivider = typeUpper.Contains("DIV") ||
                         typeUpper.Contains("SEP") ||
                         ctrl.GetCustomProperty("Style")?.Equals("Divider", StringComparison.OrdinalIgnoreCase) == true ||
                         ctrl.GetCustomProperty("Line")?.Equals("True", StringComparison.OrdinalIgnoreCase) == true ||
                         ctrl.HeightPixels <= 4;

        var font = GetEffectiveFont(ctrl);
        var foreBrush = GetEffectiveForeColor(ctrl, Brushes.Black);
        var bg = GetEffectiveBackground(ctrl, Brushes.Transparent);
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            ctrl.BorderStyle == 1 ? new SolidColorBrush(Color.FromRgb(128, 128, 128)) : Brushes.Transparent,
            defaultThickness: ctrl.BorderStyle == 1 ? 1 : 0);

        string caption = GetEffectiveCaption(ctrl);
        if (string.IsNullOrEmpty(caption))
        {
            caption = isDivider ? string.Empty : ctrl.Name;
        }

        // Si es un divisor horizontal
        if (isDivider)
        {
            var divGrid = new Grid
            {
                Background = bg,
                VerticalAlignment = VerticalAlignment.Center
            };

            var lineColor = GetColorProperty(ctrl, "LineColor", "BorderColor") ?? foreBrush;

            if (string.IsNullOrEmpty(caption))
            {
                var line = new Border
                {
                    Height = 1,
                    Background = lineColor,
                    VerticalAlignment = VerticalAlignment.Center
                };
                return line;
            }
            else
            {
                divGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                divGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var lbl = new TextBlock
                {
                    Text = caption,
                    FontFamily = GetFontFamily(font),
                    FontSize = GetFontSize(font),
                    FontWeight = FontWeights.Bold,
                    Foreground = foreBrush,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 6, 0)
                };
                Grid.SetColumn(lbl, 0);
                divGrid.Children.Add(lbl);

                var line = new Border
                {
                    Height = 1,
                    Background = lineColor,
                    VerticalAlignment = VerticalAlignment.Center
                };
                Grid.SetColumn(line, 1);
                divGrid.Children.Add(line);
                return divGrid;
            }
        }

        var border = new Border
        {
            Background = bg,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius,
            Padding = new Thickness(2)
        };

        TextAlignment align = TextAlignment.Left;
        int alignVal = ctrl.Alignment ?? ctrl.GetCustomPropertyInt("TextAlign") ?? ctrl.GetCustomPropertyInt("Alignment") ?? 0;
        if (alignVal == 1) align = TextAlignment.Right;
        else if (alignVal == 2) align = TextAlignment.Center;

        var panel = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center
        };

        bool isWordWrap = ctrl.GetCustomProperty("WordWrap")?.Equals("True", StringComparison.OrdinalIgnoreCase) == true ||
                          ctrl.GetCustomProperty("WordWrap") == "-1" ||
                          ctrl.GetCustomProperty("WordWrap") == "1";

        var text = new TextBlock
        {
            Text = caption,
            FontFamily = GetFontFamily(font),
            FontSize = GetFontSize(font),
            FontWeight = GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = foreBrush,
            TextAlignment = align,
            TextWrapping = isWordWrap ? TextWrapping.Wrap : TextWrapping.NoWrap
        };

        string? charCodeStr = ctrl.GetCustomProperty("IconCharCode") ?? ctrl.GetCustomProperty("IconCode") ?? ctrl.GetCustomProperty("CharCode");
        int? charCode = DesignerPropertyFactory.ParseCharCode(charCodeStr);

        if (charCode.HasValue && charCode.Value > 0)
        {
            string glyph = SafeConvertFromUtf32(charCode.Value);
            var iconFont = ctrl.GetFontProperty("IconFont") ?? ctrl.GetFontProperty("Font") ?? font;
            string? iconColorHex = ctrl.GetCustomProperty("IconForeColor") ?? ctrl.GetCustomProperty("IconColor") ?? ctrl.ForeColorHex;
            var iconBrush = VbColorHelper.ToBrush(iconColorHex, foreBrush as SolidColorBrush ?? Brushes.Black);

            var iconBlock = new TextBlock
            {
                Text = glyph,
                FontFamily = GetFontFamily(iconFont),
                FontSize = iconFont?.Size > 0 ? GetFontSize(iconFont) : GetFontSize(font) + 2,
                FontWeight = GetFontWeight(iconFont),
                FontStyle = GetFontStyle(iconFont),
                Foreground = iconBrush,
                VerticalAlignment = VerticalAlignment.Center
            };

            if (!string.IsNullOrEmpty(caption) && !caption.Equals(ctrl.Name, StringComparison.OrdinalIgnoreCase))
            {
                var hPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    VerticalAlignment = VerticalAlignment.Center
                };
                if (align == TextAlignment.Center) hPanel.HorizontalAlignment = HorizontalAlignment.Center;
                else if (align == TextAlignment.Right) hPanel.HorizontalAlignment = HorizontalAlignment.Right;

                iconBlock.Margin = new Thickness(0, 0, 6, 0);
                hPanel.Children.Add(iconBlock);
                hPanel.Children.Add(text);
                panel.Children.Add(hPanel);
            }
            else
            {
                if (align == TextAlignment.Center) iconBlock.HorizontalAlignment = HorizontalAlignment.Center;
                else if (align == TextAlignment.Right) iconBlock.HorizontalAlignment = HorizontalAlignment.Right;
                panel.Children.Add(iconBlock);
            }
        }
        else
        {
            panel.Children.Add(text);
        }

        string? subCaption = GetEffectiveSubCaption(ctrl);
        if (!string.IsNullOrEmpty(subCaption))
        {
            panel.Children.Add(new TextBlock
            {
                Text = subCaption,
                FontFamily = GetFontFamily(font),
                FontSize = Math.Max(8, GetFontSize(font) - 2),
                Foreground = Brushes.DimGray,
                TextAlignment = align,
                Margin = new Thickness(0, 2, 0, 0)
            });
        }

        border.Child = panel;
        return border;
    }

    private static FrameworkElement CreateMultiColumnCombo(FormDesignerControl ctrl)
    {
        var (borderBrush, borderThickness, cornerRadius) = GetEffectiveBorder(
            ctrl,
            new SolidColorBrush(Color.FromRgb(127, 157, 185)),
            defaultThickness: 1);

        var bg = GetEffectiveBackground(ctrl, Brushes.White);
        var foreBrush = GetEffectiveForeColor(ctrl, Brushes.Black);
        var font = GetEffectiveFont(ctrl);

        var border = new Border
        {
            Background = bg,
            BorderBrush = borderBrush,
            BorderThickness = borderThickness,
            CornerRadius = cornerRadius
        };

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });

        string display = GetEffectiveCaption(ctrl);
        if (string.IsNullOrEmpty(display)) display = !string.IsNullOrEmpty(ctrl.Text) ? ctrl.Text : ctrl.Name;

        var text = new TextBlock
        {
            Text = display,
            FontSize = GetFontSize(font),
            FontFamily = GetFontFamily(font),
            FontWeight = GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = foreBrush,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 2, 0)
        };
        var arrowBtn = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(236, 233, 216)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
            BorderThickness = new Thickness(1, 0, 0, 0),
            Child = new TextBlock { Text = "▼", FontSize = 7.5, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
        };

        Grid.SetColumn(text, 0);
        Grid.SetColumn(arrowBtn, 1);
        grid.Children.Add(text);
        grid.Children.Add(arrowBtn);

        border.Child = grid;
        return border;
    }

    private static FrameworkElement CreateMonthView(FormDesignerControl ctrl)
    {
        var border = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(1),
            Width = Math.Max(160, ctrl.WidthPixels),
            Height = Math.Max(140, ctrl.HeightPixels)
        };

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(22) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(16) });
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var header = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(10, 36, 106)),
            Padding = new Thickness(4, 2, 4, 2)
        };
        var headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });

        var prevBtn = new TextBlock { Text = "◄", FontSize = 8, Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center };
        var monthTitle = new TextBlock { Text = DateTime.Now.ToString("MMMM yyyy"), FontSize = 9, FontWeight = FontWeights.Bold, Foreground = Brushes.White, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
        var nextBtn = new TextBlock { Text = "►", FontSize = 8, Foreground = Brushes.White, HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Center };

        Grid.SetColumn(prevBtn, 0);
        Grid.SetColumn(monthTitle, 1);
        Grid.SetColumn(nextBtn, 2);
        headerGrid.Children.Add(prevBtn);
        headerGrid.Children.Add(monthTitle);
        headerGrid.Children.Add(nextBtn);
        header.Child = headerGrid;
        Grid.SetRow(header, 0);
        grid.Children.Add(header);

        var daysHeader = new StackPanel { Orientation = Orientation.Horizontal, Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)), HorizontalAlignment = HorizontalAlignment.Center };
        string[] dayLabels = ["Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sá"];
        foreach (var d in dayLabels)
        {
            daysHeader.Children.Add(new TextBlock { Text = d, Width = 22, FontSize = 8, FontWeight = FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, TextAlignment = TextAlignment.Center });
        }
        Grid.SetRow(daysHeader, 1);
        grid.Children.Add(daysHeader);

        var daysGrid = new Grid { Margin = new Thickness(4) };
        for (int r = 0; r < 5; r++) daysGrid.RowDefinitions.Add(new RowDefinition());
        for (int c = 0; c < 7; c++) daysGrid.ColumnDefinitions.Add(new ColumnDefinition());

        int dayNum = 1;
        for (int r = 0; r < 5; r++)
        {
            for (int c = 0; c < 7; c++)
            {
                if (dayNum <= 31)
                {
                    var dBlock = new TextBlock
                    {
                        Text = dayNum.ToString(),
                        FontSize = 8,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    if (dayNum == DateTime.Now.Day)
                    {
                        dBlock.FontWeight = FontWeights.Bold;
                        dBlock.Foreground = new SolidColorBrush(Color.FromRgb(10, 36, 106));
                    }
                    Grid.SetRow(dBlock, r);
                    Grid.SetColumn(dBlock, c);
                    daysGrid.Children.Add(dBlock);
                    dayNum++;
                }
            }
        }
        Grid.SetRow(daysGrid, 2);
        grid.Children.Add(daysGrid);

        border.Child = grid;
        return border;
    }

    /// <summary>
    /// Construye un TextBlock interpretando aceleradores '&' de VB6 con subrayado.
    /// </summary>
    public static TextBlock BuildTextBlockWithHotkey(
        string rawText,
        FormFontDescription? font,
        string? foreColorHex,
        TextAlignment alignment)
    {
        var textBlock = new TextBlock
        {
            FontFamily = GetFontFamily(font),
            FontSize = GetFontSize(font),
            FontWeight = GetFontWeight(font),
            FontStyle = GetFontStyle(font),
            Foreground = VbColorHelper.ToBrush(foreColorHex, Brushes.Black),
            TextAlignment = alignment
        };

        if (string.IsNullOrEmpty(rawText))
        {
            return textBlock;
        }

        bool nextIsHotkey = false;

        for (int i = 0; i < rawText.Length; i++)
        {
            char c = rawText[i];

            if (c == '&' && !nextIsHotkey)
            {
                if (i + 1 < rawText.Length && rawText[i + 1] == '&')
                {
                    // Ampersand doble '&&' representa '&' literal
                    textBlock.Inlines.Add(new Run("&"));
                    i++;
                }
                else
                {
                    nextIsHotkey = true;
                }
                continue;
            }

            if (nextIsHotkey)
            {
                var run = new Run(c.ToString());
                run.TextDecorations.Add(TextDecorations.Underline);
                textBlock.Inlines.Add(run);
                nextIsHotkey = false;
            }
            else
            {
                textBlock.Inlines.Add(new Run(c.ToString()));
            }
        }

        return textBlock;
    }

    /// <summary>
    /// Convierte de forma segura un punto de código Unicode a su representación de texto UTF-16,
    /// evitando excepciones por caracteres fuera de rango o subrogados no válidos.
    /// </summary>
    public static string SafeConvertFromUtf32(int code)
    {
        if (code >= 0 && code <= 0x10FFFF && (code < 0xD800 || code > 0xDFFF))
        {
            try
            {
                return char.ConvertFromUtf32(code);
            }
            catch
            {
            }
        }
        return "\uFFFD";
    }

    private static FontFamily GetFontFamily(FormFontDescription? font)
    {
        if (font != null && !string.IsNullOrWhiteSpace(font.Name))
        {
            try
            {
                return new FontFamily($"{font.Name}, Tahoma, Arial");
            }
            catch
            {
                return DefaultFontFamily;
            }
        }
        return DefaultFontFamily;
    }

    private static double GetFontSize(FormFontDescription? font)
    {
        if (font != null && font.Size > 0)
        {
            // Tamaño en puntos a tamaño en píxeles aproximado (points * 96 / 72)
            return Math.Max(9.0, Math.Min(48.0, font.Size * (96.0 / 72.0)));
        }
        return 11.0;
    }

    private static FontWeight GetFontWeight(FormFontDescription? font)
    {
        return font != null && font.IsBold ? FontWeights.Bold : FontWeights.Normal;
    }

    private static FontStyle GetFontStyle(FormFontDescription? font)
    {
        return font != null && font.Italic ? FontStyles.Italic : FontStyles.Normal;
    }
}
