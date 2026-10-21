using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using ViBa7.ActiveXHost;
using ViBa7.IDE.Designer;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Execution;

namespace ViBa7.IDE.Execution;

/// <summary>
/// Ventana anfitriona para la ejecución interactiva en caliente (Hot-Run) de formularios VB6.
/// Instancia controles WPF dinámicos y ActiveXRuntimeHost (HwndHost), enlaza eventos en vivo con el motor de ejecución
/// y mantiene sincronizada la interfaz visual con las sentencias del script.
/// </summary>
public partial class LiveFormWindow : Window
{
    private readonly FormDesignModel _model;
    private readonly LiveExecutionHost _host;
    private readonly LiveFormProxy _formProxy;
    private readonly List<DispatcherTimer> _activeTimers = new();
    private readonly List<IDisposable> _disposableRuntimeHosts = new();
    private readonly Dictionary<string, DispatcherTimer> _controlTimers = new(StringComparer.OrdinalIgnoreCase);
    private readonly Func<string, FormDesignModel?>? _userControlResolver;

    public LiveFormProxy FormProxy => _formProxy;

    public LiveFormWindow(
        FormDesignModel model,
        LiveExecutionHost host,
        string? sourceCode = null,
        Func<string, FormDesignModel?>? userControlResolver = null)
    {
        InitializeComponent();

        _model = model ?? throw new ArgumentNullException(nameof(model));
        _host = host ?? throw new ArgumentNullException(nameof(host));
        _userControlResolver = userControlResolver;
        _formProxy = new LiveFormProxy(_model, this);

        if (_model.IsUserControl)
        {
            Title = $"[Contenedor de Prueba de UserControl: {_model.Name}]";
            Width = Math.Max(340, _model.ClientWidthPixels + 50);
            Height = Math.Max(260, _model.ClientHeightPixels + 70);
            var testBg = new SolidColorBrush(Color.FromRgb(128, 128, 128));
            Background = testBg;
            FormCanvas.Background = testBg;
        }
        else
        {
            Title = string.IsNullOrEmpty(_model.Caption) ? _model.Name : _model.Caption;
            Width = Math.Max(200, _model.ClientWidthPixels);
            Height = Math.Max(150, _model.ClientHeightPixels);

            if (!string.IsNullOrWhiteSpace(_model.BackColorHex))
            {
                var brush = VbColorHelper.ToBrush(_model.BackColorHex, SystemColors.ControlBrush);
                Background = brush;
                FormCanvas.Background = brush;
            }
        }

        if (sourceCode != null)
        {
            _host.LoadSourceCode(sourceCode, _model.Name);
        }

        BuildControls();
        WireEvents();
    }

    private void BuildControls()
    {
        // 1. Registrar el formulario y la variable Me en el contexto
        _host.Context.ActiveForms[_model.Name] = _formProxy;
        _host.Context.SetVariable(_model.Name, VbValue.FromObject(_formProxy));
        _host.Context.SetVariable("Me", VbValue.FromObject(_formProxy));

        // 2. Construir menús de formulario si existen
        BuildMenus();

        // 3. Normalizar jerarquía del modelo para asegurar que ningún control hijo permanezca en la raíz
        NormalizeModelHierarchy();

        // 4. Si se está ejecutando directamente un UserControl en modo Test Container:
        if (_model.IsUserControl)
        {
            var ucBorder = new Border
            {
                Width = _model.ClientWidthPixels,
                Height = _model.ClientHeightPixels,
                Background = VbColorHelper.ToBrush(_model.BackColorHex, SystemColors.ControlBrush),
                BorderBrush = new SolidColorBrush(Color.FromRgb(64, 64, 64)),
                BorderThickness = new Thickness(1),
                ClipToBounds = true
            };
            Canvas.SetLeft(ucBorder, 20);
            Canvas.SetTop(ucBorder, 20);
            var innerCanvas = new Canvas { ClipToBounds = true, Background = Brushes.Transparent };
            ucBorder.Child = innerCanvas;
            FormCanvas.Children.Add(ucBorder);

            foreach (var ctrl in _model.Children)
            {
                if (ctrl.TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase)) continue;
                BuildControlRecursive(ctrl, innerCanvas);
            }
            return;
        }

        // 5. Construir controles jerárquicamente a partir de los controles de primer nivel
        foreach (var ctrl in _model.Children)
        {
            if (ctrl.TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase))
            {
                continue; // Ya procesado en BuildMenus
            }

            if (ctrl.ParentControl != null)
            {
                continue; // Los controles anidados solo se construyen dentro de su contenedor padre
            }

            BuildControlRecursive(ctrl, FormCanvas);
        }
    }

    private void NormalizeModelHierarchy()
    {
        var allControls = _model.GetAllControls();
        foreach (var ctrl in allControls)
        {
            if (ctrl.ParentControl != null)
            {
                _model.Children.Remove(ctrl);
                if (!ctrl.ParentControl.Children.Contains(ctrl))
                {
                    ctrl.ParentControl.Children.Add(ctrl);
                }
            }
        }
    }

    private void BuildControlRecursive(FormDesignerControl ctrl, Canvas parentCanvas)
    {
        var visual = CreateLiveControlVisual(ctrl);
        if (visual != null)
        {
            Canvas.SetLeft(visual, ctrl.LeftPixels);
            Canvas.SetTop(visual, ctrl.TopPixels);
            visual.Width = ctrl.WidthPixels;
            visual.Height = ctrl.HeightPixels;

            parentCanvas.Children.Add(visual);

            var proxy = new LiveControlProxy(ctrl, visual);
            RegisterControlWithContext(ctrl, proxy);

            // Si es un contenedor, obtener su lienzo interno para añadir sus hijos
            Canvas? innerCanvas = GetContainerInnerCanvas(visual);
            if (innerCanvas != null)
            {
                foreach (var child in ctrl.Children)
                {
                    if (child.TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    BuildControlRecursive(child, innerCanvas);
                }
            }
        }
        else
        {
            // Para controles sin representación visual en runtime (ej. Timer)
            var dummyVisual = new FrameworkElement { Name = ctrl.Name };
            _controlTimers.TryGetValue(ctrl.Name, out var timer);
            var proxy = new LiveControlProxy(ctrl, dummyVisual, timer);
            RegisterControlWithContext(ctrl, proxy);
        }
    }

    private void RegisterControlWithContext(FormDesignerControl ctrl, LiveControlProxy proxy)
    {
        if (ctrl.Index.HasValue)
        {
            _formProxy.RegisterControl(proxy, ctrl.Index.Value);
            var existingVar = _host.Context.GetVariable(ctrl.Name);
            if (existingVar.RawValue is LiveControlArray arrProxy)
            {
                arrProxy.Add(ctrl.Index.Value, proxy);
            }
            else
            {
                var newArr = new LiveControlArray(ctrl.Name);
                if (existingVar.RawValue is ILiveControl prevCtrl && existingVar.RawValue is not LiveControlArray)
                {
                    newArr.Add(0, prevCtrl);
                }
                newArr.Add(ctrl.Index.Value, proxy);
                _host.Context.SetVariable(ctrl.Name, VbValue.FromObject(newArr));
            }
            _host.Context.SetVariable($"{ctrl.Name}({ctrl.Index.Value})", VbValue.FromObject(proxy));
        }
        else
        {
            _formProxy.RegisterControl(proxy);
            _host.Context.SetVariable(ctrl.Name, VbValue.FromObject(proxy));
        }
    }

    private static Canvas? GetContainerInnerCanvas(FrameworkElement visual)
    {
        if (visual is GroupBox gb && gb.Content is Canvas c1)
        {
            return c1;
        }
        if (visual is Border b && b.Child is Canvas c2)
        {
            return c2;
        }
        if (visual is Canvas c3)
        {
            return c3;
        }
        return null;
    }

    private void BuildMenus()
    {
        var menuControls = _model.GetAllControls()
            .FindAll(c => c.TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase));

        if (menuControls.Count == 0) return;

        FormMenuBar.Visibility = Visibility.Visible;
        FormMenuBar.Items.Clear();

        var topLevelMenus = menuControls.FindAll(c => c.ParentControl == null);
        foreach (var topMenu in topLevelMenus)
        {
            var menuItem = CreateMenuItemHierarchy(topMenu, menuControls);
            FormMenuBar.Items.Add(menuItem);
        }
    }

    private MenuItem CreateMenuItemHierarchy(FormDesignerControl menuCtrl, List<FormDesignerControl> allMenus)
    {
        var item = new MenuItem
        {
            Header = menuCtrl.Caption,
            IsEnabled = menuCtrl.Enabled,
            IsChecked = menuCtrl.Checked
        };

        if (!string.IsNullOrEmpty(menuCtrl.Shortcut))
        {
            item.InputGestureText = menuCtrl.Shortcut;
        }

        string ctrlName = menuCtrl.Name;
        item.Click += (s, e) => _host.InvokeProcedure($"{ctrlName}_Click", _model.Name);

        var children = allMenus.FindAll(c => c.ParentControl == menuCtrl);
        foreach (var child in children)
        {
            item.Items.Add(CreateMenuItemHierarchy(child, allMenus));
        }

        return item;
    }

    private FrameworkElement? CreateLiveControlVisual(
        FormDesignerControl ctrl,
        string? targetModuleName = null,
        string? parentControlPrefix = null)
    {
        // 1. Si es un UserControl registrado o resuelto, componerlo en vivo a partir de su definición .ctl
        if (_userControlResolver != null)
        {
            string shortName = ctrl.TypeName.Contains('.')
                ? ctrl.TypeName.Substring(ctrl.TypeName.LastIndexOf('.') + 1)
                : ctrl.TypeName;
            string normTypeName = ctrl.TypeName.ToUpperInvariant();
            var ucModel = _userControlResolver(ctrl.TypeName)
                       ?? _userControlResolver(shortName)
                       ?? _userControlResolver(normTypeName);
            if (ucModel != null)
            {
                return CreateLiveUserControlVisual(ctrl, ucModel);
            }
        }

        string normType = ctrl.TypeName.ToUpperInvariant();
        if (normType.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        {
            normType = normType[3..];
        }

        string ctrlName = ctrl.Name;
        string modName = targetModuleName ?? _model.Name;

        void DispatchEvent(string eventName)
        {
            try
            {
                if (ctrl.Index.HasValue)
                {
                    _host.InvokeProcedure($"{ctrlName}_{eventName}", modName, VbValue.FromLong(ctrl.Index.Value));
                }
                else
                {
                    _host.InvokeProcedure($"{ctrlName}_{eventName}", modName);
                }

                if (!string.IsNullOrEmpty(parentControlPrefix))
                {
                    if (ctrl.Index.HasValue)
                    {
                        _host.InvokeProcedure($"{parentControlPrefix}_{ctrlName}_{eventName}", _model.Name, VbValue.FromLong(ctrl.Index.Value));
                    }
                    else
                    {
                        _host.InvokeProcedure($"{parentControlPrefix}_{ctrlName}_{eventName}", _model.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                _host.Context.WriteDebugPrint($"[Error al ejecutar {ctrlName}_{eventName}]: {ex.Message}");
            }
        }

        switch (normType)
        {
            case "COMMANDBUTTON":
                var btn = new Button
                {
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    Background = VbColorHelper.ToBrush(ctrl.BackColorHex, SystemColors.ControlBrush),
                    Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, SystemColors.ControlTextBrush)
                };

                var liveBmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
                if (liveBmp != null && (ctrl.Style == 1 || !string.IsNullOrEmpty(ctrl.Picture) || ctrl.PictureData != null))
                {
                    var liveImg = new System.Windows.Controls.Image
                    {
                        Source = liveBmp,
                        Stretch = Stretch.None,
                        Width = liveBmp.PixelWidth,
                        Height = liveBmp.PixelHeight,
                        SnapsToDevicePixels = true,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center
                    };

                    string cap = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption;
                    var txtBlock = new TextBlock
                    {
                        Text = cap,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center
                    };

                    if (string.IsNullOrEmpty(ctrl.Caption))
                    {
                        btn.Content = liveImg;
                    }
                    else
                    {
                        switch (ctrl.PictureAlignment)
                        {
                            case PictureAlignment.Bottom:
                                var spBottom = new StackPanel { Orientation = Orientation.Vertical, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                                spBottom.Children.Add(txtBlock);
                                liveImg.Margin = new Thickness(0, 2, 0, 0);
                                spBottom.Children.Add(liveImg);
                                btn.Content = spBottom;
                                break;
                            case PictureAlignment.Left:
                                var spLeft = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                                liveImg.Margin = new Thickness(0, 0, 4, 0);
                                spLeft.Children.Add(liveImg);
                                spLeft.Children.Add(txtBlock);
                                btn.Content = spLeft;
                                break;
                            case PictureAlignment.Right:
                                var spRight = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                                spRight.Children.Add(txtBlock);
                                liveImg.Margin = new Thickness(4, 0, 0, 0);
                                spRight.Children.Add(liveImg);
                                btn.Content = spRight;
                                break;
                            case PictureAlignment.Center:
                                var gridCenter = new Grid { HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                                gridCenter.Children.Add(liveImg);
                                gridCenter.Children.Add(txtBlock);
                                btn.Content = gridCenter;
                                break;
                            case PictureAlignment.Top:
                            default:
                                var spTop = new StackPanel { Orientation = Orientation.Vertical, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
                                liveImg.Margin = new Thickness(0, 0, 0, 2);
                                spTop.Children.Add(liveImg);
                                spTop.Children.Add(txtBlock);
                                btn.Content = spTop;
                                break;
                        }
                    }
                }
                else
                {
                    btn.Content = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption;
                }

                btn.Click += (s, e) => DispatchEvent("Click");
                return btn;

            case "TEXTBOX":
                var tb = new TextBox
                {
                    Text = ctrl.Text,
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    Background = VbColorHelper.ToBrush(ctrl.BackColorHex, SystemColors.WindowBrush),
                    Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, SystemColors.WindowTextBrush)
                };
                if (ctrl.MultiLine == true)
                {
                    tb.AcceptsReturn = true;
                    tb.TextWrapping = TextWrapping.Wrap;
                    tb.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                }
                tb.TextChanged += (s, e) => DispatchEvent("Change");
                return tb;

            case "LABEL":
                var lbl = new TextBlock
                {
                    Text = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, SystemColors.ControlTextBrush),
                    VerticalAlignment = VerticalAlignment.Center
                };
                lbl.MouseDown += (s, e) => DispatchEvent("Click");
                return lbl;

            case "CHECKBOX":
                var chk = new CheckBox
                {
                    Content = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, SystemColors.ControlTextBrush)
                };
                chk.Click += (s, e) => DispatchEvent("Click");
                return chk;

            case "OPTIONBUTTON":
                var rad = new RadioButton
                {
                    Content = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    Foreground = VbColorHelper.ToBrush(ctrl.ForeColorHex, SystemColors.ControlTextBrush)
                };
                rad.Click += (s, e) => DispatchEvent("Click");
                return rad;

            case "LISTBOX":
                var lb = new ListBox
                {
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden
                };
                lb.SelectionChanged += (s, e) => DispatchEvent("Click");
                lb.MouseDoubleClick += (s, e) => DispatchEvent("DblClick");
                return lb;

            case "COMBOBOX":
                var cb = new ComboBox
                {
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    IsEditable = true
                };
                cb.SelectionChanged += (s, e) => DispatchEvent("Click");
                return cb;

            case "FRAME":
                var gb = new GroupBox
                {
                    Header = string.IsNullOrEmpty(ctrl.Caption) ? ctrl.Name : ctrl.Caption,
                    IsEnabled = ctrl.Enabled,
                    Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(208, 206, 191)),
                    BorderThickness = new Thickness(1),
                    Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.Transparent),
                    Padding = new Thickness(2)
                };
                var frameCanvas = new Canvas
                {
                    ClipToBounds = true,
                    Background = Brushes.Transparent
                };
                gb.Content = frameCanvas;
                return gb;

            case "PICTUREBOX":
                var picBorder = new Border
                {
                    Background = VbColorHelper.ToBrush(ctrl.BackColorHex, Brushes.White),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(128, 128, 128)),
                    BorderThickness = new Thickness(ctrl.BorderStyle == 0 ? 0 : 1),
                    ClipToBounds = true
                };
                var picCanvas = new Canvas
                {
                    ClipToBounds = true,
                    Background = Brushes.Transparent
                };

                var livePicBmp = VbImageHelper.LoadBitmapFromBytes(ctrl.PictureData);
                if (livePicBmp != null)
                {
                    if (ctrl.AutoSize)
                    {
                        ctrl.Width = (int)(livePicBmp.PixelWidth * FormDesignerControl.TwipsPerPixel);
                        ctrl.Height = (int)(livePicBmp.PixelHeight * FormDesignerControl.TwipsPerPixel);
                    }

                    var imgElem = new System.Windows.Controls.Image
                    {
                        Source = livePicBmp,
                        Stretch = Stretch.None,
                        HorizontalAlignment = HorizontalAlignment.Left,
                        VerticalAlignment = VerticalAlignment.Top
                    };
                    Canvas.SetLeft(imgElem, 0);
                    Canvas.SetTop(imgElem, 0);
                    picCanvas.Children.Add(imgElem);
                }

                picBorder.Child = picCanvas;
                picBorder.MouseDown += (s, e) => DispatchEvent("Click");
                return picBorder;

            case "TIMER":
                int intervalMs = ctrl.Interval;
                if (intervalMs <= 0)
                {
                    var intervalProp = ctrl.CustomProperties.Find(p => p.Name.Equals("Interval", StringComparison.OrdinalIgnoreCase));
                    if (intervalProp != null && int.TryParse(intervalProp.Value, out int iv) && iv > 0)
                    {
                        intervalMs = iv;
                    }
                }

                var timer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(Math.Max(1, intervalMs > 0 ? intervalMs : 1)),
                    IsEnabled = ctrl.Enabled && intervalMs > 0
                };
                timer.Tick += (s, e) =>
                {
                    if (ctrl.Enabled && ctrl.Interval > 0)
                    {
                        DispatchEvent("Timer");
                    }
                };
                _activeTimers.Add(timer);
                _controlTimers[ctrl.Name] = timer;
                if (ctrl.Enabled && intervalMs > 0)
                {
                    timer.Start();
                }

                // Un Timer no es visible en tiempo de ejecución en VB6
                return null;

            default:
                // Para controles OCX ActiveX (alojados interactivamente en ActiveXRuntimeHost con UserMode = True) y otros controles gráficos
                var genericVisual = ControlVisualRenderer.CreateRuntimeVisual(ctrl);
                if (genericVisual != null)
                {
                    if (genericVisual is IDisposable dispHost)
                    {
                        _disposableRuntimeHosts.Add(dispHost);
                    }
                    else if (genericVisual is Grid containerGrid && containerGrid.Children.Count > 0 && containerGrid.Children[0] is IDisposable childDisp)
                    {
                        _disposableRuntimeHosts.Add(childDisp);
                    }

                    ActiveXRuntimeHost? axRtHost = genericVisual as ActiveXRuntimeHost
                        ?? (genericVisual is Grid g && g.Children.Count > 0 ? g.Children[0] as ActiveXRuntimeHost : null);

                    if (axRtHost != null)
                    {
                        axRtHost.RuntimeExtenderBoundsChanged += host =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                Canvas.SetLeft(genericVisual, host.AxControl.Extender.Left / FormDesignModel.TwipsPerPixel);
                                Canvas.SetTop(genericVisual, host.AxControl.Extender.Top / FormDesignModel.TwipsPerPixel);
                                genericVisual.Width = Math.Max(16, host.AxControl.Extender.Width / FormDesignModel.TwipsPerPixel);
                                genericVisual.Height = Math.Max(16, host.AxControl.Extender.Height / FormDesignModel.TwipsPerPixel);
                            });
                        };

                        // Fase 4: Enlazar eventos COM en vivo (IConnectionPoint / ActiveXEventSink)
                        axRtHost.ComEventFired += invocation =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    var procArgs = new VbValue[invocation.Arguments.Length];
                                    for (int i = 0; i < invocation.Arguments.Length; i++)
                                    {
                                        var arg = invocation.Arguments[i];
                                        if (arg.RawValue != null && Marshal.IsComObject(arg.RawValue))
                                        {
                                            procArgs[i] = LiveControlProxy.ConvertComResultToVbValue(arg.RawValue, $"Arg{i}", Dispatcher);
                                        }
                                        else
                                        {
                                            procArgs[i] = arg;
                                        }
                                    }

                                    string procName = $"{ctrlName}_{invocation.EventName}";
                                    _host.InvokeProcedure(procName, modName, procArgs);

                                    if (!string.IsNullOrEmpty(parentControlPrefix))
                                    {
                                        _host.InvokeProcedure($"{parentControlPrefix}_{ctrlName}_{invocation.EventName}", _model.Name, procArgs);
                                    }

                                    // Propagar parámetros ByRef modificados de vuelta al OCX nativo
                                    for (int i = 0; i < invocation.Arguments.Length; i++)
                                    {
                                        if (invocation.IsByRef[i])
                                        {
                                            invocation.UpdateArgumentValue(i, procArgs[i]);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _host.Context.WriteDebugPrint($"[Error al ejecutar {ctrlName}_{invocation.EventName}]: {ex.Message}");
                                }
                            });
                        };
                    }
                    else
                    {
                        genericVisual.MouseLeftButtonDown += (s, e) =>
                        {
                            if (ctrl.Enabled)
                            {
                                if (e.ClickCount == 2)
                                {
                                    DispatchEvent("DblClick");
                                }
                                else
                                {
                                    DispatchEvent("Click");
                                }
                            }
                        };
                    }

                    genericVisual.IsEnabled = ctrl.Enabled;
                    genericVisual.Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden;
                }
                return genericVisual;
        }
    }

    private FrameworkElement CreateLiveUserControlVisual(FormDesignerControl ctrl, FormDesignModel ucModel)
    {
        var ucBorder = new Border
        {
            Background = VbColorHelper.ToBrush(ctrl.BackColorHex ?? ucModel.BackColorHex, SystemColors.ControlBrush),
            BorderBrush = new SolidColorBrush(Color.FromRgb(140, 140, 140)),
            BorderThickness = new Thickness(ctrl.BorderStyle == 0 ? 0 : 1),
            ClipToBounds = true,
            Visibility = ctrl.Visible ? Visibility.Visible : Visibility.Hidden,
            IsEnabled = ctrl.Enabled
        };

        var ucCanvas = new Canvas
        {
            ClipToBounds = true,
            Background = Brushes.Transparent
        };
        ucBorder.Child = ucCanvas;

        // Ambient y Extender específicos para esta instancia de UserControl
        var ambient = new VbUserControlAmbient(isUserMode: true, displayName: ctrl.Name);
        var extender = new VbUserControlExtender(ctrl);
        _host.Context.SetVariable("Ambient", VbValue.FromObject(ambient));
        _host.Context.SetVariable("Extender", VbValue.FromObject(extender));

        // 1. Inicialización
        _host.Context.ActiveControlInstanceName = ctrl.Name;
        _host.InvokeProcedure("UserControl_Initialize", ucModel.Name);
        _host.InvokeProcedure($"{ctrl.Name}_Initialize", _model.Name);

        // 2. Persistencia PropertyBag: ReadProperties o InitProperties
        if (ctrl.CustomProperties.Count > 0 || ctrl.BinaryProperties.Count > 0)
        {
            var propBag = new VbPropertyBagRuntime();
            foreach (var cp in ctrl.CustomProperties)
            {
                propBag.SetPropertyDirect(cp.Name, VbValue.FromString(cp.Value));
            }
            foreach (var bp in ctrl.BinaryProperties)
            {
                propBag.SetPropertyDirect(bp.Key, VbValue.FromObject(bp.Value));
            }
            _host.Context.SetVariable("PropBag", VbValue.FromObject(propBag));
            _host.InvokeProcedure("UserControl_ReadProperties", ucModel.Name, VbValue.FromObject(propBag));
        }
        else
        {
            _host.InvokeProcedure("UserControl_InitProperties", ucModel.Name);
        }

        // 3. Resize y Show inicial
        _host.InvokeProcedure("UserControl_Resize", ucModel.Name);
        _host.InvokeProcedure("UserControl_Show", ucModel.Name);
        _host.InvokeProcedure("UserControl_Paint", ucModel.Name);

        // Construir cada control constitutivo del UserControl dentro de su lienzo
        if (ucModel.Children.Count == 0)
        {
            // UserControl sin controles constitutivos (dibujado por código, GDI o skins)
            var genericVisual = ControlVisualRenderer.CreateRuntimeVisual(ctrl);
            if (genericVisual != null)
            {
                genericVisual.IsHitTestVisible = false;
                ucCanvas.Children.Add(genericVisual);
            }
        }
        else
        {
            foreach (var child in ucModel.Children)
            {
                if (child.TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase)) continue;

                var childVisual = CreateLiveControlVisual(child, ucModel.Name, ctrl.Name);
                if (childVisual != null)
                {
                    Canvas.SetLeft(childVisual, child.LeftPixels);
                    Canvas.SetTop(childVisual, child.TopPixels);
                    childVisual.Width = child.WidthPixels;
                    childVisual.Height = child.HeightPixels;
                    ucCanvas.Children.Add(childVisual);

                    // Registrar proxy constitutivo accesible tanto como uc1_btn como btn
                    var childProxy = new LiveControlProxy(child, childVisual);
                    _formProxy.RegisterControl(childProxy);
                    _host.Context.SetVariable($"{ctrl.Name}_{child.Name}", VbValue.FromObject(childProxy));
                    _host.Context.SetVariable(child.Name, VbValue.FromObject(childProxy));
                }
            }
        }

        // Renderizar controles hijos alojados dentro de esta instancia si actúa como contenedor
        if (ctrl.Children.Count > 0)
        {
            foreach (var hostedChild in ctrl.Children)
            {
                var hostedVisual = CreateLiveControlVisual(hostedChild);
                if (hostedVisual != null)
                {
                    Canvas.SetLeft(hostedVisual, hostedChild.LeftPixels);
                    Canvas.SetTop(hostedVisual, hostedChild.TopPixels);
                    hostedVisual.Width = hostedChild.WidthPixels;
                    hostedVisual.Height = hostedChild.HeightPixels;
                    ucCanvas.Children.Add(hostedVisual);
                }
            }
        }

        // Clic en el área del UserControl
        ucBorder.MouseDown += (s, e) =>
        {
            _host.Context.ActiveControlInstanceName = ctrl.Name;
            _host.InvokeProcedure($"{ctrl.Name}_Click", _model.Name);
            _host.InvokeProcedure("UserControl_Click", ucModel.Name);
        };

        return ucBorder;
    }

    private void WireEvents()
    {
        // Enlazar despachador global de RaiseEvent desde UserControls hacia el formulario
        _host.Context.OnRaiseEvent += (eventName, eventArgs, containerName, activeInstanceName) =>
        {
            Dispatcher.Invoke(() =>
            {
                try
                {
                    string targetInstance = !string.IsNullOrEmpty(activeInstanceName)
                        ? activeInstanceName
                        : _model.GetAllControls().FirstOrDefault(c => c.TypeName.Equals(containerName, StringComparison.OrdinalIgnoreCase))?.Name ?? string.Empty;

                    if (string.IsNullOrEmpty(targetInstance))
                    {
                        targetInstance = containerName ?? string.Empty;
                    }

                    var ctrl = _model.GetAllControls().Find(c => c.Name.Equals(targetInstance, StringComparison.OrdinalIgnoreCase));
                    VbValue[] finalArgs;
                    if (ctrl?.Index.HasValue == true)
                    {
                        finalArgs = new VbValue[eventArgs.Length + 1];
                        finalArgs[0] = VbValue.FromLong(ctrl.Index.Value);
                        Array.Copy(eventArgs, 0, finalArgs, 1, eventArgs.Length);
                    }
                    else
                    {
                        finalArgs = eventArgs;
                    }

                    string procName = $"{targetInstance}_{eventName}";
                    _host.InvokeProcedure(procName, _model.Name, finalArgs);
                }
                catch (Exception ex)
                {
                    _host.Context.WriteDebugPrint($"[Error al despachar RaiseEvent {eventName}]: {ex.Message}");
                }
            });
        };

        Loaded += (s, e) =>
        {
            if (_model.IsUserControl)
            {
                _host.InvokeProcedure("UserControl_Initialize", _model.Name);
                _host.InvokeProcedure("UserControl_Resize", _model.Name);
                _host.InvokeProcedure("UserControl_Show", _model.Name);
                _host.InvokeProcedure("UserControl_Paint", _model.Name);
            }
            else
            {
                _host.InvokeProcedure("Form_Load", _model.Name);
            }
        };

        FormCanvas.MouseDown += (s, e) =>
        {
            if (_model.IsUserControl)
            {
                _host.InvokeProcedure("UserControl_Click", _model.Name);
            }
            else
            {
                _host.InvokeProcedure("Form_Click", _model.Name);
            }
        };

        SizeChanged += (s, e) =>
        {
            if (_model.IsUserControl)
            {
                _host.InvokeProcedure("UserControl_Resize", _model.Name);
            }
            else
            {
                _host.InvokeProcedure("Form_Resize", _model.Name);
            }
        };

        Closing += (s, e) =>
        {
            if (_model.IsUserControl)
            {
                _host.InvokeProcedure("UserControl_Hide", _model.Name);
                _host.InvokeProcedure("UserControl_Terminate", _model.Name);
            }
            else
            {
                foreach (var child in _model.GetAllControls())
                {
                    if (_userControlResolver != null)
                    {
                        var ucModel = _userControlResolver(child.TypeName);
                        if (ucModel != null)
                        {
                            _host.InvokeProcedure("UserControl_Hide", ucModel.Name);
                            _host.InvokeProcedure("UserControl_Terminate", ucModel.Name);
                        }
                    }
                }
                _host.InvokeProcedure("Form_Unload", _model.Name, VbValue.FromLong(0));
            }
        };

        Closed += (s, e) =>
        {
            foreach (var timer in _activeTimers)
            {
                timer.Stop();
            }
            _activeTimers.Clear();

            foreach (var disp in _disposableRuntimeHosts)
            {
                try { disp.Dispose(); } catch { }
            }
            _disposableRuntimeHosts.Clear();

            _host.Context.ActiveForms.Remove(_model.Name);
            _host.Context.State = ExecutionState.Stopped;
        };
    }
}
