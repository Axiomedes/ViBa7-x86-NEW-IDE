using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using ViBa7.ProjectModel.Designers;
using ViBa7.ActiveXHost;

namespace ViBa7.IDE.Designer;

/// <summary>
/// Superficie de diseño interactiva para formularios de Visual Basic 6.0 con soporte completo para
/// selección simple/múltiple, 8 tiradores de redimensión, arrastre con ajuste a cuadrícula, selección por área y teclado.
/// </summary>
public partial class FormDesignSurface : UserControl
{
    private const double TitleBarHeight = 26.0;
    private const double WindowBorderThickness = 3.0;

    public FormDesignModel? CurrentModel { get; private set; }

    public event Action<FormDesignerControl>? ControlSelected;
    public event Action<List<FormDesignerControl>>? SelectionChanged;
    public event Action<FormDesignerControl>? ControlDoubleClicked;
    public event Action? FormDoubleClicked;
    public event Action<FormDesignerControl>? ControlAdded;
    public event Action<FormDesignerControl>? ControlDeleted;
    public event Action? ToolCreationFinished;
    public event Action? ModelModified;
    public event Action? RequestCodeView;
    public event Action? RequestMenuEditor;

    // Herramienta activa para creación desde la Toolbox
    public string? ActiveCreationTool { get; set; }
    private bool _isCreatingControl;
    private Point _creationStartPoint;

    // Colección de controles seleccionados actualmente
    private readonly HashSet<FormDesignerControl> _selectedControls = new();
    private FormDesignerControl? _primarySelectedControl;

    public IReadOnlyCollection<FormDesignerControl> SelectedControls => _selectedControls;
    public FormDesignerControl? SelectedControl => _primarySelectedControl;

    // Mapa de control a su elemento visual en WPF
    private readonly Dictionary<FormDesignerControl, FrameworkElement> _visualMap = new();

    // Estado de arrastre de controles (Drag & Move)
    private bool _isDraggingControls;
    private Point _dragStartPoint;
    private readonly Dictionary<FormDesignerControl, (double Left, double Top)> _initialPositions = new();

    // Estado de selección por área (Rubber-band marquee)
    private bool _isMarqueeSelecting;
    private Point _marqueeStartPoint;

    public FormDesignSurface()
    {
        InitializeComponent();
        Loaded += FormDesignSurface_Loaded;

        SelectionFrame.SnapFunction = (val) => GridCanvas.SnapValue(val);
        SelectionFrame.Resizing += (ctrl) =>
        {
            UpdateDimensionText();
        };
        SelectionFrame.ResizeCompleted += (ctrl) =>
        {
            ctrl.UpdateExtentsFromDimensions();
            if (_visualMap.TryGetValue(ctrl, out var v))
            {
                if (v is ActiveXInPlaceHost axHost)
                {
                    axHost.RefreshVisualSnapshot((int)Math.Round(ctrl.WidthPixels), (int)Math.Round(ctrl.HeightPixels));
                }
                else if (ControlVisualRenderer.UserControlModelResolver != null)
                {
                    string shortName = ctrl.TypeName.Contains('.')
                        ? ctrl.TypeName.Substring(ctrl.TypeName.LastIndexOf('.') + 1)
                        : ctrl.TypeName;
                    var ucModel = ControlVisualRenderer.UserControlModelResolver(ctrl.TypeName)
                               ?? ControlVisualRenderer.UserControlModelResolver(shortName);
                    if (ucModel != null && ucModel.Children.Count > 0)
                    {
                        RefreshSingleControlVisual(ctrl);
                    }
                }
            }
            UpdateDimensionText();
            ModelModified?.Invoke();
        };

        FormWindowBorder.SizeChanged += (s, e) => UpdateResizeThumbs();
        SetupCanvasContextMenu();
    }

    private void FormDesignSurface_Loaded(object sender, RoutedEventArgs e)
    {
        UpdateResizeThumbs();
    }

    private void BtnSwitchToCode_Click(object sender, RoutedEventArgs e)
    {
        RequestCodeView?.Invoke();
    }

    private void BtnOpenMenuEditor_Click(object sender, RoutedEventArgs e)
    {
        RequestMenuEditor?.Invoke();
    }

    /// <summary>
    /// Carga y visualiza el modelo de diseño del formulario.
    /// </summary>
    public void LoadModel(FormDesignModel model)
    {
        CurrentModel = model;
        ClearSelection();

        TxtFormCaption.Text = string.IsNullOrEmpty(model.Caption) ? model.Name : model.Caption;
        TxtFormTitle.Text = model.IsUserControl ? $"Diseñador UserControl: {model.Name}" : $"Diseñador: {model.Name}";

        // Diferenciar entre Formulario estándar y UserControl (sin barra de título de ventana)
        if (!model.HasTitleBar)
        {
            TitleBarRow.Height = new GridLength(0);
            BorderTitleBar.Visibility = Visibility.Collapsed;
            FormWindowBorder.BorderThickness = new Thickness(1);
            FormWindowBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(128, 128, 128));
        }
        else
        {
            TitleBarRow.Height = new GridLength(26);
            BorderTitleBar.Visibility = Visibility.Visible;
            FormWindowBorder.BorderThickness = new Thickness(3);
            FormWindowBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(10, 36, 106));
        }

        // Configurar dimensiones del área cliente
        GridCanvas.Width = model.ClientWidthPixels;
        GridCanvas.Height = model.ClientHeightPixels;

        // Configurar color de fondo del formulario
        FormClientArea.Background = VbColorHelper.ToBrush(
            model.BackColorHex,
            new SolidColorBrush(Color.FromRgb(240, 240, 240))
        );

        UpdateDimensionText();
        RebuildMenuBar();
        RefreshCanvas();
        UpdateResizeThumbs();
    }

    /// <summary>
    /// Sincroniza el estado de persistencia de todos los controles ActiveX alojados hacia el modelo de diseño.
    /// </summary>
    public void SyncAllActiveXHostsToModel()
    {
        foreach (var kv in _visualMap)
        {
            if (kv.Value is ActiveXInPlaceHost axHost)
            {
                axHost.SyncToDesignerModel();
            }
            else if (kv.Value is Grid g && g.Children.Count > 0 && g.Children[0] is ActiveXInPlaceHost gridHost)
            {
                gridHost.SyncToDesignerModel();
            }
        }
    }

    /// <summary>
    /// Regenera todos los elementos visuales de los controles en el lienzo.
    /// </summary>
    public void RefreshCanvas()
    {
        // Guardar elementos especiales (Marquee y SelectionFrame)
        var specialElements = new List<UIElement> { MarqueeRectangle, SelectionFrame };

        if (CurrentModel == null)
        {
            foreach (var kv in _visualMap)
            {
                if (kv.Value is IDisposable disposable) disposable.Dispose();
            }
            var toRemoveAll = GridCanvas.Children.Cast<UIElement>().Where(el => !specialElements.Contains(el)).ToList();
            foreach (var el in toRemoveAll) GridCanvas.Children.Remove(el);
            _visualMap.Clear();
            UpdateSelectionFrame();
            return;
        }

        var currentControls = CurrentModel.NonMenuControls.ToHashSet();

        // 1. Liberar y remover controles que ya no existen en el modelo
        var removedControls = _visualMap.Keys.Where(c => !currentControls.Contains(c)).ToList();
        foreach (var rem in removedControls)
        {
            if (_visualMap.TryGetValue(rem, out var remVisual))
            {
                if (remVisual is IDisposable disposable) disposable.Dispose();
                GridCanvas.Children.Remove(remVisual);
                _visualMap.Remove(rem);
            }
        }

        // 2. Para cada control del modelo: si ya tiene un ActiveXInPlaceHost existente, reutilizarlo y actualizarlo.
        // Si no existe o es un control visual de WPF, crearlo/refrescarlo.
        foreach (var control in CurrentModel.NonMenuControls)
        {
            if (_visualMap.TryGetValue(control, out var existingVisual))
            {
                ActiveXInPlaceHost? axHost = existingVisual as ActiveXInPlaceHost
                    ?? (existingVisual is Grid g && g.Children.Count > 0 ? g.Children[0] as ActiveXInPlaceHost : null);

                if (axHost != null)
                {
                    Canvas.SetLeft(existingVisual, control.LeftPixels);
                    Canvas.SetTop(existingVisual, control.TopPixels);
                    existingVisual.Width = control.WidthPixels;
                    existingVisual.Height = control.HeightPixels;
                    axHost.Width = control.WidthPixels;
                    axHost.Height = control.HeightPixels;
                    axHost.SyncFromDesignerModel();
                    if (!GridCanvas.Children.Contains(existingVisual))
                    {
                        GridCanvas.Children.Add(existingVisual);
                    }
                    continue;
                }
                else
                {
                    // Control WPF: reemplazar visual
                    if (existingVisual is IDisposable disposable) disposable.Dispose();
                    GridCanvas.Children.Remove(existingVisual);
                }
            }

            var visual = ControlVisualRenderer.CreateVisual(control, OnChildVisualCreated);
            AttachControlEvents(visual, control);
            _visualMap[control] = visual;
            GridCanvas.Children.Add(visual);
        }

        // Actualizar marco de selección
        UpdateSelectionFrame();
    }

    /// <summary>
    /// Regenera únicamente el elemento visual de un control individual sin reconstruir todo el lienzo,
    /// evitando parpadeos (flicker) en el resto de los controles del formulario.
    /// Para controles ActiveX in-place, preserva la instancia COM nativa y solo sincroniza propiedades y dimensiones.
    /// </summary>
    public void RefreshSingleControlVisual(FormDesignerControl control)
    {
        if (!_visualMap.TryGetValue(control, out var oldVisual)) return;

        // Si es un control ActiveX in-place (o un contenedor Grid con ActiveXInPlaceHost), no destruir ni recrear
        // la instancia COM nativa; simplemente actualizar dimensiones, sincronizar propiedades y refrescar la captura.
        ActiveXInPlaceHost? axHost = oldVisual as ActiveXInPlaceHost
            ?? (oldVisual is Grid g && g.Children.Count > 0 ? g.Children[0] as ActiveXInPlaceHost : null);

        if (axHost != null)
        {
            Canvas.SetLeft(oldVisual, control.LeftPixels);
            Canvas.SetTop(oldVisual, control.TopPixels);
            oldVisual.Width = control.WidthPixels;
            oldVisual.Height = control.HeightPixels;
            axHost.Width = control.WidthPixels;
            axHost.Height = control.HeightPixels;

            axHost.SyncFromDesignerModel();
            UpdateSelectionFrame();
            return;
        }

        if (oldVisual.Parent is Panel parentPanel)
        {
            int idx = parentPanel.Children.IndexOf(oldVisual);
            if (oldVisual is IDisposable disposable)
            {
                disposable.Dispose();
            }

            var newVisual = ControlVisualRenderer.CreateVisual(control, OnChildVisualCreated);
            AttachControlEvents(newVisual, control);
            _visualMap[control] = newVisual;

            parentPanel.Children.RemoveAt(idx);
            parentPanel.Children.Insert(idx, newVisual);

            UpdateSelectionFrame();
        }
        else
        {
            RefreshCanvas();
        }
    }

    private void OnChildVisualCreated(FormDesignerControl childControl, FrameworkElement childVisual)
    {
        _visualMap[childControl] = childVisual;
        AttachControlEvents(childVisual, childControl);
    }

    /// <summary>
    /// Retorna el elemento visual de WPF asociado a un control de diseño.
    /// </summary>
    public FrameworkElement? GetVisualForControl(FormDesignerControl control)
    {
        return _visualMap.TryGetValue(control, out var v) ? v : null;
    }

    private void AttachControlEvents(FrameworkElement visual, FormDesignerControl control)
    {
        if (visual is ActiveXInPlaceHost axHost)
        {
            axHost.HostSelected += (host, pt) =>
            {
                Dispatcher.Invoke(() =>
                {
                    GridCanvas.Focus();
                    bool isShiftOrCtrl = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift) ||
                                         Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

                    if (isShiftOrCtrl)
                    {
                        ToggleControlSelection(control);
                    }
                    else
                    {
                        if (!_selectedControls.Contains(control))
                        {
                            SelectSingleControl(control);
                        }
                        else
                        {
                            _primarySelectedControl = control;
                            UpdateSelectionFrame();
                        }
                    }
                });
            };

            axHost.HostDoubleClicked += (host, pt) =>
            {
                Dispatcher.Invoke(() => ControlDoubleClicked?.Invoke(control));
            };

            axHost.HostContextMenuRequested += (host, pt) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (GridCanvas.ContextMenu != null)
                    {
                        GridCanvas.ContextMenu.PlacementTarget = visual;
                        GridCanvas.ContextMenu.IsOpen = true;
                    }
                });
            };

            axHost.HostDragStarted += (host, pt) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (_selectedControls.Count > 0)
                    {
                        _isDraggingControls = true;
                        _dragStartPoint = visual.TransformToAncestor(GridCanvas).Transform(pt);

                        _initialPositions.Clear();
                        foreach (var sel in _selectedControls)
                        {
                            _initialPositions[sel] = (sel.LeftPixels, sel.TopPixels);
                        }
                    }
                });
            };

            axHost.HostDragging += (host, pt) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (_isDraggingControls && _selectedControls.Count > 0)
                    {
                        var currentCanvasPos = visual.TransformToAncestor(GridCanvas).Transform(pt);
                        double deltaX = currentCanvasPos.X - _dragStartPoint.X;
                        double deltaY = currentCanvasPos.Y - _dragStartPoint.Y;

                        foreach (var sel in _selectedControls)
                        {
                            if (_initialPositions.TryGetValue(sel, out var initPos))
                            {
                                double newLeft = GridCanvas.SnapValue(initPos.Left + deltaX);
                                double newTop = GridCanvas.SnapValue(initPos.Top + deltaY);
                                sel.LeftPixels = Math.Max(0, newLeft);
                                sel.TopPixels = Math.Max(0, newTop);

                                if (_visualMap.TryGetValue(sel, out var v))
                                {
                                    Canvas.SetLeft(v, sel.LeftPixels);
                                    Canvas.SetTop(v, sel.TopPixels);
                                }
                            }
                        }

                        UpdateSelectionFrame();
                        UpdateDimensionText();
                    }
                });
            };

            axHost.HostDragEnded += (host) =>
            {
                Dispatcher.Invoke(() =>
                {
                    if (_isDraggingControls)
                    {
                        _isDraggingControls = false;
                        ModelModified?.Invoke();
                    }
                });
            };

            axHost.HostPropertyChangedByOcx += (host, dispId, propName) =>
            {
                Dispatcher.Invoke(() =>
                {
                    ModelModified?.Invoke();
                });
            };

            axHost.HostExtenderBoundsChanged += (host) =>
            {
                Dispatcher.Invoke(() =>
                {
                    Canvas.SetLeft(visual, control.LeftPixels);
                    Canvas.SetTop(visual, control.TopPixels);
                    visual.Width = control.WidthPixels;
                    visual.Height = control.HeightPixels;
                    axHost.RefreshVisualSnapshot((int)Math.Round(control.WidthPixels), (int)Math.Round(control.HeightPixels));
                    UpdateSelectionFrame();
                    UpdateDimensionText();
                    ModelModified?.Invoke();
                });
            };
        }

        visual.MouseLeftButtonDown += (s, e) =>
        {
            GridCanvas.Focus();

            // Si hay una herramienta de creación activa en la Toolbox (distinta del puntero), permitir dibujar en este control/contenedor
            if (!string.IsNullOrEmpty(ActiveCreationTool) && ActiveCreationTool != ControlToolboxRegistry.PointerTool)
            {
                GridCanvas_MouseLeftButtonDown(GridCanvas, e);
                return;
            }

            if (e.ClickCount == 2)
            {
                ControlDoubleClicked?.Invoke(control);
                e.Handled = true;
                return;
            }

            bool isShiftOrCtrl = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift) ||
                                 Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

            if (isShiftOrCtrl)
            {
                ToggleControlSelection(control);
            }
            else
            {
                if (!_selectedControls.Contains(control))
                {
                    SelectSingleControl(control);
                }
                else
                {
                    _primarySelectedControl = control;
                    UpdateSelectionFrame();
                }
            }

            // Iniciar arrastre si hay controles seleccionados
            if (_selectedControls.Count > 0)
            {
                _isDraggingControls = true;
                _dragStartPoint = e.GetPosition(GridCanvas);

                _initialPositions.Clear();
                foreach (var sel in _selectedControls)
                {
                    _initialPositions[sel] = (sel.LeftPixels, sel.TopPixels);
                }

                GridCanvas.CaptureMouse();
            }

            e.Handled = true;
        };

        visual.MouseRightButtonDown += (s, e) =>
        {
            if (!_selectedControls.Contains(control))
            {
                SelectSingleControl(control);
            }
        };
    }

    #region Gestión de Selección
    private void ExitActiveXInteractiveEditMode(FormDesignerControl? exceptControl = null)
    {
        bool exitedAny = false;
        foreach (var kv in _visualMap)
        {
            if (kv.Key == exceptControl) continue;

            ActiveXInPlaceHost? axHost = kv.Value as ActiveXInPlaceHost;
            if (axHost == null && kv.Value is Grid g && g.Children.Count > 0)
            {
                axHost = g.Children[0] as ActiveXInPlaceHost;
            }

            if (axHost is { IsInteractiveEditMode: true })
            {
                axHost.ExitInteractiveEditMode();
                exitedAny = true;
            }
        }

        if (exitedAny)
        {
            ModelModified?.Invoke();
        }
    }

    public void SelectSingleControl(FormDesignerControl control)
    {
        ExitActiveXInteractiveEditMode(control);
        _selectedControls.Clear();
        _selectedControls.Add(control);
        _primarySelectedControl = control;

        UpdateSelectionFrame();
        NotifySelectionChanged();
    }

    public void ToggleControlSelection(FormDesignerControl control)
    {
        if (_selectedControls.Contains(control))
        {
            _selectedControls.Remove(control);
            if (_primarySelectedControl == control)
            {
                _primarySelectedControl = _selectedControls.LastOrDefault();
            }
        }
        else
        {
            _selectedControls.Add(control);
            _primarySelectedControl = control;
        }

        UpdateSelectionFrame();
        NotifySelectionChanged();
    }

    public void ClearSelection()
    {
        ExitActiveXInteractiveEditMode(null);
        _selectedControls.Clear();
        _primarySelectedControl = null;

        UpdateSelectionFrame();
        NotifySelectionChanged();
    }

    public void SelectAllControls()
    {
        if (CurrentModel == null) return;

        _selectedControls.Clear();
        foreach (var ctrl in CurrentModel.GetAllControls())
        {
            _selectedControls.Add(ctrl);
        }

        _primarySelectedControl = _selectedControls.FirstOrDefault();
        UpdateSelectionFrame();
        NotifySelectionChanged();
    }

    private void UpdateSelectionFrame()
    {
        if (_primarySelectedControl != null && _visualMap.TryGetValue(_primarySelectedControl, out var visual))
        {
            SelectionFrame.BindTo(_primarySelectedControl, visual);
        }
        else
        {
            SelectionFrame.BindTo(null, null);
        }
    }

    private void NotifySelectionChanged()
    {
        if (_primarySelectedControl != null)
        {
            TxtDesignerStatus.Text = _selectedControls.Count == 1
                ? $"Seleccionado: {_primarySelectedControl.Name} ({_primarySelectedControl.DisplayType}) | Pos: X={_primarySelectedControl.LeftPixels:F0}, Y={_primarySelectedControl.TopPixels:F0} | Dim: {_primarySelectedControl.WidthPixels:F0} × {_primarySelectedControl.HeightPixels:F0} px"
                : $"{_selectedControls.Count} controles seleccionados. Primario: {_primarySelectedControl.Name}";

            ControlSelected?.Invoke(_primarySelectedControl);
        }
        else
        {
            TxtDesignerStatus.Text = "Formulario seleccionado.";
        }

        SelectionChanged?.Invoke(_selectedControls.ToList());
    }
    #endregion

    #region Manejo del Ratón en GridCanvas (Mover y Selección por Área)
    private void GridCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        GridCanvas.Focus();

        if (e.ClickCount == 2)
        {
            FormDoubleClicked?.Invoke();
            e.Handled = true;
            return;
        }

        // 1. Si hay una herramienta activa de Toolbox, iniciar creación de control
        if (!string.IsNullOrEmpty(ActiveCreationTool) && ActiveCreationTool != ControlToolboxRegistry.PointerTool)
        {
            ClearSelection();
            _isCreatingControl = true;
            _creationStartPoint = e.GetPosition(GridCanvas);

            Canvas.SetLeft(MarqueeRectangle, _creationStartPoint.X);
            Canvas.SetTop(MarqueeRectangle, _creationStartPoint.Y);
            MarqueeRectangle.Width = 0;
            MarqueeRectangle.Height = 0;
            MarqueeRectangle.Visibility = Visibility.Visible;

            GridCanvas.CaptureMouse();
            e.Handled = true;
            return;
        }

        // Clic en espacio vacío: deseleccionar o iniciar selección por área
        bool isShiftOrCtrl = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift) ||
                             Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

        if (!isShiftOrCtrl)
        {
            ClearSelection();
        }

        // Iniciar Rubber-band Marquee
        _isMarqueeSelecting = true;
        _marqueeStartPoint = e.GetPosition(GridCanvas);

        Canvas.SetLeft(MarqueeRectangle, _marqueeStartPoint.X);
        Canvas.SetTop(MarqueeRectangle, _marqueeStartPoint.Y);
        MarqueeRectangle.Width = 0;
        MarqueeRectangle.Height = 0;
        MarqueeRectangle.Visibility = Visibility.Visible;

        GridCanvas.CaptureMouse();
    }

    private void GridCanvas_MouseMove(object sender, MouseEventArgs e)
    {
        Point curPoint = e.GetPosition(GridCanvas);

        // 0. Creación interactiva de control desde Toolbox
        if (_isCreatingControl)
        {
            double x = Math.Min(_creationStartPoint.X, curPoint.X);
            double y = Math.Min(_creationStartPoint.Y, curPoint.Y);
            double w = Math.Abs(curPoint.X - _creationStartPoint.X);
            double h = Math.Abs(curPoint.Y - _creationStartPoint.Y);

            Canvas.SetLeft(MarqueeRectangle, x);
            Canvas.SetTop(MarqueeRectangle, y);
            MarqueeRectangle.Width = w;
            MarqueeRectangle.Height = h;
            return;
        }

        // Si hay una herramienta activa, cursor en cruz
        if (!string.IsNullOrEmpty(ActiveCreationTool) && ActiveCreationTool != ControlToolboxRegistry.PointerTool)
        {
            GridCanvas.Cursor = Cursors.Cross;
        }
        else
        {
            GridCanvas.Cursor = Cursors.Arrow;
        }

        // 1. Arrastre de controles seleccionados
        if (_isDraggingControls && _selectedControls.Count > 0)
        {
            double rawDeltaX = curPoint.X - _dragStartPoint.X;
            double rawDeltaY = curPoint.Y - _dragStartPoint.Y;

            double snappedDeltaX = GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawDeltaX) : rawDeltaX;
            double snappedDeltaY = GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawDeltaY) : rawDeltaY;

            foreach (var ctrl in _selectedControls)
            {
                if (_initialPositions.TryGetValue(ctrl, out var initial))
                {
                    double newLeft = Math.Max(0, initial.Left + snappedDeltaX);
                    double newTop = Math.Max(0, initial.Top + snappedDeltaY);

                    if (GridCanvas.SnapToGrid)
                    {
                        newLeft = GridCanvas.SnapValue(newLeft);
                        newTop = GridCanvas.SnapValue(newTop);
                    }

                    ctrl.LeftPixels = newLeft;
                    ctrl.TopPixels = newTop;

                    if (_visualMap.TryGetValue(ctrl, out var vis))
                    {
                        Canvas.SetLeft(vis, newLeft);
                        Canvas.SetTop(vis, newTop);
                    }
                }
            }

            SelectionFrame.UpdateLayoutFromControl();
            if (_primarySelectedControl != null)
            {
                TxtDesignerStatus.Text = $"Moviendo {_selectedControls.Count} control(es) | X={_primarySelectedControl.LeftPixels:F0}, Y={_primarySelectedControl.TopPixels:F0}";
            }
            return;
        }

        // 2. Selección por área (Marquee)
        if (_isMarqueeSelecting)
        {
            double x = Math.Min(_marqueeStartPoint.X, curPoint.X);
            double y = Math.Min(_marqueeStartPoint.Y, curPoint.Y);
            double w = Math.Abs(curPoint.X - _marqueeStartPoint.X);
            double h = Math.Abs(curPoint.Y - _marqueeStartPoint.Y);

            Canvas.SetLeft(MarqueeRectangle, x);
            Canvas.SetTop(MarqueeRectangle, y);
            MarqueeRectangle.Width = w;
            MarqueeRectangle.Height = h;
        }
    }

    private void GridCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // 0. Finalizar creación de control desde Toolbox
        if (_isCreatingControl)
        {
            _isCreatingControl = false;
            MarqueeRectangle.Visibility = Visibility.Collapsed;
            GridCanvas.ReleaseMouseCapture();

            Point endPoint = e.GetPosition(GridCanvas);
            double rawW = Math.Abs(endPoint.X - _creationStartPoint.X);
            double rawH = Math.Abs(endPoint.Y - _creationStartPoint.Y);
            double left = Math.Min(_creationStartPoint.X, endPoint.X);
            double top = Math.Min(_creationStartPoint.Y, endPoint.Y);

            if (GridCanvas.SnapToGrid)
            {
                left = GridCanvas.SnapValue(left);
                top = GridCanvas.SnapValue(top);
                rawW = GridCanvas.SnapValue(rawW);
                rawH = GridCanvas.SnapValue(rawH);
            }

            int leftTwips = (int)Math.Round(left * FormDesignerControl.TwipsPerPixel);
            int topTwips = (int)Math.Round(top * FormDesignerControl.TwipsPerPixel);
            int? widthTwips = (rawW > 15) ? (int)Math.Round(rawW * FormDesignerControl.TwipsPerPixel) : null;
            int? heightTwips = (rawH > 15) ? (int)Math.Round(rawH * FormDesignerControl.TwipsPerPixel) : null;

            if (CurrentModel != null && !string.IsNullOrEmpty(ActiveCreationTool))
            {
                var targetContainer = FindContainerAt(_creationStartPoint);
                FormDesignerControl newCtrl;

                if (targetContainer != null)
                {
                    Point innerOrigin = GetContainerInnerOrigin(targetContainer);
                    double relLeft = Math.Max(0, left - innerOrigin.X);
                    double relTop = Math.Max(0, top - innerOrigin.Y);

                    if (GridCanvas.SnapToGrid)
                    {
                        relLeft = GridCanvas.SnapValue(relLeft);
                        relTop = GridCanvas.SnapValue(relTop);
                    }

                    int leftTw = (int)Math.Round(relLeft * FormDesignerControl.TwipsPerPixel);
                    int topTw = (int)Math.Round(relTop * FormDesignerControl.TwipsPerPixel);

                    newCtrl = ControlToolboxRegistry.CreateControl(ActiveCreationTool, CurrentModel, leftTw, topTw, widthTwips, heightTwips);
                    newCtrl.ParentControl = targetContainer;
                    if (newCtrl.IsNonResizable)
                    {
                        newCtrl.Width = 480;
                        newCtrl.Height = 480;
                    }
                    targetContainer.Children.Add(newCtrl);
                }
                else
                {
                    newCtrl = ControlToolboxRegistry.CreateControl(ActiveCreationTool, CurrentModel, leftTwips, topTwips, widthTwips, heightTwips);
                    newCtrl.ParentControl = null;
                    if (newCtrl.IsNonResizable)
                    {
                        newCtrl.Width = 480;
                        newCtrl.Height = 480;
                    }
                    CurrentModel.Children.Add(newCtrl);
                }

                RefreshCanvas();
                SelectSingleControl(newCtrl);
                ControlAdded?.Invoke(newCtrl);
                ModelModified?.Invoke();
            }

            ActiveCreationTool = null;
            GridCanvas.Cursor = Cursors.Arrow;
            ToolCreationFinished?.Invoke();
            e.Handled = true;
            return;
        }

        if (_isDraggingControls)
        {
            _isDraggingControls = false;
            GridCanvas.ReleaseMouseCapture();
            CheckAndReparentDraggedControls();
            NotifySelectionChanged();
            ModelModified?.Invoke();
        }

        if (_isMarqueeSelecting)
        {
            _isMarqueeSelecting = false;
            MarqueeRectangle.Visibility = Visibility.Collapsed;
            GridCanvas.ReleaseMouseCapture();

            double mx = Canvas.GetLeft(MarqueeRectangle);
            double my = Canvas.GetTop(MarqueeRectangle);
            var marqueeRect = new Rect(mx, my, MarqueeRectangle.Width, MarqueeRectangle.Height);

            if (marqueeRect.Width > 4 && marqueeRect.Height > 4 && CurrentModel != null)
            {
                foreach (var ctrl in CurrentModel.GetAllControls())
                {
                    var ctrlRect = new Rect(ctrl.LeftPixels, ctrl.TopPixels, ctrl.WidthPixels, ctrl.HeightPixels);
                    if (marqueeRect.IntersectsWith(ctrlRect))
                    {
                        _selectedControls.Add(ctrl);
                        _primarySelectedControl = ctrl;
                    }
                }

                UpdateSelectionFrame();
                NotifySelectionChanged();
            }
        }
    }
    #endregion

    #region Detección de Contenedores (Frame, PictureBox) y Operaciones de Portapapeles
    private FormDesignerControl? FindContainerAt(Point canvasPoint, IEnumerable<FormDesignerControl>? exclude = null)
    {
        if (CurrentModel == null) return null;

        var excludeSet = exclude != null ? new HashSet<FormDesignerControl>(exclude) : null;
        FormDesignerControl? bestContainer = null;
        int bestDepth = -1;

        void SearchInList(IEnumerable<FormDesignerControl> controls, int depth)
        {
            foreach (var ctrl in controls)
            {
                if (excludeSet != null && excludeSet.Contains(ctrl))
                {
                    continue;
                }

                if (ctrl.IsContainer)
                {
                    double ox = ctrl.AbsoluteLeftPixels;
                    double oy = ctrl.AbsoluteTopPixels;
                    var bounds = new Rect(ox, oy, Math.Max(10, ctrl.WidthPixels), Math.Max(10, ctrl.HeightPixels));
                    if (bounds.Contains(canvasPoint))
                    {
                        if (depth > bestDepth)
                        {
                            bestDepth = depth;
                            bestContainer = ctrl;
                        }
                    }
                }

                if (ctrl.Children.Count > 0)
                {
                    SearchInList(ctrl.Children, depth + 1);
                }
            }
        }

        SearchInList(CurrentModel.Children, 0);
        return bestContainer;
    }

    private Point GetContainerInnerOrigin(FormDesignerControl container)
    {
        double baseLeft = container.AbsoluteLeftPixels;
        double baseTop = container.AbsoluteTopPixels;

        string normType = container.TypeName.ToUpperInvariant();
        if (normType.StartsWith("VB.", StringComparison.OrdinalIgnoreCase))
        {
            normType = normType[3..];
        }

        double fallbackOffsetX = 2.0;
        double fallbackOffsetY = 2.0;

        if (normType == "FRAME")
        {
            // Un GroupBox en WPF tiene bisel exterior y encabezado de texto
            fallbackOffsetX = 4.0;
            fallbackOffsetY = 16.0;
        }
        else if (normType == "PICTUREBOX" || container.IsContainer)
        {
            fallbackOffsetX = (container.BorderStyle == 0) ? 0.0 : 2.0;
            fallbackOffsetY = (container.BorderStyle == 0) ? 0.0 : 2.0;
        }

        if (_visualMap.TryGetValue(container, out var containerVis))
        {
            Canvas? innerCanvas = null;
            if (containerVis is GroupBox gb && gb.Content is Canvas c1)
            {
                innerCanvas = c1;
            }
            else if (containerVis is Border b && b.Child is Canvas c2)
            {
                innerCanvas = c2;
            }

            if (innerCanvas != null && innerCanvas.IsLoaded)
            {
                try
                {
                    Point pt = innerCanvas.TranslatePoint(new Point(0, 0), GridCanvas);
                    if (pt.X >= baseLeft && pt.Y >= baseTop &&
                        pt.X < baseLeft + container.WidthPixels && pt.Y < baseTop + container.HeightPixels)
                    {
                        return pt;
                    }
                }
                catch { }
            }
        }

        return new Point(baseLeft + fallbackOffsetX, baseTop + fallbackOffsetY);
    }

    private void CheckAndReparentDraggedControls()
    {
        if (CurrentModel == null || _selectedControls.Count == 0) return;

        bool reparentedAny = false;

        foreach (var ctrl in _selectedControls.ToList())
        {
            // Colección de controles excluidos (el propio control y toda su descendencia)
            var exclude = new HashSet<FormDesignerControl>();
            CollectDescendants(ctrl, exclude);

            // Centro geométrico absoluto del control en GridCanvas
            Point center = new Point(
                ctrl.AbsoluteLeftPixels + (ctrl.WidthPixels / 2.0),
                ctrl.AbsoluteTopPixels + (ctrl.HeightPixels / 2.0)
            );

            // Buscar contenedor en esa posición absoluta
            var targetContainer = FindContainerAt(center, exclude);

            if (targetContainer != ctrl.ParentControl)
            {
                double currentAbsLeft = ctrl.AbsoluteLeftPixels;
                double currentAbsTop = ctrl.AbsoluteTopPixels;

                // 1. Remover de su contenedor/padre actual
                if (ctrl.ParentControl != null)
                {
                    ctrl.ParentControl.Children.Remove(ctrl);
                }
                else
                {
                    CurrentModel.Children.Remove(ctrl);
                }

                // 2. Insertar en el nuevo contenedor o en el lienzo raíz
                if (targetContainer != null)
                {
                    ctrl.ParentControl = targetContainer;
                    if (!targetContainer.Children.Contains(ctrl))
                    {
                        targetContainer.Children.Add(ctrl);
                    }

                    Point innerOrigin = GetContainerInnerOrigin(targetContainer);
                    double relLeft = Math.Max(0, currentAbsLeft - innerOrigin.X);
                    double relTop = Math.Max(0, currentAbsTop - innerOrigin.Y);

                    if (GridCanvas.SnapToGrid)
                    {
                        relLeft = GridCanvas.SnapValue(relLeft);
                        relTop = GridCanvas.SnapValue(relTop);
                    }

                    ctrl.LeftPixels = relLeft;
                    ctrl.TopPixels = relTop;
                }
                else
                {
                    ctrl.ParentControl = null;
                    if (!CurrentModel.Children.Contains(ctrl))
                    {
                        CurrentModel.Children.Add(ctrl);
                    }

                    if (GridCanvas.SnapToGrid)
                    {
                        currentAbsLeft = GridCanvas.SnapValue(currentAbsLeft);
                        currentAbsTop = GridCanvas.SnapValue(currentAbsTop);
                    }

                    ctrl.LeftPixels = Math.Max(0, currentAbsLeft);
                    ctrl.TopPixels = Math.Max(0, currentAbsTop);
                }

                reparentedAny = true;
            }
        }

        if (reparentedAny)
        {
            RefreshCanvas();
        }
    }

    private static void CollectDescendants(FormDesignerControl ctrl, HashSet<FormDesignerControl> set)
    {
        set.Add(ctrl);
        foreach (var child in ctrl.Children)
        {
            CollectDescendants(child, set);
        }
    }

    public void CopySelectedControls()
    {
        if (_selectedControls.Count == 0) return;
        try
        {
            string serialized = FormDesignSerializer.SerializeControls(_selectedControls);
            Clipboard.SetText(serialized);
            TxtDesignerStatus.Text = $"{_selectedControls.Count} control(es) copiado(s) al portapapeles.";
        }
        catch (Exception ex)
        {
            TxtDesignerStatus.Text = $"Error al copiar: {ex.Message}";
        }
    }

    public void CutSelectedControls()
    {
        if (_selectedControls.Count == 0 || CurrentModel == null) return;
        try
        {
            CopySelectedControls();
            var controlsToDelete = _selectedControls.ToList();
            foreach (var ctrl in controlsToDelete)
            {
                CurrentModel.RemoveControl(ctrl);
                ControlDeleted?.Invoke(ctrl);
            }
            ClearSelection();
            RefreshCanvas();
            ModelModified?.Invoke();
            TxtDesignerStatus.Text = $"{controlsToDelete.Count} control(es) cortado(s).";
        }
        catch (Exception ex)
        {
            TxtDesignerStatus.Text = $"Error al cortar: {ex.Message}";
        }
    }

    public void PasteControls()
    {
        if (CurrentModel == null) return;
        try
        {
            if (!Clipboard.ContainsText()) return;
            string text = Clipboard.GetText();
            var parsedList = FormDesignParser.ParseControlSnippet(text);
            if (parsedList.Count == 0) return;

            // Determinar si hay un contenedor seleccionado para pegar dentro de él
            FormDesignerControl? targetContainer = null;
            if (_selectedControls.Count == 1 && _primarySelectedControl != null)
            {
                if (_primarySelectedControl.IsContainer)
                {
                    targetContainer = _primarySelectedControl;
                }
                else if (_primarySelectedControl.ParentControl != null)
                {
                    targetContainer = _primarySelectedControl.ParentControl;
                }
            }

            _selectedControls.Clear();
            foreach (var ctrl in parsedList)
            {
                // Desplazamiento sutil para no superponer exactamente
                ctrl.Left += 120; // 8 px
                ctrl.Top += 120;

                // Verificar si ya existe un control con este nombre (Matriz de controles como en VB6)
                var existingControls = CurrentModel.GetAllControls()
                    .Where(c => string.Equals(c.Name, ctrl.Name, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (existingControls.Count > 0)
                {
                    var msgRes = MessageBox.Show(
                        $"Ya tiene un control llamado '{ctrl.Name}'. ¿Desea crear una matriz de controles?",
                        "Microsoft Visual Basic",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (msgRes == MessageBoxResult.Yes)
                    {
                        var firstExisting = existingControls[0];
                        if (!firstExisting.Index.HasValue)
                        {
                            firstExisting.Index = 0;
                        }

                        var usedIndices = existingControls.Select(c => c.Index ?? 0).ToHashSet();
                        int nextIdx = 0;
                        while (usedIndices.Contains(nextIdx)) nextIdx++;

                        ctrl.Index = nextIdx;
                    }
                    else
                    {
                        string basePrefix = System.Text.RegularExpressions.Regex.Replace(ctrl.Name, @"\d+$", "");
                        if (string.IsNullOrWhiteSpace(basePrefix)) basePrefix = ctrl.DisplayType;
                        ctrl.Name = CurrentModel.GenerateUniqueName(basePrefix);
                        ctrl.Index = null;
                    }
                }
                else if (ctrl.Index.HasValue)
                {
                    ctrl.Index = null;
                }

                if (ctrl.IsNonResizable)
                {
                    ctrl.Width = 480;
                    ctrl.Height = 480;
                }

                if (targetContainer != null)
                {
                    ctrl.ParentControl = targetContainer;
                    targetContainer.Children.Add(ctrl);
                }
                else
                {
                    ctrl.ParentControl = null;
                    CurrentModel.Children.Add(ctrl);
                }

                _selectedControls.Add(ctrl);
                ControlAdded?.Invoke(ctrl);
            }

            _primarySelectedControl = parsedList.LastOrDefault();
            RefreshCanvas();
            UpdateSelectionFrame();
            NotifySelectionChanged();
            ModelModified?.Invoke();
            TxtDesignerStatus.Text = $"{parsedList.Count} control(es) pegado(s).";
        }
        catch (Exception ex)
        {
            TxtDesignerStatus.Text = $"Error al pegar: {ex.Message}";
        }
    }

    public void DuplicateSelectedControls()
    {
        if (_selectedControls.Count == 0 || CurrentModel == null) return;
        CopySelectedControls();
        PasteControls();
    }

    public void DeleteSelectedControls()
    {
        if (_selectedControls.Count == 0 || CurrentModel == null) return;
        var controlsToDelete = _selectedControls.ToList();
        foreach (var ctrl in controlsToDelete)
        {
            CurrentModel.RemoveControl(ctrl);
            ControlDeleted?.Invoke(ctrl);
        }

        ClearSelection();
        RefreshCanvas();
        ModelModified?.Invoke();
        TxtDesignerStatus.Text = $"{controlsToDelete.Count} control(es) eliminado(s).";
    }
    #endregion

    #region Manejo de Teclado (Flechas para Mover/Redimensionar, Supr, Ctrl+C/X/V/D, Ctrl+A, Esc)
    private void GridCanvas_KeyDown(object sender, KeyEventArgs e)
    {
        if (CurrentModel == null) return;

        double step = GridCanvas.SnapToGrid ? GridCanvas.GridSize : 1.0;
        bool isShift = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);
        bool isCtrl = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

        if (e.Key == Key.V && isCtrl)
        {
            PasteControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.A && isCtrl)
        {
            SelectAllControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Escape)
        {
            ClearSelection();
            e.Handled = true;
            return;
        }

        if (_selectedControls.Count == 0) return;

        if (e.Key == Key.C && isCtrl)
        {
            CopySelectedControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.X && isCtrl)
        {
            CutSelectedControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.D && isCtrl)
        {
            DuplicateSelectedControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Delete)
        {
            DeleteSelectedControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.A && isCtrl)
        {
            SelectAllControls();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Escape)
        {
            ClearSelection();
            e.Handled = true;
            return;
        }

        // Flechas del cursor
        double dx = 0, dy = 0, dw = 0, dh = 0;

        switch (e.Key)
        {
            case Key.Left:
                if (isShift) dw = -step; else dx = -step;
                break;
            case Key.Right:
                if (isShift) dw = step; else dx = step;
                break;
            case Key.Up:
                if (isShift) dh = -step; else dy = -step;
                break;
            case Key.Down:
                if (isShift) dh = step; else dy = step;
                break;
            default:
                return;
        }

        foreach (var ctrl in _selectedControls)
        {
            if (isShift)
            {
                ctrl.WidthPixels = Math.Max(15, ctrl.WidthPixels + dw);
                ctrl.HeightPixels = Math.Max(15, ctrl.HeightPixels + dh);

                if (_visualMap.TryGetValue(ctrl, out var vis))
                {
                    vis.Width = ctrl.WidthPixels;
                    vis.Height = ctrl.HeightPixels;
                }
            }
            else
            {
                ctrl.LeftPixels = Math.Max(0, ctrl.LeftPixels + dx);
                ctrl.TopPixels = Math.Max(0, ctrl.TopPixels + dy);

                if (_visualMap.TryGetValue(ctrl, out var vis))
                {
                    Canvas.SetLeft(vis, ctrl.LeftPixels);
                    Canvas.SetTop(vis, ctrl.TopPixels);
                }
            }
        }

        SelectionFrame.UpdateLayoutFromControl();
        NotifySelectionChanged();
        ModelModified?.Invoke();
        e.Handled = true;
    }
    #endregion

    private void UpdateDimensionText()
    {
        if (CurrentModel == null) return;
        TxtFormDimensions.Text = $"{CurrentModel.ClientWidth} × {CurrentModel.ClientHeight} twips ({CurrentModel.ClientWidthPixels:F0} × {CurrentModel.ClientHeightPixels:F0} px)";
    }

    private double _formDragStartWidth;
    private double _formDragStartHeight;
    private double _formAccumulatedDeltaX;
    private double _formAccumulatedDeltaY;

    private void UpdateResizeThumbs()
    {
        double titleH = (CurrentModel?.HasTitleBar == false) ? 0.0 : TitleBarHeight;
        double menuH = (BorderMenuBar.Visibility == Visibility.Visible) ? (BorderMenuBar.ActualHeight > 0 ? BorderMenuBar.ActualHeight : 24.0) : 0.0;
        double borderThick = (CurrentModel?.HasTitleBar == false) ? 1.0 : WindowBorderThickness;
        double formW = GridCanvas.Width > 0 ? (GridCanvas.Width + (borderThick * 2)) : (FormWindowBorder.ActualWidth > 0 ? FormWindowBorder.ActualWidth : 318.0);
        double formH = GridCanvas.Height > 0 ? (GridCanvas.Height + titleH + menuH + (borderThick * 2)) : (FormWindowBorder.ActualHeight > 0 ? FormWindowBorder.ActualHeight : 245.0);

        double thumbSize = 10.0;
        double thumbHalf = thumbSize / 2.0;

        ThumbResizeCorner.Width = thumbSize;
        ThumbResizeCorner.Height = thumbSize;
        ThumbResizeCorner.Margin = new Thickness(formW - thumbHalf, formH - thumbHalf, 0, 0);

        ThumbResizeRight.Width = thumbSize;
        ThumbResizeRight.Height = thumbSize;
        ThumbResizeRight.Margin = new Thickness(formW - thumbHalf, (formH / 2.0) - thumbHalf, 0, 0);

        ThumbResizeBottom.Width = thumbSize;
        ThumbResizeBottom.Height = thumbSize;
        ThumbResizeBottom.Margin = new Thickness((formW / 2.0) - thumbHalf, formH - thumbHalf, 0, 0);

        // Tiras de arrastre en los bordes derecho e inferior
        ThumbResizeBorderRight.Height = Math.Max(20, formH - titleH - 12);
        ThumbResizeBorderRight.Margin = new Thickness(formW - 4, titleH, 0, 0);

        ThumbResizeBorderBottom.Width = Math.Max(20, formW - 12);
        ThumbResizeBorderBottom.Margin = new Thickness(0, formH - 4, 0, 0);
    }

    #region Redimensión del Formulario con Tiradores
    private void ThumbResize_DragStarted(object sender, DragStartedEventArgs e)
    {
        _formDragStartWidth = GridCanvas.Width;
        _formDragStartHeight = GridCanvas.Height;
        _formAccumulatedDeltaX = 0;
        _formAccumulatedDeltaY = 0;
    }

    private void ThumbResizeCorner_DragDelta(object sender, DragDeltaEventArgs e)
    {
        if (CurrentModel == null) return;

        _formAccumulatedDeltaX += e.HorizontalChange;
        _formAccumulatedDeltaY += e.VerticalChange;

        double rawWidth = _formDragStartWidth + _formAccumulatedDeltaX;
        double rawHeight = _formDragStartHeight + _formAccumulatedDeltaY;

        double newWidth = Math.Max(120, GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawWidth) : rawWidth);
        double newHeight = Math.Max(60, GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawHeight) : rawHeight);

        GridCanvas.Width = newWidth;
        GridCanvas.Height = newHeight;

        CurrentModel.ClientWidthPixels = newWidth;
        CurrentModel.ClientHeightPixels = newHeight;

        UpdateDimensionText();
        UpdateResizeThumbs();
        ModelModified?.Invoke();
    }

    private void ThumbResizeRight_DragDelta(object sender, DragDeltaEventArgs e)
    {
        if (CurrentModel == null) return;

        _formAccumulatedDeltaX += e.HorizontalChange;
        double rawWidth = _formDragStartWidth + _formAccumulatedDeltaX;
        double newWidth = Math.Max(120, GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawWidth) : rawWidth);

        GridCanvas.Width = newWidth;
        CurrentModel.ClientWidthPixels = newWidth;

        UpdateDimensionText();
        UpdateResizeThumbs();
        ModelModified?.Invoke();
    }

    private void ThumbResizeBottom_DragDelta(object sender, DragDeltaEventArgs e)
    {
        if (CurrentModel == null) return;

        _formAccumulatedDeltaY += e.VerticalChange;
        double rawHeight = _formDragStartHeight + _formAccumulatedDeltaY;
        double newHeight = Math.Max(60, GridCanvas.SnapToGrid ? GridCanvas.SnapValue(rawHeight) : rawHeight);

        GridCanvas.Height = newHeight;
        CurrentModel.ClientHeightPixels = newHeight;

        UpdateDimensionText();
        UpdateResizeThumbs();
        ModelModified?.Invoke();
    }
    #endregion

    #region Barra de Herramientas del Diseñador
    private void BtnToggleGrid_Click(object sender, RoutedEventArgs e)
    {
        GridCanvas.ShowGrid = BtnToggleGrid.IsChecked == true;
    }

    private void BtnToggleSnap_Click(object sender, RoutedEventArgs e)
    {
        GridCanvas.SnapToGrid = BtnToggleSnap.IsChecked == true;
    }

    private void CmbGridSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (GridCanvas == null) return;

        GridCanvas.GridSize = CmbGridSize.SelectedIndex switch
        {
            0 => 4.0,
            1 => 8.0,
            2 => 12.0,
            3 => 16.0,
            _ => 8.0
        };
    }
    #endregion

    #region Métodos de Integración con Toolbox y Property Grid
    /// <summary>
    /// Crea e inserta un control centrado en el formulario.
    /// </summary>
    public FormDesignerControl? CreateControlInCenter(string typeName)
    {
        if (CurrentModel == null) return null;

        var info = ControlToolboxRegistry.Find(typeName);
        int defW = info?.DefaultWidthTwips ?? 1200;
        int defH = info?.DefaultHeightTwips ?? 400;

        FormDesignerControl? targetContainer = (_selectedControls.Count == 1 && _primarySelectedControl?.IsContainer == true)
            ? _primarySelectedControl
            : null;

        int containerW = targetContainer?.Width ?? CurrentModel.ClientWidth;
        int containerH = targetContainer?.Height ?? CurrentModel.ClientHeight;

        int leftTwips = Math.Max(0, (containerW - defW) / 2);
        int topTwips = Math.Max(0, (containerH - defH) / 2);

        if (GridCanvas.SnapToGrid)
        {
            double leftPx = GridCanvas.SnapValue(leftTwips / FormDesignerControl.TwipsPerPixel);
            double topPx = GridCanvas.SnapValue(topTwips / FormDesignerControl.TwipsPerPixel);
            leftTwips = (int)Math.Round(leftPx * FormDesignerControl.TwipsPerPixel);
            topTwips = (int)Math.Round(topPx * FormDesignerControl.TwipsPerPixel);
        }

        var newCtrl = ControlToolboxRegistry.CreateControl(typeName, CurrentModel, leftTwips, topTwips, defW, defH);
        if (newCtrl.IsNonResizable)
        {
            newCtrl.Width = 480;
            newCtrl.Height = 480;
        }

        if (targetContainer != null)
        {
            newCtrl.ParentControl = targetContainer;
            targetContainer.Children.Add(newCtrl);
        }
        else
        {
            newCtrl.ParentControl = null;
            CurrentModel.Children.Add(newCtrl);
        }

        RefreshCanvas();
        SelectSingleControl(newCtrl);
        ControlAdded?.Invoke(newCtrl);
        ModelModified?.Invoke();
        return newCtrl;
    }

    /// <summary>
    /// Selecciona programáticamente un control o deselecciona si es nulo.
    /// </summary>
    public void SelectControl(FormDesignerControl? control)
    {
        if (control == null)
        {
            ClearSelection();
        }
        else
        {
            SelectSingleControl(control);
        }
    }

    /// <summary>
    /// Actualiza la representación visual de los controles y del formulario tras cambios en la Ventana de Propiedades.
    /// Si hay un único control seleccionado, actualiza solo ese elemento visual in-situ para evitar parpadeos (flicker).
    /// </summary>
    public void RefreshVisuals()
    {
        if (CurrentModel == null) return;

        TxtFormCaption.Text = string.IsNullOrEmpty(CurrentModel.Caption) ? CurrentModel.Name : CurrentModel.Caption;
        TxtFormTitle.Text = CurrentModel.IsUserControl ? $"Diseñador UserControl: {CurrentModel.Name}" : $"Diseñador: {CurrentModel.Name}";

        GridCanvas.Width = CurrentModel.ClientWidthPixels;
        GridCanvas.Height = CurrentModel.ClientHeightPixels;

        FormClientArea.Background = VbColorHelper.ToBrush(
            CurrentModel.BackColorHex,
            new SolidColorBrush(Color.FromRgb(240, 240, 240))
        );

        UpdateDimensionText();
        RebuildMenuBar();
        UpdateResizeThumbs();

        if (_selectedControls.Count == 1 && _primarySelectedControl != null && _visualMap.ContainsKey(_primarySelectedControl))
        {
            RefreshSingleControlVisual(_primarySelectedControl);
        }
        else
        {
            RefreshCanvas();
        }
    }
    #endregion

    #region Barra de Menús del Formulario (VB.Menu)
    /// <summary>
    /// Reconstruye la barra de menús visual del formulario según los menús definidos en el modelo.
    /// </summary>
    public void RebuildMenuBar()
    {
        FormDesignMenuBar.Items.Clear();

        if (CurrentModel == null || CurrentModel.TopLevelMenus.Count == 0)
        {
            BorderMenuBar.Visibility = Visibility.Collapsed;
            MenuBarRow.Height = new GridLength(0);
            return;
        }

        BorderMenuBar.Visibility = Visibility.Visible;
        MenuBarRow.Height = GridLength.Auto;

        foreach (var menuCtrl in CurrentModel.TopLevelMenus)
        {
            var wpfItem = new MenuItem
            {
                Header = VbCaptionToWpfHeader(menuCtrl.Caption),
                InputGestureText = VbMenuShortcutHelper.ToDisplayString(menuCtrl.Shortcut),
                IsChecked = menuCtrl.Checked,
                IsEnabled = menuCtrl.Enabled,
                Visibility = menuCtrl.Visible ? Visibility.Visible : Visibility.Collapsed,
                Tag = menuCtrl
            };

            SetupMenuItemContextMenu(wpfItem, menuCtrl);

            if (menuCtrl.Children.Any(c => c.IsMenu))
            {
                PopulateMenuChildren(wpfItem, menuCtrl);
            }
            else
            {
                wpfItem.Click += (s, e) =>
                {
                    e.Handled = true;
                    OnMenuItemClicked(menuCtrl);
                };
            }

            FormDesignMenuBar.Items.Add(wpfItem);
        }

        SetupMenuBarContextMenu();
    }

    private void PopulateMenuChildren(MenuItem parentWpfItem, FormDesignerControl parentCtrl)
    {
        foreach (var childCtrl in parentCtrl.Children.Where(c => c.IsMenu))
        {
            if (childCtrl.Caption == "-")
            {
                var sep = new Separator { Tag = childCtrl };
                parentWpfItem.Items.Add(sep);
            }
            else
            {
                var childWpfItem = new MenuItem
                {
                    Header = VbCaptionToWpfHeader(childCtrl.Caption),
                    InputGestureText = VbMenuShortcutHelper.ToDisplayString(childCtrl.Shortcut),
                    IsChecked = childCtrl.Checked,
                    IsEnabled = childCtrl.Enabled,
                    Visibility = childCtrl.Visible ? Visibility.Visible : Visibility.Collapsed,
                    Tag = childCtrl
                };

                SetupMenuItemContextMenu(childWpfItem, childCtrl);

                if (childCtrl.Children.Any(c => c.IsMenu))
                {
                    PopulateMenuChildren(childWpfItem, childCtrl);
                }
                else
                {
                    childWpfItem.Click += (s, e) =>
                    {
                        e.Handled = true;
                        OnMenuItemClicked(childCtrl);
                    };
                }

                parentWpfItem.Items.Add(childWpfItem);
            }
        }
    }

    private void OnMenuItemClicked(FormDesignerControl menuCtrl)
    {
        SelectSingleControl(menuCtrl);
        ControlDoubleClicked?.Invoke(menuCtrl);
    }

    private static string VbCaptionToWpfHeader(string caption)
    {
        if (string.IsNullOrEmpty(caption)) return string.Empty;
        // VB6 '&' para acelerador, '&&' para ampersand literal
        return caption.Replace("_", "__").Replace("&&", "\0").Replace('&', '_').Replace("\0", "&");
    }

    private void SetupMenuBarContextMenu()
    {
        var cm = new ContextMenu();
        var miEditor = new MenuItem { Header = "Editor de Menús (Ctrl+E)..." };
        miEditor.Click += (_, _) => RequestMenuEditor?.Invoke();
        cm.Items.Add(miEditor);

        BorderMenuBar.ContextMenu = cm;
    }

    private void SetupMenuItemContextMenu(MenuItem wpfItem, FormDesignerControl menuCtrl)
    {
        var cm = new ContextMenu();
        var miCode = new MenuItem { Header = "Ver Código de Evento (Click)" };
        miCode.Click += (_, _) => OnMenuItemClicked(menuCtrl);
        cm.Items.Add(miCode);

        var miEditor = new MenuItem { Header = "Editor de Menús (Ctrl+E)..." };
        miEditor.Click += (_, _) => RequestMenuEditor?.Invoke();
        cm.Items.Add(miEditor);

        wpfItem.ContextMenu = cm;
    }

    private ActiveXInPlaceHost? GetActiveXHostForControl(FormDesignerControl? control)
    {
        if (control == null || !_visualMap.TryGetValue(control, out var vis))
            return null;

        if (vis is ActiveXInPlaceHost axHost)
            return axHost;

        if (vis is Grid g && g.Children.Count > 0 && g.Children[0] is ActiveXInPlaceHost gridHost)
            return gridHost;

        return null;
    }

    private void SetupCanvasContextMenu()
    {
        var cm = new ContextMenu();

        var miCode = new MenuItem { Header = "Ver Código", InputGestureText = "F7" };
        miCode.Click += (_, _) => RequestCodeView?.Invoke();
        cm.Items.Add(miCode);

        var miMenuEditor = new MenuItem { Header = "Editor de Menús...", InputGestureText = "Ctrl+E" };
        miMenuEditor.Click += (_, _) => RequestMenuEditor?.Invoke();
        cm.Items.Add(miMenuEditor);

        cm.Items.Add(new Separator());

        var miCut = new MenuItem { Header = "Cortar", InputGestureText = "Ctrl+X" };
        miCut.Click += (_, _) => CutSelectedControls();
        cm.Items.Add(miCut);

        var miCopy = new MenuItem { Header = "Copiar", InputGestureText = "Ctrl+C" };
        miCopy.Click += (_, _) => CopySelectedControls();
        cm.Items.Add(miCopy);

        var miPaste = new MenuItem { Header = "Pegar", InputGestureText = "Ctrl+V" };
        miPaste.Click += (_, _) => PasteControls();
        cm.Items.Add(miPaste);

        var miDelete = new MenuItem { Header = "Eliminar", InputGestureText = "Supr" };
        miDelete.Click += (_, _) => DeleteSelectedControls();
        cm.Items.Add(miDelete);

        cm.Items.Add(new Separator());

        var miEditActiveX = new MenuItem { Header = "Editar control ActiveX (UI-Active)", Visibility = Visibility.Collapsed };
        miEditActiveX.Click += (_, _) =>
        {
            var sel = _primarySelectedControl ?? _selectedControls.FirstOrDefault();
            var axHost = GetActiveXHostForControl(sel);
            if (axHost != null)
            {
                if (axHost.IsInteractiveEditMode)
                {
                    axHost.ExitInteractiveEditMode();
                    ModelModified?.Invoke();
                }
                else
                {
                    axHost.EnterInteractiveEditMode();
                }
            }
        };
        cm.Items.Add(miEditActiveX);

        var miOleVerbs = new MenuItem { Header = "Verbos ActiveX", Visibility = Visibility.Collapsed };
        cm.Items.Add(miOleVerbs);

        var miProperties = new MenuItem { Header = "Propiedades...", FontWeight = FontWeights.Bold };
        miProperties.Click += (_, _) => OpenSelectedControlProperties();
        cm.Items.Add(miProperties);

        cm.Opened += (s, e) =>
        {
            var sel = _primarySelectedControl ?? _selectedControls.FirstOrDefault();
            miProperties.IsEnabled = sel != null;
            if (sel != null)
            {
                miProperties.Header = $"Propiedades de '{sel.Name}'...";
            }
            else
            {
                miProperties.Header = "Propiedades...";
            }

            var axHost = GetActiveXHostForControl(sel);
            if (axHost != null)
            {
                miEditActiveX.Visibility = Visibility.Visible;
                miEditActiveX.Header = axHost.IsInteractiveEditMode
                    ? $"Finalizar edición de '{sel!.Name}'"
                    : $"Editar '{sel!.Name}' (UI-Active)";

                miOleVerbs.Items.Clear();
                var verbs = axHost.GetOleVerbs();
                foreach (var verb in verbs)
                {
                    int verbId = verb.VerbId;
                    var verbItem = new MenuItem { Header = verb.Name };
                    verbItem.Click += (_, _) =>
                    {
                        axHost.ExecuteOleVerb(verbId, Window.GetWindow(this));
                        ModelModified?.Invoke();
                    };
                    miOleVerbs.Items.Add(verbItem);
                }
                miOleVerbs.Visibility = miOleVerbs.Items.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            else
            {
                miEditActiveX.Visibility = Visibility.Collapsed;
                miOleVerbs.Visibility = Visibility.Collapsed;
            }
        };

        GridCanvas.ContextMenu = cm;
    }

    public void OpenSelectedControlProperties()
    {
        var sel = _primarySelectedControl ?? _selectedControls.FirstOrDefault();
        if (sel != null)
        {
            var dlg = new Dialogs.OcxPropertyPagesWindow(sel, CurrentModel, () =>
            {
                RefreshVisuals();
                ModelModified?.Invoke();
            });
            dlg.Owner = Window.GetWindow(this);
            dlg.ShowDialog();
        }
    }
    #endregion
}
