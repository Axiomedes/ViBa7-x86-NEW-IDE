using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.ActiveXHost;

/// <summary>
/// HwndHost nativo de Win32 que aloja la ventana real (HWND) de un control ActiveX (OCX)
/// e implementa el enrutamiento de teclado OLE (IKeyboardInputSink: TranslateAcceleratorCore, TabIntoCore, OnMnemonicCore)
/// y el redimensionamiento sincronizado (OnWindowPositionChanged -> IOleObject::SetExtent + IOleInPlaceObject::SetObjectRects).
/// </summary>
public sealed class ActiveXWin32HwndHost : HwndHost
{
    private readonly DynamicActiveXControl _axControl;
    private readonly bool _attachExistingControl;
    private IntPtr _hwndContainer;

    public DynamicActiveXControl AxControl => _axControl;
    public IntPtr ContainerHwnd => _hwndContainer;
    public bool IsControlAttached { get; private set; }

    public ActiveXWin32HwndHost(DynamicActiveXControl axControl, bool attachExistingControl = false)
    {
        _axControl = axControl ?? throw new ArgumentNullException(nameof(axControl));
        _attachExistingControl = attachExistingControl;
        Focusable = true;
    }

    protected override HandleRef BuildWindowCore(HandleRef hwndParent)
    {
        int w = Math.Max(16, (int)Math.Round(Width > 0 ? Width : _axControl.Width));
        int h = Math.Max(16, (int)Math.Round(Height > 0 ? Height : _axControl.Height));

        _hwndContainer = Ole32Interop.CreateWindowExW(
            0,
            "STATIC",
            string.Empty,
            Ole32Interop.WS_CHILD | Ole32Interop.WS_VISIBLE | Ole32Interop.WS_CLIPCHILDREN | Ole32Interop.WS_CLIPSIBLINGS | Ole32Interop.WS_TABSTOP,
            0,
            0,
            w,
            h,
            hwndParent.Handle,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero);

        if (_hwndContainer != IntPtr.Zero)
        {
            if (_attachExistingControl || _axControl.GetActiveXObject() != null)
            {
                IsControlAttached = _axControl.AttachToInteractiveParent(_hwndContainer, w, h);
            }
            else
            {
                IsControlAttached = _axControl.CreateRuntimeControl(_hwndContainer, w, h) && !_axControl.HasInitializationError;
            }
        }

        return new HandleRef(this, _hwndContainer);
    }

    protected override void OnWindowPositionChanged(Rect rcBoundingBox)
    {
        base.OnWindowPositionChanged(rcBoundingBox);

        if (_hwndContainer == IntPtr.Zero) return;

        int w = Math.Max(16, (int)Math.Round(rcBoundingBox.Width));
        int h = Math.Max(16, (int)Math.Round(rcBoundingBox.Height));

        Ole32Interop.MoveWindow(_hwndContainer, (int)Math.Round(rcBoundingBox.X), (int)Math.Round(rcBoundingBox.Y), w, h, true);
        _axControl.Size = new System.Drawing.Size(w, h);
    }

    protected override bool TranslateAcceleratorCore(ref System.Windows.Interop.MSG msg, ModifierKeys modifiers)
    {
        if (msg.message is >= (int)Ole32Interop.WM_KEYDOWN and <= (int)Ole32Interop.WM_SYSCHAR)
        {
            var nativeMsg = ConvertToNativeMsg(ref msg);
            if (_axControl.TranslateAccelerator(ref nativeMsg))
            {
                return true;
            }
        }

        return base.TranslateAcceleratorCore(ref msg, modifiers);
    }

    protected override bool TabIntoCore(TraversalRequest request)
    {
        return _axControl.FocusControl();
    }

    protected override bool OnMnemonicCore(ref System.Windows.Interop.MSG msg, ModifierKeys modifiers)
    {
        var nativeMsg = ConvertToNativeMsg(ref msg);
        if (_axControl.OnMnemonic(ref nativeMsg))
        {
            return true;
        }

        return base.OnMnemonicCore(ref msg, modifiers);
    }

    protected override void DestroyWindowCore(HandleRef hwnd)
    {
        if (_attachExistingControl)
        {
            _axControl.DetachToSharedParking(_axControl.Width, _axControl.Height);
        }

        if (hwnd.Handle != IntPtr.Zero)
        {
            Ole32Interop.DestroyWindow(hwnd.Handle);
        }
        _hwndContainer = IntPtr.Zero;
        IsControlAttached = false;
    }

    private static Native.MSG ConvertToNativeMsg(ref System.Windows.Interop.MSG msg)
    {
        return new Native.MSG
        {
            hwnd = msg.hwnd,
            message = unchecked((uint)msg.message),
            wParam = msg.wParam,
            lParam = msg.lParam,
            time = unchecked((uint)msg.time),
            pt = new Native.POINT(msg.pt_x, msg.pt_y)
        };
    }
}

/// <summary>
/// Contenedor interactivo de tiempo de ejecución (F5 / LiveFormWindow) para controles ActiveX (OCX).
/// Inicializa el control con DISPID_AMBIENT_USERMODE = True (IsDesignMode = false), restaura su estado
/// binario (.frx / .ctx) y propiedades de diseño, y aloja su HWND real mediante ActiveXWin32HwndHost
/// para permitir interacción directa de ratón, foco y teclado.
/// </summary>
public sealed class ActiveXRuntimeHost : Border, IDisposable
{
    private readonly DynamicActiveXControl _axControl;
    private readonly Image _windowlessImage;
    private ActiveXWin32HwndHost? _hwndHost;
    private FrameworkElement? _fallbackVisual;
    private bool _initialized;
    private bool _disposed;

    public string Clsid => _axControl.Clsid;
    public string ProgId { get; }
    public DynamicActiveXControl AxControl => _axControl;
    public FormDesignerControl? BoundControl { get; set; }
    public ActiveXWin32HwndHost? HwndHostControl => _hwndHost;

    /// <summary>
    /// Indica si el control OCX se activó exitosamente en modo ejecución.
    /// </summary>
    public bool IsRuntimeActive { get; private set; }

    /// <summary>
    /// Elemento visual de respaldo (WPF) en caso de que el OCX no esté registrado en el equipo.
    /// </summary>
    public FrameworkElement? FallbackVisual
    {
        get => _fallbackVisual;
        set
        {
            _fallbackVisual = value;
            if (_initialized && !IsRuntimeActive && _fallbackVisual != null)
            {
                Child = _fallbackVisual;
            }
        }
    }

    public event Action<ActiveXRuntimeHost, int>? RuntimePropertyChanged;
    public event Action<ActiveXRuntimeHost>? RuntimeExtenderBoundsChanged;
    public event Action<ActiveXEventInvocation>? ComEventFired;

    public ActiveXRuntimeHost(string clsid, string progId = "", int initialWidthPx = 120, int initialHeightPx = 36)
    {
        ProgId = progId;
        int safeW = Math.Max(24, initialWidthPx);
        int safeH = Math.Max(20, initialHeightPx);

        _axControl = new DynamicActiveXControl(clsid, safeW, safeH)
        {
            IsDesignMode = false,
            Name = !string.IsNullOrWhiteSpace(progId) ? progId : "ActiveXControl1",
            ProgId = progId
        };

        _windowlessImage = new Image
        {
            Stretch = Stretch.Fill,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };
        RenderOptions.SetBitmapScalingMode(_windowlessImage, BitmapScalingMode.NearestNeighbor);

        Width = safeW;
        Height = safeH;
        Focusable = true;
        SnapsToDevicePixels = true;
        UseLayoutRounding = true;
        ClipToBounds = true;

        _axControl.PropertyChangedByOcx += dispId =>
        {
            if (_disposed) return;
            if (_axControl.IsWindowless)
            {
                RefreshWindowlessSnapshot();
            }
            RuntimePropertyChanged?.Invoke(this, dispId);
        };

        _axControl.ViewInvalidatedByOcx += () =>
        {
            if (_disposed) return;
            if (_axControl.IsWindowless)
            {
                RefreshWindowlessSnapshot();
            }
        };

        _axControl.ExtenderBoundsChanged += () =>
        {
            if (_disposed) return;
            RuntimeExtenderBoundsChanged?.Invoke(this);
        };

        _axControl.ComEventFired += invocation =>
        {
            if (_disposed) return;
            ComEventFired?.Invoke(invocation);
        };

        SizeChanged += OnHostSizeChanged;
    }

    /// <summary>
    /// Configura las propiedades del modelo de diseño y activa el control ActiveX en modo ejecución (UserMode = True).
    /// </summary>
    public bool TryInitializeRuntime()
    {
        if (_disposed) return false;
        _initialized = true;

        int w = BoundControl != null ? Math.Max(24, (int)Math.Round(BoundControl.WidthPixels)) : Math.Max(24, (int)Width);
        int h = BoundControl != null ? Math.Max(20, (int)Math.Round(BoundControl.HeightPixels)) : Math.Max(20, (int)Height);

        Width = w;
        Height = h;

        if (BoundControl != null)
        {
            ApplyModelStateToAxControl(BoundControl, _axControl);
        }

        if (_axControl.IsWindowless)
        {
            IsRuntimeActive = true;
            RefreshWindowlessSnapshot(w, h);
            return true;
        }

        _hwndHost = new ActiveXWin32HwndHost(_axControl, attachExistingControl: false)
        {
            Width = w,
            Height = h
        };

        // Montar el HwndHost como hijo; si la ventana padre ya está cargada, BuildWindowCore se ejecuta al cargarse en el árbol visual.
        // Si queremos verificar inmediatamente si el CLSID es instanciable, probamos o dejamos que HwndHost lo inicialice.
        Child = _hwndHost;

        _hwndHost.Loaded += (_, _) =>
        {
            if (_disposed) return;

            if (_axControl.HasInitializationError || (!_hwndHost.IsControlAttached && !_axControl.IsWindowless))
            {
                IsRuntimeActive = false;
                if (_fallbackVisual != null)
                {
                    _fallbackVisual.Width = Width;
                    _fallbackVisual.Height = Height;
                    Child = _fallbackVisual;
                }
            }
            else if (_axControl.IsWindowless)
            {
                IsRuntimeActive = true;
                RefreshWindowlessSnapshot();
            }
            else
            {
                IsRuntimeActive = true;
            }
        };

        return true;
    }

    /// <summary>
    /// Inicializa la instancia COM en modo ejecución de forma inmediata (útil para pruebas o acceso a propiedades en Form_Load
    /// antes de que el evento Loaded de WPF haya disparado BuildWindowCore).
    /// </summary>
    public bool EnsureControlCreatedEarly()
    {
        if (_disposed) return false;
        if (_axControl.GetActiveXObject() != null) return !_axControl.HasInitializationError;

        if (BoundControl != null)
        {
            ApplyModelStateToAxControl(BoundControl, _axControl);
        }

        _axControl.IsDesignMode = false;
        bool ok = _axControl.CreateControl() && !_axControl.HasInitializationError;
        if (ok)
        {
            IsRuntimeActive = true;
            if (_hwndHost != null && _hwndHost.ContainerHwnd != IntPtr.Zero && !_axControl.IsWindowless)
            {
                _axControl.AttachToInteractiveParent(_hwndHost.ContainerHwnd, (int)Math.Max(16, Width), (int)Math.Max(16, Height));
            }
        }
        return ok;
    }

    private void OnHostSizeChanged(object sender, SizeChangedEventArgs e)
    {
        if (_disposed) return;
        int w = Math.Max(16, (int)Math.Round(e.NewSize.Width));
        int h = Math.Max(16, (int)Math.Round(e.NewSize.Height));

        if (_hwndHost != null)
        {
            _hwndHost.Width = w;
            _hwndHost.Height = h;
        }

        _axControl.Size = new System.Drawing.Size(w, h);
        if (_axControl.IsWindowless)
        {
            RefreshWindowlessSnapshot(w, h);
        }
    }

    public void RefreshWindowlessSnapshot(int widthPx = 0, int heightPx = 0)
    {
        if (_disposed || _axControl.HasInitializationError) return;
        int w = widthPx > 0 ? widthPx : Math.Max(24, (int)Math.Round(ActualWidth > 0 ? ActualWidth : Width));
        int h = heightPx > 0 ? heightPx : Math.Max(20, (int)Math.Round(ActualHeight > 0 ? ActualHeight : Height));

        var bmp = _axControl.TryCaptureVisualSnapshot(w, h);
        if (bmp != null)
        {
            _windowlessImage.Source = bmp;
            if (Child != _windowlessImage)
            {
                Child = _windowlessImage;
            }
        }
        else if (_fallbackVisual != null && Child != _fallbackVisual)
        {
            _fallbackVisual.Width = Width;
            _fallbackVisual.Height = Height;
            Child = _fallbackVisual;
        }
    }

    internal static void ApplyModelStateToAxControl(FormDesignerControl model, DynamicActiveXControl axControl)
    {
        if (!string.IsNullOrWhiteSpace(model.Name))
        {
            axControl.Name = model.Name;
        }

        axControl.Extender.Name = model.Name;
        axControl.Extender.Index = (short)(model.Index ?? -1);
        axControl.Extender.SetBoundsDirect(model.Left, model.Top, model.Width, model.Height);
        axControl.Extender.Visible = model.Visible;
        axControl.Extender.Enabled = model.Enabled;
        axControl.Extender.TabIndex = (short)(model.TabIndex ?? 0);
        axControl.Extender.TabStop = model.TabStop ?? true;
        axControl.Extender.HelpContextID = model.HelpContextID;
        axControl.Extender.Tag = model.GetCustomPropertyString("Tag") ?? string.Empty;
        axControl.Extender.ToolTipText = model.GetCustomPropertyString("ToolTipText") ?? string.Empty;

        if (!string.IsNullOrEmpty(model.LicenseKey))
        {
            axControl.LicenseKey = model.LicenseKey;
        }

        if (model.ActiveXStateBlob is { Length: > 0 })
        {
            axControl.InitialStreamBlob = model.ActiveXStateBlob;
        }

        if (model.BinaryProperties.Count > 0)
        {
            axControl.SetBinaryProperties(model.BinaryProperties);
            foreach (var binKvp in model.BinaryProperties)
            {
                if (ActiveXInPlaceHost.IsPicturePropertyName(binKvp.Key) && binKvp.Value is { Length: > 0 })
                {
                    axControl.SetPictureProperty(binKvp.Key, binKvp.Value);
                }
            }
        }

        if (model.Font != null)
        {
            axControl.ConfigureAmbientFont(
                model.Font.Name,
                (float)model.Font.Size,
                model.Font.IsBold,
                model.Font.Italic,
                model.Font.Underline,
                model.Font.Strikethrough,
                (short)model.Font.Charset);

            axControl.SetFontProperty(
                "Font",
                model.Font.Name,
                (float)model.Font.Size,
                model.Font.IsBold,
                model.Font.Italic,
                model.Font.Underline,
                model.Font.Strikethrough,
                (short)model.Font.Charset);
        }

        if (model.PictureData is { Length: > 0 })
        {
            axControl.SetPictureProperty("Picture", model.PictureData);
        }
        if (model.DownPictureData is { Length: > 0 })
        {
            axControl.SetPictureProperty("DownPicture", model.DownPictureData);
        }
        if (model.DisabledPictureData is { Length: > 0 })
        {
            axControl.SetPictureProperty("DisabledPicture", model.DisabledPictureData);
        }

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in model.CustomProperties)
        {
            if (p.IsBeginProperty)
            {
                if (ActiveXInPlaceHost.TryParseStdFontBeginProperty(p, out var subFont))
                {
                    axControl.SetFontProperty(
                        p.Name,
                        subFont.Name,
                        (float)subFont.Size,
                        subFont.IsBold,
                        subFont.Italic,
                        subFont.Underline,
                        subFont.Strikethrough,
                        (short)subFont.Charset);
                }
                continue;
            }

            if (!string.IsNullOrWhiteSpace(p.Name))
            {
                props[p.Name] = p.Value;
            }
        }

        if (!string.IsNullOrEmpty(model.Caption)) props["Caption"] = model.Caption;
        if (!string.IsNullOrEmpty(model.Text)) props["Text"] = model.Text;
        if (model.BackColorHex != null) props["BackColor"] = model.BackColorHex;
        if (model.ForeColorHex != null) props["ForeColor"] = model.ForeColorHex;
        if (model.BorderStyle.HasValue) props["BorderStyle"] = model.BorderStyle.Value.ToString();

        axControl.ApplyProperties(props);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_hwndHost != null)
            {
                Child = null;
                _hwndHost.Dispose();
                _hwndHost = null;
            }
        }
        catch { }

        try
        {
            _axControl.Dispose();
        }
        catch { }
    }
}
