using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.ActiveXHost;

/// <summary>
/// Contenedor visual de WPF para controles ActiveX (OCX) en la superficie de diseño de ViBa7.
/// Mantiene la instancia COM real en proceso (para reflexión IDispatch, IPersistPropertyBag y hojas
/// de propiedades OLE nativas) y renderiza su apariencia verificada libre de parpadeos (flicker)
/// y sin agujeros transparentes de Airspace, con respaldo visual automático de alta fidelidad.
/// </summary>
public class ActiveXInPlaceHost : Border, IDisposable
{
    private readonly DynamicActiveXControl _axControl;
    private readonly Image _snapshotImage;
    private FrameworkElement? _fallbackVisual;
    private bool _nativeActivated;
    private bool _disposed;
    private bool _isSyncingProperties;

    public string Clsid => _axControl.Clsid;
    public string ProgId { get; }
    public DynamicActiveXControl AxControl => _axControl;
    public FormDesignerControl? BoundControl { get; set; }

    /// <summary>
    /// Indica si el control OCX nativo produjo una captura gráfica válida en modo diseño.
    /// Si es false, se muestra el FallbackVisual de alta fidelidad de WPF.
    /// </summary>
    public bool HasValidNativeVisual { get; private set; }

    /// <summary>
    /// Fábrica global opcional para generar respaldos visuales de WPF cuando el control OCX no se puede capturar nativamente.
    /// </summary>
    public static Func<FormDesignerControl, FrameworkElement?>? FallbackVisualFactory { get; set; }

    /// <summary>
    /// Elemento visual de respaldo de alta fidelidad (WPF) utilizado cuando el OCX de terceros
    /// no puede dibujarse nativamente en modo diseño (por falta de licencia de diseño, estado .frx propietario, etc.).
    /// </summary>
    public FrameworkElement? FallbackVisual
    {
        get => _fallbackVisual;
        set
        {
            _fallbackVisual = value;
            if (!HasValidNativeVisual && _fallbackVisual != null)
            {
                Child = _fallbackVisual;
            }
        }
    }

    public bool IsDesignMode
    {
        get => _axControl.IsDesignMode;
        set => _axControl.IsDesignMode = value;
    }

    public event Action<ActiveXInPlaceHost, Point>? HostSelected;
    public event Action<ActiveXInPlaceHost, Point>? HostDoubleClicked;
    public event Action<ActiveXInPlaceHost, Point>? HostContextMenuRequested;
    public event Action<ActiveXInPlaceHost, Point>? HostDragStarted;
    public event Action<ActiveXInPlaceHost, Point>? HostDragging;
    public event Action<ActiveXInPlaceHost>? HostDragEnded;
    public event Action<ActiveXInPlaceHost, int, string>? HostPropertyChangedByOcx;
    public event Action<ActiveXInPlaceHost>? HostExtenderBoundsChanged;

    public ActiveXDesignSite DesignSite { get; }

    public ActiveXInPlaceHost(string clsid, string progId = "", int initialWidthPx = 120, int initialHeightPx = 36)
    {
        ProgId = progId;
        int safeW = Math.Max(24, initialWidthPx);
        int safeH = Math.Max(20, initialHeightPx);

        _axControl = new DynamicActiveXControl(clsid, safeW, safeH)
        {
            IsDesignMode = true,
            Name = !string.IsNullOrWhiteSpace(progId) ? progId : "ActiveXControl1"
        };
        DesignSite = new ActiveXDesignSite(_axControl, _axControl.Name);
        _axControl.Site = DesignSite;

        _snapshotImage = new Image
        {
            Stretch = Stretch.Fill,
            SnapsToDevicePixels = true,
            UseLayoutRounding = true
        };
        RenderOptions.SetBitmapScalingMode(_snapshotImage, BitmapScalingMode.NearestNeighbor);

        Background = Brushes.Transparent;
        SnapsToDevicePixels = true;
        UseLayoutRounding = true;
        ClipToBounds = true;

        HookAxEvents();
    }

    private void HookAxEvents()
    {
        _axControl.DesignMouseDown += (s, e) =>
        {
            var pt = new Point(e.X, e.Y);
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                HostSelected?.Invoke(this, pt);
                HostDragStarted?.Invoke(this, pt);
            }
        };

        _axControl.DesignMouseMove += (s, e) =>
        {
            var pt = new Point(e.X, e.Y);
            HostDragging?.Invoke(this, pt);
        };

        _axControl.DesignMouseUp += (s, e) =>
        {
            HostDragEnded?.Invoke(this);
        };

        _axControl.DesignMouseDoubleClick += (s, e) =>
        {
            var pt = new Point(e.X, e.Y);
            HostDoubleClicked?.Invoke(this, pt);
        };

        _axControl.DesignContextMenu += (s, e) =>
        {
            var pt = new Point(e.X, e.Y);
            HostContextMenuRequested?.Invoke(this, pt);
        };

        _axControl.PropertyChangedByOcx += dispId =>
        {
            if (_disposed || !_nativeActivated || _axControl.HasInitializationError || _isSyncingProperties) return;

            string propName = ResolvePropertyNameFromDispId(dispId);
            if (BoundControl != null)
            {
                if (!string.IsNullOrEmpty(propName))
                {
                    var liveVal = _axControl.GetPropertyValue(propName);
                    if (liveVal != null && !System.Runtime.InteropServices.Marshal.IsComObject(liveVal))
                    {
                        string valStr = liveVal.ToString() ?? string.Empty;
                        BoundControl.SetCustomProperty(propName, valStr);
                        DesignerPropertyFactory.SynchronizeStandardProperties(BoundControl, propName, valStr);
                    }
                }
                else
                {
                    SyncToDesignerModel();
                }
            }

            RefreshVisualSnapshot();
            HostPropertyChangedByOcx?.Invoke(this, dispId, propName);
        };

        _axControl.ViewInvalidatedByOcx += () =>
        {
            if (_disposed || !_nativeActivated || _axControl.HasInitializationError) return;
            RefreshVisualSnapshot();
        };

        _axControl.ExtenderBoundsChanged += () =>
        {
            if (_disposed || BoundControl == null || _isSyncingProperties) return;
            BoundControl.Left = (int)Math.Round(_axControl.Extender.Left);
            BoundControl.Top = (int)Math.Round(_axControl.Extender.Top);
            BoundControl.Width = Math.Max(15, (int)Math.Round(_axControl.Extender.Width));
            BoundControl.Height = Math.Max(15, (int)Math.Round(_axControl.Extender.Height));
            BoundControl.UpdateExtentsFromDimensions();
            HostExtenderBoundsChanged?.Invoke(this);
        };
    }

    private string ResolvePropertyNameFromDispId(int dispId)
    {
        // Stock OLE / ActiveX DispIds estándar (Olectl.h)
        switch (dispId)
        {
            case -501: return "BackColor";
            case -502: return "BackStyle";
            case -503: return "BorderColor";
            case -504: return "BorderStyle";
            case -505: return "BorderWidth";
            case -512: return "Font";
            case -513: return "ForeColor";
            case -514: return "Enabled";
            case -515: return "hWnd";
            case -516: return "TabStop";
            case -517: return "Text";
            case -518: return "Caption";
            case -520: return "Appearance";
            case -521: return "MousePointer";
            case -522: return "Picture";
            case -523: return "MouseIcon";
            case -525: return "ReadyState";
        }

        string ctrlType = BoundControl?.TypeName ?? ProgId;
        if (!string.IsNullOrWhiteSpace(ctrlType) &&
            ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.TryResolvePropertyNameByDispId(ctrlType, dispId, out string resolved))
        {
            return resolved;
        }

        return string.Empty;
    }

    /// <summary>
    /// Inicializa de forma segura la instancia COM OLE (con o sin ventana)
    /// y captura su representación visual, conmutando automáticamente al FallbackVisual
    /// si el control produce un cuadro vacío o lanza un error interno (ej. División por cero).
    /// </summary>
    public bool TryInitializeAndRender()
    {
        if (_disposed) return false;

        int w = BoundControl != null ? Math.Max(24, (int)Math.Round(BoundControl.WidthPixels)) : Math.Max(24, (int)Width);
        int h = BoundControl != null ? Math.Max(20, (int)Math.Round(BoundControl.HeightPixels)) : Math.Max(20, (int)Height);

        if (BoundControl != null && !string.IsNullOrWhiteSpace(BoundControl.Name))
        {
            _axControl.Name = BoundControl.Name;
            DesignSite.Name = BoundControl.Name;
        }

        SyncPropertiesDictionaryToAx();

        if (!_nativeActivated && !_axControl.HasInitializationError)
        {
            _axControl.Size = new System.Drawing.Size(w, h);
            _nativeActivated = _axControl.CreateControl() && !_axControl.HasInitializationError;

            if (_nativeActivated && BoundControl != null)
            {
                if (!string.IsNullOrEmpty(_axControl.LicenseKey))
                {
                    BoundControl.LicenseKey = _axControl.LicenseKey;
                }

                // Sincronizar propiedades descubiertas en vivo desde el OCX hacia el modelo de diseño para el PropertyGrid
                if (_axControl.DiscoveredProperties.Count > 0)
                {
                    foreach (var kvp in _axControl.DiscoveredProperties)
                    {
                        if (!BoundControl.CustomProperties.Exists(p => p.Name.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase)))
                        {
                            BoundControl.CustomProperties.Add(new FormPropertyItem
                            {
                                Name = kvp.Key,
                                Value = kvp.Value
                            });
                        }
                    }
                }
            }
        }

        return RefreshVisualSnapshot(w, h);
    }

    /// <summary>
    /// Refresca la imagen visual del control ActiveX para las dimensiones dadas,
    /// o activa el FallbackVisual si la captura nativa no es válida.
    /// </summary>
    public bool RefreshVisualSnapshot(int widthPx = 0, int heightPx = 0)
    {
        if (widthPx <= 0)
        {
            widthPx = BoundControl != null ? Math.Max(24, (int)Math.Round(BoundControl.WidthPixels)) : Math.Max(24, (int)ActualWidth);
        }
        if (heightPx <= 0)
        {
            heightPx = BoundControl != null ? Math.Max(20, (int)Math.Round(BoundControl.HeightPixels)) : Math.Max(20, (int)ActualHeight);
        }

        if (IsInteractiveEditMode && _interactiveHwndHost != null)
        {
            _interactiveHwndHost.Width = widthPx;
            _interactiveHwndHost.Height = heightPx;
            _axControl.Size = new System.Drawing.Size(widthPx, heightPx);
            HasValidNativeVisual = true;
            return true;
        }

        if (_nativeActivated && !_axControl.HasInitializationError)
        {
            var bmpSource = _axControl.TryCaptureVisualSnapshot(widthPx, heightPx);
            if (bmpSource != null)
            {
                _snapshotImage.Source = bmpSource;
                Child = _snapshotImage;
                HasValidNativeVisual = true;
                return true;
            }
        }

        HasValidNativeVisual = false;
        if (BoundControl != null && FallbackVisualFactory != null)
        {
            try
            {
                var generated = FallbackVisualFactory(BoundControl);
                if (generated != null)
                {
                    _fallbackVisual = generated;
                }
            }
            catch
            {
            }
        }

        if (_fallbackVisual != null)
        {
            _fallbackVisual.Width = widthPx;
            _fallbackVisual.Height = heightPx;
            Child = _fallbackVisual;
        }
        return false;
    }

    private void SyncPropertiesDictionaryToAx()
    {
        if (BoundControl == null) return;

        _isSyncingProperties = true;
        try
        {
            if (!BoundControl.ExtentX.HasValue || BoundControl.ExtentX.Value <= 0 ||
                !BoundControl.ExtentY.HasValue || BoundControl.ExtentY.Value <= 0)
            {
                BoundControl.UpdateExtentsFromDimensions();
            }

            // 1. Sincronizar dimensiones y coordenadas del Extender de forma atómica y segura
            _axControl.Extender.SetBoundsDirect(BoundControl.Left, BoundControl.Top, BoundControl.Width, BoundControl.Height);
            _axControl.Extender.Name = BoundControl.Name;
            _axControl.Extender.Index = (short)(BoundControl.Index ?? -1);
            _axControl.Extender.Visible = BoundControl.Visible;
            _axControl.Extender.Enabled = BoundControl.Enabled;
            _axControl.Extender.TabIndex = (short)(BoundControl.TabIndex ?? 0);
            _axControl.Extender.TabStop = BoundControl.TabStop ?? true;
            _axControl.Extender.HelpContextID = BoundControl.HelpContextID;
            _axControl.Extender.Tag = BoundControl.GetCustomPropertyString("Tag") ?? string.Empty;
            _axControl.Extender.ToolTipText = BoundControl.GetCustomPropertyString("ToolTipText") ?? string.Empty;

        // 2. Configurar Licencia y Estado Binario (.frx / .ctx)
        if (!string.IsNullOrEmpty(BoundControl.LicenseKey))
        {
            _axControl.LicenseKey = BoundControl.LicenseKey;
        }

        if (BoundControl.ActiveXStateBlob is { Length: > 0 })
        {
            _axControl.InitialStreamBlob = BoundControl.ActiveXStateBlob;
        }

        if (BoundControl.BinaryProperties.Count > 0)
        {
            _axControl.SetBinaryProperties(BoundControl.BinaryProperties);

            // Si alguna propiedad binaria corresponde a un StdPicture (Picture, DownPicture, DisabledPicture, MouseIcon, Image, MaskPicture),
            // instanciar IPictureDisp vía OleCreatePictureIndirect
            foreach (var binKvp in BoundControl.BinaryProperties)
            {
                if (IsPicturePropertyName(binKvp.Key) && binKvp.Value is { Length: > 0 })
                {
                    _axControl.SetPictureProperty(binKvp.Key, binKvp.Value);
                }
            }
        }

        // 3. Configurar StdFont (IFontDisp) y StdPicture (IPictureDisp) principales
        if (BoundControl.Font != null)
        {
            _axControl.ConfigureAmbientFont(
                BoundControl.Font.Name,
                (float)BoundControl.Font.Size,
                BoundControl.Font.IsBold,
                BoundControl.Font.Italic,
                BoundControl.Font.Underline,
                BoundControl.Font.Strikethrough,
                (short)BoundControl.Font.Charset);

            _axControl.SetFontProperty(
                "Font",
                BoundControl.Font.Name,
                (float)BoundControl.Font.Size,
                BoundControl.Font.IsBold,
                BoundControl.Font.Italic,
                BoundControl.Font.Underline,
                BoundControl.Font.Strikethrough,
                (short)BoundControl.Font.Charset);
        }

        // Configurar fuentes expuestas adicionales (ej. IconFont)
        foreach (var fontKvp in BoundControl.FontProperties)
        {
            var f = fontKvp.Value;
            if (f != null && !string.IsNullOrWhiteSpace(f.Name) && !fontKvp.Key.Equals("Font", StringComparison.OrdinalIgnoreCase))
            {
                _axControl.SetFontProperty(
                    fontKvp.Key,
                    f.Name,
                    (float)f.Size,
                    f.IsBold,
                    f.Italic,
                    f.Underline,
                    f.Strikethrough,
                    (short)f.Charset);
            }
        }

        if (BoundControl.PictureData is { Length: > 0 })
        {
            _axControl.SetPictureProperty("Picture", BoundControl.PictureData);
        }
        if (BoundControl.DownPictureData is { Length: > 0 })
        {
            _axControl.SetPictureProperty("DownPicture", BoundControl.DownPictureData);
        }
        if (BoundControl.DisabledPictureData is { Length: > 0 })
        {
            _axControl.SetPictureProperty("DisabledPicture", BoundControl.DisabledPictureData);
        }

        // 4. Extraer sub-bloques BeginProperty de tipo StdFont ({0BE35203-8F91-11CE-9DE3-00AA004BB851}) como HeadFont, FixedFont, etc.
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in BoundControl.CustomProperties)
        {
            if (p.IsBeginProperty)
            {
                if (TryParseStdFontBeginProperty(p, out var subFont))
                {
                    _axControl.SetFontProperty(
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

        if (!string.IsNullOrEmpty(BoundControl.Caption)) props["Caption"] = BoundControl.Caption;
        if (!string.IsNullOrEmpty(BoundControl.Text)) props["Text"] = BoundControl.Text;
        if (BoundControl.BackColorHex != null) props["BackColor"] = BoundControl.BackColorHex;
        if (BoundControl.ForeColorHex != null) props["ForeColor"] = BoundControl.ForeColorHex;
        if (BoundControl.BorderStyle.HasValue) props["BorderStyle"] = BoundControl.BorderStyle.Value.ToString();

        _axControl.ApplyProperties(props);
        }
        finally
        {
            _isSyncingProperties = false;
        }
    }

    private ActiveXWin32HwndHost? _interactiveHwndHost;

    /// <summary>
    /// Indica si el control se encuentra actualmente en modo de edición interactiva in-situ (OLEIVERB_UIACTIVATE).
    /// </summary>
    public bool IsInteractiveEditMode { get; private set; }

    internal static bool IsPicturePropertyName(string propName)
    {
        if (string.IsNullOrWhiteSpace(propName)) return false;
        if (propName.Contains("Font", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Color", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Code", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Size", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return propName.Contains("Picture", StringComparison.OrdinalIgnoreCase) ||
               propName.Equals("MouseIcon", StringComparison.OrdinalIgnoreCase) ||
               propName.Equals("Icon", StringComparison.OrdinalIgnoreCase) ||
               propName.Equals("Image", StringComparison.OrdinalIgnoreCase) ||
               propName.Equals("Palette", StringComparison.OrdinalIgnoreCase);
    }

    internal static bool TryParseStdFontBeginProperty(FormPropertyItem item, out FormFontDescription font)
    {
        font = new FormFontDescription();
        if (!item.IsBeginProperty || item.RawPropertyLines.Count < 2)
            return false;

        string header = item.RawPropertyLines[0];
        bool isStdFontGuid = header.Contains("0BE35203-8F91-11CE-9DE3-00AA004BB851", StringComparison.OrdinalIgnoreCase);
        bool isFontName = item.Name.EndsWith("Font", StringComparison.OrdinalIgnoreCase);
        if (!isStdFontGuid && !isFontName)
            return false;

        bool foundAny = false;
        for (int i = 1; i < item.RawPropertyLines.Count - 1; i++)
        {
            string line = item.RawPropertyLines[i].Trim();
            int eq = line.IndexOf('=');
            if (eq <= 0) continue;

            string key = line[..eq].Trim();
            string val = line[(eq + 1)..].Trim();
            int commentIdx = val.IndexOf('\'');
            if (commentIdx >= 0 && !val.StartsWith('\"'))
            {
                val = val[..commentIdx].Trim();
            }

            if (key.Equals("Name", StringComparison.OrdinalIgnoreCase))
            {
                font.Name = val.Trim('\"');
                foundAny = true;
            }
            else if (key.Equals("Size", StringComparison.OrdinalIgnoreCase) &&
                     double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double sz))
            {
                font.Size = sz;
                foundAny = true;
            }
            else if (key.Equals("Charset", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(val, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int cs))
            {
                font.Charset = cs;
            }
            else if (key.Equals("Weight", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(val, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int wt))
            {
                font.Weight = wt;
            }
            else if (key.Equals("Underline", StringComparison.OrdinalIgnoreCase))
            {
                font.Underline = val == "-1" || val.Equals("True", StringComparison.OrdinalIgnoreCase);
            }
            else if (key.Equals("Italic", StringComparison.OrdinalIgnoreCase))
            {
                font.Italic = val == "-1" || val.Equals("True", StringComparison.OrdinalIgnoreCase);
            }
            else if (key.Equals("Strikethrough", StringComparison.OrdinalIgnoreCase))
            {
                font.Strikethrough = val == "-1" || val.Equals("True", StringComparison.OrdinalIgnoreCase);
            }
        }

        return foundAny;
    }

    /// <summary>
    /// Activa el modo de edición interactiva in-situ (OLEIVERB_UIACTIVATE / verbo "Editar" en VB6),
    /// montando temporalmente el HWND nativo del OCX dentro del contenedor para permitir cambiar pestañas,
    /// redimensionar columnas o editar celdas directamente en el diseñador.
    /// </summary>
    public bool EnterInteractiveEditMode()
    {
        if (_disposed || !_nativeActivated || _axControl.HasInitializationError) return false;
        if (IsInteractiveEditMode) return true;

        int w = BoundControl != null ? Math.Max(24, (int)Math.Round(BoundControl.WidthPixels)) : Math.Max(24, (int)ActualWidth);
        int h = BoundControl != null ? Math.Max(20, (int)Math.Round(BoundControl.HeightPixels)) : Math.Max(20, (int)ActualHeight);

        _interactiveHwndHost = new ActiveXWin32HwndHost(_axControl, attachExistingControl: true)
        {
            Width = w,
            Height = h
        };

        Child = _interactiveHwndHost;
        IsInteractiveEditMode = true;
        return true;
    }

    /// <summary>
    /// Sale del modo de edición interactiva in-situ (UIDeactivate), sincronizando el estado modificado
    /// hacia el modelo del formulario y restaurando el snapshot libre de Airspace.
    /// </summary>
    public void ExitInteractiveEditMode()
    {
        if (!IsInteractiveEditMode) return;
        IsInteractiveEditMode = false;

        int w = BoundControl != null ? Math.Max(24, (int)Math.Round(BoundControl.WidthPixels)) : Math.Max(24, (int)ActualWidth);
        int h = BoundControl != null ? Math.Max(20, (int)Math.Round(BoundControl.HeightPixels)) : Math.Max(20, (int)ActualHeight);

        if (_interactiveHwndHost != null)
        {
            try
            {
                Child = null;
                _interactiveHwndHost.Dispose();
            }
            catch { }
            _interactiveHwndHost = null;
        }

        _axControl.DetachToSharedParking(w, h);
        SyncToDesignerModel();
        RefreshVisualSnapshot(w, h);
    }

    /// <summary>
    /// Obtiene la lista de verbos OLE expuestos por el control ActiveX (IOleObject::EnumVerbs).
    /// </summary>
    public IReadOnlyList<OleVerbEntry> GetOleVerbs() => _axControl.GetOleVerbs();

    /// <summary>
    /// Ejecuta un verbo OLE sobre el control ActiveX alojado.
    /// </summary>
    public bool ExecuteOleVerb(int verbId, Window? ownerWindow = null)
    {
        if (!_nativeActivated || _axControl.HasInitializationError) return false;

        if (verbId == Ole32Interop.OLEIVERB_PROPERTIES)
        {
            return ShowPropertyPages(ownerWindow);
        }

        if (verbId == Ole32Interop.OLEIVERB_UIACTIVATE || verbId == Ole32Interop.OLEIVERB_PRIMARY || verbId == Ole32Interop.OLEIVERB_OPEN)
        {
            return EnterInteractiveEditMode();
        }

        IntPtr hwndOwner = ownerWindow != null ? new WindowInteropHelper(ownerWindow).Handle : IntPtr.Zero;
        bool ok = _axControl.ExecuteOleVerb(verbId, hwndOwner);
        if (ok)
        {
            SyncToDesignerModel();
            RefreshVisualSnapshot();
        }
        return ok;
    }

    /// <summary>
    /// Muestra la hoja de propiedades nativa OLE del control ActiveX alojado.
    /// </summary>
    public bool ShowPropertyPages(Window? ownerWindow = null)
    {
        if (!_nativeActivated || _axControl.HasInitializationError) return false;

        IntPtr hwndOwner = IntPtr.Zero;
        if (ownerWindow != null)
        {
            hwndOwner = new WindowInteropHelper(ownerWindow).Handle;
        }

        string caption = BoundControl != null
            ? $"Propiedades de {BoundControl.Name} ({BoundControl.TypeName})"
            : $"Propiedades de ActiveX {ProgId}";

        bool shown = _axControl.ShowPropertyPages(hwndOwner, caption);
        if (shown)
        {
            SyncToDesignerModel();
            RefreshVisualSnapshot();
        }
        return shown;
    }

    /// <summary>
    /// Sincroniza las propiedades modificadas en la hoja de propiedades nativa OLE hacia el modelo BoundControl,
    /// incluyendo el estado binario IPersistStreamInit / IPersistPropertyBag para su persistencia en .frx / .ctx.
    /// </summary>
    public void SyncToDesignerModel()
    {
        if (BoundControl == null || !_nativeActivated || _axControl.HasInitializationError) return;

        _axControl.SavePersistenceState(clearDirty: true);

        if (_axControl.SavedStreamBlob is { Length: > 0 })
        {
            BoundControl.ActiveXStateBlob = _axControl.SavedStreamBlob;
            BoundControl.HasModifiedActiveXState = true;
            if (string.IsNullOrWhiteSpace(BoundControl.ActiveXStatePropertyName))
            {
                BoundControl.ActiveXStatePropertyName = "Object.zyx";
            }
        }

        if (_axControl.SavedBinaryProperties.Count > 0)
        {
            foreach (var binKvp in _axControl.SavedBinaryProperties)
            {
                BoundControl.BinaryProperties[binKvp.Key] = binKvp.Value;
                BoundControl.HasModifiedActiveXState = true;
            }
        }

        foreach (var p in BoundControl.CustomProperties)
        {
            if (!p.IsBeginProperty &&
                !p.Value.Contains(".frx\":", StringComparison.OrdinalIgnoreCase) &&
                !p.Value.Contains(".ctx\":", StringComparison.OrdinalIgnoreCase))
            {
                var liveVal = _axControl.GetPropertyValue(p.Name);
                if (liveVal != null)
                {
                    p.Value = liveVal.ToString() ?? string.Empty;
                }
            }
        }
        foreach (var kv in _axControl.DiscoveredProperties)
        {
            var liveVal = _axControl.GetPropertyValue(kv.Key);
            if (liveVal != null)
            {
                var existing = BoundControl.CustomProperties.Find(p => p.Name.Equals(kv.Key, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    if (!existing.Value.Contains(".frx\":", StringComparison.OrdinalIgnoreCase) &&
                        !existing.Value.Contains(".ctx\":", StringComparison.OrdinalIgnoreCase))
                    {
                        existing.Value = liveVal.ToString() ?? string.Empty;
                    }
                }
                else
                {
                    BoundControl.CustomProperties.Add(new FormPropertyItem
                    {
                        Name = kv.Key,
                        Value = liveVal.ToString() ?? string.Empty
                    });
                }
            }
        }
    }

    /// <summary>
    /// Sincroniza las propiedades del modelo de diseño FormDesignerControl hacia el control ActiveX nativo.
    /// </summary>
    public void SyncFromDesignerModel()
    {
        if (BoundControl == null) return;
        SyncPropertiesDictionaryToAx();

        if (_nativeActivated && !_axControl.HasInitializationError)
        {
            RefreshVisualSnapshot();
        }
    }

    /// <summary>
    /// Sincroniza un valor de propiedad individual modificado en el PropertyGrid hacia el control nativo.
    /// </summary>
    public void SetLiveProperty(string propertyName, object? value)
    {
        _axControl.SetPropertyValue(propertyName, value);
        if (_nativeActivated && !_axControl.HasInitializationError)
        {
            RefreshVisualSnapshot();
        }
    }

    /// <summary>
    /// Lee el valor de una propiedad directamente del control COM IDispatch.
    /// </summary>
    public object? GetLiveProperty(string propertyName)
    {
        return _axControl.GetPropertyValue(propertyName);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            if (_interactiveHwndHost != null)
            {
                Child = null;
                _interactiveHwndHost.Dispose();
                _interactiveHwndHost = null;
            }
        }
        catch { }

        try
        {
            _axControl.Dispose();
        }
        catch
        {
            // Ignorar errores al liberar recursos COM
        }
    }
}
