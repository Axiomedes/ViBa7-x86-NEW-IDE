using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using ViBa7.ActiveXHost.Diagnostics;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.ActiveX;

namespace ViBa7.ActiveXHost;

public class DesignMouseEventArgs : EventArgs
{
    public int X { get; }
    public int Y { get; }
    public MouseButtons Button { get; }
    public int Clicks { get; }

    public DesignMouseEventArgs(int x, int y, MouseButtons button, int clicks)
    {
        X = x;
        Y = y;
        Button = button;
        Clicks = clicks;
    }
}

public enum ActiveXLifecycleState
{
    Uninitialized = 0,
    Created = 1,
    Sited = 2,
    Loaded = 3,
    Visible = 4,
    InPlaceActive = 5,
    UIActive = 6,
    Faulted = 7,
    Disposed = 8
}

public sealed record ActiveXDiagnosticEntry(
    DateTime TimestampUtc,
    string Clsid,
    string ControlName,
    string Stage,
    int HResult,
    string Message)
{
    public string HResultHex => $"0x{HResult:X8}";
}

public sealed record OleVerbEntry(
    int VerbId,
    string Name,
    int Flags,
    int Attributes);

/// <summary>
/// Implementación COM de IPropertyBag e IErrorLog para inicializar controles OCX de VB6 y terceros
/// mediante IPersistPropertyBag::Load, garantizando que UserControl_ReadProperties reciba los valores
/// del archivo .frm (incluyendo blobs binarios .frx/.ctx) o utilice los valores por defecto del autor del control.
/// </summary>
public sealed class VbPropertyBag : IPropertyBag, IErrorLog
{
    private const int S_OK = 0;
    private const int E_INVALIDARG = unchecked((int)0x80070057);

    private readonly Dictionary<string, object> _values = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyDictionary<string, object> SavedProperties => _values;

    public VbPropertyBag(IDictionary<string, object>? initialValues = null)
    {
        if (initialValues != null)
        {
            foreach (var kvp in initialValues)
            {
                _values[kvp.Key] = kvp.Value;
            }
        }
    }

    public int Read(string pszPropName, ref object pVar, IErrorLog? pErrorLog)
    {
        if (string.IsNullOrWhiteSpace(pszPropName))
            return E_INVALIDARG;

        if (_values.TryGetValue(pszPropName, out var storedVal) && storedVal != null)
        {
            if (storedVal is byte[] rawBytes)
            {
                pVar = rawBytes;
                return S_OK;
            }

            try
            {
                Type? targetType = pVar?.GetType();
                if (targetType != null && targetType != typeof(object))
                {
                    if (targetType == typeof(bool))
                    {
                        string str = storedVal.ToString()?.Trim() ?? string.Empty;
                        pVar = str == "-1" || str.Equals("True", StringComparison.OrdinalIgnoreCase);
                        return S_OK;
                    }
                    if (targetType == typeof(int) || targetType == typeof(short) || targetType == typeof(long))
                    {
                        string str = storedVal.ToString()?.Trim() ?? string.Empty;
                        if (str.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
                        {
                            long hexVal = Convert.ToInt64(str[2..].TrimEnd('&'), 16);
                            pVar = Convert.ChangeType(hexVal, targetType, CultureInfo.InvariantCulture);
                            return S_OK;
                        }
                    }

                    pVar = Convert.ChangeType(storedVal, targetType, CultureInfo.InvariantCulture);
                }
                else
                {
                    pVar = storedVal;
                }
                return S_OK;
            }
            catch
            {
                pVar = storedVal;
                return S_OK;
            }
        }

        // Retornar E_INVALIDARG indica a msvbvm60.dll que debe usar el DefaultValue de ReadProperty(Name, Default)
        return E_INVALIDARG;
    }

    public int Write(string pszPropName, ref object pVar)
    {
        if (!string.IsNullOrWhiteSpace(pszPropName) && pVar != null)
        {
            _values[pszPropName] = pVar;
        }
        return S_OK;
    }

    public int AddError(string pszPropName, IntPtr pExcepInfo) => S_OK;
}

/// <summary>
/// Sitio de diseño (ISite) que informa al contenedor y al control ActiveX que se encuentra en Modo Diseño
/// (DISPID_AMBIENT_USERMODE = False).
/// </summary>
public sealed class ActiveXDesignSite : ISite
{
    private readonly DynamicActiveXControl _control;
    private string _name;

    public ActiveXDesignSite(DynamicActiveXControl control, string name = "ActiveXControl1")
    {
        _control = control;
        _name = name;
    }

    public IComponent Component => _control;
    public IContainer? Container => null;
    public bool DesignMode => _control.IsDesignMode;

    public string? Name
    {
        get => _name;
        set => _name = value ?? "ActiveXControl1";
    }

    public object? GetService(Type serviceType) => null;
}

/// <summary>
/// Sitio OLE nativo (IOleClientSite + IOleInPlaceSite + IOleInPlaceFrame + IOleControlSite + ISimpleFrameSite + IPropertyNotifySink + IAdviseSink + Ambient IDispatch)
/// que hospeda controles ActiveX tanto con ventana (Windowed) como sin ventana (Windowless) sin depender
/// de System.Windows.Forms.AxHost, proveyendo punteros de marco OLE reales para controles de MFC, Delphi y Sheridan.
/// </summary>
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class OleDesignClientSite : IOleClientSite, IOleInPlaceSite, IOleInPlaceFrame, IOleControlSite, ISimpleFrameSite, IPropertyNotifySink, IAdviseSink
{
    private const int S_OK = 0;
    private const int S_FALSE = 1;
    private const int E_NOTIMPL = unchecked((int)0x80004001);
    private const int E_NOINTERFACE = unchecked((int)0x80004002);

    private readonly DynamicActiveXControl _owner;

    public OleDesignClientSite(DynamicActiveXControl owner)
    {
        _owner = owner;
    }

    // Propiedades Ambientales estándar OLE (DISPID_AMBIENT_*) consultadas por msvbvm60.dll y controles vía IDispatch
    [DispId(-701)]
    public int BackColor => _owner.AmbientBackColorOle;

    [DispId(-702)]
    public string DisplayName => !string.IsNullOrWhiteSpace(_owner.Name) ? _owner.Name : "ActiveXControl1";

    [DispId(-703)]
    public object? Font => _owner.GetOrCreateAmbientStdFont();

    [DispId(-704)]
    public int ForeColor => _owner.AmbientForeColorOle;

    [DispId(-705)]
    public int LocaleID => CultureInfo.CurrentUICulture.LCID;

    [DispId(-706)]
    public bool MessageReflect => true;

    [DispId(-707)]
    public string ScaleUnits => "Twip";

    [DispId(-708)]
    public short TextAlign => 0;

    [DispId(-709)]
    public bool UserMode => !_owner.IsDesignMode;

    [DispId(-710)]
    public bool UIDead => false;

    [DispId(-711)]
    public bool ShowGrabHandles => false;

    [DispId(-712)]
    public bool ShowHatching => false;

    [DispId(-713)]
    public bool DisplayAsDefault => false;

    [DispId(-714)]
    public bool SupportsMnemonics => true;

    [DispId(-715)]
    public bool AutoClip => true;

    [DispId(-716)]
    public short Appearance => _owner.AmbientAppearance;

    [DispId(-725)]
    public bool RightToLeftLegacy => _owner.AmbientRightToLeft;

    [DispId(-726)]
    public IntPtr Palette => IntPtr.Zero;

    [DispId(-732)]
    public bool RightToLeft => _owner.AmbientRightToLeft;

    // IOleClientSite
    public int SaveObject()
    {
        _owner.SavePersistenceState(clearDirty: true);
        return S_OK;
    }

    public int GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk)
    {
        ppmk = IntPtr.Zero;
        return E_NOTIMPL;
    }

    public int GetContainer(out IntPtr ppContainer)
    {
        ppContainer = IntPtr.Zero;
        return E_NOINTERFACE;
    }

    public int ShowObject() => S_OK;

    public int OnShowWindow(bool fShow) => S_OK;

    public int RequestNewObjectLayout() => E_NOTIMPL;

    // IOleInPlaceSite & IOleInPlaceUIWindow & IOleInPlaceFrame
    public int GetWindow(out IntPtr phwnd)
    {
        phwnd = _owner.GetHostWindowHandle();
        return phwnd != IntPtr.Zero ? S_OK : E_NOTIMPL;
    }

    public int ContextSensitiveHelp(bool fEnterMode) => S_OK;

    public int CanInPlaceActivate() => S_OK;

    public int OnInPlaceActivate() => S_OK;

    public int OnUIActivate() => S_OK;

    public int GetWindowContext(
        out IntPtr ppFrame,
        out IntPtr ppDoc,
        out COMRECT lprcPosRect,
        out COMRECT lprcClipRect,
        ref OLEINPLACEFRAMEINFO lpFrameInfo)
    {
        // Vital: Controles de MFC (COleControl::OnActivateInPlace) y Delphi desreferencian ppFrame sin validar null.
        ppFrame = Marshal.GetComInterfaceForObject(this, typeof(IOleInPlaceFrame));
        ppDoc = IntPtr.Zero;
        int w = Math.Max(16, _owner.Width);
        int h = Math.Max(16, _owner.Height);
        lprcPosRect = new COMRECT(0, 0, w, h);
        lprcClipRect = new COMRECT(0, 0, w, h);

        if (lpFrameInfo.cb >= Marshal.SizeOf<OLEINPLACEFRAMEINFO>())
        {
            lpFrameInfo.fMDIApp = 0;
            lpFrameInfo.hwndFrame = _owner.GetHostWindowHandle();
            lpFrameInfo.hAccel = IntPtr.Zero;
            lpFrameInfo.cAccelEntries = 0;
        }
        return S_OK;
    }

    public int Scroll(SIZEL scrollExtant) => S_OK;

    public int OnUIDeactivate(bool fUndoable) => S_OK;

    public int OnInPlaceDeactivate() => S_OK;

    public int DiscardUndoState() => S_OK;

    public int DeactivateAndUndo() => S_OK;

    public int OnPosRectChange(ref COMRECT lprcPosRect) => S_OK;

    // IOleInPlaceUIWindow
    public int GetBorder(ref COMRECT lprectBorder) => S_OK;
    public int RequestBorderSpace(ref COMRECT pborderwidths) => S_OK;
    public int SetBorderSpace(ref COMRECT pborderwidths) => S_OK;
    public int SetActiveObject(object? pActiveObj, string? pszObjName) => S_OK;

    // IOleInPlaceFrame
    public int InsertMenus(IntPtr hmenuShared, IntPtr lpMenuWidths) => S_OK;
    public int SetMenu(IntPtr hmenuShared, IntPtr holemenu, IntPtr hwndActiveObject) => S_OK;
    public int RemoveMenus(IntPtr hmenuShared) => S_OK;
    public int SetStatusText(string pszStatusText) => S_OK;
    public int EnableModeless(bool fEnable) => S_OK;
    public int TranslateAccelerator(ref Native.MSG lpmsg, ushort wID) => S_FALSE;

    // IOleControlSite
    public int OnControlInfoChanged() => S_OK;
    public int LockInPlaceActive(bool fLock) => S_OK;

    public int GetExtendedControl(out object? ppDisp)
    {
        ppDisp = _owner.Extender;
        return ppDisp != null ? S_OK : E_NOTIMPL;
    }

    public int TransformCoords(ref POINT pPtlHimetric, ref POINTF pPtfContainer, uint dwFlags)
    {
        // Conversión matemática real entre HIMETRIC (2540 unidades/pulgada) y Twips (1440 unidades/pulgada)
        if ((dwFlags & Ole32Interop.XFORMCOORDS_HIMETRICTOCONTAINER) != 0)
        {
            pPtfContainer.x = (float)(pPtlHimetric.x * 1440.0 / 2540.0);
            pPtfContainer.y = (float)(pPtlHimetric.y * 1440.0 / 2540.0);
            return S_OK;
        }

        if ((dwFlags & Ole32Interop.XFORMCOORDS_CONTAINERTOHIMETRIC) != 0)
        {
            pPtlHimetric.x = (int)Math.Round(pPtfContainer.x * 2540.0 / 1440.0);
            pPtlHimetric.y = (int)Math.Round(pPtfContainer.y * 2540.0 / 1440.0);
            return S_OK;
        }

        return E_NOTIMPL;
    }

    public int TranslateAccelerator(ref Native.MSG pMsg, uint grfModifiers) => S_FALSE;
    public int OnFocus(bool fGotFocus) => S_OK;
    public int ShowPropertyFrame() => S_OK;

    // ISimpleFrameSite (para controles contenedores como SSTab y Sheridan Frame/Panel)
    public int PreMessageFilter(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp, out IntPtr plResult, out uint pdwCookie)
    {
        plResult = IntPtr.Zero;
        pdwCookie = 0;
        return S_OK;
    }

    public int PostMessageFilter(IntPtr hWnd, uint msg, IntPtr wp, IntPtr lp, out IntPtr plResult, uint dwCookie)
    {
        plResult = IntPtr.Zero;
        return S_OK;
    }

    // IPropertyNotifySink
    public int OnChanged(int dispID)
    {
        _owner.RaisePropertyChangedByOcx(dispID);
        return S_OK;
    }

    public int OnRequestEdit(int dispID) => S_OK;

    // IAdviseSink
    public void OnDataChange(IntPtr pFormatetc, IntPtr pStgmed) { }

    public void OnViewChange(int dwAspect, int lindex)
    {
        _owner.RaiseViewInvalidatedByOcx();
    }

    public void OnRename(IntPtr pmk) { }
    public void OnSave() { }
    public void OnClose() { }
}

/// <summary>
/// Contenedor OLE directo que aloja dinámicamente cualquier control ActiveX (OCX) con o sin ventana
/// (Windowed y Windowless) por su CLSID, sin heredar de System.Windows.Forms.AxHost (evitando el fallo
/// NotSupportedException en EnsureWindowPresent y los cuadros de excepción de NativeWindow.Callback).
/// </summary>
public class DynamicActiveXControl : IComponent, IDisposable
{
    [ThreadStatic]
    private static IntPtr _sharedParkingHwnd;

    private static readonly HashSet<string> _faultedClsids = new(StringComparer.OrdinalIgnoreCase);

    private readonly Guid _clsidGuid;
    private readonly OleDesignClientSite _oleClientSite;
    private readonly Dictionary<string, string> _pendingProperties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _pendingBinaryProperties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, object> _pendingObjectProperties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _discoveredProperties = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _savedBinaryProperties = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<ActiveXDiagnosticEntry> _diagnostics = new();

    private object? _ocxInstance;
    private object? _cachedAmbientStdFont;
    private IConnectionPoint? _propertyNotifyConnectionPoint;
    private int _propertyNotifyCookie;
    private IConnectionPoint? _eventConnectionPoint;
    private int _eventCookie;
    private ActiveXEventSink? _eventSink;
    private readonly List<OcxEventInfo> _knownEvents = new();
    private IntPtr _customHostHwnd;
    private bool _persistenceInitialized;
    private bool _disposed;
    private int _width;
    private int _height;

    public string Clsid { get; }

    /// <summary>
    /// ProgID o TypeName asociado a este control ActiveX.
    /// </summary>
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string? ProgId { get; set; }

    /// <summary>
    /// GUID de la interfaz Source de eventos COM configurada para este control (e.g. "{...}").
    /// </summary>
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string? EventsInterfaceIid { get; set; }

    /// <summary>
    /// Sumidero de eventos COM actualmente conectado a este control ActiveX.
    /// </summary>
    [Browsable(false)]
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public ActiveXEventSink? EventSink => _eventSink;

    /// <summary>
    /// Objeto extendido VBControlExtender asociado a este control ActiveX.
    /// </summary>
    public VbControlExtender Extender { get; }

    /// <summary>
    /// Evento disparado cuando el control ActiveX notifica cambio en una propiedad mediante IPropertyNotifySink::OnChanged.
    /// </summary>
    public event Action<int>? PropertyChangedByOcx;

    /// <summary>
    /// Evento disparado cuando el control ActiveX solicita repintado visual mediante IAdviseSink::OnViewChange.
    /// </summary>
    public event Action? ViewInvalidatedByOcx;

    /// <summary>
    /// Evento disparado cuando el objeto VbControlExtender cambia sus coordenadas o dimensiones.
    /// </summary>
    public event Action? ExtenderBoundsChanged;

    /// <summary>
    /// Evento disparado cada vez que el control ActiveX emite un evento COM a través de su interfaz Source (IConnectionPoint).
    /// </summary>
    public event Action<ActiveXEventInvocation>? ComEventFired;

    /// <summary>
    /// Estado actual del ciclo de vida OLE del control ActiveX.
    /// </summary>
    public ActiveXLifecycleState LifecycleState { get; private set; } = ActiveXLifecycleState.Uninitialized;

    /// <summary>
    /// Registro estructurado de eventos y diagnósticos OLE/COM durante la vida del control.
    /// </summary>
    public IReadOnlyList<ActiveXDiagnosticEntry> Diagnostics => _diagnostics;

    /// <summary>
    /// Clave de licencia de diseño/ejecución obtenida o proporcionada mediante IClassFactory2.
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string? LicenseKey { get; set; }

    /// <summary>
    /// Indica si el control implementa IClassFactory2 y requiere licencia.
    /// </summary>
    public bool RequiresLicense { get; private set; }

    /// <summary>
    /// Indica si la licencia de diseño está verificada en la máquina actual.
    /// </summary>
    public bool IsLicenseVerified { get; private set; } = true;

    /// <summary>
    /// Blob binario inicial (.frx / .ctx) para restaurar mediante IPersistStreamInit::Load o IPersistStream::Load.
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public byte[]? InitialStreamBlob { get; set; }

    /// <summary>
    /// Último blob binario capturado desde IPersistStreamInit::Save o IPersistStream::Save.
    /// </summary>
    public byte[]? SavedStreamBlob { get; private set; }

    /// <summary>
    /// Propiedades binarias guardadas por el control durante IPersistPropertyBag::Save.
    /// </summary>
    public IReadOnlyDictionary<string, byte[]> SavedBinaryProperties => _savedBinaryProperties;

    /// <summary>
    /// Propiedades persistentes o expuestas descubiertas en vivo mediante IPersistPropertyBag e IDispatch.
    /// </summary>
    public IReadOnlyDictionary<string, string> DiscoveredProperties => _discoveredProperties;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string Name { get; set; } = "ActiveXControl1";

    /// <summary>
    /// Cuando es true, el control está en Modo Diseño (DISPID_AMBIENT_USERMODE = False).
    /// </summary>
    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool IsDesignMode { get; set; } = true;

    /// <summary>
    /// Indica si ocurrió un error durante la inicialización o activación COM del control.
    /// </summary>
    public bool HasInitializationError { get; private set; }

    /// <summary>
    /// Indica si el control OCX es sin ventana (Windowless) o no creó un HWND propio.
    /// </summary>
    public bool IsWindowless { get; private set; }

    /// <summary>
    /// Indica si el control declara OLEMISC_SIMPLEFRAME (contenedor de otros controles como SSTab o Sheridan Frame).
    /// </summary>
    public bool IsSimpleFrame { get; private set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int AmbientBackColorOle { get; set; } = unchecked((int)0x8000000F);

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int AmbientForeColorOle { get; set; } = unchecked((int)0x80000012);

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public string AmbientFontName { get; set; } = "MS Sans Serif";

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public float AmbientFontSize { get; set; } = 8.25f;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientFontBold { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientFontItalic { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientFontUnderline { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientFontStrikethrough { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public short AmbientFontCharset { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public short AmbientAppearance { get; set; } = 1;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientRightToLeft { get; set; }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public bool AmbientEnabled { get; set; } = true;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public ISite? Site { get; set; }

    public event EventHandler? Disposed;

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int Width
    {
        get => _width;
        set => SetBounds(value, _height);
    }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public int Height
    {
        get => _height;
        set => SetBounds(_width, value);
    }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public System.Drawing.Size Size
    {
        get => new(_width, _height);
        set => SetBounds(value.Width, value.Height);
    }

#pragma warning disable CS0067
    public event EventHandler<DesignMouseEventArgs>? DesignMouseDown;
    public event EventHandler<DesignMouseEventArgs>? DesignMouseMove;
    public event EventHandler<DesignMouseEventArgs>? DesignMouseUp;
    public event EventHandler<DesignMouseEventArgs>? DesignMouseDoubleClick;
    public event EventHandler<DesignMouseEventArgs>? DesignContextMenu;
#pragma warning restore CS0067

    public DynamicActiveXControl(string clsid, int initialWidth = 120, int initialHeight = 36)
    {
        Ole32Interop.MaskFpuExceptions();
        Clsid = NormalizeClsid(clsid);
        _clsidGuid = Guid.TryParse(Clsid, out var g) ? g : Guid.Empty;
        _width = Math.Max(16, initialWidth);
        _height = Math.Max(16, initialHeight);
        _oleClientSite = new OleDesignClientSite(this);
        Extender = new VbControlExtender(this);
    }

    internal object? GetOrCreateAmbientStdFont()
    {
        _cachedAmbientStdFont ??= Ole32Interop.CreateStdFont(
            AmbientFontName,
            AmbientFontSize,
            AmbientFontBold,
            AmbientFontItalic,
            AmbientFontUnderline,
            AmbientFontStrikethrough,
            AmbientFontCharset);
        return _cachedAmbientStdFont;
    }

    public void ConfigureAmbientFont(
        string fontName,
        float sizePoints,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        bool strikethrough = false,
        short charset = 0)
    {
        AmbientFontName = string.IsNullOrWhiteSpace(fontName) ? "MS Sans Serif" : fontName;
        AmbientFontSize = Math.Max(1.0f, sizePoints);
        AmbientFontBold = bold;
        AmbientFontItalic = italic;
        AmbientFontUnderline = underline;
        AmbientFontStrikethrough = strikethrough;
        AmbientFontCharset = charset;
        _cachedAmbientStdFont = null;
        NotifyAmbientPropertyChanged(-703);
    }

    public void SetFontProperty(
        string propertyName,
        string fontName,
        float sizePoints,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        bool strikethrough = false,
        short charset = 0)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return;
        object? stdFont = Ole32Interop.CreateStdFont(fontName, sizePoints, bold, italic, underline, strikethrough, charset);
        if (stdFont != null)
        {
            _pendingObjectProperties[propertyName] = stdFont;
            if (_ocxInstance != null && !HasInitializationError)
            {
                TryAssignComObjectProperty(_ocxInstance, propertyName, stdFont);
            }
        }
    }

    public void SetPictureProperty(string propertyName, byte[]? imageBytes)
    {
        if (string.IsNullOrWhiteSpace(propertyName) || imageBytes == null || imageBytes.Length == 0) return;
        object? stdPict = Ole32Interop.CreateStdPictureFromBytes(imageBytes);
        if (stdPict != null)
        {
            _pendingObjectProperties[propertyName] = stdPict;
            if (_ocxInstance != null && !HasInitializationError)
            {
                TryAssignComObjectProperty(_ocxInstance, propertyName, stdPict);
            }
        }
    }

    public void NotifyAmbientPropertyChanged(int dispId = -1)
    {
        if (_ocxInstance is IOleControl oleControl)
        {
            try
            {
                oleControl.OnAmbientPropertyChange(dispId);
            }
            catch
            {
                // Ignorar si el control no procesa la notificación ambiental
            }
        }
    }

    internal void RaisePropertyChangedByOcx(int dispId)
    {
        PropertyChangedByOcx?.Invoke(dispId);
    }

    internal void RaiseViewInvalidatedByOcx()
    {
        ViewInvalidatedByOcx?.Invoke();
    }

    internal void NotifyExtenderBoundsChanged()
    {
        ExtenderBoundsChanged?.Invoke();
    }

    internal void RequestFocusFromExtender()
    {
        FocusControl();
    }

    internal void RequestZOrderFromExtender(int position)
    {
    }

    internal void RequestRefreshFromExtender()
    {
        RaiseViewInvalidatedByOcx();
    }

    private void RecordDiagnostic(string stage, int hResult, string? detail = null)
    {
        Ole32Interop.TryGetComErrorDescription(out string? errDesc, out string? errSrc);
        string msg = detail ?? errDesc ?? (hResult == 0 ? "OK" : $"HRESULT 0x{hResult:X8}");
        if (!string.IsNullOrWhiteSpace(errSrc))
        {
            msg = $"[{errSrc}] {msg}";
        }

        _diagnostics.Add(new ActiveXDiagnosticEntry(
            DateTime.UtcNow,
            Clsid,
            Name,
            stage,
            hResult,
            msg));

        ActiveXDiagnosticsLogger.Log(stage, Clsid, Name, hResult, detail: msg, errorInfo: errDesc);
    }

    public void SetBinaryProperties(IDictionary<string, byte[]>? binaryProperties)
    {
        _pendingBinaryProperties.Clear();
        if (binaryProperties == null) return;
        foreach (var kvp in binaryProperties)
        {
            if (kvp.Value != null && kvp.Value.Length > 0)
            {
                _pendingBinaryProperties[kvp.Key] = kvp.Value;
            }
        }
    }

    internal IntPtr GetHostWindowHandle()
    {
        if (_customHostHwnd != IntPtr.Zero)
        {
            return _customHostHwnd;
        }
        return EnsureSharedParkingHwnd();
    }

    public IntPtr GetControlWindowHandle()
    {
        if (_ocxInstance is IOleInPlaceObject inPlace &&
            inPlace.GetWindow(out IntPtr ctrlHwnd) == 0 && ctrlHwnd != IntPtr.Zero)
        {
            return ctrlHwnd;
        }
        return IntPtr.Zero;
    }

    internal static IntPtr EnsureSharedParkingHwnd()
    {
        if (_sharedParkingHwnd == IntPtr.Zero)
        {
            _sharedParkingHwnd = Ole32Interop.CreateWindowExW(
                Ole32Interop.WS_EX_TOOLWINDOW | Ole32Interop.WS_EX_NOACTIVATE,
                "STATIC",
                "ViBa7OleParkingHost",
                Ole32Interop.WS_POPUP | Ole32Interop.WS_CLIPCHILDREN | Ole32Interop.WS_CLIPSIBLINGS,
                -32000,
                -32000,
                1024,
                1024,
                IntPtr.Zero,
                IntPtr.Zero,
                IntPtr.Zero,
                IntPtr.Zero);
        }
        return _sharedParkingHwnd;
    }

    private void SetBounds(int width, int height)
    {
        int safeW = Math.Max(16, width);
        int safeH = Math.Max(16, height);
        bool changed = (safeW != _width || safeH != _height);
        _width = safeW;
        _height = safeH;

        if (changed && _ocxInstance != null && !HasInitializationError)
        {
            UpdateOleBoundsAndExtent(safeW, safeH);
        }
    }

    private void UpdateOleBoundsAndExtent(int widthPx, int heightPx)
    {
        using var guard = new NativeDialogGuard();
        try
        {
            if (_ocxInstance is IOleObject oleObj)
            {
                // Conversión de píxeles (96 DPI) a unidades HIMETRIC (2540 por pulgada)
                int cxHimetric = Math.Max(100, (int)Math.Round(widthPx * 2540.0 / 96.0));
                int cyHimetric = Math.Max(100, (int)Math.Round(heightPx * 2540.0 / 96.0));
                var sizel = new SIZEL(cxHimetric, cyHimetric);
                oleObj.SetExtent(Ole32Interop.DVASPECT_CONTENT, ref sizel);
            }

            if (_ocxInstance is IOleInPlaceObject inPlaceObj)
            {
                var rect = new COMRECT(0, 0, widthPx, heightPx);
                inPlaceObj.SetObjectRects(ref rect, ref rect);

                if (inPlaceObj.GetWindow(out IntPtr ctrlHwnd) == 0 && ctrlHwnd != IntPtr.Zero)
                {
                    Ole32Interop.MoveWindow(ctrlHwnd, 0, 0, widthPx, heightPx, true);
                }
            }
        }
        catch
        {
            // Ignorar errores de redimensionamiento en controles sin ventana o pasivos
        }
    }

    private static string NormalizeClsid(string clsid)
    {
        string clean = clsid.Trim();
        if (Guid.TryParse(clean, out Guid g))
        {
            return g.ToString("B").ToUpperInvariant();
        }
        return clean;
    }

    /// <summary>
    /// Registra la lista de eventos conocidos de la TypeLib para este control.
    /// </summary>
    public void RegisterKnownEvents(IEnumerable<OcxEventInfo>? events)
    {
        _knownEvents.Clear();
        if (events != null)
        {
            _knownEvents.AddRange(events);
        }
    }

    private void ConnectSinks(object ocx)
    {
        // 1. Conectar IPropertyNotifySink para cambios en páginas de propiedades o cambios internos del control
        if (ocx is IConnectionPointContainer cpc)
        {
            try
            {
                Guid iidPropNotify = Ole32Interop.IID_IPropertyNotifySink;
                if (cpc.FindConnectionPoint(ref iidPropNotify, out IConnectionPoint? cp) >= 0 && cp != null)
                {
                    if (cp.Advise(_oleClientSite, out int cookie) >= 0 && cookie != 0)
                    {
                        _propertyNotifyConnectionPoint = cp;
                        _propertyNotifyCookie = cookie;
                        RecordDiagnostic("IPropertyNotifySink.Advise", 0);
                    }
                    else
                    {
                        Marshal.ReleaseComObject(cp);
                    }
                }
            }
            catch
            {
                // Ignorar si el control no soporta IPropertyNotifySink
            }
        }

        // 2. Conectar IViewObject::SetAdvise para notificaciones asíncronas de repintado (OnViewChange)
        if (ocx is IViewObject viewObj)
        {
            try
            {
                int hrAdv = viewObj.SetAdvise(Ole32Interop.DVASPECT_CONTENT, 0, _oleClientSite);
                RecordDiagnostic("IViewObject.SetAdvise", hrAdv);
            }
            catch
            {
                // Ignorar si SetAdvise no está implementado
            }
        }

        // 3. Conectar ActiveXEventSink para eventos Source del control COM (IConnectionPoint)
        try
        {
            if (_eventSink != null)
            {
                _eventSink.Dispose();
                _eventSink = null;
            }

            Guid? preferredIid = null;
            if (!string.IsNullOrWhiteSpace(EventsInterfaceIid) && Guid.TryParse(EventsInterfaceIid, out var parsedGuid))
            {
                preferredIid = parsedGuid;
            }

            if (ActiveXEventSink.TryConnect(
                ocx,
                Name,
                ProgId ?? string.Empty,
                Extender.Index >= 0 ? Extender.Index : null,
                preferredIid,
                _knownEvents,
                out var sink,
                out var eventCp,
                out int eventCookie))
            {
                _eventSink = sink;
                _eventConnectionPoint = eventCp;
                _eventCookie = eventCookie;
                if (_eventSink != null)
                {
                    _eventSink.EventTriggered += invocation => ComEventFired?.Invoke(invocation);
                    RecordDiagnostic("IConnectionPoint.Advise", 0, $"EventSink conectado para {Name} ({_eventSink.EventIid}, Cookie={eventCookie})");
                }
            }
        }
        catch (Exception ex)
        {
            RecordDiagnostic("ActiveXEventSink.TryConnect", ex.HResult, ex.Message);
        }
    }

    private void DisconnectSinks()
    {
        if (_propertyNotifyConnectionPoint != null)
        {
            try
            {
                if (_propertyNotifyCookie != 0)
                {
                    _propertyNotifyConnectionPoint.Unadvise(_propertyNotifyCookie);
                }
            }
            catch { }

            try
            {
                if (Marshal.IsComObject(_propertyNotifyConnectionPoint))
                {
                    Marshal.ReleaseComObject(_propertyNotifyConnectionPoint);
                }
            }
            catch { }

            _propertyNotifyConnectionPoint = null;
            _propertyNotifyCookie = 0;
        }

        if (_eventConnectionPoint != null)
        {
            try
            {
                if (_eventCookie != 0)
                {
                    _eventConnectionPoint.Unadvise(_eventCookie);
                }
            }
            catch { }

            try
            {
                if (Marshal.IsComObject(_eventConnectionPoint))
                {
                    Marshal.ReleaseComObject(_eventConnectionPoint);
                }
            }
            catch { }

            _eventConnectionPoint = null;
            _eventCookie = 0;
        }

        if (_eventSink != null)
        {
            _eventSink.Dispose();
            _eventSink = null;
        }

        if (_ocxInstance is IViewObject viewObj)
        {
            try
            {
                viewObj.SetAdvise(Ole32Interop.DVASPECT_CONTENT, 0, null);
            }
            catch { }
        }
    }

    /// <summary>
    /// Instancia e inicializa el objeto COM ActiveX (tanto con ventana como sin ventana) en modo diseño o ejecución,
    /// ejecutando la secuencia formal del Ciclo de Vida OLE:
    /// Uninitialized -> Created (CoGetClassObject/IClassFactory2/CoCreateInstance) -> Sited (SetClientSite)
    /// -> Loaded (IPersistStreamInit/IPersistPropertyBag) -> InPlaceActive / Visible.
    /// </summary>
    public bool CreateControl()
    {
        if (_disposed || HasInitializationError) return false;
        if (_ocxInstance != null) return true;
        if (_clsidGuid == Guid.Empty)
        {
            HasInitializationError = true;
            LifecycleState = ActiveXLifecycleState.Faulted;
            RecordDiagnostic("ValidateClsid", unchecked((int)0x800401F3), "CLSID vacío o con formato inválido.");
            return false;
        }

        // Si el control ya falló catastróficamente en este CLSID previamente, omitir para proteger la estabilidad del IDE
        if (_faultedClsids.Contains(Clsid))
        {
            HasInitializationError = true;
            LifecycleState = ActiveXLifecycleState.Faulted;
            RecordDiagnostic("BlacklistCheck", unchecked((int)0x80004005), "CLSID previamente marcado como inestable en esta sesión.");
            return false;
        }

        using var guard = new NativeDialogGuard();
        try
        {
            if (!TryCreateInstanceWithLicensing(out _ocxInstance, out int createHr) || _ocxInstance == null)
            {
                HasInitializationError = true;
                LifecycleState = ActiveXLifecycleState.Faulted;
                RecordDiagnostic("CoCreateInstance", createHr, "No se pudo instanciar el objeto COM ActiveX.");
                return false;
            }

            LifecycleState = ActiveXLifecycleState.Created;
            ActiveXResourceMonitor.TrackComObjectCreated();
            RecordDiagnostic("Created", 0, RequiresLicense ? "Instanciado con soporte IClassFactory2." : "Instanciado vía CoCreateInstance.");

            // 1. Consultar OLEMISC y vincular el sitio ANTES de InitNew/Load
            if (_ocxInstance is IOleObject oleObj)
            {
                try
                {
                    if (oleObj.GetMiscStatus(Ole32Interop.DVASPECT_CONTENT, out int miscStatus) >= 0)
                    {
                        IsSimpleFrame = (miscStatus & Ole32Interop.OLEMISC_SIMPLEFRAME) != 0;
                    }
                }
                catch { }

                try
                {
                    int hrSite = oleObj.SetClientSite(_oleClientSite);
                    RecordDiagnostic("SetClientSite", hrSite);
                }
                catch (Exception ex)
                {
                    RecordDiagnostic("SetClientSite", ex.HResult, ex.Message);
                }

                try { oleObj.SetHostNames("ViBa7 IDE", Name); } catch { }
            }

            LifecycleState = ActiveXLifecycleState.Sited;

            // Conectar sumideros de notificación de propiedades y vista (IPropertyNotifySink + IAdviseSink)
            ConnectSinks(_ocxInstance);

            // 2. Inicializar estado de persistencia (IPersistStreamInit / IPersistPropertyBag) ANTES de disparar Resize/SetExtent
            InitializePersistenceState(_ocxInstance);
            LifecycleState = ActiveXLifecycleState.Loaded;

            // 3. Aplicar propiedades escalares e indexadas del .frm vía IDispatch
            if (!HasInitializationError)
            {
                ApplyPendingPropertiesToOcx(_ocxInstance);
            }

            // 4. Establecer dimensiones HIMETRIC tras la correcta inicialización de propiedades
            UpdateOleBoundsAndExtent(_width, _height);
            LifecycleState = ActiveXLifecycleState.Visible;

            // 5. Activación In-Place sobre el HWND de alojamiento actual (Parking HWND o HwndHost interactivo)
            if (_ocxInstance is IOleObject oleActivate && !guard.DialogSuppressed)
            {
                try
                {
                    IntPtr hostHwnd = GetHostWindowHandle();
                    var posRect = new COMRECT(0, 0, _width, _height);
                    int hrVerb = oleActivate.DoVerb(
                        Ole32Interop.OLEIVERB_INPLACEACTIVATE,
                        IntPtr.Zero,
                        _oleClientSite,
                        0,
                        hostHwnd,
                        ref posRect);
                    if (hrVerb >= 0)
                    {
                        LifecycleState = ActiveXLifecycleState.InPlaceActive;
                    }
                    RecordDiagnostic("DoVerb(INPLACEACTIVATE)", hrVerb);

                    if (!IsDesignMode && _customHostHwnd != IntPtr.Zero)
                    {
                        int hrShow = oleActivate.DoVerb(
                            Ole32Interop.OLEIVERB_SHOW,
                            IntPtr.Zero,
                            _oleClientSite,
                            0,
                            hostHwnd,
                            ref posRect);
                        RecordDiagnostic("DoVerb(SHOW)", hrShow);
                    }
                }
                catch (Exception ex)
                {
                    // Los controles ActiveX sin ventana (Windowless) o solo-diseño funcionan vía IViewObject/OleDraw sin HWND
                    IsWindowless = true;
                    RecordDiagnostic("DoVerb(WindowlessFallback)", ex.HResult, ex.Message);
                }
            }

            // Determinar si el control creó HWND o es Windowless
            if (_ocxInstance is IOleInPlaceObject inPlace)
            {
                if (inPlace.GetWindow(out IntPtr ctrlHwnd) == 0 && ctrlHwnd != IntPtr.Zero)
                {
                    IsWindowless = false;
                    if (_customHostHwnd != IntPtr.Zero)
                    {
                        Ole32Interop.MoveWindow(ctrlHwnd, 0, 0, _width, _height, true);
                        Ole32Interop.ShowWindow(ctrlHwnd, Ole32Interop.SW_SHOW);
                    }
                }
                else
                {
                    IsWindowless = true;
                }
            }
            else
            {
                IsWindowless = true;
            }

            if (guard.DialogSuppressed)
            {
                HasInitializationError = true;
                LifecycleState = ActiveXLifecycleState.Faulted;
                RecordDiagnostic("DialogSuppressed", unchecked((int)0x80004005), "El control intentó bloquear el hilo UI con un diálogo modal durante la creación.");
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            _faultedClsids.Add(Clsid);
            HasInitializationError = true;
            LifecycleState = ActiveXLifecycleState.Faulted;
            RecordDiagnostic("CreateControlException", ex.HResult, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Crea e inicializa el control ActiveX directamente en modo ejecución (Ambient.UserMode = True)
    /// emparentado a la ventana Win32 de un HwndHost interactivo.
    /// </summary>
    public bool CreateRuntimeControl(IntPtr parentHwnd, int widthPx, int heightPx)
    {
        if (parentHwnd == IntPtr.Zero) return false;
        _customHostHwnd = parentHwnd;
        IsDesignMode = false;
        _width = Math.Max(16, widthPx);
        _height = Math.Max(16, heightPx);
        if (_ocxInstance != null)
        {
            return AttachToInteractiveParent(parentHwnd, widthPx, heightPx);
        }
        return CreateControl();
    }

    /// <summary>
    /// Activa el modo interactivo (o modo ejecución en tiempo real) emparentando el HWND del OCX
    /// al contenedor proporcionado y activándolo in-place.
    /// </summary>
    public bool AttachToInteractiveParent(IntPtr parentHwnd, int widthPx, int heightPx)
    {
        if (_disposed || HasInitializationError || _ocxInstance == null || parentHwnd == IntPtr.Zero)
            return false;

        _customHostHwnd = parentHwnd;
        SetBounds(widthPx, heightPx);

        using var guard = new NativeDialogGuard();
        try
        {
            IntPtr ctrlHwnd = GetControlWindowHandle();
            if (ctrlHwnd != IntPtr.Zero)
            {
                Ole32Interop.SetParent(ctrlHwnd, parentHwnd);
                Ole32Interop.MoveWindow(ctrlHwnd, 0, 0, _width, _height, true);
                Ole32Interop.ShowWindow(ctrlHwnd, Ole32Interop.SW_SHOW);
            }

            if (_ocxInstance is IOleObject oleObj)
            {
                var posRect = new COMRECT(0, 0, _width, _height);
                int verb = IsDesignMode ? Ole32Interop.OLEIVERB_UIACTIVATE : Ole32Interop.OLEIVERB_INPLACEACTIVATE;
                int hr = oleObj.DoVerb(
                    verb,
                    IntPtr.Zero,
                    _oleClientSite,
                    0,
                    parentHwnd,
                    ref posRect);
                if (!IsDesignMode)
                {
                    oleObj.DoVerb(Ole32Interop.OLEIVERB_SHOW, IntPtr.Zero, _oleClientSite, 0, parentHwnd, ref posRect);
                }
                if (hr >= 0)
                {
                    LifecycleState = IsDesignMode ? ActiveXLifecycleState.UIActive : ActiveXLifecycleState.InPlaceActive;
                }
                RecordDiagnostic("DoVerb(Attach)", hr);
            }

            ctrlHwnd = GetControlWindowHandle();
            if (ctrlHwnd != IntPtr.Zero)
            {
                Ole32Interop.SetParent(ctrlHwnd, parentHwnd);
                Ole32Interop.MoveWindow(ctrlHwnd, 0, 0, _width, _height, true);
                Ole32Interop.ShowWindow(ctrlHwnd, Ole32Interop.SW_SHOW);
                Ole32Interop.SetFocus(ctrlHwnd);
                return true;
            }
            return false;
        }
        catch (Exception ex)
        {
            RecordDiagnostic("AttachToInteractiveParent", ex.HResult, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Desactiva el modo UI-Active ("Editar") devolviendo el HWND del control al parking HWND compartido
    /// y sincronizando su estado de persistencia.
    /// </summary>
    public void DetachToSharedParking(int widthPx, int heightPx)
    {
        if (_disposed || _ocxInstance == null) return;

        using var guard = new NativeDialogGuard();
        try
        {
            if (_ocxInstance is IOleInPlaceObject inPlace)
            {
                try { inPlace.UIDeactivate(); } catch { }
            }

            IntPtr parking = EnsureSharedParkingHwnd();
            _customHostHwnd = IntPtr.Zero;

            IntPtr ctrlHwnd = GetControlWindowHandle();
            if (ctrlHwnd != IntPtr.Zero && parking != IntPtr.Zero)
            {
                Ole32Interop.SetParent(ctrlHwnd, parking);
            }

            SetBounds(widthPx, heightPx);
            if (LifecycleState == ActiveXLifecycleState.UIActive)
            {
                LifecycleState = ActiveXLifecycleState.InPlaceActive;
            }
            SavePersistenceState(clearDirty: false);
        }
        catch { }
    }

    /// <summary>
    /// Transfiere el foco de teclado Win32 a la ventana nativa del control ActiveX.
    /// </summary>
    public bool FocusControl()
    {
        if (_disposed || HasInitializationError || _ocxInstance == null) return false;

        try
        {
            IntPtr ctrlHwnd = GetControlWindowHandle();
            if (ctrlHwnd != IntPtr.Zero)
            {
                Ole32Interop.SetFocus(ctrlHwnd);
                return true;
            }

            if (_ocxInstance is IOleObject oleObj && _customHostHwnd != IntPtr.Zero)
            {
                var posRect = new COMRECT(0, 0, _width, _height);
                oleObj.DoVerb(Ole32Interop.OLEIVERB_UIACTIVATE, IntPtr.Zero, _oleClientSite, 0, _customHostHwnd, ref posRect);
                return true;
            }
        }
        catch { }
        return false;
    }

    /// <summary>
    /// Enruta aceleradores de teclado (flechas, Tab, Enter, Delete) hacia IOleInPlaceActiveObject::TranslateAccelerator.
    /// Retorna true si el control ActiveX consumió la tecla (S_OK == 0).
    /// </summary>
    public bool TranslateAccelerator(ref Native.MSG msg)
    {
        if (_disposed || HasInitializationError || _ocxInstance is not IOleInPlaceActiveObject activeObj)
            return false;

        try
        {
            return activeObj.TranslateAccelerator(ref msg) == 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Enruta teclas mnemónicas (Alt+Letra) hacia IOleControl::OnMnemonic.
    /// </summary>
    public bool OnMnemonic(ref Native.MSG msg)
    {
        if (_disposed || HasInitializationError || _ocxInstance is not IOleControl oleCtrl)
            return false;

        try
        {
            return oleCtrl.OnMnemonic(ref msg) == 0;
        }
        catch
        {
            return false;
        }
    }

    private bool TryCreateInstanceWithLicensing(out object? instance, out int hr)
    {
        instance = null;
        Guid clsid = _clsidGuid;
        Guid iidUnknown = Ole32Interop.IID_IUnknown;
        Guid iidClassFactory2 = Ole32Interop.IID_IClassFactory2;
        uint clsCtx = Ole32Interop.CLSCTX_INPROC_SERVER | Ole32Interop.CLSCTX_LOCAL_SERVER;

        // Intentar primero IClassFactory2 para consultar licencia de diseño y capturar/usar LicenseKey
        try
        {
            int hrCf2 = Ole32Interop.CoGetClassObject(
                ref clsid,
                clsCtx,
                IntPtr.Zero,
                ref iidClassFactory2,
                out object? cfObj);

            if (hrCf2 >= 0 && cfObj is IClassFactory2 cf2)
            {
                try
                {
                    var licInfo = new LICINFO { cbLicInfo = Marshal.SizeOf<LICINFO>() };
                    if (cf2.GetLicInfo(ref licInfo) >= 0)
                    {
                        RequiresLicense = licInfo.fRuntimeKeyAvail || !licInfo.fLicVerified;
                        IsLicenseVerified = licInfo.fLicVerified;

                        if (string.IsNullOrEmpty(LicenseKey) && licInfo.fRuntimeKeyAvail)
                        {
                            if (cf2.RequestLicKey(0, out string? requestedKey) >= 0 && !string.IsNullOrEmpty(requestedKey))
                            {
                                LicenseKey = requestedKey;
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(LicenseKey))
                    {
                        hr = cf2.CreateInstanceLic(IntPtr.Zero, IntPtr.Zero, ref iidUnknown, LicenseKey, out instance);
                        if (hr >= 0 && instance != null)
                        {
                            return true;
                        }
                    }

                    hr = cf2.CreateInstance(IntPtr.Zero, ref iidUnknown, out instance);
                    if (hr >= 0 && instance != null)
                    {
                        return true;
                    }

                    if (hr == Ole32Interop.CLASS_E_NOTLICENSED)
                    {
                        IsLicenseVerified = false;
                        RecordDiagnostic("IClassFactory2", hr, "CLASS_E_NOTLICENSED: El control ActiveX requiere licencia de tiempo de diseño en el equipo.");
                        return false;
                    }
                }
                finally
                {
                    if (Marshal.IsComObject(cfObj))
                    {
                        try { Marshal.ReleaseComObject(cfObj); } catch { }
                    }
                }
            }
        }
        catch
        {
            // Fallback a CoCreateInstance estándar
        }

        hr = Ole32Interop.CoCreateInstance(
            ref clsid,
            IntPtr.Zero,
            clsCtx,
            ref iidUnknown,
            out instance);

        if (hr == Ole32Interop.CLASS_E_NOTLICENSED)
        {
            RequiresLicense = true;
            IsLicenseVerified = false;
            RecordDiagnostic("CoCreateInstance", hr, "CLASS_E_NOTLICENSED: El control ActiveX requiere licencia de diseño.");
        }

        return hr >= 0 && instance != null;
    }

    private void InitializePersistenceState(object ocx)
    {
        if (_persistenceInitialized) return;
        _persistenceInitialized = true;

        // 1. Si tenemos un blob binario de estado (.frx / .ctx, ej. Object.zyx), priorizar IPersistStreamInit::Load / IPersistStream::Load
        bool streamLoaded = false;
        if (InitialStreamBlob != null && InitialStreamBlob.Length > 0)
        {
            if (ocx is IPersistStreamInit persistStreamLoad)
            {
                IntPtr pStream = Ole32Interop.CreateComStreamFromBytes(InitialStreamBlob);
                if (pStream != IntPtr.Zero)
                {
                    try
                    {
                        int hrLoad = persistStreamLoad.Load(pStream);
                        streamLoaded = hrLoad >= 0;
                        RecordDiagnostic("IPersistStreamInit.Load", hrLoad, $"Bytes={InitialStreamBlob.Length}");
                    }
                    catch (Exception ex)
                    {
                        RecordDiagnostic("IPersistStreamInit.Load", ex.HResult, ex.Message);
                    }
                    finally
                    {
                        Marshal.Release(pStream);
                    }
                }
            }
            else if (ocx is IPersistStream persistLegacyStream)
            {
                IntPtr pStream = Ole32Interop.CreateComStreamFromBytes(InitialStreamBlob);
                if (pStream != IntPtr.Zero)
                {
                    try
                    {
                        int hrLoad = persistLegacyStream.Load(pStream);
                        streamLoaded = hrLoad >= 0;
                        RecordDiagnostic("IPersistStream.Load", hrLoad, $"Bytes={InitialStreamBlob.Length}");
                    }
                    catch (Exception ex)
                    {
                        RecordDiagnostic("IPersistStream.Load", ex.HResult, ex.Message);
                    }
                    finally
                    {
                        Marshal.Release(pStream);
                    }
                }
            }
        }

        var typedProps = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        int extentX = Math.Max(423, (int)Math.Round(_width * 2540.0 / 96.0));
        int extentY = Math.Max(423, (int)Math.Round(_height * 2540.0 / 96.0));
        typedProps["_ExtentX"] = extentX;
        typedProps["_ExtentY"] = extentY;

        foreach (var kvp in _pendingProperties)
        {
            if (TryParseIndexedPropertyExpression(kvp.Key, out _, out _))
            {
                // Las propiedades indexadas se aplican después vía IDispatch con argumentos
                continue;
            }

            var converted = ConvertPropertyStringToOcxType(kvp.Key, kvp.Value);
            if (converted != null)
            {
                typedProps[kvp.Key] = converted;
            }
        }

        foreach (var binKvp in _pendingBinaryProperties)
        {
            typedProps[binKvp.Key] = binKvp.Value;
        }

        foreach (var objKvp in _pendingObjectProperties)
        {
            if (objKvp.Value != null)
            {
                typedProps[objKvp.Key] = objKvp.Value;
            }
        }

        if (ocx is IPersistPropertyBag persistBag)
        {
            if (!streamLoaded)
            {
                try
                {
                    int hrInit = persistBag.InitNew();
                    RecordDiagnostic("IPersistPropertyBag.InitNew", hrInit);
                }
                catch
                {
                    // Ignorar si requiere Load directo
                }
            }

            if (typedProps.Count > 0)
            {
                try
                {
                    var bag = new VbPropertyBag(typedProps);
                    int hrLoad = persistBag.Load(bag, bag);
                    RecordDiagnostic("IPersistPropertyBag.Load", hrLoad, $"Props={typedProps.Count}");
                }
                catch (Exception ex)
                {
                    RecordDiagnostic("IPersistPropertyBag.Load", ex.HResult, ex.Message);
                }
            }

            // Descubrir todas las propiedades persistidas reales y sus valores por defecto
            SavePersistenceState(clearDirty: false);
        }
        else if (!streamLoaded && ocx is IPersistStreamInit persistStream)
        {
            try
            {
                int hrInit = persistStream.InitNew();
                RecordDiagnostic("IPersistStreamInit.InitNew", hrInit);
            }
            catch (Exception ex)
            {
                RecordDiagnostic("IPersistStreamInit.InitNew", ex.HResult, ex.Message);
            }
        }
    }

    /// <summary>
    /// Extrae y sincroniza el estado actual de persistencia del control ActiveX mediante
    /// IPersistPropertyBag::Save e IPersistStreamInit::Save / IPersistStream::Save.
    /// </summary>
    public bool SavePersistenceState(bool clearDirty = true)
    {
        var ocx = GetActiveXObject();
        if (ocx == null) return false;

        bool savedAny = false;

        if (ocx is IPersistPropertyBag persistBag)
        {
            try
            {
                var saveBag = new VbPropertyBag();
                int hr = persistBag.Save(saveBag, fClearDirty: clearDirty, fSaveAllProperties: true);
                if (hr >= 0)
                {
                    savedAny = true;
                    foreach (var kvp in saveBag.SavedProperties)
                    {
                        if (kvp.Value is byte[] binBytes)
                        {
                            _savedBinaryProperties[kvp.Key] = binBytes;
                        }
                        else if (kvp.Value != null && !Marshal.IsComObject(kvp.Value))
                        {
                            _discoveredProperties[kvp.Key] = kvp.Value.ToString() ?? string.Empty;
                        }
                    }
                }
            }
            catch
            {
                // Ignorar si Save no está soportado en este control
            }
        }

        if (ocx is IPersistStreamInit persistStreamInit)
        {
            IntPtr pStream = Ole32Interop.CreateComStreamFromBytes(null);
            if (pStream != IntPtr.Zero)
            {
                try
                {
                    int hr = persistStreamInit.Save(pStream, clearDirty);
                    if (hr >= 0)
                    {
                        byte[]? bytes = Ole32Interop.ReadBytesFromComStream(pStream);
                        if (bytes != null && bytes.Length > 0)
                        {
                            SavedStreamBlob = bytes;
                            savedAny = true;
                        }
                    }
                }
                catch
                {
                    // Ignorar si el control falla al guardar en stream vacío
                }
                finally
                {
                    Marshal.Release(pStream);
                }
            }
        }
        else if (ocx is IPersistStream persistStream)
        {
            IntPtr pStream = Ole32Interop.CreateComStreamFromBytes(null);
            if (pStream != IntPtr.Zero)
            {
                try
                {
                    int hr = persistStream.Save(pStream, clearDirty);
                    if (hr >= 0)
                    {
                        byte[]? bytes = Ole32Interop.ReadBytesFromComStream(pStream);
                        if (bytes != null && bytes.Length > 0)
                        {
                            SavedStreamBlob = bytes;
                            savedAny = true;
                        }
                    }
                }
                catch
                {
                    // Ignorar
                }
                finally
                {
                    Marshal.Release(pStream);
                }
            }
        }

        return savedAny;
    }

    /// <summary>
    /// Obtiene la referencia al objeto COM subyacente (IDispatch / IUnknown) del control ActiveX.
    /// </summary>
    public object? GetActiveXObject()
    {
        if (HasInitializationError) return null;
        return _ocxInstance;
    }

    #region Captura Visual Libre de Airspace y Verificación de Pintado

    /// <summary>
    /// Renderiza el control ActiveX (con o sin ventana) a un BitmapSource de WPF utilizando OleDraw (IViewObject::Draw)
    /// y WM_PRINT, verificando que el resultado no sea un recuadro transparente o sin pintar.
    /// </summary>
    public BitmapSource? TryCaptureVisualSnapshot(int widthPx, int heightPx)
    {
        if (_disposed || HasInitializationError || _ocxInstance == null) return null;

        int w = Math.Max(16, Math.Min(4096, widthPx));
        int h = Math.Max(16, Math.Min(4096, heightPx));

        using var guard = new NativeDialogGuard();
        try
        {
            SetBounds(w, h);

            using var bmp = new Bitmap(w, h, PixelFormat.Format32bppPArgb);
            Color baseBg = ResolveBackgroundColor();

            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(baseBg);
                var rect = new COMRECT(0, 0, w, h);
                IntPtr hdc = g.GetHdc();
                try
                {
                    // 1. OleDraw (IViewObject::Draw): método universal de OLE para controles con ventana y sin ventana (Windowless)
                    int hr = Ole32Interop.OleDraw(_ocxInstance, Ole32Interop.DVASPECT_CONTENT, hdc, ref rect);

                    // 2. Si el control tiene HWND propio y OleDraw falló, intentar captura de ventana Win32
                    if (hr < 0 && _ocxInstance is IOleInPlaceObject inPlace &&
                        inPlace.GetWindow(out IntPtr ctrlHwnd) == 0 && ctrlHwnd != IntPtr.Zero)
                    {
                        Ole32Interop.PrintWindow(ctrlHwnd, hdc, 0);
                        Ole32Interop.SendMessageTimeoutW(
                            ctrlHwnd,
                            Ole32Interop.WM_PRINT,
                            hdc,
                            (IntPtr)(Ole32Interop.PRF_CLIENT | Ole32Interop.PRF_NONCLIENT | Ole32Interop.PRF_CHILDREN | Ole32Interop.PRF_ERASEBKGND),
                            Ole32Interop.SMTO_ABORTIFHUNG | Ole32Interop.SMTO_NOTIMEOUTIFNOTHUNG,
                            500,
                            out _);
                    }
                }
                finally
                {
                    g.ReleaseHdc(hdc);
                }

                // Si OleDraw devolvió S_OK pero el control pintó en su propio HWND hijo en lugar de IViewObject
                if (!IsBitmapVisuallyValid(bmp) && _ocxInstance is IOleInPlaceObject inPlaceFallback &&
                    inPlaceFallback.GetWindow(out IntPtr hwndChild) == 0 && hwndChild != IntPtr.Zero)
                {
                    IntPtr hdc2 = g.GetHdc();
                    try
                    {
                        Ole32Interop.SendMessageTimeoutW(
                            hwndChild,
                            Ole32Interop.WM_PRINT,
                            hdc2,
                            (IntPtr)(Ole32Interop.PRF_CLIENT | Ole32Interop.PRF_NONCLIENT | Ole32Interop.PRF_CHILDREN | Ole32Interop.PRF_ERASEBKGND),
                            Ole32Interop.SMTO_ABORTIFHUNG | Ole32Interop.SMTO_NOTIMEOUTIFNOTHUNG,
                            500,
                            out _);
                    }
                    finally
                    {
                        g.ReleaseHdc(hdc2);
                    }
                }
            }

            if (guard.DialogSuppressed)
            {
                HasInitializationError = true;
                return null;
            }

            if (!IsBitmapVisuallyValid(bmp))
            {
                return null;
            }

            IntPtr hBitmap = bmp.GetHbitmap();
            try
            {
                var source = Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
                source.Freeze();
                return source;
            }
            finally
            {
                Ole32Interop.DeleteObject(hBitmap);
            }
        }
        catch
        {
            HasInitializationError = true;
            return null;
        }
    }

    private Color ResolveBackgroundColor()
    {
        if (_pendingProperties.TryGetValue("BackColor", out var bgStr) && !string.IsNullOrWhiteSpace(bgStr))
        {
            if (bgStr.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
            {
                string hex = bgStr[2..].TrimEnd('&');
                if (uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint oleCol))
                {
                    AmbientBackColorOle = unchecked((int)oleCol);
                    return Ole32Interop.TranslateOleColor(oleCol);
                }
            }
            else if (long.TryParse(bgStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out long decCol))
            {
                uint oleCol = unchecked((uint)decCol);
                AmbientBackColorOle = unchecked((int)oleCol);
                return Ole32Interop.TranslateOleColor(oleCol);
            }
        }
        return Ole32Interop.TranslateOleColor(unchecked((uint)AmbientBackColorOle));
    }

    /// <summary>
    /// Verifica si el bitmap capturado contiene un dibujo real (bordes, texto, iconos o gradientes)
    /// y no es simplemente un rectángulo vacío/transparente de un solo color plano.
    /// </summary>
    public static bool IsBitmapVisuallyValid(Bitmap bmp)
    {
        if (bmp.Width < 4 || bmp.Height < 4) return false;

        Color first = bmp.GetPixel(bmp.Width / 2, bmp.Height / 2);
        int distinctPixels = 0;

        int stepX = Math.Max(1, bmp.Width / 16);
        int stepY = Math.Max(1, bmp.Height / 16);

        for (int y = 0; y < bmp.Height; y += stepY)
        {
            for (int x = 0; x < bmp.Width; x += stepX)
            {
                Color c = bmp.GetPixel(x, y);
                if (c.A > 10 && (Math.Abs(c.R - first.R) > 3 || Math.Abs(c.G - first.G) > 3 || Math.Abs(c.B - first.B) > 3))
                {
                    distinctPixels++;
                    if (distinctPixels >= 2)
                    {
                        return true;
                    }
                }
            }
        }

        // Revisar los bordes (1px interior y exterior) por si el control solo dibuja un marco 3D alrededor
        for (int x = 0; x < bmp.Width; x += Math.Max(1, bmp.Width / 20))
        {
            Color top = bmp.GetPixel(x, 0);
            Color bottom = bmp.GetPixel(x, bmp.Height - 1);
            if (Math.Abs(top.R - first.R) > 3 || Math.Abs(bottom.R - first.R) > 3 ||
                Math.Abs(top.G - first.G) > 3 || Math.Abs(bottom.G - first.G) > 3)
            {
                return true;
            }
        }

        return false;
    }

    #endregion

    #region Sincronización de Propiedades vía IDispatch (Escalares, Objetos COM e Indexadas)

    public object? GetPropertyValue(string propertyName)
    {
        var ocx = GetActiveXObject();
        if (ocx == null || string.IsNullOrWhiteSpace(propertyName)) return null;

        if (TryParseIndexedPropertyExpression(propertyName, out string baseName, out object[] indices))
        {
            return GetIndexedPropertyValue(baseName, indices);
        }

        try
        {
            return ocx.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                null,
                ocx,
                null,
                CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    public object? GetIndexedPropertyValue(string propertyName, params object[] indices)
    {
        var ocx = GetActiveXObject();
        if (ocx == null || string.IsNullOrWhiteSpace(propertyName)) return null;

        try
        {
            return ocx.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                null,
                ocx,
                indices,
                CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    public bool SetPropertyValue(string propertyName, object? value)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return false;

        if (TryParseIndexedPropertyExpression(propertyName, out string baseName, out object[] indices))
        {
            if (value != null)
            {
                _pendingProperties[propertyName] = value.ToString() ?? string.Empty;
            }
            return SetIndexedPropertyValue(baseName, value, indices);
        }

        if (value != null && !Marshal.IsComObject(value))
        {
            _pendingProperties[propertyName] = value.ToString() ?? string.Empty;
        }

        var ocx = GetActiveXObject();
        if (ocx == null) return false;

        if (value != null && Marshal.IsComObject(value))
        {
            _pendingObjectProperties[propertyName] = value;
            return TryAssignComObjectProperty(ocx, propertyName, value);
        }

        using var guard = new NativeDialogGuard();
        try
        {
            object? invokeValue = value;
            if (value is string strVal)
            {
                invokeValue = ConvertPropertyStringToOcxType(propertyName, strVal) ?? strVal;
            }

            ocx.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                null,
                ocx,
                new object?[] { invokeValue },
                CultureInfo.InvariantCulture);
            return !guard.DialogSuppressed;
        }
        catch
        {
            return false;
        }
    }

    public bool SetIndexedPropertyValue(string propertyName, object? value, params object[] indices)
    {
        var ocx = GetActiveXObject();
        if (ocx == null || string.IsNullOrWhiteSpace(propertyName) || indices == null || indices.Length == 0)
            return false;

        using var guard = new NativeDialogGuard();
        try
        {
            object?[] args = new object?[indices.Length + 1];
            Array.Copy(indices, args, indices.Length);
            args[indices.Length] = value;

            ocx.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                null,
                ocx,
                args,
                CultureInfo.InvariantCulture);
            return !guard.DialogSuppressed;
        }
        catch
        {
            return false;
        }
    }

    private bool TryAssignComObjectProperty(object ocx, string propertyName, object? comObject)
    {
        using var guard = new NativeDialogGuard();
        try
        {
            // En COM/OLE Automation (VB6 Set obj.Font = ... / Set obj.Picture = ...), las propiedades de objeto
            // utilizan DISPATCH_PROPERTYPUTREF (PutRefDispProperty) o SetProperty estándar.
            try
            {
                ocx.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.PutRefDispProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                    null,
                    ocx,
                    new object?[] { comObject },
                    CultureInfo.InvariantCulture);
                return !guard.DialogSuppressed;
            }
            catch
            {
                ocx.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                    null,
                    ocx,
                    new object?[] { comObject },
                    CultureInfo.InvariantCulture);
                return !guard.DialogSuppressed;
            }
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Analiza expresiones de propiedades indexadas de archivos .frm (ej. "TabCaption(0)", "ColWidth(1)", "Cell(0, 1)").
    /// </summary>
    public static bool TryParseIndexedPropertyExpression(string expression, out string propertyName, out object[] indices)
    {
        propertyName = string.Empty;
        indices = Array.Empty<object>();

        if (string.IsNullOrWhiteSpace(expression)) return false;

        int openParen = expression.IndexOf('(');
        int closeParen = expression.LastIndexOf(')');
        if (openParen <= 0 || closeParen <= openParen || closeParen != expression.Length - 1)
            return false;

        propertyName = expression[..openParen].Trim();
        if (string.IsNullOrEmpty(propertyName)) return false;

        string inside = expression[(openParen + 1)..closeParen].Trim();
        if (string.IsNullOrEmpty(inside)) return false;

        string[] parts = inside.Split(',');
        var parsedIndices = new List<object>(parts.Length);
        foreach (string part in parts)
        {
            string trimmed = part.Trim();
            if (short.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out short sVal))
            {
                parsedIndices.Add((int)sVal);
            }
            else if (int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out int iVal))
            {
                parsedIndices.Add(iVal);
            }
            else
            {
                parsedIndices.Add(trimmed.Trim('\"'));
            }
        }

        indices = parsedIndices.ToArray();
        return indices.Length > 0;
    }

    public void ApplyProperties(IDictionary<string, string> properties)
    {
        if (properties == null) return;

        foreach (var kvp in properties)
        {
            _pendingProperties[kvp.Key] = kvp.Value;
        }

        if (_ocxInstance != null && !HasInitializationError)
        {
            ApplyPendingPropertiesToOcx(_ocxInstance);
        }
    }

    private void ApplyPendingPropertiesToOcx(object ocx)
    {
        using var guard = new NativeDialogGuard();

        // 1. Asignar objetos COM pendientes (StdFont / StdPicture)
        foreach (var objKvp in _pendingObjectProperties)
        {
            if (objKvp.Value != null)
            {
                TryAssignComObjectProperty(ocx, objKvp.Key, objKvp.Value);
                if (guard.DialogSuppressed)
                {
                    HasInitializationError = true;
                    return;
                }
            }
        }

        if (_pendingProperties.Count == 0) return;

        // 2. Asignar propiedades escalares primero y luego las indexadas (ej. Tabs antes que TabCaption(0))
        var indexedQueue = new List<(string BaseName, object[] Indices, string RawVal)>();

        foreach (var kvp in _pendingProperties)
        {
            string propName = kvp.Key;
            string rawVal = kvp.Value;

            if (propName.StartsWith("_", StringComparison.Ordinal) ||
                propName.Equals("Name", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("Index", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("Left", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("Top", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("Width", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("Height", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("TabIndex", StringComparison.OrdinalIgnoreCase) ||
                propName.Equals("TabStop", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (TryParseIndexedPropertyExpression(propName, out string baseName, out object[] indices))
            {
                indexedQueue.Add((baseName, indices, rawVal));
                continue;
            }

            try
            {
                object? converted = ConvertPropertyStringToOcxType(propName, rawVal);
                if (converted != null)
                {
                    ocx.GetType().InvokeMember(
                        propName,
                        BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                        null,
                        ocx,
                        new object?[] { converted },
                        CultureInfo.InvariantCulture);
                }

                if (guard.DialogSuppressed)
                {
                    HasInitializationError = true;
                    return;
                }
            }
            catch
            {
                // Ignorar propiedades de solo lectura o con tipos especiales
            }
        }

        // 3. Aplicar propiedades indexadas (ej. TabCaption(0), TabPicture(0), ColWidth(1))
        foreach (var item in indexedQueue)
        {
            object? converted = ConvertPropertyStringToOcxType(item.BaseName, item.RawVal);
            if (converted != null)
            {
                SetIndexedPropertyValue(item.BaseName, converted, item.Indices);
                if (guard.DialogSuppressed)
                {
                    HasInitializationError = true;
                    return;
                }
            }
        }
    }

    public static object? ConvertPropertyStringToOcxType(string propertyName, string strValue)
    {
        if (string.IsNullOrWhiteSpace(strValue)) return strValue;

        if (strValue.Contains(".frx\":", StringComparison.OrdinalIgnoreCase) ||
            strValue.Contains(".ctx\":", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (strValue.Equals("True", StringComparison.OrdinalIgnoreCase))
            return true;
        if (strValue.Equals("False", StringComparison.OrdinalIgnoreCase))
            return false;

        // Solo interpretar -1 / 0 como booleano si la propiedad es conocida por ser booleana.
        // De lo contrario, propiedades como BorderStyle = 0, Style = 0, Appearance = 0, Alignment = 0 se corrompen a false.
        if (IsKnownBooleanProperty(propertyName))
        {
            if (strValue == "-1") return true;
            if (strValue == "0") return false;
        }

        // Propiedades de códigos de caracteres de iconos (ej. IconCharCode, CharCode) deben permanecer como texto
        if (propertyName.Contains("CharCode", StringComparison.OrdinalIgnoreCase) ||
            propertyName.Equals("CharCode", StringComparison.OrdinalIgnoreCase))
        {
            return strValue.Trim('\"');
        }

        if (strValue.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            string hexClean = strValue[2..].TrimEnd('&');
            if (long.TryParse(hexClean, NumberStyles.HexNumber, null, out long hexVal))
            {
                return (uint)hexVal;
            }
        }

        if (int.TryParse(strValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intVal))
            return intVal;

        if (double.TryParse(strValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double dblVal))
            return dblVal;

        return strValue.Trim('\"');
    }

    private static bool IsKnownBooleanProperty(string propertyName)
    {
        return propertyName.Equals("Enabled", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("Locked", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("TabStop", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("AutoSize", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("WordWrap", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("MultiLine", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("Default", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("Cancel", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("CausesValidation", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("AutoRedraw", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("ClipControls", StringComparison.OrdinalIgnoreCase) ||
               propertyName.Equals("RightToLeft", StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Hojas de Propiedades OLE Nativas (OleCreatePropertyFrame)

    public bool HasPropertyPages()
    {
        var ocx = GetActiveXObject();
        if (ocx is not ISpecifyPropertyPages specifyPages) return false;

        try
        {
            int hr = specifyPages.GetPages(out var pages);
            bool has = hr == 0 && pages.cElems > 0;
            if (pages.pElems != IntPtr.Zero)
            {
                Ole32Interop.CoTaskMemFree(pages.pElems);
            }
            return has;
        }
        catch
        {
            return false;
        }
    }

    public bool ShowPropertyPages(IntPtr ownerHwnd, string? caption = null)
    {
        var ocx = GetActiveXObject();
        if (ocx is not ISpecifyPropertyPages specifyPages) return false;

        try
        {
            int hr = specifyPages.GetPages(out var pages);
            if (hr != 0 || pages.cElems == 0 || pages.pElems == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                string dlgCaption = caption ?? $"Propiedades de {Name}";
                object[] objs = { ocx };

                hr = Ole32Interop.OleCreatePropertyFrame(
                    ownerHwnd,
                    0,
                    0,
                    dlgCaption,
                    1,
                    objs,
                    pages.cElems,
                    pages.pElems,
                    0,
                    0,
                    IntPtr.Zero);

                if (hr == 0)
                {
                    SavePersistenceState(clearDirty: true);
                }

                return hr == 0;
            }
            finally
            {
                Ole32Interop.CoTaskMemFree(pages.pElems);
            }
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Enumera los verbos OLE expuestos por el control ActiveX mediante IOleObject::EnumVerbs.
    /// </summary>
    public IReadOnlyList<OleVerbEntry> GetOleVerbs()
    {
        var list = new List<OleVerbEntry>();
        var ocx = GetActiveXObject();
        if (ocx is not IOleObject oleObj) return list;

        try
        {
            if (oleObj.EnumVerbs(out IntPtr ppEnum) == 0 && ppEnum != IntPtr.Zero)
            {
                try
                {
                    if (Marshal.GetObjectForIUnknown(ppEnum) is IEnumOLEVERB enumVerb)
                    {
                        var buffer = new OLEVERB[1];
                        while (enumVerb.Next(1, buffer, out uint fetched) == 0 && fetched == 1)
                        {
                            string name = (buffer[0].lpszVerbName ?? string.Empty).Replace("&", string.Empty).Trim();
                            if (!string.IsNullOrEmpty(name))
                            {
                                list.Add(new OleVerbEntry(buffer[0].lVerb, name, buffer[0].fuFlags, buffer[0].grfAttribs));
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.Release(ppEnum);
                }
            }
        }
        catch
        {
            // Ignorar si EnumVerbs no está implementado
        }

        return list;
    }

    /// <summary>
    /// Ejecuta un verbo OLE específico sobre el control ActiveX mediante IOleObject::DoVerb.
    /// </summary>
    public bool ExecuteOleVerb(int verbId, IntPtr parentHwnd = default)
    {
        var ocx = GetActiveXObject();
        if (ocx is not IOleObject oleObj) return false;

        if (verbId == Ole32Interop.OLEIVERB_PROPERTIES)
        {
            return ShowPropertyPages(parentHwnd);
        }

        IntPtr targetHwnd = parentHwnd != IntPtr.Zero ? parentHwnd : GetHostWindowHandle();
        var posRect = new COMRECT(0, 0, _width, _height);

        using var guard = new NativeDialogGuard();
        try
        {
            int hr = oleObj.DoVerb(verbId, IntPtr.Zero, _oleClientSite, 0, targetHwnd, ref posRect);
            if (hr >= 0 && verbId == Ole32Interop.OLEIVERB_UIACTIVATE)
            {
                LifecycleState = ActiveXLifecycleState.UIActive;
            }
            SavePersistenceState(clearDirty: false);
            return hr >= 0 && !guard.DialogSuppressed;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Invoca un método o propiedad parametrizada sobre la instancia COM del control ActiveX vía IDispatch.
    /// </summary>
    public object? InvokeMethod(string methodName, params object?[]? args)
    {
        var ocx = GetActiveXObject();
        if (ocx == null || string.IsNullOrWhiteSpace(methodName)) return null;
        return InvokeComDispatchMember(ocx, methodName, args);
    }

    /// <summary>
    /// Invoca un método o propiedad (incluyendo el miembro por defecto DISPID_VALUE si memberName es vacío)
    /// sobre cualquier objeto COM IDispatch (tanto el OCX principal como sub-objetos de colección Nodes, ListItems, etc.).
    /// </summary>
    public static object? InvokeComDispatchMember(object comTarget, string memberName, params object?[]? args)
    {
        if (comTarget == null) return null;

        using var guard = new NativeDialogGuard();
        try
        {
            object?[]? invokeArgs = NormalizeComInvokeArguments(args);
            string targetName = string.IsNullOrWhiteSpace(memberName) ? "[DISPID=0]" : memberName;

            return comTarget.GetType().InvokeMember(
                targetName,
                BindingFlags.InvokeMethod | BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase | BindingFlags.OptionalParamBinding,
                null,
                comTarget,
                invokeArgs,
                CultureInfo.InvariantCulture);
        }
        catch
        {
            // Fallback: si se invocó sin nombre o falló con "[DISPID=0]", intentar con "Item"
            if (string.IsNullOrWhiteSpace(memberName) && args is { Length: > 0 })
            {
                try
                {
                    return comTarget.GetType().InvokeMember(
                        "Item",
                        BindingFlags.InvokeMethod | BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase | BindingFlags.OptionalParamBinding,
                        null,
                        comTarget,
                        NormalizeComInvokeArguments(args),
                        CultureInfo.InvariantCulture);
                }
                catch { }
            }
            return null;
        }
    }

    /// <summary>
    /// Asigna una propiedad simple o indexada sobre cualquier objeto COM IDispatch (incluyendo sub-objetos).
    /// </summary>
    public static bool SetComDispatchProperty(object comTarget, string propertyName, object? value, params object?[]? indices)
    {
        if (comTarget == null || string.IsNullOrWhiteSpace(propertyName)) return false;

        using var guard = new NativeDialogGuard();
        try
        {
            object?[] invokeArgs;
            if (indices is { Length: > 0 })
            {
                invokeArgs = new object?[indices.Length + 1];
                for (int i = 0; i < indices.Length; i++)
                {
                    invokeArgs[i] = indices[i] ?? Type.Missing;
                }
                invokeArgs[indices.Length] = value;
            }
            else
            {
                invokeArgs = new object?[] { value };
            }

            if (value != null && Marshal.IsComObject(value))
            {
                try
                {
                    comTarget.GetType().InvokeMember(
                        propertyName,
                        BindingFlags.PutRefDispProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                        null,
                        comTarget,
                        invokeArgs,
                        CultureInfo.InvariantCulture);
                    return !guard.DialogSuppressed;
                }
                catch
                {
                    // Fallback a SetProperty estándar
                }
            }

            comTarget.GetType().InvokeMember(
                propertyName,
                BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase,
                null,
                comTarget,
                invokeArgs,
                CultureInfo.InvariantCulture);
            return !guard.DialogSuppressed;
        }
        catch
        {
            return false;
        }
    }

    private static object?[]? NormalizeComInvokeArguments(object?[]? args)
    {
        if (args == null || args.Length == 0) return null;
        var normalized = new object?[args.Length];
        for (int i = 0; i < args.Length; i++)
        {
            normalized[i] = args[i] ?? Type.Missing;
        }
        return normalized;
    }

    #endregion

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        try
        {
            DisconnectSinks();

            if (_ocxInstance != null)
            {
                // Teardown estricto OLE: UIDeactivate -> InPlaceDeactivate -> Close(OLECLOSE_NOSAVE) -> SetClientSite(null) -> ReleaseComObject
                if (_ocxInstance is IOleInPlaceObject inPlace)
                {
                    try { inPlace.UIDeactivate(); } catch { }
                    try { inPlace.InPlaceDeactivate(); } catch { }
                }
                if (_ocxInstance is IOleObject oleObj)
                {
                    try { oleObj.Close(Ole32Interop.OLECLOSE_NOSAVE); } catch { }
                    try { oleObj.SetClientSite(null); } catch { }
                }
                if (Marshal.IsComObject(_ocxInstance))
                {
                    try { Marshal.ReleaseComObject(_ocxInstance); } catch { }
                    ActiveXResourceMonitor.TrackComObjectReleased();
                }
                _ocxInstance = null;
            }

            if (_cachedAmbientStdFont != null && Marshal.IsComObject(_cachedAmbientStdFont))
            {
                try { Marshal.ReleaseComObject(_cachedAmbientStdFont); } catch { }
                _cachedAmbientStdFont = null;
            }

            foreach (var obj in _pendingObjectProperties.Values)
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    try { Marshal.ReleaseComObject(obj); } catch { }
                }
            }
            _pendingObjectProperties.Clear();
        }
        catch
        {
            // Ignorar errores al liberar COM
        }

        LifecycleState = ActiveXLifecycleState.Disposed;
        Disposed?.Invoke(this, EventArgs.Empty);
    }
}
