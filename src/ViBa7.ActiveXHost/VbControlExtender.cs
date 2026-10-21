using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace ViBa7.ActiveXHost;

/// <summary>
/// Implementación COM del objeto extendido de contenedor de Visual Basic 6.0 (VBControlExtender),
/// devuelto por IOleControlSite::GetExtendedControl. Expone las propiedades y métodos de posicionamiento,
/// foco y contenedor que el control ActiveX o el código VB6 invocan sobre el sitio del control.
/// </summary>
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class VbControlExtender : IDisposable
{
    private DynamicActiveXControl _control;
    private float _leftTwips;
    private float _topTwips;
    private bool _suppressBoundsNotifications;

    public VbControlExtender(DynamicActiveXControl control)
    {
        _control = control ?? throw new ArgumentNullException(nameof(control));
    }

    /// <summary>
    /// Convierte unidades Twips (1440 por pulgada) a HIMETRIC (2540 por pulgada).
    /// </summary>
    public static int TwipsToHimetric(float twips) => (int)Math.Round(twips * 2540.0 / 1440.0);

    /// <summary>
    /// Convierte unidades HIMETRIC (2540 por pulgada) a Twips (1440 por pulgada).
    /// </summary>
    public static float HimetricToTwips(int himetric) => (float)(himetric * 1440.0 / 2540.0);

    public void Dispose()
    {
        // Limpieza de referencias
    }

    /// <summary>
    /// Configura las dimensiones y coordenadas del Extender directamente sin disparar notificaciones de cambio.
    /// Esto previene bucles de retroalimentación destructivos y sobreescrituras de coordenadas (Top = 0) durante la sincronización.
    /// </summary>
    public void SetBoundsDirect(float left, float top, float width, float height)
    {
        _suppressBoundsNotifications = true;
        try
        {
            _leftTwips = left;
            _topTwips = top;
            int pxW = Math.Max(16, (int)Math.Round(width / 15.0f));
            int pxH = Math.Max(16, (int)Math.Round(height / 15.0f));
            _control.Width = pxW;
            _control.Height = pxH;
        }
        finally
        {
            _suppressBoundsNotifications = false;
        }
    }

    [DispId(unchecked((int)0x80010000))]
    public string Name
    {
        get => _control.Name;
        set => _control.Name = value ?? string.Empty;
    }

    public short Index { get; set; } = -1;

    public float Left
    {
        get => _leftTwips;
        set
        {
            _leftTwips = value;
            if (!_suppressBoundsNotifications)
            {
                _control.NotifyExtenderBoundsChanged();
            }
        }
    }

    public float Top
    {
        get => _topTwips;
        set
        {
            _topTwips = value;
            if (!_suppressBoundsNotifications)
            {
                _control.NotifyExtenderBoundsChanged();
            }
        }
    }

    public float Width
    {
        get => _control.Width * 15.0f;
        set
        {
            int px = Math.Max(16, (int)Math.Round(value / 15.0f));
            _control.Width = px;
            if (!_suppressBoundsNotifications)
            {
                _control.NotifyExtenderBoundsChanged();
            }
        }
    }

    public float Height
    {
        get => _control.Height * 15.0f;
        set
        {
            int px = Math.Max(16, (int)Math.Round(value / 15.0f));
            _control.Height = px;
            if (!_suppressBoundsNotifications)
            {
                _control.NotifyExtenderBoundsChanged();
            }
        }
    }

    public bool Visible { get; set; } = true;

    public bool Enabled
    {
        get => _control.AmbientEnabled;
        set => _control.AmbientEnabled = value;
    }

    public short TabIndex { get; set; }

    public bool TabStop { get; set; } = true;

    public string Tag { get; set; } = string.Empty;

    public string ToolTipText { get; set; } = string.Empty;

    public bool CausesValidation { get; set; } = true;

    public int HelpContextID { get; set; }

    public int WhatsThisHelpID { get; set; }

    public short DragMode { get; set; }

    public object? Object => _control.GetActiveXObject();

    public object? Parent { get; set; }

    public object? Container { get; set; }

    public void Move(float left, object? top = null, object? width = null, object? height = null)
    {
        _leftTwips = left;
        if (top != null && top != Type.Missing && TryConvertToFloat(top, out float t))
        {
            _topTwips = t;
        }
        if (width != null && width != Type.Missing && TryConvertToFloat(width, out float w))
        {
            _control.Width = Math.Max(16, (int)Math.Round(w / 15.0f));
        }
        if (height != null && height != Type.Missing && TryConvertToFloat(height, out float h))
        {
            _control.Height = Math.Max(16, (int)Math.Round(h / 15.0f));
        }
        _control.NotifyExtenderBoundsChanged();
    }

    public void SetFocus()
    {
        _control.RequestFocusFromExtender();
    }

    public void ZOrder(object? position = null)
    {
        int pos = 0;
        if (position != null && position != Type.Missing)
        {
            try { pos = Convert.ToInt32(position, CultureInfo.InvariantCulture); } catch { }
        }
        _control.RequestZOrderFromExtender(pos);
    }

    public void Refresh()
    {
        _control.RequestRefreshFromExtender();
    }

    private static bool TryConvertToFloat(object value, out float result)
    {
        try
        {
            result = Convert.ToSingle(value, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            result = 0f;
            return false;
        }
    }
}
