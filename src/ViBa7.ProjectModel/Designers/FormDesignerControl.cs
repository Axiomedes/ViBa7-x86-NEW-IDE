using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ViBa7.ProjectModel.Designers;

public enum PictureAlignment
{
    Top = 0,     // 0 - Superior (arriba del texto)
    Bottom = 1,  // 1 - Inferior (abajo del texto)
    Left = 2,    // 2 - Izquierda (a la izquierda del texto)
    Right = 3,   // 3 - Derecha (a la derecha del texto)
    Center = 4   // 4 - Centro (centrado)
}

/// <summary>
/// Representa un control individual o contenedor dentro del formulario VB6 en modo diseño.
/// </summary>
public class FormDesignerControl
{
    public const double TwipsPerPixel = 15.0;

    public string TypeName { get; set; } = "VB.CommandButton";
    public string Name { get; set; } = "Command1";
    public int? Index { get; set; }

    // Geometría en Twips (unidades nativas de VB6)
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; } = 1500;
    public int Height { get; set; } = 495;
    public bool HasExplicitWidth { get; set; }
    public bool HasExplicitHeight { get; set; }
    public int? ExtentX { get; set; }
    public int? ExtentY { get; set; }

    /// <summary>
    /// Convierte dimensiones de HiMetric (1/100 de mm, usado en _ExtentX/_ExtentY de OLE) a Twips (1/20 de punto).
    /// Factor exacto: 1 pulgada = 2540 HiMetric = 1440 Twips => Twips = HiMetric * 1440 / 2540 = HiMetric * 72 / 127.
    /// </summary>
    public static int HiMetricToTwips(double hiMetric) => (int)Math.Round(hiMetric * (1440.0 / 2540.0));

    /// <summary>
    /// Convierte dimensiones de Twips a HiMetric (usado en _ExtentX/_ExtentY de OLE).
    /// Factor exacto: HiMetric = Twips * 2540 / 1440 = Twips * 127 / 72.
    /// </summary>
    public static int TwipsToHiMetric(double twips) => (int)Math.Round(twips * (2540.0 / 1440.0));

    /// <summary>
    /// Sincroniza los valores ExtentX y ExtentY en HiMetric a partir del ancho y alto actuales en Twips.
    /// </summary>
    public void UpdateExtentsFromDimensions()
    {
        ExtentX = TwipsToHiMetric(Width);
        ExtentY = TwipsToHiMetric(Height);
    }

    // Propiedades comunes
    public string Caption { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public int? TabIndex { get; set; }
    public bool? TabStop { get; set; }
    public bool Enabled { get; set; } = true;
    public bool Visible { get; set; } = true;

    /// <summary>
    /// Intervalo en milisegundos entre eventos Timer para el control VB.Timer (0 a 65535).
    /// </summary>
    public int Interval { get; set; } = 0;

    // Propiedades específicas de menú (VB.Menu)
    public bool Checked { get; set; }
    public string Shortcut { get; set; } = string.Empty;
    public bool WindowList { get; set; }
    public int NegotiatePosition { get; set; }
    public int HelpContextID { get; set; }

    // Apariencia
    public string? BackColorHex { get; set; }
    public string? ForeColorHex { get; set; }
    public int? BorderStyle { get; set; }
    public int? Alignment { get; set; }
    public bool? MultiLine { get; set; }
    public int? ScrollBars { get; set; }
    public FormFontDescription? Font { get; set; }

    /// <summary>
    /// Diccionario de fuentes con nombre para el control (ej: Font, IconFont, CaptionFont).
    /// </summary>
    public Dictionary<string, FormFontDescription> FontProperties { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Obtiene una fuente tipográfica por nombre de propiedad (ej: "Font", "IconFont").
    /// </summary>
    public FormFontDescription? GetFontProperty(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        if (FontProperties.TryGetValue(name, out var font)) return font;
        if (name.Equals("Font", StringComparison.OrdinalIgnoreCase)) return Font;
        return null;
    }

    /// <summary>
    /// Establece o actualiza una fuente tipográfica por nombre de propiedad.
    /// </summary>
    public void SetFontProperty(string name, FormFontDescription? font)
    {
        if (string.IsNullOrWhiteSpace(name)) return;
        if (font == null)
        {
            FontProperties.Remove(name);
            if (name.Equals("Font", StringComparison.OrdinalIgnoreCase))
            {
                Font = null;
            }
        }
        else
        {
            FontProperties[name] = font;
            if (name.Equals("Font", StringComparison.OrdinalIgnoreCase))
            {
                Font = font;
            }
        }
    }

    // Soporte de imágenes y alineación (PictureBox, CommandButton, etc.)
    public string? Picture { get; set; }
    public byte[]? PictureData { get; set; }
    public PictureAlignment PictureAlignment { get; set; } = PictureAlignment.Top;
    public int Style { get; set; } = 0; // 0 = Standard, 1 = Graphical
    public bool AutoSize { get; set; }
    public string? DownPicture { get; set; }
    public byte[]? DownPictureData { get; set; }
    public string? DisabledPicture { get; set; }
    public byte[]? DisabledPictureData { get; set; }

    public bool HasPicture => !string.IsNullOrEmpty(Picture) || (PictureData != null && PictureData.Length > 0);

    // Jerarquía de contenedores (para Frame, PictureBox, UserControls con ControlContainer, etc.)
    public bool? ControlContainer { get; set; }
    public FormDesignerControl? ParentControl { get; set; }
    public List<FormDesignerControl> Children { get; set; } = new();

    // Colección de todas las propiedades originales y no mapeadas (para round-trip sin pérdidas)
    public List<FormPropertyItem> CustomProperties { get; set; } = new();

    /// <summary>
    /// Propiedades binarias del control cargadas desde .frx / .ctx (indexadas por nombre de propiedad).
    /// </summary>
    public Dictionary<string, byte[]> BinaryProperties { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Estado binario opaco de persistencia OLE (IPersistStreamInit / IPersistStream / IPersistStorage) del control ActiveX.
    /// </summary>
    public byte[]? ActiveXStateBlob { get; set; }

    /// <summary>
    /// Nombre de la propiedad en el archivo .frm que referencia al ActiveXStateBlob en el .frx (por defecto "Object.zyx").
    /// </summary>
    public string? ActiveXStatePropertyName { get; set; }

    /// <summary>
    /// Indica si el estado binario ActiveX fue modificado en diseño (ej. mediante páginas de propiedades nativas) y debe reescribirse en el .frx.
    /// </summary>
    public bool HasModifiedActiveXState { get; set; }

    /// <summary>
    /// Clave de licencia COM en tiempo de ejecución obtenida vía IClassFactory2::RequestLicKey.
    /// </summary>
    public string? LicenseKey { get; set; }

    /// <summary>
    /// Indica si el control es un componente ActiveX / OCX no registrado o ausente en el sistema local,
    /// cuyo estado y propiedades se conservan intactos para permitir round-trip sin pérdida de datos.
    /// </summary>
    public bool IsUnresolved { get; set; }

    /// <summary>
    /// Mensaje de diagnóstico si el control no pudo resolverse en el sistema.
    /// </summary>
    public string? UnresolvedErrorMessage { get; set; }

    /// <summary>
    /// Líneas textuales originales de la declaración del control en el archivo .frm (preservación verbatim).
    /// </summary>
    public List<string> RawControlLines { get; set; } = new();

    private bool? _hasPropertyPages;

    /// <summary>
    /// Indica si el control soporta o expone una página de propiedades avanzada del fabricante ((Custom)).
    /// </summary>
    public bool HasPropertyPages
    {
        get => _hasPropertyPages ?? (!IsIntrinsic || CustomProperties.Any(p => p.Name.Equals("(Custom)", StringComparison.OrdinalIgnoreCase) || p.Name.Equals("Custom", StringComparison.OrdinalIgnoreCase)));
        set => _hasPropertyPages = value;
    }

    #region Conversión Twips ↔ Píxeles
    public double LeftPixels
    {
        get => Left / TwipsPerPixel;
        set => Left = (int)Math.Round(value * TwipsPerPixel);
    }

    public double TopPixels
    {
        get => Top / TwipsPerPixel;
        set => Top = (int)Math.Round(value * TwipsPerPixel);
    }

    public double WidthPixels
    {
        get => Math.Max(1, Width / TwipsPerPixel);
        set => Width = Math.Max(15, (int)Math.Round(value * TwipsPerPixel));
    }

    public double HeightPixels
    {
        get => Math.Max(1, Height / TwipsPerPixel);
        set => Height = Math.Max(15, (int)Math.Round(value * TwipsPerPixel));
    }

    public void SetBoundsPixels(double left, double top, double width, double height)
    {
        LeftPixels = left;
        TopPixels = top;
        WidthPixels = width;
        HeightPixels = height;
    }

    public void SetPositionPixels(double left, double top)
    {
        LeftPixels = left;
        TopPixels = top;
    }

    public double AbsoluteLeftPixels
    {
        get
        {
            double l = LeftPixels;
            var p = ParentControl;
            while (p != null)
            {
                l += p.LeftPixels;
                p = p.ParentControl;
            }
            return l;
        }
    }

    public double AbsoluteTopPixels
    {
        get
        {
            double t = TopPixels;
            var p = ParentControl;
            while (p != null)
            {
                t += p.TopPixels;
                p = p.ParentControl;
            }
            return t;
        }
    }

    public void SetSizePixels(double width, double height)
    {
        WidthPixels = width;
        HeightPixels = height;
    }
    #endregion

    #region Propiedades de Diagnóstico y UI
    public string DisplayType => TypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase)
        ? TypeName.Substring(3)
        : TypeName;

    public string DisplayNameWithIndex => Index.HasValue ? $"{Name}({Index.Value})" : Name;

    /// <summary>
    /// Delegado global para resolver si un UserControl del proyecto actúa como contenedor en tiempo de diseño.
    /// </summary>
    public static Func<string, bool>? IsUserControlContainerResolver { get; set; }

    public bool IsContainer =>
        ControlContainer == true ||
        TypeName.Equals("VB.Frame", StringComparison.OrdinalIgnoreCase) ||
        TypeName.Equals("Frame", StringComparison.OrdinalIgnoreCase) ||
        TypeName.Equals("VB.PictureBox", StringComparison.OrdinalIgnoreCase) ||
        TypeName.Equals("PictureBox", StringComparison.OrdinalIgnoreCase) ||
        (IsUserControlContainerResolver?.Invoke(TypeName) ?? false);

    public bool IsMenu =>
        TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase);

    public bool IsTimer =>
        TypeName.Equals("VB.Timer", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Indica si el control es un control intrínseco estándar de VB6.
    /// </summary>
    public bool IsIntrinsic =>
        TypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase) ||
        TypeName is "PictureBox" or "Label" or "TextBox" or "Frame" or "CommandButton" or
                    "CheckBox" or "OptionButton" or "ComboBox" or "ListBox" or "HScrollBar" or
                    "VScrollBar" or "Timer" or "DriveListBox" or "DirListBox" or "FileListBox" or
                    "Shape" or "Line" or "Image" or "Data" or "OLE" or "Menu";

    /// <summary>
    /// Obtiene el valor de una propiedad personalizada o no estándar de la cabecera del control.
    /// </summary>
    public string? GetCustomProperty(string propName)
    {
        return CustomProperties.FirstOrDefault(p => p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase))?.Value;
    }

    /// <summary>
    /// Obtiene el valor textual de una propiedad personalizada removiendo comillas delimitadoras si existen.
    /// </summary>
    public string? GetCustomPropertyString(string propName)
    {
        string? val = GetCustomProperty(propName);
        if (val == null) return null;
        string trimmed = val.Trim();
        if (trimmed.Length >= 2 && trimmed.StartsWith('\"') && trimmed.EndsWith('\"'))
        {
            return trimmed[1..^1];
        }
        return trimmed;
    }

    /// <summary>
    /// Establece o actualiza el valor de una propiedad personalizada de la cabecera del control.
    /// </summary>
    public void SetCustomProperty(string propName, string value)
    {
        var existing = CustomProperties.FirstOrDefault(p => p.Name.Equals(propName, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            existing.Value = value;
        }
        else
        {
            CustomProperties.Add(new FormPropertyItem(propName, value));
        }
    }

    /// <summary>
    /// Obtiene el valor numérico entero de una propiedad personalizada si existe y es válida.
    /// </summary>
    public int? GetCustomPropertyInt(string propName)
    {
        string? val = GetCustomProperty(propName);
        if (val != null && int.TryParse(val.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int intVal))
        {
            return intVal;
        }
        return null;
    }

    /// <summary>
    /// Obtiene el valor numérico largo de una propiedad personalizada si existe y es válida.
    /// </summary>
    public long? GetCustomPropertyLong(string propName)
    {
        string? val = GetCustomProperty(propName);
        if (val != null && long.TryParse(val.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long lVal))
        {
            return lVal;
        }
        return null;
    }

    /// <summary>
    /// Indica si el control tiene un tamaño fijo en diseño y no puede redimensionarse (ej: Timer).
    /// </summary>
    public bool IsNonResizable => IsTimer;
    #endregion

    /// <summary>
    /// Clona el control y su jerarquía de hijos.
    /// </summary>
    public FormDesignerControl Clone(FormDesignerControl? newParent = null)
    {
        var clone = new FormDesignerControl
        {
            TypeName = TypeName,
            Name = Name,
            Index = Index,
            Left = Left,
            Top = Top,
            Width = Width,
            Height = Height,
            HasExplicitWidth = HasExplicitWidth,
            HasExplicitHeight = HasExplicitHeight,
            ExtentX = ExtentX,
            ExtentY = ExtentY,
            Caption = Caption,
            Text = Text,
            TabIndex = TabIndex,
            TabStop = TabStop,
            Enabled = Enabled,
            Visible = Visible,
            Checked = Checked,
            Shortcut = Shortcut,
            WindowList = WindowList,
            NegotiatePosition = NegotiatePosition,
            HelpContextID = HelpContextID,
            BackColorHex = BackColorHex,
            ForeColorHex = ForeColorHex,
            BorderStyle = BorderStyle,
            Alignment = Alignment,
            MultiLine = MultiLine,
            ScrollBars = ScrollBars,
            Font = Font?.Clone(),
            Interval = Interval,
            Picture = Picture,
            PictureData = PictureData != null ? (byte[])PictureData.Clone() : null,
            PictureAlignment = PictureAlignment,
            Style = Style,
            AutoSize = AutoSize,
            DownPicture = DownPicture,
            DownPictureData = DownPictureData != null ? (byte[])DownPictureData.Clone() : null,
            DisabledPicture = DisabledPicture,
            DisabledPictureData = DisabledPictureData != null ? (byte[])DisabledPictureData.Clone() : null,
            ParentControl = newParent,
            CustomProperties = CustomProperties.Select(p => p.Clone()).ToList(),
            BinaryProperties = BinaryProperties.ToDictionary(kv => kv.Key, kv => (byte[])kv.Value.Clone(), StringComparer.OrdinalIgnoreCase),
            ActiveXStateBlob = ActiveXStateBlob != null ? (byte[])ActiveXStateBlob.Clone() : null,
            ActiveXStatePropertyName = ActiveXStatePropertyName,
            HasModifiedActiveXState = HasModifiedActiveXState,
            LicenseKey = LicenseKey,
            IsUnresolved = IsUnresolved,
            UnresolvedErrorMessage = UnresolvedErrorMessage,
            RawControlLines = new List<string>(RawControlLines),
            _hasPropertyPages = _hasPropertyPages
        };

        foreach (var kvp in FontProperties)
        {
            clone.FontProperties[kvp.Key] = kvp.Value.Clone();
        }

        foreach (var child in Children)
        {
            clone.Children.Add(child.Clone(clone));
        }

        return clone;
    }
}
