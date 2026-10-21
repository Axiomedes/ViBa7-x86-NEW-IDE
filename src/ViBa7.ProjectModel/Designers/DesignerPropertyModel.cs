using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ViBa7.ProjectModel.Designers;

public enum DesignerPropertyType
{
    Text,
    IntegerTwips,
    IntegerPixels,
    IntegerRaw,
    Boolean,
    ColorHex,
    EnumChoice,
    FontDescription,
    ReadOnlyText,
    Picture,
    CustomDialog
}

public record DesignerPropertyChoice(string DisplayText, string RawValue);

/// <summary>
/// Representa un elemento de propiedad individual listo para ser inspeccionado y editado en la Ventana de Propiedades.
/// </summary>
public class DesignerPropertyItem
{
    public string Name { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string Category { get; init; } = "General";
    public string Description { get; init; } = string.Empty;
    public DesignerPropertyType Type { get; init; } = DesignerPropertyType.Text;
    public string CurrentValue { get; set; } = string.Empty;
    public List<DesignerPropertyChoice> Choices { get; init; } = new();
    public bool IsReadOnly { get; init; }
    public Func<string, bool>? ValueSetter { get; init; }

    public bool SetValue(string newValue)
    {
        if (IsReadOnly || ValueSetter == null) return false;
        if (ValueSetter(newValue))
        {
            CurrentValue = newValue;
            return true;
        }
        return false;
    }
}

/// <summary>
/// Fábrica que genera descriptores de propiedades para formularios y controles VB6.
/// </summary>
public static class DesignerPropertyFactory
{
    /// <summary>
    /// Delegado global para obtener la lista de propiedades expuestas de un UserControl por su TypeName.
    /// </summary>
    public static Func<string, List<UserControlExposedProperty>?>? UserControlPropertiesResolver { get; set; }

    private static readonly List<DesignerPropertyChoice> BoolChoices = new()
    {
        new DesignerPropertyChoice("True", "True"),
        new DesignerPropertyChoice("False", "False")
    };

    private static readonly List<DesignerPropertyChoice> FormBorderChoices = new()
    {
        new DesignerPropertyChoice("0 - None", "0"),
        new DesignerPropertyChoice("1 - Fixed Single", "1"),
        new DesignerPropertyChoice("2 - Sizable", "2"),
        new DesignerPropertyChoice("3 - Fixed Dialog", "3"),
        new DesignerPropertyChoice("4 - Fixed ToolWindow", "4"),
        new DesignerPropertyChoice("5 - Sizable ToolWindow", "5")
    };

    private static readonly List<DesignerPropertyChoice> StartUpChoices = new()
    {
        new DesignerPropertyChoice("0 - Manual", "0"),
        new DesignerPropertyChoice("1 - CenterOwner", "1"),
        new DesignerPropertyChoice("2 - CenterScreen", "2"),
        new DesignerPropertyChoice("3 - Windows Default", "3")
    };

    private static readonly List<DesignerPropertyChoice> AlignmentChoices = new()
    {
        new DesignerPropertyChoice("0 - Left Justify", "0"),
        new DesignerPropertyChoice("1 - Right Justify", "1"),
        new DesignerPropertyChoice("2 - Center", "2")
    };

    private static readonly List<DesignerPropertyChoice> PictureAlignmentChoices = new()
    {
        new DesignerPropertyChoice("0 - Superior (Top)", "0"),
        new DesignerPropertyChoice("1 - Inferior (Bottom)", "1"),
        new DesignerPropertyChoice("2 - Izquierda (Left)", "2"),
        new DesignerPropertyChoice("3 - Derecha (Right)", "3"),
        new DesignerPropertyChoice("4 - Centro (Center)", "4")
    };

    private static readonly List<DesignerPropertyChoice> ButtonStyleChoices = new()
    {
        new DesignerPropertyChoice("0 - Standard", "0"),
        new DesignerPropertyChoice("1 - Graphical", "1")
    };

    /// <summary>
    /// Proveedor de lista de tipografías instaladas en el sistema (configurado desde la capa UI de WPF).
    /// </summary>
    public static Func<List<DesignerPropertyChoice>>? SystemFontChoicesProvider { get; set; }

    private static readonly List<DesignerPropertyChoice> FallbackFontChoices = new()
    {
        new("Arial", "Arial"),
        new("Calibri", "Calibri"),
        new("Comic Sans MS", "Comic Sans MS"),
        new("Consolas", "Consolas"),
        new("Courier New", "Courier New"),
        new("Georgia", "Georgia"),
        new("Impact", "Impact"),
        new("Lucida Console", "Lucida Console"),
        new("Microsoft Sans Serif", "Microsoft Sans Serif"),
        new("MS Sans Serif", "MS Sans Serif"),
        new("Segoe UI", "Segoe UI"),
        new("Tahoma", "Tahoma"),
        new("Times New Roman", "Times New Roman"),
        new("Trebuchet MS", "Trebuchet MS"),
        new("Verdana", "Verdana")
    };

    public static List<DesignerPropertyChoice> GetSystemFontChoices()
    {
        if (SystemFontChoicesProvider != null)
        {
            try
            {
                var choices = SystemFontChoicesProvider();
                if (choices != null && choices.Count > 0) return choices;
            }
            catch { }
        }
        return FallbackFontChoices;
    }

    /// <summary>
    /// Determina si una propiedad representa un color OLE / VB6 (&amp;H00BBGGRR&amp;, OLE_COLOR, colores de sistema).
    /// </summary>
    public static bool IsColorProperty(string propName, string? typeName = null, string? currentValue = null)
    {
        if (string.IsNullOrWhiteSpace(propName)) return false;

        // 1. Tipo explícito OLE_COLOR o que contiene COLOR
        if (!string.IsNullOrEmpty(typeName))
        {
            string t = typeName.Trim().ToUpperInvariant();
            if (t.Contains("OLE_COLOR") || t.Contains("COLOR") || t.Equals("COLORCONSTANTS") || t.Contains("COLORREF"))
            {
                return true;
            }
        }

        // 2. Nombre de propiedad que contiene Color o Colour (BackColor, ForeColor1, BorderColorList, GridLineColor, etc.)
        if (propName.Contains("Color", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Colour", StringComparison.OrdinalIgnoreCase) ||
            propName.EndsWith("Clr", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("BkColor", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("FgColor", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Background", StringComparison.OrdinalIgnoreCase) ||
            propName.Contains("Foreground", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // 3. Nombres típicos de propiedades de color en controles VB6
        if (propName.Equals("ColorFace", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("ColorHighlight", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("ColorShadow", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("ColorLight", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("ColorDkShadow", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("FillColor", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("LineColor", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("BorderColor", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("Fill", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("Tint", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        // 4. Valor actual en formato de color VB6
        if (!string.IsNullOrWhiteSpace(currentValue))
        {
            string val = currentValue.Trim();
            // Formato hexadecimal VB6: &H00BBGGRR& o &H8000000F& o &HFFFFFF
            if (val.StartsWith("&H", StringComparison.OrdinalIgnoreCase) || val.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                string hexOnly = val.TrimStart('&').TrimStart('H', 'h').TrimEnd('&');
                if (hexOnly.Length >= 1 && hexOnly.Length <= 8 && uint.TryParse(hexOnly, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
                {
                    return true;
                }
            }
            // Color de sistema de 32-bit con signo en decimal (e.g. -2147483633)
            else if (val.StartsWith("-2147483") && long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out long sysCol) &&
                     sysCol >= -2147483648 && sysCol <= -2147483600)
            {
                return true;
            }
            // Valor decimal positivo típico de RGB (0 a 16777215) si el nombre tiene indicio de color
            else if (long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out long rgbVal) && rgbVal >= 0 && rgbVal <= 16777215)
            {
                if (propName.Contains("Color", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Clr", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Bk", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Fg", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Back", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Fore", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Face", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Shadow", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Highlight", StringComparison.OrdinalIgnoreCase) ||
                    propName.Contains("Fill", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Sincroniza campos estándar de FormDesignerControl cuando se modifica una propiedad expuesta o de CustomProperties.
    /// </summary>
    public static void SynchronizeStandardProperties(FormDesignerControl ctrl, string propName, string val)
    {
        if (string.Equals(propName, "Caption", StringComparison.OrdinalIgnoreCase)) ctrl.Caption = val;
        else if (string.Equals(propName, "Text", StringComparison.OrdinalIgnoreCase)) ctrl.Text = val;
        else if (string.Equals(propName, "BackColor", StringComparison.OrdinalIgnoreCase)) ctrl.BackColorHex = val;
        else if (string.Equals(propName, "ForeColor", StringComparison.OrdinalIgnoreCase)) ctrl.ForeColorHex = val;
        else if (string.Equals(propName, "BorderStyle", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int bs)) ctrl.BorderStyle = bs;
        else if (string.Equals(propName, "Alignment", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int al)) ctrl.Alignment = al;
        else if (string.Equals(propName, "Enabled", StringComparison.OrdinalIgnoreCase)) ctrl.Enabled = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        else if (string.Equals(propName, "Visible", StringComparison.OrdinalIgnoreCase)) ctrl.Visible = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        else if (string.Equals(propName, "Left", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int left)) ctrl.Left = left;
        else if (string.Equals(propName, "Top", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int top)) ctrl.Top = top;
        else if (string.Equals(propName, "Width", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int width))
        {
            ctrl.Width = width;
            ctrl.UpdateExtentsFromDimensions();
        }
        else if (string.Equals(propName, "Height", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out int height))
        {
            ctrl.Height = height;
            ctrl.UpdateExtentsFromDimensions();
        }
        else if (string.Equals(propName, "FontName", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(val))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.Name = val.Trim();
        }
        else if (string.Equals(propName, "FontSize", StringComparison.OrdinalIgnoreCase) && double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out double fsz))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.Size = fsz;
        }
        else if (string.Equals(propName, "FontBold", StringComparison.OrdinalIgnoreCase))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.IsBold = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        }
        else if (string.Equals(propName, "FontItalic", StringComparison.OrdinalIgnoreCase))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.Italic = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        }
        else if (string.Equals(propName, "FontUnderline", StringComparison.OrdinalIgnoreCase))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.Underline = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        }
        else if (string.Equals(propName, "FontStrikethru", StringComparison.OrdinalIgnoreCase))
        {
            ctrl.Font ??= new FormFontDescription();
            ctrl.Font.Strikethrough = val.Equals("True", StringComparison.OrdinalIgnoreCase) || val == "-1" || val == "1";
        }
        else if (propName.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
        {
            string norm = NormalizeIconCharCode(val);
            ctrl.SetCustomProperty(propName, norm);
        }
        else if (propName.EndsWith("Font", StringComparison.OrdinalIgnoreCase) || propName.Contains("Font", StringComparison.OrdinalIgnoreCase))
        {
            var f = ctrl.GetFontProperty(propName) ?? new FormFontDescription();
            if (!string.IsNullOrWhiteSpace(val))
            {
                var parts = val.Split(',');
                f.Name = parts[0].Trim();
                if (parts.Length > 1 && double.TryParse(parts[1].Replace("pt", "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double sz))
                {
                    f.Size = sz;
                }
            }
            ctrl.SetFontProperty(propName, f);
            if (propName.Equals("Font", StringComparison.OrdinalIgnoreCase))
            {
                ctrl.Font = f;
            }
        }
    }

    /// <summary>
    /// Valida y normaliza un código de carácter de icono (ej. axLabelPlus IconCharCode) a formato canónico &amp;HXXXX.
    /// Soporta entradas como: "EA01", "ea01", "&amp;HEA01", "U+EA01", "0xEA01", o valor decimal "59905".
    /// </summary>
    public static string NormalizeIconCharCode(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        string s = raw.Trim();
        if (s.StartsWith("\"") && s.EndsWith("\"") && s.Length >= 2)
        {
            s = s[1..^1].Trim();
        }
        s = s.Replace(" ", "");
        if (s.StartsWith("U+", StringComparison.OrdinalIgnoreCase) || s.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
        {
            s = "&H" + s[2..];
        }
        else if (!s.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            // Si es puramente numérico decimal
            if (long.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out long decVal))
            {
                s = "&H" + decVal.ToString("X");
            }
            else
            {
                // Puede ser un hex directo sin prefijo, ej "EA01"
                bool isHex = s.Length > 0 && s.All(Uri.IsHexDigit);
                if (isHex)
                {
                    s = "&H" + s;
                }
            }
        }

        if (s.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            string hexPart = s[2..].TrimEnd('&');
            return "&H" + hexPart.ToUpperInvariant();
        }

        return s.ToUpperInvariant();
    }

    /// <summary>
    /// Convierte un código de carácter de icono (ej. "&amp;HEA01", "EA01", "59905") a su punto de código entero Unicode.
    /// </summary>
    public static int? ParseCharCode(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        string norm = NormalizeIconCharCode(raw);
        int? val = null;
        if (norm.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            string hex = norm[2..].TrimEnd('&');
            if (int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hVal))
            {
                val = hVal;
            }
        }
        else if (int.TryParse(norm, NumberStyles.Integer, CultureInfo.InvariantCulture, out int dVal))
        {
            val = dVal;
        }

        if (val.HasValue && val.Value >= 0 && val.Value <= 0x10FFFF && !(val.Value >= 0xD800 && val.Value <= 0xDFFF))
        {
            return val.Value;
        }
        return null;
    }

    /// <summary>
    /// Infiere el tipo de propiedad de diseñador y sus opciones para propiedades de CustomProperties o TypeLib.
    /// </summary>
    public static (DesignerPropertyType Type, List<DesignerPropertyChoice> Choices) InferPropertyType(
        string propName,
        string? typeName,
        string? currentValue)
    {
        // 1. Color
        if (IsColorProperty(propName, typeName, currentValue))
        {
            return (DesignerPropertyType.ColorHex, new List<DesignerPropertyChoice>());
        }

        string t = (typeName ?? string.Empty).Trim().ToUpperInvariant();

        // 2. Booleano
        if (t.Equals("BOOLEAN") ||
            propName.Equals("Enabled", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("Visible", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("Sorted", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("WordWrap", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("FullRow", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("MultiSelect", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("DrawEmpty", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("Striped", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("AutoRedraw", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("HandCursor", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("ShowDropBack", StringComparison.OrdinalIgnoreCase) ||
            propName.Equals("FocusRect", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.Boolean, BoolChoices);
        }

        // 3. Picture
        if (t.Contains("PICTURE") || t.Contains("IPIC") || t.Contains("STDPIC") || propName.Contains("Picture", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.Picture, new List<DesignerPropertyChoice>());
        }

        // 4. FontName (lista desplegable con las tipografías instaladas en el sistema)
        if (propName.Equals("FontName", StringComparison.OrdinalIgnoreCase) ||
            (propName.Contains("Font", StringComparison.OrdinalIgnoreCase) && propName.EndsWith("Name", StringComparison.OrdinalIgnoreCase)))
        {
            return (DesignerPropertyType.EnumChoice, GetSystemFontChoices());
        }

        // 4.5 Códigos de carácter de iconos (ej. IconCharCode, CharCode)
        if (propName.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.Text, new List<DesignerPropertyChoice>());
        }

        // 5. Font (objeto StdFont / bloque BeginProperty Font con selector de fuente)
        if (t.Contains("FONT") || propName.Equals("Font", StringComparison.OrdinalIgnoreCase) || propName.EndsWith("Font", StringComparison.OrdinalIgnoreCase) || propName.EndsWith(".Font", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.FontDescription, new List<DesignerPropertyChoice>());
        }

        // 5. Enumeraciones de TypeLibs registradas
        if (!string.IsNullOrEmpty(typeName))
        {
            var enumChoices = IntelliSense.VbControlMemberCatalog.GetEnumChoices(typeName);
            if (enumChoices != null && enumChoices.Count > 0)
            {
                return (DesignerPropertyType.EnumChoice, enumChoices);
            }
        }

        // 6. Propiedades con listas fijas estándar de VB6
        if (propName.Equals("BorderStyle", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.EnumChoice, new List<DesignerPropertyChoice>
            {
                new("0 - None", "0"),
                new("1 - Fixed Single", "1")
            });
        }
        if (propName.Equals("Appearance", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.EnumChoice, new List<DesignerPropertyChoice>
            {
                new("0 - Flat", "0"),
                new("1 - 3D", "1")
            });
        }
        if (propName.Equals("BackStyle", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.EnumChoice, new List<DesignerPropertyChoice>
            {
                new("0 - Transparent", "0"),
                new("1 - Opaque", "1")
            });
        }
        if (propName.Equals("Alignment", StringComparison.OrdinalIgnoreCase) || propName.Equals("TextAlign", StringComparison.OrdinalIgnoreCase))
        {
            return (DesignerPropertyType.EnumChoice, AlignmentChoices);
        }

        // 7. Números enteros
        if (t.Equals("INTEGER") || t.Equals("LONG") || t.Equals("BYTE") || t.Equals("SINGLE") || t.Equals("DOUBLE"))
        {
            return (DesignerPropertyType.IntegerRaw, new List<DesignerPropertyChoice>());
        }

        if (!string.IsNullOrWhiteSpace(currentValue))
        {
            string cv = currentValue.Trim();
            if (cv.Equals("True", StringComparison.OrdinalIgnoreCase) || cv.Equals("False", StringComparison.OrdinalIgnoreCase))
            {
                return (DesignerPropertyType.Boolean, BoolChoices);
            }
            if (long.TryParse(cv, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) && !cv.StartsWith("-2147483"))
            {
                return (DesignerPropertyType.IntegerRaw, new List<DesignerPropertyChoice>());
            }
        }

        return (DesignerPropertyType.Text, new List<DesignerPropertyChoice>());
    }

    public static List<DesignerPropertyItem> GetPropertiesForForm(FormDesignModel form, Action onModelChanged)
    {
        var list = new List<DesignerPropertyItem>
        {
            new DesignerPropertyItem
            {
                Name = "Name",
                DisplayName = "(Name)",
                Category = "General",
                Description = "Devuelve o establece el nombre utilizado en el código para identificar el formulario.",
                Type = DesignerPropertyType.Text,
                CurrentValue = form.Name,
                ValueSetter = val =>
                {
                    val = val.Trim();
                    if (string.IsNullOrWhiteSpace(val)) return false;
                    form.Name = val;
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "Caption",
                DisplayName = "Caption",
                Category = "Appearance",
                Description = "Devuelve o establece el texto que se muestra en la barra de título del formulario.",
                Type = DesignerPropertyType.Text,
                CurrentValue = form.Caption,
                ValueSetter = val =>
                {
                    form.Caption = val;
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "BackColor",
                DisplayName = "BackColor",
                Category = "Appearance",
                Description = "Devuelve o establece el color de fondo del formulario.",
                Type = DesignerPropertyType.ColorHex,
                CurrentValue = form.BackColorHex ?? "&H8000000F&",
                ValueSetter = val =>
                {
                    form.BackColorHex = val;
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "ForeColor",
                DisplayName = "ForeColor",
                Category = "Appearance",
                Description = "Devuelve o establece el color de primer plano para texto y gráficos en el formulario.",
                Type = DesignerPropertyType.ColorHex,
                CurrentValue = form.ForeColorHex ?? "&H80000012&",
                ValueSetter = val =>
                {
                    form.ForeColorHex = val;
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "Picture",
                DisplayName = "Picture",
                Category = "Appearance",
                Description = "Devuelve o establece el gráfico que se va a mostrar como fondo del formulario.",
                Type = DesignerPropertyType.Picture,
                CurrentValue = Resources.FrxResourceFile.GetFormatDescription(form.PictureData, form.Picture),
                ValueSetter = val =>
                {
                    if (string.IsNullOrWhiteSpace(val) || val.Equals("(None)", StringComparison.OrdinalIgnoreCase))
                    {
                        form.Picture = null;
                        form.PictureData = null;
                        onModelChanged();
                        return true;
                    }
                    if (val.StartsWith('(') && val.EndsWith(')'))
                    {
                        return true;
                    }
                    form.Picture = val;
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "BorderStyle",
                DisplayName = "BorderStyle",
                Category = "Appearance",
                Description = "Devuelve o establece el estilo del borde del formulario (redimensionable, fijo, diálogo, etc.).",
                Type = DesignerPropertyType.EnumChoice,
                CurrentValue = form.BorderStyle?.ToString() ?? "2",
                Choices = FormBorderChoices,
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int b))
                    {
                        form.BorderStyle = b;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            },
            new DesignerPropertyItem
            {
                Name = "StartUpPosition",
                DisplayName = "StartUpPosition",
                Category = "Position",
                Description = "Devuelve o establece la posición en pantalla en la que aparecerá el formulario al cargarse.",
                Type = DesignerPropertyType.EnumChoice,
                CurrentValue = form.StartUpPosition?.ToString() ?? "3",
                Choices = StartUpChoices,
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int sp))
                    {
                        form.StartUpPosition = sp;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            },
            new DesignerPropertyItem
            {
                Name = "ClientWidth",
                DisplayName = "ClientWidth",
                Category = "Position",
                Description = "Devuelve o establece el ancho del área cliente del formulario en Twips.",
                Type = DesignerPropertyType.IntegerTwips,
                CurrentValue = form.ClientWidth.ToString(),
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int w) && w > 0)
                    {
                        form.ClientWidth = w;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            },
            new DesignerPropertyItem
            {
                Name = "ClientHeight",
                DisplayName = "ClientHeight",
                Category = "Position",
                Description = "Devuelve o establece el alto del área cliente del formulario en Twips.",
                Type = DesignerPropertyType.IntegerTwips,
                CurrentValue = form.ClientHeight.ToString(),
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int h) && h > 0)
                    {
                        form.ClientHeight = h;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            },
            new DesignerPropertyItem
            {
                Name = "Font",
                DisplayName = "Font",
                Category = "Font",
                Description = "Devuelve o establece la tipografía predeterminada para el formulario.",
                Type = DesignerPropertyType.FontDescription,
                CurrentValue = form.Font?.Name != null ? $"{form.Font.Name}, {form.Font.Size}pt" : "MS Sans Serif, 8pt",
                IsReadOnly = true
            }
        };

        return list;
    }

    public static List<DesignerPropertyItem> CreateForControl(
        FormDesignerControl ctrl,
        Action onModelChanged,
        FormDesignModel? form = null)
    {
        return GetPropertiesForControl(ctrl, form, onModelChanged);
    }

    public static List<DesignerPropertyItem> GetPropertiesForControl(
        FormDesignerControl ctrl,
        FormDesignModel? form,
        Action onModelChanged,
        Action<FormDesignerControl, string, string>? onControlRenamed = null,
        Func<string, bool>? confirmCreateControlArray = null,
        Action<string>? showWarning = null)
    {
        var list = new List<DesignerPropertyItem>
        {
            new DesignerPropertyItem
            {
                Name = "Name",
                DisplayName = "(Name)",
                Category = "General",
                Description = "Devuelve o establece el nombre utilizado en el código para identificar el control.",
                Type = DesignerPropertyType.Text,
                CurrentValue = ctrl.Name,
                ValueSetter = val =>
                {
                    val = val.Trim();
                    if (string.IsNullOrWhiteSpace(val)) return false;
                    if (string.Equals(ctrl.Name, val, StringComparison.OrdinalIgnoreCase))
                    {
                        ctrl.Name = val;
                        onModelChanged();
                        return true;
                    }

                    // Buscar controles existentes con ese mismo nombre
                    var existingControls = form?.GetAllControls()
                        .Where(c => c != ctrl && string.Equals(c.Name, val, StringComparison.OrdinalIgnoreCase))
                        .ToList() ?? new List<FormDesignerControl>();

                    if (existingControls.Count > 0)
                    {
                        if (!ctrl.Index.HasValue)
                        {
                            bool shouldCreate = confirmCreateControlArray?.Invoke(val) ?? false;
                            if (!shouldCreate)
                            {
                                return false;
                            }
                        }

                        var firstExisting = existingControls[0];
                        if (!firstExisting.Index.HasValue)
                        {
                            firstExisting.Index = 0;
                        }

                        var usedIndices = existingControls.Where(c => c.Index.HasValue).Select(c => c.Index!.Value).ToHashSet();
                        if (firstExisting.Index.HasValue) usedIndices.Add(firstExisting.Index.Value);

                        int nextIdx = 0;
                        while (usedIndices.Contains(nextIdx)) nextIdx++;

                        string oldName = ctrl.Name;
                        ctrl.Name = val;
                        if (!ctrl.Index.HasValue || usedIndices.Contains(ctrl.Index.Value))
                        {
                            ctrl.Index = nextIdx;
                        }
                        onControlRenamed?.Invoke(ctrl, oldName, val);
                        onModelChanged();
                        return true;
                    }

                    string prevName = ctrl.Name;
                    ctrl.Name = val;
                    onControlRenamed?.Invoke(ctrl, prevName, val);
                    onModelChanged();
                    return true;
                }
            },
            new DesignerPropertyItem
            {
                Name = "Index",
                DisplayName = "Index",
                Category = "General",
                Description = "Devuelve o establece el número que identifica a un control en una matriz de controles.",
                Type = DesignerPropertyType.IntegerRaw,
                CurrentValue = ctrl.Index?.ToString() ?? "",
                ValueSetter = val =>
                {
                    val = val?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(val))
                    {
                        var siblings = form?.GetAllControls()
                            .Where(c => c != ctrl && string.Equals(c.Name, ctrl.Name, StringComparison.OrdinalIgnoreCase))
                            .ToList() ?? new List<FormDesignerControl>();
                        if (siblings.Count > 0)
                        {
                            showWarning?.Invoke($"El control '{ctrl.Name}' pertenece a una matriz de controles y debe tener un índice válido.");
                            return false;
                        }

                        ctrl.Index = null;
                        onModelChanged();
                        return true;
                    }

                    if (int.TryParse(val, out int idx) && idx >= 0)
                    {
                        var duplicate = form?.GetAllControls()
                            .FirstOrDefault(c => c != ctrl &&
                                                 string.Equals(c.Name, ctrl.Name, StringComparison.OrdinalIgnoreCase) &&
                                                 c.Index == idx);

                        if (duplicate != null)
                        {
                            showWarning?.Invoke($"Ya existe un control '{ctrl.Name}' con el índice {idx} en la matriz.");
                            return false;
                        }

                        ctrl.Index = idx;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            }
        };

        // Si es un Timer, mostrar exclusivamente las propiedades soportadas por VB6 para temporizadores no visibles
        if (ctrl.IsTimer)
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "Interval",
                DisplayName = "Interval",
                Category = "Behavior",
                Description = "Devuelve o establece el número de milisegundos entre llamadas a un evento Timer (Tik). 0 desactiva el temporizador.",
                Type = DesignerPropertyType.IntegerRaw,
                CurrentValue = ctrl.Interval.ToString(),
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int iv) && iv >= 0 && iv <= 65535)
                    {
                        ctrl.Interval = iv;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });

            list.Add(new DesignerPropertyItem
            {
                Name = "Enabled",
                DisplayName = "Enabled",
                Category = "Behavior",
                Description = "Determina si el temporizador se ejecuta y produce eventos Timer periódicamente.",
                Type = DesignerPropertyType.Boolean,
                CurrentValue = ctrl.Enabled ? "True" : "False",
                Choices = BoolChoices,
                ValueSetter = val =>
                {
                    if (bool.TryParse(val, out bool b))
                    {
                        ctrl.Enabled = b;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });

            list.Add(new DesignerPropertyItem
            {
                Name = "Left",
                DisplayName = "Left",
                Category = "Position",
                Description = "Posición horizontal del icono del temporizador en diseño en Twips.",
                Type = DesignerPropertyType.IntegerTwips,
                CurrentValue = ctrl.Left.ToString(),
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int v)) { ctrl.Left = v; onModelChanged(); return true; }
                    return false;
                }
            });

            list.Add(new DesignerPropertyItem
            {
                Name = "Top",
                DisplayName = "Top",
                Category = "Position",
                Description = "Posición vertical del icono del temporizador en diseño en Twips.",
                Type = DesignerPropertyType.IntegerTwips,
                CurrentValue = ctrl.Top.ToString(),
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int v)) { ctrl.Top = v; onModelChanged(); return true; }
                    return false;
                }
            });

            list.Add(new DesignerPropertyItem
            {
                Name = "Tag",
                DisplayName = "Tag",
                Category = "General",
                Description = "Almacena cualquier dato adicional necesario para su programa.",
                Type = DesignerPropertyType.Text,
                CurrentValue = ctrl.CustomProperties.FirstOrDefault(p => p.Name.Equals("Tag", StringComparison.OrdinalIgnoreCase))?.Value ?? "",
                ValueSetter = val =>
                {
                    var p = ctrl.CustomProperties.FirstOrDefault(cp => cp.Name.Equals("Tag", StringComparison.OrdinalIgnoreCase));
                    if (p != null) p.Value = val;
                    else ctrl.CustomProperties.Add(new FormPropertyItem { Name = "Tag", Value = val });
                    onModelChanged();
                    return true;
                }
            });

            return list;
        }

        // Caption o Text según el tipo de control (exclusivo de controles intrínsecos VB6; UserControls y OCX los exponen desde su código/TypeLib)
        if (ctrl.IsIntrinsic)
        {
            var toolboxInfo = ControlToolboxRegistry.Find(ctrl.TypeName);
            if (toolboxInfo?.HasCaption == true || !string.IsNullOrEmpty(ctrl.Caption) ||
                ctrl.TypeName.EndsWith("Button", StringComparison.OrdinalIgnoreCase) ||
                ctrl.TypeName.EndsWith("Label", StringComparison.OrdinalIgnoreCase) ||
                ctrl.TypeName.EndsWith("Frame", StringComparison.OrdinalIgnoreCase) ||
                ctrl.TypeName.EndsWith("CheckBox", StringComparison.OrdinalIgnoreCase) ||
                ctrl.TypeName.EndsWith("OptionButton", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new DesignerPropertyItem
                {
                    Name = "Caption",
                    DisplayName = "Caption",
                    Category = "Appearance",
                    Description = "Devuelve o establece el texto que se muestra en el control.",
                    Type = DesignerPropertyType.Text,
                    CurrentValue = ctrl.Caption,
                    ValueSetter = val =>
                    {
                        ctrl.Caption = val;
                        onModelChanged();
                        return true;
                    }
                });
            }

            if (toolboxInfo?.HasText == true || !string.IsNullOrEmpty(ctrl.Text) ||
                ctrl.TypeName.EndsWith("TextBox", StringComparison.OrdinalIgnoreCase) ||
                ctrl.TypeName.EndsWith("ComboBox", StringComparison.OrdinalIgnoreCase))
            {
                list.Add(new DesignerPropertyItem
                {
                    Name = "Text",
                    DisplayName = "Text",
                    Category = "Appearance",
                    Description = "Devuelve o establece el texto contenido en el control.",
                    Type = DesignerPropertyType.Text,
                    CurrentValue = ctrl.Text,
                    ValueSetter = val =>
                    {
                        ctrl.Text = val;
                        onModelChanged();
                        return true;
                    }
                });
            }
        }

        // Posicionamiento y dimensiones en Twips
        list.Add(new DesignerPropertyItem
        {
            Name = "Left",
            DisplayName = "Left",
            Category = "Position",
            Description = "Devuelve o establece la distancia entre el borde izquierdo del control y el contenedor en Twips.",
            Type = DesignerPropertyType.IntegerTwips,
            CurrentValue = ctrl.Left.ToString(),
            ValueSetter = val =>
            {
                if (int.TryParse(val, out int v))
                {
                    ctrl.Left = v;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "Top",
            DisplayName = "Top",
            Category = "Position",
            Description = "Devuelve o establece la distancia entre el borde superior del control y el contenedor en Twips.",
            Type = DesignerPropertyType.IntegerTwips,
            CurrentValue = ctrl.Top.ToString(),
            ValueSetter = val =>
            {
                if (int.TryParse(val, out int v))
                {
                    ctrl.Top = v;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "Width",
            DisplayName = "Width",
            Category = "Position",
            Description = "Devuelve o establece el ancho del control en Twips.",
            Type = DesignerPropertyType.IntegerTwips,
            CurrentValue = ctrl.Width.ToString(),
            ValueSetter = val =>
            {
                if (int.TryParse(val, out int v) && v >= 0)
                {
                    ctrl.Width = v;
                    ctrl.UpdateExtentsFromDimensions();
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "Height",
            DisplayName = "Height",
            Category = "Position",
            Description = "Devuelve o establece la altura del control en Twips.",
            Type = DesignerPropertyType.IntegerTwips,
            CurrentValue = ctrl.Height.ToString(),
            ValueSetter = val =>
            {
                if (int.TryParse(val, out int v) && v >= 0)
                {
                    ctrl.Height = v;
                    ctrl.UpdateExtentsFromDimensions();
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        // Comportamiento
        list.Add(new DesignerPropertyItem
        {
            Name = "Enabled",
            DisplayName = "Enabled",
            Category = "Behavior",
            Description = "Determina si el control puede responder a eventos generados por el usuario.",
            Type = DesignerPropertyType.Boolean,
            CurrentValue = ctrl.Enabled ? "True" : "False",
            Choices = BoolChoices,
            ValueSetter = val =>
            {
                if (bool.TryParse(val, out bool b))
                {
                    ctrl.Enabled = b;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "Visible",
            DisplayName = "Visible",
            Category = "Behavior",
            Description = "Determina si el control está visible u oculto.",
            Type = DesignerPropertyType.Boolean,
            CurrentValue = ctrl.Visible ? "True" : "False",
            Choices = BoolChoices,
            ValueSetter = val =>
            {
                if (bool.TryParse(val, out bool b))
                {
                    ctrl.Visible = b;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "TabIndex",
            DisplayName = "TabIndex",
            Category = "Behavior",
            Description = "Devuelve o establece el orden de tabulación del control dentro de su formulario principal.",
            Type = DesignerPropertyType.IntegerRaw,
            CurrentValue = ctrl.TabIndex?.ToString() ?? "",
            ValueSetter = val =>
            {
                if (string.IsNullOrWhiteSpace(val))
                {
                    ctrl.TabIndex = null;
                    onModelChanged();
                    return true;
                }
                if (int.TryParse(val, out int t) && t >= 0)
                {
                    ctrl.TabIndex = t;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "TabStop",
            DisplayName = "TabStop",
            Category = "Behavior",
            Description = "Determina si el usuario puede utilizar la tecla TAB para pasar el foco al control.",
            Type = DesignerPropertyType.Boolean,
            CurrentValue = (ctrl.TabStop ?? true) ? "True" : "False",
            Choices = BoolChoices,
            ValueSetter = val =>
            {
                if (bool.TryParse(val, out bool b))
                {
                    ctrl.TabStop = b;
                    onModelChanged();
                    return true;
                }
                return false;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "Tag",
            DisplayName = "Tag",
            Category = "General",
            Description = "Almacena cualquier dato adicional necesario para su programa.",
            Type = DesignerPropertyType.Text,
            CurrentValue = ctrl.CustomProperties.FirstOrDefault(p => p.Name.Equals("Tag", StringComparison.OrdinalIgnoreCase))?.Value ?? "",
            ValueSetter = val =>
            {
                var p = ctrl.CustomProperties.FirstOrDefault(cp => cp.Name.Equals("Tag", StringComparison.OrdinalIgnoreCase));
                if (p != null) p.Value = val;
                else ctrl.CustomProperties.Add(new FormPropertyItem { Name = "Tag", Value = val });
                onModelChanged();
                return true;
            }
        });

        // Propiedades estándar intrínsecas de VB6 (Colores, Font, Alineación, Imágenes)
        // Solo aplican por defecto a controles intrínsecos de VB. Para UserControls y OCX,
        // estas propiedades deben provenir exclusivamente de su código fuente (.ctl) o TypeLib.
        if (ctrl.IsIntrinsic)
        {
            // Colores
            list.Add(new DesignerPropertyItem
            {
                Name = "BackColor",
                DisplayName = "BackColor",
            Category = "Appearance",
            Description = "Devuelve o establece el color de fondo utilizado para mostrar texto y gráficos en el control.",
            Type = DesignerPropertyType.ColorHex,
            CurrentValue = ctrl.BackColorHex ?? "&H8000000F&",
            ValueSetter = val =>
            {
                ctrl.BackColorHex = val;
                onModelChanged();
                return true;
            }
        });

        list.Add(new DesignerPropertyItem
        {
            Name = "ForeColor",
            DisplayName = "ForeColor",
            Category = "Appearance",
            Description = "Devuelve o establece el color de primer plano utilizado para mostrar texto y gráficos en el control.",
            Type = DesignerPropertyType.ColorHex,
            CurrentValue = ctrl.ForeColorHex ?? "&H80000012&",
            ValueSetter = val =>
            {
                ctrl.ForeColorHex = val;
                onModelChanged();
                return true;
            }
        });

        // Fuente tipográfica (Font / StdFont) para controles con capacidad de texto
        bool hasFont = !ctrl.IsTimer &&
                       !ctrl.TypeName.EndsWith("Line", StringComparison.OrdinalIgnoreCase) &&
                       !ctrl.TypeName.EndsWith("Shape", StringComparison.OrdinalIgnoreCase);
        if (hasFont)
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "Font",
                DisplayName = "Font",
                Category = "Font",
                Description = "Devuelve o establece el objeto de fuente (StdFont) utilizado para mostrar texto en el control.",
                Type = DesignerPropertyType.FontDescription,
                CurrentValue = ctrl.Font?.Name != null
                    ? $"{ctrl.Font.Name}, {ctrl.Font.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt"
                    : (form?.Font?.Name != null ? $"{form.Font.Name}, {form.Font.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt" : "MS Sans Serif, 8pt"),
                IsReadOnly = true,
                ValueSetter = _ => true
            });
        }

        // Alineación (si aplica)
        if (ctrl.TypeName.EndsWith("Label", StringComparison.OrdinalIgnoreCase) ||
            ctrl.TypeName.EndsWith("TextBox", StringComparison.OrdinalIgnoreCase))
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "Alignment",
                DisplayName = "Alignment",
                Category = "Appearance",
                Description = "Determina la alineación del texto dentro del control (Izquierda, Derecha o Centro).",
                Type = DesignerPropertyType.EnumChoice,
                CurrentValue = ctrl.Alignment?.ToString() ?? "0",
                Choices = AlignmentChoices,
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int a))
                    {
                        ctrl.Alignment = a;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });
        }

        // MultiLine para TextBox
        if (ctrl.TypeName.EndsWith("TextBox", StringComparison.OrdinalIgnoreCase))
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "MultiLine",
                DisplayName = "MultiLine",
                Category = "Behavior",
                Description = "Indica si un control TextBox puede aceptar y mostrar múltiples líneas de texto.",
                Type = DesignerPropertyType.Boolean,
                CurrentValue = (ctrl.MultiLine ?? false) ? "True" : "False",
                Choices = BoolChoices,
                ValueSetter = val =>
                {
                    if (bool.TryParse(val, out bool b))
                    {
                        ctrl.MultiLine = b;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });
        }

        // Soporte de imágenes para PictureBox, Image, CommandButton, etc.
        bool isPictureBox = ctrl.TypeName.EndsWith("PictureBox", StringComparison.OrdinalIgnoreCase);
        bool isImageControl = ctrl.TypeName.EndsWith("Image", StringComparison.OrdinalIgnoreCase);
        bool isButton = ctrl.TypeName.EndsWith("Button", StringComparison.OrdinalIgnoreCase);

        if (isPictureBox || isImageControl || isButton || !string.IsNullOrEmpty(ctrl.Picture) || ctrl.PictureData != null)
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "Picture",
                DisplayName = "Picture",
                Category = "Appearance",
                Description = "Devuelve o establece el gráfico que se va a mostrar en el control.",
                Type = DesignerPropertyType.Picture,
                CurrentValue = Resources.FrxResourceFile.GetFormatDescription(ctrl.PictureData, ctrl.Picture),
                ValueSetter = val =>
                {
                    if (string.IsNullOrWhiteSpace(val) || val.Equals("(None)", StringComparison.OrdinalIgnoreCase))
                    {
                        ctrl.Picture = null;
                        ctrl.PictureData = null;
                        onModelChanged();
                        return true;
                    }
                    if (val.StartsWith('(') && val.EndsWith(')'))
                    {
                        return true;
                    }
                    ctrl.Picture = val;
                    onModelChanged();
                    return true;
                }
            });
        }

        if (isPictureBox)
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "AutoSize",
                DisplayName = "AutoSize",
                Category = "Behavior",
                Description = "Determina si el control PictureBox se redimensiona automáticamente para ajustarse al tamaño de la imagen cargada.",
                Type = DesignerPropertyType.Boolean,
                CurrentValue = ctrl.AutoSize ? "True" : "False",
                Choices = BoolChoices,
                ValueSetter = val =>
                {
                    if (bool.TryParse(val, out bool b))
                    {
                        ctrl.AutoSize = b;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });
        }

        if (isButton)
        {
            list.Add(new DesignerPropertyItem
            {
                Name = "Style",
                DisplayName = "Style",
                Category = "Appearance",
                Description = "Devuelve o establece si el botón se muestra en estilo estándar (0) o gráfico (1) para soportar iconos e imágenes.",
                Type = DesignerPropertyType.EnumChoice,
                CurrentValue = ctrl.Style.ToString(),
                Choices = ButtonStyleChoices,
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int s))
                    {
                        ctrl.Style = s;
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });

            list.Add(new DesignerPropertyItem
            {
                Name = "PictureAlignment",
                DisplayName = "PictureAlignment",
                Category = "Appearance",
                Description = "Determina la alineación de la imagen en relación al texto del botón (0-Superior, 1-Inferior, 2-Izquierda, 3-Derecha, 4-Centro).",
                Type = DesignerPropertyType.EnumChoice,
                CurrentValue = ((int)ctrl.PictureAlignment).ToString(),
                Choices = PictureAlignmentChoices,
                ValueSetter = val =>
                {
                    if (int.TryParse(val, out int pa))
                    {
                        ctrl.PictureAlignment = (PictureAlignment)Math.Clamp(pa, 0, 4);
                        onModelChanged();
                        return true;
                    }
                    return false;
                }
            });
        }
        }

        // Extraer propiedades expuestas públicas de UserControls del proyecto
        var exposedProps = UserControlPropertiesResolver?.Invoke(ctrl.TypeName);
        if (exposedProps == null && ctrl.TypeName.Contains('.'))
        {
            string shortName = ctrl.TypeName.Substring(ctrl.TypeName.LastIndexOf('.') + 1);
            exposedProps = UserControlPropertiesResolver?.Invoke(shortName);
        }

        var handledPropNames = new HashSet<string>(list.Select(p => p.Name), StringComparer.OrdinalIgnoreCase);

        if (exposedProps != null && exposedProps.Count > 0)
        {
            foreach (var exp in exposedProps)
            {
                if (handledPropNames.Contains(exp.Name)) continue;
                handledPropNames.Add(exp.Name);

                var existingCustom = ctrl.CustomProperties.FirstOrDefault(p => string.Equals(p.Name, exp.Name, StringComparison.OrdinalIgnoreCase));
                string curVal = existingCustom != null ? existingCustom.Value : (exp.DefaultValue ?? string.Empty);

                var (inferredType, inferredChoices) = InferPropertyType(exp.Name, exp.TypeName, curVal);
                var finalType = exp.InferredType != DesignerPropertyType.Text ? exp.InferredType : inferredType;
                var finalChoices = exp.Choices.Count > 0 ? exp.Choices : (inferredChoices.Count > 0 ? inferredChoices : (finalType == DesignerPropertyType.Boolean ? BoolChoices : new List<DesignerPropertyChoice>()));

                if (exp.Name.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
                {
                    finalType = DesignerPropertyType.Text;
                    curVal = NormalizeIconCharCode(curVal);
                }

                if (finalType == DesignerPropertyType.FontDescription)
                {
                    var f = ctrl.GetFontProperty(exp.Name);
                    if (f != null && !string.IsNullOrEmpty(f.Name))
                    {
                        curVal = $"{f.Name}, {f.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt";
                    }
                    else if (exp.Name.Equals("Font", StringComparison.OrdinalIgnoreCase) && ctrl.Font?.Name != null)
                    {
                        curVal = $"{ctrl.Font.Name}, {ctrl.Font.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt";
                    }
                }

                list.Add(new DesignerPropertyItem
                {
                    Name = exp.Name,
                    DisplayName = exp.Name,
                    Category = exp.Category,
                    Description = !string.IsNullOrEmpty(exp.Description) ? exp.Description : $"Propiedad expuesta por el UserControl {ctrl.TypeName}.",
                    Type = finalType,
                    Choices = finalChoices,
                    CurrentValue = curVal,
                    IsReadOnly = finalType == DesignerPropertyType.FontDescription,
                    ValueSetter = val =>
                    {
                        if (exp.Name.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
                        {
                            val = NormalizeIconCharCode(val);
                        }
                        ctrl.SetCustomProperty(exp.Name, val);
                        SynchronizeStandardProperties(ctrl, exp.Name, val);
                        onModelChanged();
                        return true;
                    }
                });
            }
        }

        // Extraer propiedades COM / TypeLib registradas para controles ActiveX / OCX
        // Si el control ya fue resuelto como UserControl de proyecto (.ctl), sus propiedades
        // provienen exclusivamente de su código y no deben mezclarse con TypeLibs externas de OCX.
        if (exposedProps == null)
        {
            var typeLibMembers = IntelliSense.VbControlMemberCatalog.GetMembersForControl(ctrl.TypeName);
            if (typeLibMembers.Count > 0)
            {
                foreach (var mem in typeLibMembers)
                {
                    if (mem.Kind != IntelliSense.VbSymbolKind.Property) continue;
                    if (mem.IsHidden || !mem.IsBrowsable || mem.IsReadOnly) continue;
                    if (mem.Parameters.Any(p => !p.IsOptional)) continue;
                    if (handledPropNames.Contains(mem.Name)) continue;
                    handledPropNames.Add(mem.Name);

                    var existingCustom = ctrl.CustomProperties.FirstOrDefault(p => string.Equals(p.Name, mem.Name, StringComparison.OrdinalIgnoreCase));
                    string curVal = existingCustom != null ? existingCustom.Value : string.Empty;

                    var (propType, choices) = InferPropertyType(mem.Name, mem.ReturnType, curVal);

                    if (mem.Name.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
                    {
                        propType = DesignerPropertyType.Text;
                        curVal = NormalizeIconCharCode(curVal);
                    }

                    if (propType == DesignerPropertyType.FontDescription)
                    {
                        var f = ctrl.GetFontProperty(mem.Name);
                        if (f != null && !string.IsNullOrEmpty(f.Name))
                        {
                            curVal = $"{f.Name}, {f.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt";
                        }
                        else if (mem.Name.Equals("Font", StringComparison.OrdinalIgnoreCase) && ctrl.Font?.Name != null)
                        {
                            curVal = $"{ctrl.Font.Name}, {ctrl.Font.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt";
                        }
                    }

                    list.Add(new DesignerPropertyItem
                    {
                        Name = mem.Name,
                        DisplayName = mem.Name,
                        Category = "ActiveX / TypeLib",
                        Description = !string.IsNullOrEmpty(mem.Description) ? mem.Description : $"Propiedad COM '{mem.Name}' de {ctrl.TypeName}.",
                        Type = propType,
                        Choices = choices,
                        CurrentValue = curVal,
                        IsReadOnly = propType == DesignerPropertyType.FontDescription,
                        ValueSetter = val =>
                        {
                            if (mem.Name.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
                            {
                                val = NormalizeIconCharCode(val);
                            }
                            ctrl.SetCustomProperty(mem.Name, val);
                            SynchronizeStandardProperties(ctrl, mem.Name, val);
                            onModelChanged();
                            return true;
                        }
                    });
                }
            }
        }

        // Exponer propiedades adicionales personalizadas presentes en CustomProperties (de OCX o UserControls)
        foreach (var cust in ctrl.CustomProperties)
        {
            if (handledPropNames.Contains(cust.Name)) continue;
            if (cust.Name.Equals("_ExtentX", StringComparison.OrdinalIgnoreCase) ||
                cust.Name.Equals("_ExtentY", StringComparison.OrdinalIgnoreCase) ||
                cust.Name.Equals("_Version", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (cust.IsBeginProperty)
            {
                // Si es un bloque BeginProperty pero de fuente (ej. IconFont o Font guardado en el .frm), exponerlo con selector
                if (cust.Name.EndsWith("Font", StringComparison.OrdinalIgnoreCase) || cust.Name.Contains("Font", StringComparison.OrdinalIgnoreCase))
                {
                    handledPropNames.Add(cust.Name);
                    var f = ctrl.GetFontProperty(cust.Name);
                    string fontDescVal = f != null && !string.IsNullOrEmpty(f.Name)
                        ? $"{f.Name}, {f.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt"
                        : "MS Sans Serif, 8pt";

                    list.Add(new DesignerPropertyItem
                    {
                        Name = cust.Name,
                        DisplayName = cust.Name,
                        Category = "Font",
                        Description = $"Fuente tipográfica '{cust.Name}' del control {ctrl.Name}.",
                        Type = DesignerPropertyType.FontDescription,
                        CurrentValue = fontDescVal,
                        IsReadOnly = true,
                        ValueSetter = val =>
                        {
                            ctrl.SetCustomProperty(cust.Name, val);
                            SynchronizeStandardProperties(ctrl, cust.Name, val);
                            onModelChanged();
                            return true;
                        }
                    });
                }
                continue;
            }

            handledPropNames.Add(cust.Name);
            string pName = cust.Name;
            var (propType, choices) = InferPropertyType(pName, null, cust.Value);

            string curCustomVal = cust.Value;
            if (pName.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
            {
                propType = DesignerPropertyType.Text;
                curCustomVal = NormalizeIconCharCode(curCustomVal);
            }
            else if (propType == DesignerPropertyType.FontDescription)
            {
                var f = ctrl.GetFontProperty(pName);
                if (f != null && !string.IsNullOrEmpty(f.Name))
                {
                    curCustomVal = $"{f.Name}, {f.Size.ToString("0.##", CultureInfo.InvariantCulture)}pt";
                }
            }

            list.Add(new DesignerPropertyItem
            {
                Name = pName,
                DisplayName = pName,
                Category = "Custom / OLE",
                Description = $"Propiedad personalizada del control {ctrl.Name}.",
                Type = propType,
                Choices = choices,
                CurrentValue = curCustomVal,
                IsReadOnly = propType == DesignerPropertyType.FontDescription,
                ValueSetter = val =>
                {
                    if (pName.Contains("CharCode", StringComparison.OrdinalIgnoreCase))
                    {
                        val = NormalizeIconCharCode(val);
                    }
                    cust.Value = val;
                    ctrl.SetCustomProperty(pName, val);
                    SynchronizeStandardProperties(ctrl, pName, val);
                    onModelChanged();
                    return true;
                }
            });
        }

        // Si el control admite páginas de propiedades o es un UserControl/OCX de terceros, exponer (Custom) al inicio
        if (ctrl.HasPropertyPages || !ctrl.IsIntrinsic)
        {
            list.Insert(0, new DesignerPropertyItem
            {
                Name = "Custom",
                DisplayName = "(Custom)",
                Category = "General",
                Description = "Abre las Páginas de Propiedades avanzadas del fabricante para este control.",
                Type = DesignerPropertyType.CustomDialog,
                CurrentValue = "(Custom)",
                IsReadOnly = true
            });
        }

        return list;
    }
}
