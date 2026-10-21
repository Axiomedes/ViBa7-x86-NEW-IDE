using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ViBa7.ProjectModel.Designers;

/// <summary>
/// Serializador de modelos de diseño de formularios VB6 (FormDesignModel) a sintaxis .frm clásica.
/// Preserva sangrías estándar de VB6, alineación de propiedades y atributos de módulo.
/// </summary>
public static class FormDesignSerializer
{
    private const string PropertyPad = "                "; // 16 espacios

    /// <summary>
    /// Serializa el modelo de diseño a texto completo de cabecera visual de formulario.
    /// </summary>
    public static string Serialize(FormDesignModel model)
    {
        if (model == null) return string.Empty;

        var sb = new StringBuilder();

        // 1. Cabecera de versión
        if (!string.IsNullOrWhiteSpace(model.VersionHeader))
        {
            sb.AppendLine(model.VersionHeader);
        }
        else
        {
            sb.AppendLine("VERSION 5.00");
        }

        // 1b. Declaraciones Object = "{GUID}#Major.Minor#0"; "Archivo.ocx" (preservadas + deducidas de controles ActiveX activos)
        var objectLines = CollectObjectDeclarations(model);
        foreach (var objLine in objectLines)
        {
            sb.AppendLine(objLine);
        }

        // 2. Inicio del formulario raíz
        sb.AppendLine($"Begin {model.RootType} {model.Name} ");

        // 3. Propiedades del formulario (sangría 3 espacios)
        SerializeFormProperties(sb, model, indentLevel: 1);

        // 4. Controles hijos de primer nivel
        foreach (var child in model.Children)
        {
            SerializeControl(sb, child, indentLevel: 1);
        }

        // 5. Cierre del formulario raíz
        sb.AppendLine("End");

        // 6. Atributos de módulo posteriores (Attribute VB_*)
        if (model.Attributes.Count > 0)
        {
            foreach (var attr in model.Attributes)
            {
                // Si el nombre del formulario cambió, sincronizar Attribute VB_Name
                if (attr.StartsWith("Attribute VB_Name =", StringComparison.OrdinalIgnoreCase))
                {
                    sb.AppendLine($"Attribute VB_Name = \"{model.Name}\"");
                }
                else
                {
                    sb.AppendLine(attr);
                }
            }
        }
        else
        {
            // Atributos por defecto si no estaban presentes
            sb.AppendLine($"Attribute VB_Name = \"{model.Name}\"");
            sb.AppendLine("Attribute VB_GlobalNameSpace = False");
            sb.AppendLine("Attribute VB_Creatable = False");
            sb.AppendLine("Attribute VB_PredeclaredId = True");
            sb.AppendLine("Attribute VB_Exposed = False");
        }

        return sb.ToString();
    }

    /// <summary>
    /// Serializa un único control a sintaxis VB6.
    /// </summary>
    public static string SerializeControl(FormDesignerControl ctrl, int indentLevel = 0)
    {
        if (ctrl == null) return string.Empty;
        var sb = new StringBuilder();
        SerializeControl(sb, ctrl, indentLevel);
        return sb.ToString();
    }

    /// <summary>
    /// Serializa una lista de controles (por ejemplo para el portapapeles).
    /// </summary>
    public static string SerializeControls(IEnumerable<FormDesignerControl> controls)
    {
        if (controls == null) return string.Empty;
        var sb = new StringBuilder();
        foreach (var ctrl in controls)
        {
            SerializeControl(sb, ctrl, 0);
        }
        return sb.ToString();
    }

    private static void SerializeFormProperties(StringBuilder sb, FormDesignModel model, int indentLevel)
    {
        string indent = new string(' ', indentLevel * 3);

        // Mapeo de propiedades prioritarias en orden estándar de VB6
        var writtenProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void WriteProp(string name, string value, string? comment = null)
        {
            writtenProps.Add(name);
            string formattedComment = string.IsNullOrEmpty(comment) ? string.Empty : $"  {comment}";
            string propPadded = name.Length < 16 ? name.PadRight(16) : name + " ";
            sb.AppendLine($"{indent}{propPadded}=   {value}{formattedComment}");
        }

        if (!string.IsNullOrEmpty(model.Caption)) WriteProp("Caption", $"\"{model.Caption}\"");
        WriteProp("ClientHeight", model.ClientHeight.ToString(CultureInfo.InvariantCulture));
        WriteProp("ClientLeft", model.ClientLeft.ToString(CultureInfo.InvariantCulture));
        WriteProp("ClientTop", model.ClientTop.ToString(CultureInfo.InvariantCulture));
        WriteProp("ClientWidth", model.ClientWidth.ToString(CultureInfo.InvariantCulture));

        if (model.ScaleHeight.HasValue) WriteProp("ScaleHeight", model.ScaleHeight.Value.ToString(CultureInfo.InvariantCulture));
        if (model.ScaleWidth.HasValue) WriteProp("ScaleWidth", model.ScaleWidth.Value.ToString(CultureInfo.InvariantCulture));
        if (model.ScaleMode.HasValue) WriteProp("ScaleMode", model.ScaleMode.Value.ToString(CultureInfo.InvariantCulture));
        if (model.StartUpPosition.HasValue) WriteProp("StartUpPosition", model.StartUpPosition.Value.ToString(CultureInfo.InvariantCulture), "'Windows Default");
        if (model.BorderStyle.HasValue) WriteProp("BorderStyle", model.BorderStyle.Value.ToString(CultureInfo.InvariantCulture));

        if (!string.IsNullOrEmpty(model.BackColorHex)) WriteProp("BackColor", model.BackColorHex);
        if (!string.IsNullOrEmpty(model.ForeColorHex)) WriteProp("ForeColor", model.ForeColorHex);
        if (!string.IsNullOrEmpty(model.Picture)) WriteProp("Picture", model.Picture.StartsWith('"') ? model.Picture : $"\"{model.Picture}\"");

        // Bloque Font
        if (model.Font != null)
        {
            SerializeFont(sb, "Font", model.Font, indentLevel);
            writtenProps.Add("Font");
        }

        // Propiedades adicionales existentes en CustomProperties
        foreach (var prop in model.CustomProperties)
        {
            if (writtenProps.Contains(prop.Name)) continue;

            if (prop.IsBeginProperty)
            {
                foreach (var line in prop.RawPropertyLines)
                {
                    sb.AppendLine(line);
                }
            }
            else
            {
                WriteProp(prop.Name, prop.Value, prop.TrailingComment);
            }
        }
    }

    private static void SerializeControl(StringBuilder sb, FormDesignerControl ctrl, int indentLevel)
    {
        if (ctrl.IsMenu)
        {
            SerializeMenuControl(sb, ctrl, indentLevel);
            return;
        }

        string indent = new string(' ', indentLevel * 3);
        string propIndent = new string(' ', (indentLevel + 1) * 3);

        // Begin ControlType ControlName
        sb.AppendLine($"{indent}Begin {ctrl.TypeName} {ctrl.Name} ");

        var writtenProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void WriteProp(string name, string value, string? comment = null)
        {
            writtenProps.Add(name);
            string formattedComment = string.IsNullOrEmpty(comment) ? string.Empty : $"  {comment}";
            string propPadded = name.Length < 16 ? name.PadRight(16) : name + " ";
            sb.AppendLine($"{propIndent}{propPadded}=   {value}{formattedComment}");
        }

        // Propiedades canónicas
        if (ctrl.Index.HasValue) WriteProp("Index", ctrl.Index.Value.ToString(CultureInfo.InvariantCulture));
        if (ctrl.IsTimer || ctrl.Interval > 0) WriteProp("Interval", ctrl.Interval.ToString(CultureInfo.InvariantCulture));
        if (!string.IsNullOrEmpty(ctrl.Caption)) WriteProp("Caption", $"\"{ctrl.Caption}\"");
        if (!string.IsNullOrEmpty(ctrl.Text)) WriteProp("Text", $"\"{ctrl.Text}\"");

        WriteProp("Height", ctrl.Height.ToString(CultureInfo.InvariantCulture));
        WriteProp("Left", ctrl.Left.ToString(CultureInfo.InvariantCulture));

        if (ctrl.TabIndex.HasValue) WriteProp("TabIndex", ctrl.TabIndex.Value.ToString(CultureInfo.InvariantCulture));
        if (ctrl.TabStop.HasValue) WriteProp("TabStop", ctrl.TabStop.Value ? "-1  'True" : "0  'False");
        if (!ctrl.Enabled) WriteProp("Enabled", "0  'False");
        if (!ctrl.Visible) WriteProp("Visible", "0  'False");

        WriteProp("Top", ctrl.Top.ToString(CultureInfo.InvariantCulture));
        WriteProp("Width", ctrl.Width.ToString(CultureInfo.InvariantCulture));

        if (!string.IsNullOrEmpty(ctrl.BackColorHex)) WriteProp("BackColor", ctrl.BackColorHex);
        if (!string.IsNullOrEmpty(ctrl.ForeColorHex)) WriteProp("ForeColor", ctrl.ForeColorHex);
        if (ctrl.BorderStyle.HasValue) WriteProp("BorderStyle", ctrl.BorderStyle.Value.ToString(CultureInfo.InvariantCulture));
        if (ctrl.Alignment.HasValue) WriteProp("Alignment", ctrl.Alignment.Value.ToString(CultureInfo.InvariantCulture));
        if (ctrl.MultiLine.HasValue && ctrl.MultiLine.Value) WriteProp("MultiLine", "-1  'True");
        if (ctrl.ScrollBars.HasValue && ctrl.ScrollBars.Value > 0) WriteProp("ScrollBars", ctrl.ScrollBars.Value.ToString(CultureInfo.InvariantCulture));

        // Propiedades de imagen (PictureBox, CommandButton, etc.)
        if (!string.IsNullOrEmpty(ctrl.Picture)) WriteProp("Picture", ctrl.Picture.StartsWith('"') ? ctrl.Picture : $"\"{ctrl.Picture}\"");
        if (ctrl.Style != 0) WriteProp("Style", $"{ctrl.Style}  'Graphical");
        if (ctrl.AutoSize) WriteProp("AutoSize", "-1  'True");
        if (!string.IsNullOrEmpty(ctrl.DownPicture)) WriteProp("DownPicture", ctrl.DownPicture.StartsWith('"') ? ctrl.DownPicture : $"\"{ctrl.DownPicture}\"");
        if (!string.IsNullOrEmpty(ctrl.DisabledPicture)) WriteProp("DisabledPicture", ctrl.DisabledPicture.StartsWith('"') ? ctrl.DisabledPicture : $"\"{ctrl.DisabledPicture}\"");

        if (ctrl.PictureAlignment != PictureAlignment.Top)
        {
            sb.AppendLine($"{propIndent}' [VB7] PictureAlignment = {(int)ctrl.PictureAlignment}");
        }

        // Font si aplica (solo si es intrínseco o si fue declarado explícitamente)
        if (ctrl.Font != null && (ctrl.IsIntrinsic || ctrl.FontProperties.ContainsKey("Font") || ctrl.CustomProperties.Any(p => p.Name.Equals("Font", StringComparison.OrdinalIgnoreCase))))
        {
            SerializeFont(sb, "Font", ctrl.Font, indentLevel + 1);
            writtenProps.Add("Font");
        }

        // Otras fuentes nombradas (ej. IconFont, CaptionFont)
        foreach (var (fName, fDesc) in ctrl.FontProperties)
        {
            if (fName.Equals("Font", StringComparison.OrdinalIgnoreCase)) continue;
            if (fDesc != null)
            {
                SerializeFont(sb, fName, fDesc, indentLevel + 1);
                writtenProps.Add(fName);
            }
        }

        // Propiedades conservadas en CustomProperties que no hayan sido escritas
        foreach (var prop in ctrl.CustomProperties)
        {
            if (writtenProps.Contains(prop.Name)) continue;

            if (prop.IsBeginProperty)
            {
                foreach (var line in prop.RawPropertyLines)
                {
                    sb.AppendLine(line);
                }
            }
            else
            {
                string val = prop.Value;
                if (prop.Name.Equals("_ExtentX", StringComparison.OrdinalIgnoreCase))
                {
                    val = FormDesignerControl.TwipsToHiMetric(ctrl.Width).ToString(CultureInfo.InvariantCulture);
                }
                else if (prop.Name.Equals("_ExtentY", StringComparison.OrdinalIgnoreCase))
                {
                    val = FormDesignerControl.TwipsToHiMetric(ctrl.Height).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    val = FormatCustomPropertyValueForFrm(val, prop.TrailingComment);
                }
                WriteProp(prop.Name, val, prop.TrailingComment);
            }
        }

        // Si el control tenía ExtentX/ExtentY explícitos pero no estaban en CustomProperties, emitirlos
        if (ctrl.ExtentX.HasValue && !writtenProps.Contains("_ExtentX"))
        {
            WriteProp("_ExtentX", FormDesignerControl.TwipsToHiMetric(ctrl.Width).ToString(CultureInfo.InvariantCulture));
        }
        if (ctrl.ExtentY.HasValue && !writtenProps.Contains("_ExtentY"))
        {
            WriteProp("_ExtentY", FormDesignerControl.TwipsToHiMetric(ctrl.Height).ToString(CultureInfo.InvariantCulture));
        }

        // Controles hijos (si es contenedor, ej. Frame, PictureBox o Menú)
        foreach (var child in ctrl.Children)
        {
            SerializeControl(sb, child, indentLevel + 1);
        }

        // End del control
        sb.AppendLine($"{indent}End");
    }

    private static void SerializeFont(StringBuilder sb, string propertyName, FormFontDescription font, int indentLevel)
    {
        string indent = new string(' ', indentLevel * 3);
        string propIndent = new string(' ', (indentLevel + 1) * 3);

        sb.AppendLine($"{indent}BeginProperty {propertyName} ");
        sb.AppendLine($"{propIndent}{"Name".PadRight(16)}=   \"{font.Name}\"");
        sb.AppendLine($"{propIndent}{"Size".PadRight(16)}=   {font.Size.ToString("0.##", CultureInfo.InvariantCulture)}");
        sb.AppendLine($"{propIndent}{"Charset".PadRight(16)}=   {font.Charset}");
        sb.AppendLine($"{propIndent}{"Weight".PadRight(16)}=   {font.Weight}");
        sb.AppendLine($"{propIndent}{"Underline".PadRight(16)}=   {(font.Underline ? "-1  'True" : "0   'False")}");
        sb.AppendLine($"{propIndent}{"Italic".PadRight(16)}=   {(font.Italic ? "-1  'True" : "0   'False")}");
        sb.AppendLine($"{propIndent}{"Strikethrough".PadRight(16)}=   {(font.Strikethrough ? "-1  'True" : "0   'False")}");
        sb.AppendLine($"{indent}EndProperty");
    }

    private static void SerializeMenuControl(StringBuilder sb, FormDesignerControl ctrl, int indentLevel)
    {
        string indent = new string(' ', indentLevel * 3);
        string propIndent = new string(' ', (indentLevel + 1) * 3);

        sb.AppendLine($"{indent}Begin VB.Menu {ctrl.Name} ");

        var writtenProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void WriteProp(string name, string value, string? comment = null)
        {
            writtenProps.Add(name);
            string formattedComment = string.IsNullOrEmpty(comment) ? string.Empty : $"  {comment}";
            string propPadded = name.Length < 16 ? name.PadRight(16) : name + " ";
            sb.AppendLine($"{propIndent}{propPadded}=   {value}{formattedComment}");
        }

        if (!string.IsNullOrEmpty(ctrl.Caption)) WriteProp("Caption", $"\"{ctrl.Caption}\"");
        if (ctrl.Index.HasValue) WriteProp("Index", ctrl.Index.Value.ToString(CultureInfo.InvariantCulture));
        if (!string.IsNullOrEmpty(ctrl.Shortcut)) WriteProp("Shortcut", ctrl.Shortcut);
        if (ctrl.Checked) WriteProp("Checked", "-1  'True");
        if (!ctrl.Enabled) WriteProp("Enabled", "0  'False");
        if (!ctrl.Visible) WriteProp("Visible", "0  'False");
        if (ctrl.WindowList) WriteProp("WindowList", "-1  'True");
        if (ctrl.NegotiatePosition > 0) WriteProp("NegotiatePosition", ctrl.NegotiatePosition.ToString(CultureInfo.InvariantCulture));
        if (ctrl.HelpContextID > 0) WriteProp("HelpContextID", ctrl.HelpContextID.ToString(CultureInfo.InvariantCulture));

        // Propiedades conservadas en CustomProperties que no hayan sido escritas
        foreach (var prop in ctrl.CustomProperties)
        {
            if (writtenProps.Contains(prop.Name)) continue;
            if (prop.IsBeginProperty)
            {
                foreach (var line in prop.RawPropertyLines)
                {
                    sb.AppendLine(line);
                }
            }
            else
            {
                WriteProp(prop.Name, prop.Value, prop.TrailingComment);
            }
        }

        // Submenús anidados
        foreach (var child in ctrl.Children)
        {
            SerializeControl(sb, child, indentLevel + 1);
        }

        sb.AppendLine($"{indent}End");
    }

    private static List<string> CollectObjectDeclarations(FormDesignModel model)
    {
        var result = new List<string>();
        foreach (var existing in model.ObjectDeclarations)
        {
            string trimmed = existing.Trim();
            if (!string.IsNullOrEmpty(trimmed) &&
                !result.Any(r => string.Equals(r, trimmed, StringComparison.OrdinalIgnoreCase)))
            {
                result.Add(trimmed);
            }
        }

        foreach (var ctrl in model.GetAllControls())
        {
            if (ctrl.IsIntrinsic) continue;

            if (IntelliSense.VbControlMemberCatalog.TryGetOcxObjectHeaderForControl(ctrl.TypeName, out string headerLine) &&
                !string.IsNullOrWhiteSpace(headerLine))
            {
                // Verificar si ya existe una declaración para la misma GUID o el mismo archivo .ocx
                int openBrace = headerLine.IndexOf('{');
                int closeBrace = headerLine.IndexOf('}');
                string? guid = (openBrace >= 0 && closeBrace > openBrace)
                    ? headerLine.Substring(openBrace, closeBrace - openBrace + 1)
                    : null;

                int semiIdx = headerLine.IndexOf(';');
                string? filePart = semiIdx >= 0 ? headerLine[(semiIdx + 1)..].Trim().Trim('"') : null;

                bool alreadyExists = result.Any(r =>
                    string.Equals(r, headerLine, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrEmpty(guid) && r.Contains(guid, StringComparison.OrdinalIgnoreCase)) ||
                    (!string.IsNullOrEmpty(filePart) && r.EndsWith($"\"{filePart}\"", StringComparison.OrdinalIgnoreCase)));

                if (!alreadyExists)
                {
                    result.Add(headerLine);
                }
            }
        }

        return result;
    }

    private static string FormatCustomPropertyValueForFrm(string rawVal, string? trailingComment)
    {
        if (string.IsNullOrEmpty(rawVal)) return "\"\"";
        string trimmed = rawVal.Trim();

        // Si ya está entrecomillado o es una referencia a archivo binario ("Form1.frx":0000)
        if (trimmed.StartsWith('"'))
        {
            return trimmed;
        }

        // Si es valor hexadecimal (&H...)
        if (trimmed.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
        {
            return trimmed;
        }

        // Si es booleano textual puro (True / False) proveniente de reflexión IDispatch
        if (trimmed.Equals("True", StringComparison.OrdinalIgnoreCase))
        {
            return string.IsNullOrEmpty(trailingComment) ? "-1  'True" : "-1";
        }
        if (trimmed.Equals("False", StringComparison.OrdinalIgnoreCase))
        {
            return string.IsNullOrEmpty(trailingComment) ? "0   'False" : "0";
        }

        // Si es número entero o punto flotante
        if (long.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) ||
            double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return trimmed;
        }

        // Si contiene referencia a .frx/.ctx sin comillas iniciales
        if (Resources.FrxResourceFile.ParseReference(trimmed, out string refFile, out int refOffset) && !string.IsNullOrEmpty(refFile))
        {
            return Resources.FrxResourceFile.FormatReference(refFile, refOffset);
        }

        // Cualquier otro texto literal debe ir entre comillas en sintaxis .frm de VB6
        return $"\"{trimmed}\"";
    }
}
