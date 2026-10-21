using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ViBa7.ProjectModel.Designers;

/// <summary>
/// Parser de cabeceras visuales de formularios VB6 (.frm, .ctl, .pag).
/// Reconstruye el modelo jerárquico completo en memoria preservando cada propiedad original.
/// </summary>
public static class FormDesignParser
{
    private static readonly Regex BeginBlockRegex = new(
        @"^\s*Begin\s+([^\s]+)\s+([^\s]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex EndBlockRegex = new(
        @"^\s*End\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex BeginPropertyRegex = new(
        @"^\s*BeginProperty\s+([^\s]+)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex EndPropertyRegex = new(
        @"^\s*EndProperty\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PropertyAssignmentRegex = new(
        @"^\s*([a-zA-Z0-9_.\(\)]+)\s*=\s*(.*)$",
        RegexOptions.Compiled);

    /// <summary>
    /// Validador opcional para comprobar si un tipo de control no intrínseco está registrado o resuelto en el proyecto.
    /// </summary>
    public static Func<string, bool>? ControlResolutionChecker { get; set; }

    /// <summary>
    /// Parsea un fragmento de controles (por ejemplo pegados desde el portapapeles) y retorna la lista de controles.
    /// </summary>
    public static List<FormDesignerControl> ParseControlSnippet(string snippetText)
    {
        if (string.IsNullOrWhiteSpace(snippetText)) return new List<FormDesignerControl>();

        string trimmed = snippetText.TrimStart();
        string wrappedText;
        if (trimmed.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Begin VB.Form", StringComparison.OrdinalIgnoreCase) ||
            trimmed.StartsWith("Begin VB.UserControl", StringComparison.OrdinalIgnoreCase))
        {
            wrappedText = snippetText;
        }
        else
        {
            wrappedText = $"VERSION 5.00\r\nBegin VB.Form DummyHost\r\n{snippetText}\r\nEnd\r\n";
        }

        var model = Parse(wrappedText);
        return model.Children;
    }

    /// <summary>
    /// Parsea el texto de la cabecera visual de un formulario VB6 y retorna el modelo de diseño.
    /// </summary>
    public static FormDesignModel Parse(string headerText)
    {
        var model = new FormDesignModel();
        if (string.IsNullOrWhiteSpace(headerText))
        {
            return model;
        }

        string[] lines = headerText.Replace("\r\n", "\n").Split('\n');
        int lineIdx = 0;

        // 1. Extraer versión inicial (ej. VERSION 5.00) y declaraciones Object = "{GUID}#ver#0"; "file.ocx"
        while (lineIdx < lines.Length)
        {
            string trimmed = lines[lineIdx].Trim();
            if (trimmed.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
            {
                model.VersionHeader = trimmed;
                lineIdx++;
                continue;
            }

            if (trimmed.StartsWith("Object", StringComparison.OrdinalIgnoreCase) && trimmed.Contains('='))
            {
                if (!model.ObjectDeclarations.Any(o => string.Equals(o, trimmed, StringComparison.OrdinalIgnoreCase)))
                {
                    model.ObjectDeclarations.Add(trimmed);
                }
                TryRegisterFormObjectReference(trimmed);
                lineIdx++;
                continue;
            }

            if (trimmed.StartsWith("Begin ", StringComparison.OrdinalIgnoreCase))
            {
                break;
            }

            lineIdx++;
        }

        // Pila de control jerárquico
        var controlStack = new Stack<FormDesignerControl>();
        bool insideRootForm = false;
        bool rootFormEnded = false;

        // Estado de BeginProperty (con soporte de anidamiento para Toolbar Buttons, ListImages, ColumnHeaders, etc.)
        int beginPropertyDepth = 0;
        string? currentPropertyName = null;
        var currentPropertyLines = new List<string>();

        for (; lineIdx < lines.Length; lineIdx++)
        {
            string line = lines[lineIdx];
            string trimmed = line.Trim();

            if (string.IsNullOrWhiteSpace(trimmed)) continue;

            // Después del 'End' del formulario raíz, recopilar atributos de módulo (Attribute VB_*)
            if (rootFormEnded)
            {
                if (trimmed.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))
                {
                    model.Attributes.Add(line.TrimEnd());
                }
                continue;
            }

            // Manejo de bloques BeginProperty ... EndProperty (incluyendo bloques anidados)
            if (beginPropertyDepth > 0)
            {
                if (controlStack.Count > 0)
                {
                    controlStack.Peek().RawControlLines.Add(line);
                }

                if (BeginPropertyRegex.IsMatch(trimmed))
                {
                    beginPropertyDepth++;
                    currentPropertyLines.Add(line);
                    continue;
                }

                if (EndPropertyRegex.IsMatch(trimmed))
                {
                    beginPropertyDepth--;
                    currentPropertyLines.Add(line);

                    if (beginPropertyDepth == 0)
                    {
                        // Procesar la propiedad de nivel superior completada (especialmente Font e IconFont)
                        if (currentPropertyName != null && (string.Equals(currentPropertyName, "Font", StringComparison.OrdinalIgnoreCase) ||
                                                             currentPropertyName.EndsWith("Font", StringComparison.OrdinalIgnoreCase) ||
                                                             currentPropertyName.Contains("Font", StringComparison.OrdinalIgnoreCase)))
                        {
                            var fontDesc = ParseFontBlock(currentPropertyLines);
                            if (controlStack.Count > 0)
                            {
                                var ctrl = controlStack.Peek();
                                ctrl.SetFontProperty(currentPropertyName, fontDesc);
                                if (string.Equals(currentPropertyName, "Font", StringComparison.OrdinalIgnoreCase))
                                {
                                    ctrl.Font = fontDesc;
                                }
                            }
                            else if (insideRootForm)
                            {
                                if (string.Equals(currentPropertyName, "Font", StringComparison.OrdinalIgnoreCase))
                                {
                                    model.Font = fontDesc;
                                }
                            }
                        }

                        var propItem = new FormPropertyItem
                        {
                            Name = currentPropertyName ?? "Property",
                            IsBeginProperty = true,
                            RawPropertyLines = new List<string>(currentPropertyLines)
                        };

                        if (controlStack.Count > 0)
                        {
                            controlStack.Peek().CustomProperties.Add(propItem);
                        }
                        else if (insideRootForm)
                        {
                            model.CustomProperties.Add(propItem);
                        }

                        currentPropertyName = null;
                        currentPropertyLines.Clear();
                    }
                    continue;
                }

                currentPropertyLines.Add(line);
                continue;
            }

            var beginPropertyMatch = BeginPropertyRegex.Match(trimmed);
            if (beginPropertyMatch.Success)
            {
                beginPropertyDepth = 1;
                currentPropertyName = beginPropertyMatch.Groups[1].Value.Trim();
                currentPropertyLines.Clear();
                currentPropertyLines.Add(line);
                if (controlStack.Count > 0)
                {
                    controlStack.Peek().RawControlLines.Add(line);
                }
                continue;
            }

            // Comprobación de Begin de control o formulario
            var beginMatch = BeginBlockRegex.Match(trimmed);
            if (beginMatch.Success)
            {
                string typeName = beginMatch.Groups[1].Value.Trim();
                string ctrlName = beginMatch.Groups[2].Value.Trim();

                if (!insideRootForm)
                {
                    // Formulario raíz
                    insideRootForm = true;
                    model.RootType = typeName;
                    model.Name = ctrlName;
                    continue;
                }

                // Control hijo (de primer nivel o anidado en un contenedor existente)
                var newCtrl = new FormDesignerControl
                {
                    TypeName = typeName,
                    Name = ctrlName
                };

                // Verificación de resolución de controles de terceros / OCX (Aporte 4)
                bool isIntrinsic = newCtrl.IsIntrinsic;
                bool isResolved = isIntrinsic;
                if (!isIntrinsic)
                {
                    if (ControlResolutionChecker != null)
                    {
                        isResolved = ControlResolutionChecker(typeName);
                    }
                    else
                    {
                        isResolved = IntelliSense.VbControlMemberCatalog.IsControlRegistered(typeName);
                    }
                }

                if (!isResolved)
                {
                    newCtrl.IsUnresolved = true;
                    newCtrl.UnresolvedErrorMessage = $"El componente o biblioteca '{typeName}' no está registrado en el sistema. Se preservan intactas todas sus propiedades y estado binario.";
                }

                newCtrl.RawControlLines.Add(line);

                if (controlStack.Count > 0)
                {
                    // Control anidado dentro de un contenedor (ej. Frame, PictureBox)
                    var parent = controlStack.Peek();
                    newCtrl.ParentControl = parent;
                    parent.Children.Add(newCtrl);
                }
                else
                {
                    // Control de primer nivel en el formulario
                    model.Children.Add(newCtrl);
                }

                controlStack.Push(newCtrl);
                continue;
            }

            // Comprobación de End de control o formulario
            if (EndBlockRegex.IsMatch(trimmed))
            {
                if (controlStack.Count > 0)
                {
                    var endedCtrl = controlStack.Pop();
                    endedCtrl.RawControlLines.Add(line);
                }
                else if (insideRootForm)
                {
                    insideRootForm = false;
                    rootFormEnded = true;
                }
                continue;
            }

            // Comentarios dentro del bloque del control (ej. [VB7] PictureAlignment = ...)
            if (trimmed.StartsWith('\''))
            {
                if (controlStack.Count > 0)
                {
                    var activeCtrl = controlStack.Peek();
                    var matchAlign = Regex.Match(trimmed, @"PictureAlignment\s*=\s*([0-4]|Top|Bottom|Left|Right|Center)", RegexOptions.IgnoreCase);
                    if (matchAlign.Success)
                    {
                        string aVal = matchAlign.Groups[1].Value;
                        if (int.TryParse(aVal, out int aNum))
                        {
                            activeCtrl.PictureAlignment = (PictureAlignment)Math.Clamp(aNum, 0, 4);
                        }
                        else if (Enum.TryParse<PictureAlignment>(aVal, true, out var paEnum))
                        {
                            activeCtrl.PictureAlignment = paEnum;
                        }
                    }
                }
                continue;
            }

            // Asignación de propiedad: Property = Value ['Comment]
            var propMatch = PropertyAssignmentRegex.Match(trimmed);
            if (propMatch.Success)
            {
                string propName = propMatch.Groups[1].Value.Trim();
                string rhs = propMatch.Groups[2].Value;

                var (valStr, commentStr) = SplitPropertyAndComment(rhs);
                string cleanVal = Unquote(valStr);

                var propItem = new FormPropertyItem(propName, valStr, commentStr);

                if (controlStack.Count > 0)
                {
                    var activeCtrl = controlStack.Peek();
                    activeCtrl.RawControlLines.Add(line);
                    ApplyControlProperty(activeCtrl, propName, cleanVal, valStr);
                    activeCtrl.CustomProperties.Add(propItem);
                }
                else if (insideRootForm)
                {
                    ApplyFormProperty(model, propName, cleanVal, valStr);
                    model.CustomProperties.Add(propItem);
                }
            }
        }

        return model;
    }

    private static void ApplyFormProperty(FormDesignModel model, string name, string cleanVal, string rawVal)
    {
        switch (name.ToUpperInvariant())
        {
            case "CAPTION":
                model.Caption = cleanVal;
                break;
            case "CLIENTWIDTH":
                if (int.TryParse(cleanVal, out int cw)) model.ClientWidth = cw;
                break;
            case "CLIENTHEIGHT":
                if (int.TryParse(cleanVal, out int ch)) model.ClientHeight = ch;
                break;
            case "CLIENTLEFT":
                if (int.TryParse(cleanVal, out int cl)) model.ClientLeft = cl;
                break;
            case "CLIENTTOP":
                if (int.TryParse(cleanVal, out int ct)) model.ClientTop = ct;
                break;
            case "WIDTH":
                if (int.TryParse(cleanVal, out int w)) model.ClientWidth = w;
                break;
            case "HEIGHT":
                if (int.TryParse(cleanVal, out int h)) model.ClientHeight = h;
                break;
            case "SCALEWIDTH":
                if (int.TryParse(cleanVal, out int sw))
                {
                    model.ScaleWidth = sw;
                    if (model.ClientWidth <= 0 || model.ClientWidth == 4680) model.ClientWidth = sw;
                }
                break;
            case "SCALEHEIGHT":
                if (int.TryParse(cleanVal, out int sh))
                {
                    model.ScaleHeight = sh;
                    if (model.ClientHeight <= 0 || model.ClientHeight == 3195) model.ClientHeight = sh;
                }
                break;
            case "_EXTENTX":
                if (int.TryParse(cleanVal, out int ex) && ex > 0)
                {
                    int tw = FormDesignerControl.HiMetricToTwips(ex);
                    if (model.ClientWidth <= 0 || model.ClientWidth == 4680) model.ClientWidth = tw;
                }
                break;
            case "_EXTENTY":
                if (int.TryParse(cleanVal, out int ey) && ey > 0)
                {
                    int th = FormDesignerControl.HiMetricToTwips(ey);
                    if (model.ClientHeight <= 0 || model.ClientHeight == 3195) model.ClientHeight = th;
                }
                break;
            case "SCALEMODE":
                if (int.TryParse(cleanVal, out int sm)) model.ScaleMode = sm;
                break;
            case "STARTUPPOSITION":
                if (int.TryParse(cleanVal, out int sup)) model.StartUpPosition = sup;
                break;
            case "BORDERSTYLE":
                if (int.TryParse(cleanVal, out int bs)) model.BorderStyle = bs;
                break;
            case "BACKCOLOR":
                model.BackColorHex = rawVal.Trim();
                break;
            case "FORECOLOR":
                model.ForeColorHex = rawVal.Trim();
                break;
            case "PICTURE":
                model.Picture = cleanVal;
                break;
            case "TOOLBOXBITMAP":
                model.ToolboxBitmap = cleanVal;
                break;
            case "CONTROLCONTAINER":
                model.ControlContainer = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                         cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                         cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
        }
    }

    private static void ApplyControlProperty(FormDesignerControl ctrl, string name, string cleanVal, string rawVal)
    {
        switch (name.ToUpperInvariant())
        {
            case "CAPTION":
                ctrl.Caption = cleanVal;
                break;
            case "TEXT":
                ctrl.Text = cleanVal;
                break;
            case "LEFT":
                if (int.TryParse(cleanVal, out int left)) ctrl.Left = left;
                break;
            case "TOP":
                if (int.TryParse(cleanVal, out int top)) ctrl.Top = top;
                break;
            case "WIDTH":
                if (int.TryParse(cleanVal, out int width))
                {
                    ctrl.Width = width;
                    ctrl.HasExplicitWidth = true;
                }
                break;
            case "HEIGHT":
                if (int.TryParse(cleanVal, out int height))
                {
                    ctrl.Height = height;
                    ctrl.HasExplicitHeight = true;
                }
                break;
            case "_EXTENTX":
                if (int.TryParse(cleanVal, out int ex) && ex > 0)
                {
                    ctrl.ExtentX = ex;
                    if (!ctrl.HasExplicitWidth || ctrl.Width == 1500)
                    {
                        ctrl.Width = FormDesignerControl.HiMetricToTwips(ex);
                        ctrl.HasExplicitWidth = true;
                    }
                }
                break;
            case "_EXTENTY":
                if (int.TryParse(cleanVal, out int ey) && ey > 0)
                {
                    ctrl.ExtentY = ey;
                    if (!ctrl.HasExplicitHeight || ctrl.Height == 495)
                    {
                        ctrl.Height = FormDesignerControl.HiMetricToTwips(ey);
                        ctrl.HasExplicitHeight = true;
                    }
                }
                break;
            case "TABINDEX":
                if (int.TryParse(cleanVal, out int tidx)) ctrl.TabIndex = tidx;
                break;
            case "TABSTOP":
                ctrl.TabStop = !cleanVal.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                               !cleanVal.Equals("False", StringComparison.OrdinalIgnoreCase);
                break;
            case "ENABLED":
                ctrl.Enabled = !cleanVal.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                               !cleanVal.Equals("False", StringComparison.OrdinalIgnoreCase);
                break;
            case "VISIBLE":
                ctrl.Visible = !cleanVal.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                               !cleanVal.Equals("False", StringComparison.OrdinalIgnoreCase);
                break;
            case "CHECKED":
                ctrl.Checked = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                               cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                               cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
            case "SHORTCUT":
                ctrl.Shortcut = cleanVal;
                break;
            case "WINDOWLIST":
                ctrl.WindowList = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                  cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                  cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
            case "NEGOTIATEPOSITION":
                if (int.TryParse(cleanVal, out int np)) ctrl.NegotiatePosition = np;
                break;
            case "HELPCONTEXTID":
                if (int.TryParse(cleanVal, out int hcid)) ctrl.HelpContextID = hcid;
                break;
            case "INDEX":
                if (int.TryParse(cleanVal, out int idx)) ctrl.Index = idx;
                break;
            case "INTERVAL":
                if (int.TryParse(cleanVal, out int iv)) ctrl.Interval = iv;
                break;
            case "BACKCOLOR":
                ctrl.BackColorHex = rawVal.Trim();
                break;
            case "FORECOLOR":
                ctrl.ForeColorHex = rawVal.Trim();
                break;
            case "BORDERSTYLE":
                if (int.TryParse(cleanVal, out int bs)) ctrl.BorderStyle = bs;
                break;
            case "ALIGNMENT":
                if (int.TryParse(cleanVal, out int al)) ctrl.Alignment = al;
                break;
            case "MULTILINE":
                ctrl.MultiLine = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                 cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                 cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
            case "SCROLLBARS":
                if (int.TryParse(cleanVal, out int sb)) ctrl.ScrollBars = sb;
                break;
            case "PICTURE":
                ctrl.Picture = cleanVal;
                break;
            case "DOWNPICTURE":
                ctrl.DownPicture = cleanVal;
                break;
            case "DISABLEDPICTURE":
                ctrl.DisabledPicture = cleanVal;
                break;
            case "STYLE":
                if (int.TryParse(cleanVal, out int st)) ctrl.Style = st;
                break;
            case "AUTOSIZE":
                ctrl.AutoSize = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
            case "PICTUREALIGNMENT":
            case "PICTUREPOSITION":
                if (int.TryParse(cleanVal, out int pa))
                {
                    ctrl.PictureAlignment = (PictureAlignment)Math.Clamp(pa, 0, 4);
                }
                else if (Enum.TryParse<PictureAlignment>(cleanVal, true, out var paEnum))
                {
                    ctrl.PictureAlignment = paEnum;
                }
                break;
            case "CONTROLCONTAINER":
                ctrl.ControlContainer = cleanVal.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                        cleanVal.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                        cleanVal.Equals("True", StringComparison.OrdinalIgnoreCase);
                break;
            default:
                if ((name.EndsWith("Font", StringComparison.OrdinalIgnoreCase) || name.Contains("Font", StringComparison.OrdinalIgnoreCase)) &&
                    !string.IsNullOrWhiteSpace(cleanVal))
                {
                    var existingF = ctrl.GetFontProperty(name) ?? new FormFontDescription();
                    existingF.Name = cleanVal;
                    ctrl.SetFontProperty(name, existingF);
                }
                break;
        }
    }

    public static FormFontDescription ParseFontBlock(List<string> lines)
    {
        var font = new FormFontDescription();

        foreach (var line in lines)
        {
            var match = PropertyAssignmentRegex.Match(line.Trim());
            if (match.Success)
            {
                string prop = match.Groups[1].Value.Trim().ToUpperInvariant();
                var (val, _) = SplitPropertyAndComment(match.Groups[2].Value);
                string clean = Unquote(val);

                switch (prop)
                {
                    case "NAME":
                        font.Name = clean;
                        break;
                    case "SIZE":
                        if (double.TryParse(clean, NumberStyles.Float, CultureInfo.InvariantCulture, out double sz))
                        {
                            font.Size = sz;
                        }
                        break;
                    case "CHARSET":
                        if (int.TryParse(clean, out int cs)) font.Charset = cs;
                        break;
                    case "WEIGHT":
                        if (int.TryParse(clean, out int wt)) font.Weight = wt;
                        break;
                    case "UNDERLINE":
                        font.Underline = clean.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                         clean.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                         clean.Equals("True", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "ITALIC":
                        font.Italic = clean.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                      clean.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                      clean.Equals("True", StringComparison.OrdinalIgnoreCase);
                        break;
                    case "STRIKETHROUGH":
                        font.Strikethrough = clean.Equals("-1", StringComparison.OrdinalIgnoreCase) ||
                                             clean.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                                             clean.Equals("True", StringComparison.OrdinalIgnoreCase);
                        break;
                }
            }
        }

        return font;
    }

    public static (string Value, string? Comment) SplitPropertyAndComment(string rhs)
    {
        bool inQuotes = false;
        for (int i = 0; i < rhs.Length; i++)
        {
            if (rhs[i] == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (rhs[i] == '\'' && !inQuotes)
            {
                string val = rhs.Substring(0, i).Trim();
                string comment = rhs.Substring(i).Trim();
                return (val, comment);
            }
        }

        return (rhs.Trim(), null);
    }

    public static string Unquote(string str)
    {
        if (str.StartsWith("\"") && str.EndsWith("\"") && str.Length >= 2)
        {
            return str.Substring(1, str.Length - 2);
        }
        return str;
    }

    /// <summary>
    /// Resuelve una referencia binaria OLE StdPicture (.frx, .ctx o archivo en disco) a su carga útil en bytes.
    /// </summary>
    public static byte[]? ResolveBinaryResource(string? reference, string basePath, bool isUserControl, string modelName)
    {
        if (string.IsNullOrWhiteSpace(reference) || string.IsNullOrWhiteSpace(basePath))
        {
            return null;
        }

        if (Resources.FrxResourceFile.ParseReference(reference, out string resFileName, out int offset))
        {
            string defaultExt = isUserControl ? $"{modelName}.ctx" : $"{modelName}.frx";
            string targetRes = string.IsNullOrEmpty(resFileName) ? defaultExt : resFileName;
            string resPath = Path.Combine(basePath, targetRes);

            if (!File.Exists(resPath))
            {
                // Alternativa .ctx <-> .frx
                string altExt = isUserControl ? $"{modelName}.ctx" : $"{modelName}.frx";
                string altPath = Path.Combine(basePath, altExt);
                if (File.Exists(altPath))
                {
                    resPath = altPath;
                }
                else
                {
                    try
                    {
                        string? match = Directory.GetFiles(basePath)
                            .FirstOrDefault(f => string.Equals(Path.GetFileName(f), targetRes, StringComparison.OrdinalIgnoreCase) ||
                                                 string.Equals(Path.GetFileName(f), altExt, StringComparison.OrdinalIgnoreCase));
                        if (match != null) resPath = match;
                    }
                    catch { }
                }
            }

            if (File.Exists(resPath))
            {
                return Resources.FrxResourceFile.ReadResource(resPath, offset);
            }
        }
        else
        {
            string localPath = Path.IsPathRooted(reference)
                ? reference
                : Path.Combine(basePath, reference);

            if (File.Exists(localPath))
            {
                try { return File.ReadAllBytes(localPath); } catch { }
            }
        }

        return null;
    }

    /// <summary>
    /// Resuelve un bloque binario arbitrario (.frx / .ctx) para propiedades de estado ActiveX o PropertyBag binario.
    /// </summary>
    public static byte[]? ResolveRawBinaryBlock(string? reference, string basePath, bool isUserControl, string modelName, byte[]? cachedFrxBytes = null)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return null;
        }

        if (!Resources.FrxResourceFile.ParseReference(reference, out string resFileName, out int offset))
        {
            return null;
        }

        if (cachedFrxBytes != null && cachedFrxBytes.Length > 0)
        {
            return Resources.FrxResourceFile.ReadRawBlock(cachedFrxBytes, offset);
        }

        if (string.IsNullOrWhiteSpace(basePath))
        {
            return null;
        }

        string defaultExt = isUserControl ? $"{modelName}.ctx" : $"{modelName}.frx";
        string targetRes = string.IsNullOrEmpty(resFileName) ? defaultExt : resFileName;
        string resPath = Path.Combine(basePath, targetRes);

        if (!File.Exists(resPath))
        {
            string altPath = Path.Combine(basePath, defaultExt);
            if (File.Exists(altPath))
            {
                resPath = altPath;
            }
            else
            {
                try
                {
                    string? match = Directory.GetFiles(basePath)
                        .FirstOrDefault(f => string.Equals(Path.GetFileName(f), targetRes, StringComparison.OrdinalIgnoreCase) ||
                                             string.Equals(Path.GetFileName(f), defaultExt, StringComparison.OrdinalIgnoreCase));
                    if (match != null) resPath = match;
                }
                catch { }
            }
        }

        return File.Exists(resPath) ? Resources.FrxResourceFile.ReadRawBlock(resPath, offset) : null;
    }

    /// <summary>
    /// Resuelve y carga en memoria las imágenes binarias y streams de estado ActiveX (.frx, .ctx o rutas relativas)
    /// asociados al formulario o UserControl y sus controles.
    /// </summary>
    public static void ResolveControlPictures(FormDesignModel model, string basePath)
    {
        if (model == null || string.IsNullOrWhiteSpace(basePath)) return;

        // 0. Preservar en memoria los bytes originales del archivo .frx / .ctx acompañante para garantizar round-trip sin desfasar offsets
        if (model.OriginalFrxBytes == null)
        {
            try
            {
                string defaultCompanion = model.IsUserControl ? $"{model.Name}.ctx" : $"{model.Name}.frx";
                string companionPath = Path.Combine(basePath, defaultCompanion);
                if (!File.Exists(companionPath) && Directory.Exists(basePath))
                {
                    string? match = Directory.GetFiles(basePath)
                        .FirstOrDefault(f => string.Equals(Path.GetFileName(f), defaultCompanion, StringComparison.OrdinalIgnoreCase));
                    if (match != null) companionPath = match;
                }

                if (File.Exists(companionPath))
                {
                    model.OriginalFrxBytes = File.ReadAllBytes(companionPath);
                }
            }
            catch { }
        }

        // 1. Imagen del formulario o UserControl (fondos, texturas y skins de controles como JColCombo)
        if (model.PictureData == null && !string.IsNullOrEmpty(model.Picture))
        {
            model.PictureData = ResolveBinaryResource(model.Picture, basePath, model.IsUserControl, model.Name);
        }

        // 1b. Icono ToolboxBitmap del UserControl (.ctx)
        if (model.ToolboxBitmapData == null && !string.IsNullOrEmpty(model.ToolboxBitmap))
        {
            model.ToolboxBitmapData = ResolveBinaryResource(model.ToolboxBitmap, basePath, model.IsUserControl, model.Name);
        }

        // 2. Imágenes de los controles (Picture, DownPicture, DisabledPicture y skins en propiedades custom)
        string[] candidateSkinProps = ["PICTURE", "DOWNPICTURE", "DISABLEDPICTURE", "IMAGE", "SKIN", "TEXTURE", "MASKPICTURE", "BACKGROUNDPICTURE", "BACKPICTURE", "NORMALPICTURE", "OVERPICTURE"];

        foreach (var ctrl in model.GetAllControls())
        {
            if (ctrl.PictureData == null && !string.IsNullOrEmpty(ctrl.Picture))
            {
                ctrl.PictureData = ResolveBinaryResource(ctrl.Picture, basePath, model.IsUserControl, model.Name);
            }
            if (ctrl.DownPictureData == null && !string.IsNullOrEmpty(ctrl.DownPicture))
            {
                ctrl.DownPictureData = ResolveBinaryResource(ctrl.DownPicture, basePath, model.IsUserControl, model.Name);
            }
            if (ctrl.DisabledPictureData == null && !string.IsNullOrEmpty(ctrl.DisabledPicture))
            {
                ctrl.DisabledPictureData = ResolveBinaryResource(ctrl.DisabledPicture, basePath, model.IsUserControl, model.Name);
            }

            // 3. Resolver propiedades binarias y estado de persistencia de instancia ActiveX (ej. Object.zyx = "Form1.frx":0000)
            foreach (var prop in ctrl.CustomProperties)
            {
                if (prop.IsBeginProperty || string.IsNullOrWhiteSpace(prop.Value)) continue;

                string propUpper = prop.Name.ToUpperInvariant();
                if (ctrl.PictureData == null &&
                    (candidateSkinProps.Contains(propUpper) || propUpper.EndsWith("PICTURE") || propUpper.EndsWith("IMAGE")))
                {
                    var data = ResolveBinaryResource(prop.Value, basePath, model.IsUserControl, model.Name);
                    if (data != null && data.Length > 0)
                    {
                        ctrl.PictureData = data;
                    }
                }

                // Si el valor contiene una referencia a .frx o .ctx, extraer el bloque binario para persistencia ActiveX
                if ((prop.Value.Contains(".frx", StringComparison.OrdinalIgnoreCase) ||
                     prop.Value.Contains(".ctx", StringComparison.OrdinalIgnoreCase) ||
                     prop.Value.Contains(':')) &&
                    Resources.FrxResourceFile.ParseReference(prop.Value, out _, out _))
                {
                    var rawBlock = ResolveRawBinaryBlock(prop.Value, basePath, model.IsUserControl, model.Name, model.OriginalFrxBytes);
                    if (rawBlock != null)
                    {
                        ctrl.BinaryProperties[prop.Name] = rawBlock;

                        bool isStateProp = prop.Name.StartsWith("Object.", StringComparison.OrdinalIgnoreCase) ||
                                           prop.Name.Equals("Custom", StringComparison.OrdinalIgnoreCase) ||
                                           prop.Name.Equals("BinaryData", StringComparison.OrdinalIgnoreCase) ||
                                           (!ctrl.IsIntrinsic && ctrl.ActiveXStateBlob == null &&
                                            !propUpper.EndsWith("PICTURE") && !propUpper.EndsWith("ICON"));

                        if (isStateProp && ctrl.ActiveXStateBlob == null)
                        {
                            ctrl.ActiveXStateBlob = rawBlock;
                            ctrl.ActiveXStatePropertyName = prop.Name;
                        }
                    }
                }
            }
        }
    }

    private static void TryRegisterFormObjectReference(string objectLine)
    {
        if (!OperatingSystem.IsWindows()) return;

        try
        {
            // Formato típico: Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
            int eqIdx = objectLine.IndexOf('=');
            if (eqIdx < 0) return;

            string rhs = objectLine[(eqIdx + 1)..].Trim();
            string? guid = null;
            int openBrace = rhs.IndexOf('{');
            int closeBrace = rhs.IndexOf('}');
            if (openBrace >= 0 && closeBrace > openBrace)
            {
                guid = rhs.Substring(openBrace, closeBrace - openBrace + 1);
            }

            string? fileName = null;
            int semiIdx = rhs.IndexOf(';');
            if (semiIdx >= 0 && semiIdx + 1 < rhs.Length)
            {
                fileName = rhs[(semiIdx + 1)..].Trim().Trim('\"');
            }

            string searchTarget = !string.IsNullOrWhiteSpace(fileName) ? fileName : (guid ?? string.Empty);
            if (string.IsNullOrWhiteSpace(searchTarget) && string.IsNullOrWhiteSpace(guid))
                return;

            string? resolvedPath = ActiveX.OcxTypeLibReader.ResolveOcxPath(searchTarget, null, guid);
            if (!string.IsNullOrWhiteSpace(resolvedPath) &&
                ActiveX.OcxTypeLibReader.TryReadTypeLib(resolvedPath, out var libInfo, out _) &&
                libInfo != null)
            {
                IntelliSense.VbControlMemberCatalog.RegisterOcxLibrary(libInfo);
            }
        }
        catch
        {
            // Ignorar errores al intentar auto-registrar referencias Object de cabecera .frm
        }
    }
}
