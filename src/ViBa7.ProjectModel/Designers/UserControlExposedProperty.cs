using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace ViBa7.ProjectModel.Designers;

/// <summary>
/// Representa una propiedad pública expuesta por un UserControl en su código fuente (.ctl)
/// mediante declaraciones 'Property Get / Let / Set' y persistencia en 'PropBag.ReadProperty/WriteProperty'.
/// </summary>
public class UserControlExposedProperty
{
    public string Name { get; set; } = string.Empty;
    public string TypeName { get; set; } = string.Empty;
    public string? DefaultValue { get; set; }
    public string Category { get; set; } = "UserControl";
    public string Description { get; set; } = string.Empty;
    public List<DesignerPropertyChoice> Choices { get; set; } = new();

    public DesignerPropertyType InferredType
    {
        get
        {
            // Propiedades de códigos de caracteres de iconos (ej. IconCharCode, CharCode) deben ser de tipo Text para admitir cadenas hex ("&HEA01", "EA01", "U+EA01")
            if (Name.Contains("CharCode", StringComparison.OrdinalIgnoreCase)) return DesignerPropertyType.Text;
            if (Choices.Count > 0) return DesignerPropertyType.EnumChoice;
            if (DesignerPropertyFactory.IsColorProperty(Name, TypeName, DefaultValue)) return DesignerPropertyType.ColorHex;
            string t = (TypeName ?? string.Empty).Trim().ToUpperInvariant();
            if (t.Equals("BOOLEAN", StringComparison.OrdinalIgnoreCase)) return DesignerPropertyType.Boolean;
            if (t.Contains("PICTURE") || t.Contains("STDPIC") || t.Contains("IPICTURE") || Name.Contains("Picture", StringComparison.OrdinalIgnoreCase)) return DesignerPropertyType.Picture;
            if (Name.Equals("FontName", StringComparison.OrdinalIgnoreCase) || (Name.Contains("Font", StringComparison.OrdinalIgnoreCase) && Name.EndsWith("Name", StringComparison.OrdinalIgnoreCase))) return DesignerPropertyType.EnumChoice;
            if (t.Contains("FONT") || Name.Contains("Font", StringComparison.OrdinalIgnoreCase)) return DesignerPropertyType.FontDescription;
            if (t.Equals("INTEGER", StringComparison.OrdinalIgnoreCase) ||
                t.Equals("LONG", StringComparison.OrdinalIgnoreCase) ||
                t.Equals("BYTE", StringComparison.OrdinalIgnoreCase) ||
                t.Equals("SINGLE", StringComparison.OrdinalIgnoreCase) ||
                t.Equals("DOUBLE", StringComparison.OrdinalIgnoreCase)) return DesignerPropertyType.IntegerRaw;
            return DesignerPropertyType.Text;
        }
    }
}

/// <summary>
/// Extractor estático de propiedades públicas, enumeraciones y persistencia PropertyBag desde el código VB6 de un archivo .ctl.
/// </summary>
public static class UserControlPropertyExtractor
{
    private static readonly Regex EnumBlockRegex = new(
        @"(?:Public\s+)?Enum\s+([\w_]+)([\s\S]*?)End\s+Enum",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex EnumItemRegex = new(
        @"^\s*(?:\[([^\]]+)\]|([\w_]+))\s*(?:=\s*([^'\r\n]+))?(?:'([^\r\n]*))?",
        RegexOptions.Multiline | RegexOptions.Compiled);

    // En VB6, Property Get/Let/Set sin modificador de ámbito es Public por defecto
    private static readonly Regex PropertyDeclRegex = new(
        @"(?:^|\r?\n)\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?Property\s+(Get|Let|Set)\s+([\w_]+)\s*(?:\(([^)]*)\))?\s*(?:As\s+([\w_\.]+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ReadPropertyRegex = new(
        @"(?:PropBag|\.|\b)\s*ReadProperty\s*\(\s*""([^""]+)""\s*(?:,\s*([^,\)\r\n]+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex WritePropertyRegex = new(
        @"(?:PropBag|\.|\b)\s*WriteProperty\s*\(\s*""([^""]+)""\s*(?:,\s*[^,\)\r\n]+)?\s*(?:,\s*([^,\)\r\n]+))?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Analiza el código fuente de un UserControl (.ctl) y extrae todas sus propiedades públicas expuestas con sus tipos y valores por defecto.
    /// </summary>
    public static List<UserControlExposedProperty> ExtractProperties(string ctlSourceCode)
    {
        var result = new List<UserControlExposedProperty>();
        if (string.IsNullOrWhiteSpace(ctlSourceCode)) return result;

        // 1. Extraer definiciones de Enums públicos
        var enums = new Dictionary<string, List<DesignerPropertyChoice>>(StringComparer.OrdinalIgnoreCase);
        foreach (Match enumMatch in EnumBlockRegex.Matches(ctlSourceCode))
        {
            string enumName = enumMatch.Groups[1].Value.Trim();
            string enumBody = enumMatch.Groups[2].Value;

            var choices = new List<DesignerPropertyChoice>();
            int autoVal = 0;
            foreach (Match itemMatch in EnumItemRegex.Matches(enumBody))
            {
                string memberName = !string.IsNullOrEmpty(itemMatch.Groups[1].Value)
                    ? itemMatch.Groups[1].Value
                    : itemMatch.Groups[2].Value;

                if (string.IsNullOrWhiteSpace(memberName)) continue;

                string valStr;
                if (!string.IsNullOrWhiteSpace(itemMatch.Groups[3].Value))
                {
                    valStr = itemMatch.Groups[3].Value.Trim();
                    if (int.TryParse(valStr, out int parsed))
                    {
                        autoVal = parsed + 1;
                    }
                }
                else
                {
                    valStr = autoVal.ToString();
                    autoVal++;
                }

                choices.Add(new DesignerPropertyChoice($"{valStr} - {memberName}", valStr));
            }

            if (choices.Count > 0)
            {
                enums[enumName] = choices;
            }
        }

        // 2. Extraer llamadas ReadProperty y WriteProperty para valores por defecto y nombres persistidos
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match readMatch in ReadPropertyRegex.Matches(ctlSourceCode))
        {
            string propName = readMatch.Groups[1].Value.Trim();
            string defVal = readMatch.Groups[2].Value.Trim();
            if (!defaults.ContainsKey(propName))
            {
                if (defVal.StartsWith('"') && defVal.EndsWith('"') && defVal.Length >= 2)
                {
                    defVal = defVal[1..^1];
                }
                defaults[propName] = defVal;
            }
        }
        foreach (Match writeMatch in WritePropertyRegex.Matches(ctlSourceCode))
        {
            string propName = writeMatch.Groups[1].Value.Trim();
            string defVal = writeMatch.Groups[2].Value.Trim();
            if (!defaults.ContainsKey(propName))
            {
                if (defVal.StartsWith('"') && defVal.EndsWith('"') && defVal.Length >= 2)
                {
                    defVal = defVal[1..^1];
                }
                defaults[propName] = defVal;
            }
        }

        // 3. Extraer declaraciones Property (Get / Let / Set)
        var propsByName = new Dictionary<string, UserControlExposedProperty>(StringComparer.OrdinalIgnoreCase);
        foreach (Match propMatch in PropertyDeclRegex.Matches(ctlSourceCode))
        {
            string scope = propMatch.Groups[1].Value.Trim();
            if (scope.Equals("Private", StringComparison.OrdinalIgnoreCase)) continue;

            string kind = propMatch.Groups[2].Value.ToUpperInvariant();
            string propName = propMatch.Groups[3].Value.Trim();
            string paramList = propMatch.Groups[4].Value.Trim();
            string asType = propMatch.Groups[5].Value.Trim();

            // Filtrar propiedades internas o sobrecargadas con parámetros múltiples no indexados
            if (paramList.Contains(',')) continue;

            if (!propsByName.TryGetValue(propName, out var exposed))
            {
                exposed = new UserControlExposedProperty
                {
                    Name = propName,
                    Category = "UserControl"
                };
                propsByName[propName] = exposed;
                result.Add(exposed);
            }

            if (kind == "GET" && !string.IsNullOrEmpty(asType))
            {
                exposed.TypeName = asType;
            }
            else if ((kind == "LET" || kind == "SET") && string.IsNullOrEmpty(exposed.TypeName) && !string.IsNullOrEmpty(paramList))
            {
                int asIdx = paramList.LastIndexOf("As ", StringComparison.OrdinalIgnoreCase);
                if (asIdx >= 0)
                {
                    exposed.TypeName = paramList[(asIdx + 3)..].Trim();
                }
            }

            // Asociar opciones de Enum si coincide el tipo
            if (!string.IsNullOrEmpty(exposed.TypeName) && enums.TryGetValue(exposed.TypeName, out var enumChoices))
            {
                exposed.Choices = new List<DesignerPropertyChoice>(enumChoices);
            }
            else if (exposed.Name.Equals("FontName", StringComparison.OrdinalIgnoreCase) ||
                     (exposed.Name.Contains("Font", StringComparison.OrdinalIgnoreCase) && exposed.Name.EndsWith("Name", StringComparison.OrdinalIgnoreCase)))
            {
                exposed.Choices = DesignerPropertyFactory.GetSystemFontChoices();
            }

            // Asociar valor por defecto si existe
            if (defaults.TryGetValue(propName, out var defVal) && !string.IsNullOrEmpty(defVal))
            {
                exposed.DefaultValue = defVal;
            }
        }

        // 4. Agregar propiedades persistidas en PropertyBag (ReadProperty / WriteProperty) que no se hayan detectado por Get/Let
        foreach (var (propName, defVal) in defaults)
        {
            if (propsByName.ContainsKey(propName)) continue;
            if (propName.StartsWith("_", StringComparison.OrdinalIgnoreCase)) continue;

            var exposed = new UserControlExposedProperty
            {
                Name = propName,
                Category = "UserControl",
                DefaultValue = defVal
            };

            if (DesignerPropertyFactory.IsColorProperty(propName, null, defVal))
            {
                exposed.TypeName = "OLE_COLOR";
            }
            else if (propName.Equals("FontName", StringComparison.OrdinalIgnoreCase) ||
                     (propName.Contains("Font", StringComparison.OrdinalIgnoreCase) && propName.EndsWith("Name", StringComparison.OrdinalIgnoreCase)))
            {
                exposed.TypeName = "String";
                exposed.Choices = DesignerPropertyFactory.GetSystemFontChoices();
            }
            else if (propName.Contains("Font", StringComparison.OrdinalIgnoreCase))
            {
                exposed.TypeName = "StdFont";
            }
            else if (propName.Contains("Picture", StringComparison.OrdinalIgnoreCase))
            {
                exposed.TypeName = "StdPicture";
            }
            else if (defVal.Equals("True", StringComparison.OrdinalIgnoreCase) || defVal.Equals("False", StringComparison.OrdinalIgnoreCase))
            {
                exposed.TypeName = "Boolean";
            }
            else if (long.TryParse(defVal, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) && !defVal.StartsWith("-2147483"))
            {
                exposed.TypeName = "Long";
            }
            else
            {
                exposed.TypeName = "String";
            }

            if (enums.TryGetValue(exposed.TypeName, out var enumChoices))
            {
                exposed.Choices = new List<DesignerPropertyChoice>(enumChoices);
            }

            propsByName[propName] = exposed;
            result.Add(exposed);
        }

        // 5. Inferencia final de TypeName para propiedades que aún no tengan tipo definido
        foreach (var exp in result)
        {
            if (string.IsNullOrWhiteSpace(exp.TypeName))
            {
                if (DesignerPropertyFactory.IsColorProperty(exp.Name, null, exp.DefaultValue))
                {
                    exp.TypeName = "OLE_COLOR";
                }
                else if (exp.Name.Equals("FontName", StringComparison.OrdinalIgnoreCase) ||
                         (exp.Name.Contains("Font", StringComparison.OrdinalIgnoreCase) && exp.Name.EndsWith("Name", StringComparison.OrdinalIgnoreCase)))
                {
                    exp.TypeName = "String";
                    if (exp.Choices.Count == 0)
                    {
                        exp.Choices = DesignerPropertyFactory.GetSystemFontChoices();
                    }
                }
                else if (exp.Name.Contains("Font", StringComparison.OrdinalIgnoreCase))
                {
                    exp.TypeName = "StdFont";
                }
                else if (exp.Name.Contains("Picture", StringComparison.OrdinalIgnoreCase))
                {
                    exp.TypeName = "StdPicture";
                }
                else
                {
                    exp.TypeName = "String";
                }
            }
        }

        return result;
    }
}
