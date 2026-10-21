using System;
using System.Collections.Generic;
using System.Linq;

namespace ViBa7.ProjectModel.IntelliSense;

/// <summary>
/// Elemento de símbolo para Auto List Members, Complete Word y Quick Info.
/// </summary>
public class VbSymbolItem
{
    public string Name { get; init; } = string.Empty;
    public VbSymbolKind Kind { get; init; } = VbSymbolKind.Keyword;
    public string? ReturnType { get; init; } = null;
    public List<VbParameterInfo> Parameters { get; init; } = [];
    public string Description { get; init; } = string.Empty;
    public string? CustomSignature { get; init; } = null;
    public int DispId { get; init; } = -1;
    public bool IsHidden { get; init; } = false;
    public bool IsBrowsable { get; init; } = true;
    public bool IsReadOnly { get; init; } = false;

    public string Glyph => Kind switch
    {
        VbSymbolKind.Method => "⚡",
        VbSymbolKind.IntrinsicFunction => "ƒ",
        VbSymbolKind.Property => "📖",
        VbSymbolKind.Control => "🎛️",
        VbSymbolKind.Variable => "📦",
        VbSymbolKind.Constant => "🔒",
        VbSymbolKind.DataType => "📐",
        VbSymbolKind.Keyword => "🔑",
        VbSymbolKind.Enum => "📑",
        VbSymbolKind.Event => "🔔",
        _ => "•"
    };

    public string KindDisplayName => Kind switch
    {
        VbSymbolKind.Method => "Procedimiento Sub",
        VbSymbolKind.IntrinsicFunction => "Función del Runtime",
        VbSymbolKind.Property => "Propiedad",
        VbSymbolKind.Control => "Control de Formulario",
        VbSymbolKind.Variable => "Variable",
        VbSymbolKind.Constant => "Constante",
        VbSymbolKind.DataType => "Tipo de Datos",
        VbSymbolKind.Keyword => "Palabra Clave",
        VbSymbolKind.Enum => "Enumeración",
        VbSymbolKind.Event => "Evento",
        _ => "Símbolo"
    };

    public string Signature
    {
        get
        {
            if (!string.IsNullOrEmpty(CustomSignature))
            {
                return CustomSignature;
            }

            if (Kind == VbSymbolKind.IntrinsicFunction || Kind == VbSymbolKind.Method)
            {
                string kw = Kind == VbSymbolKind.IntrinsicFunction || !string.IsNullOrEmpty(ReturnType) ? "Function" : "Sub";
                string paramList = string.Join(", ", Parameters.Select(p => p.DisplayText));
                string ret = !string.IsNullOrEmpty(ReturnType) ? $" As {ReturnType}" : "";
                return $"{kw} {Name}({paramList}){ret}";
            }

            if (Kind == VbSymbolKind.Property)
            {
                string ret = !string.IsNullOrEmpty(ReturnType) ? $" As {ReturnType}" : "";
                return $"Property {Name}{ret}";
            }

            if (Kind == VbSymbolKind.Constant)
            {
                string ret = !string.IsNullOrEmpty(ReturnType) ? $" As {ReturnType}" : "";
                return $"Const {Name}{ret}";
            }

            if (Kind == VbSymbolKind.Variable || Kind == VbSymbolKind.Control)
            {
                string ret = !string.IsNullOrEmpty(ReturnType) ? $" As {ReturnType}" : "";
                return $"{Name}{ret}";
            }

            return Name;
        }
    }
}
