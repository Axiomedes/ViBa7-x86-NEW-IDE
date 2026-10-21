namespace ViBa7.ProjectModel.IntelliSense;

/// <summary>
/// Información de un parámetro para Parameter Info y Quick Info de VB6.
/// </summary>
public class VbParameterInfo
{
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = "Variant";
    public bool IsOptional { get; init; } = false;
    public bool IsByRef { get; init; } = false;
    public string? DefaultValue { get; init; } = null;
    public bool IsParamArray { get; init; } = false;
    public string Description { get; init; } = string.Empty;

    public string DisplayText
    {
        get
        {
            string opt = IsOptional ? "[Optional] " : "";
            if (IsParamArray) opt = "ParamArray ";
            string def = DefaultValue != null ? $" = {DefaultValue}" : "";
            return $"{opt}{Name} As {Type}{def}";
        }
    }
}
