using System;
using System.Collections.Generic;
using ViBa7.ProjectModel.IntelliSense;

namespace ViBa7.ProjectModel.ActiveX;

/// <summary>
/// Parámetro de un método, propiedad o evento de un control ActiveX/COM.
/// </summary>
public class OcxParameterInfo
{
    public string Name { get; set; } = string.Empty;
    public string Type { get; set; } = "Variant";
    public bool IsOptional { get; set; }
    public bool IsByRef { get; set; }
    public string? DefaultValue { get; set; }
    public string? Description { get; set; }
}

/// <summary>
/// Miembro (método o propiedad) de un control ActiveX/COM.
/// </summary>
public class OcxMemberInfo
{
    public string Name { get; set; } = string.Empty;
    public VbSymbolKind Kind { get; set; } = VbSymbolKind.Property;
    public string ReturnType { get; set; } = "Variant";
    public string Description { get; set; } = string.Empty;
    public List<OcxParameterInfo> Parameters { get; set; } = [];
    public bool IsPropertyGet { get; set; }
    public bool IsPropertyPut { get; set; }
    public int DispId { get; set; }
    public bool IsHidden { get; set; }
    public bool IsBrowsable { get; set; } = true;
    public bool IsBindable { get; set; }
    public bool IsDefaultBind { get; set; }
}

/// <summary>
/// Evento expuesto por una interfaz Source de un control ActiveX/COM.
/// </summary>
public class OcxEventInfo
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public List<OcxParameterInfo> Parameters { get; set; } = [];
    public int DispId { get; set; }
}

/// <summary>
/// Elemento individual de una enumeración COM (TKIND_ENUM).
/// </summary>
public class OcxEnumItemInfo
{
    public string Name { get; set; } = string.Empty;
    public long Value { get; set; }
    public string? Description { get; set; }
}

/// <summary>
/// Enumeración declarada en la Type Library del OCX.
/// </summary>
public class OcxEnumInfo
{
    public string Name { get; set; } = string.Empty;
    public string? Description { get; set; }
    public List<OcxEnumItemInfo> Items { get; set; } = [];
}

/// <summary>
/// Control visual o CoClass expuesto por la Type Library del OCX.
/// </summary>
public class OcxControlInfo
{
    public string Name { get; set; } = string.Empty;
    public string Clsid { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string? DefaultInterfaceName { get; set; }
    public string? DefaultInterfaceIid { get; set; }
    public string? EventsInterfaceName { get; set; }
    public string? EventsInterfaceIid { get; set; }
    public List<OcxMemberInfo> Members { get; set; } = [];
    public List<OcxEventInfo> Events { get; set; } = [];
    public bool IsControl { get; set; } = true;
}

/// <summary>
/// Información completa extraída de la Type Library de un archivo OCX/DLL/TLB.
/// </summary>
public class OcxLibraryInfo
{
    public string FilePath { get; set; } = string.Empty;
    public string LibraryName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Guid { get; set; } = string.Empty;
    public int MajorVersion { get; set; }
    public int MinorVersion { get; set; }
    public List<OcxControlInfo> Controls { get; set; } = [];
    public List<OcxEnumInfo> Enums { get; set; } = [];
}
