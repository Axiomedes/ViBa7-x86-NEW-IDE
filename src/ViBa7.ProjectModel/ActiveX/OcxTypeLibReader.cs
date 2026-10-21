using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;
using Microsoft.Win32;
using ViBa7.ProjectModel.IntelliSense;

namespace ViBa7.ProjectModel.ActiveX;

[Flags]
public enum ComImplTypeFlags
{
    Default = 1,
    Source = 2,
    Restricted = 4,
    DefaultVTable = 8
}

/// <summary>
/// Lector de Type Libraries COM (.ocx, .dll, .tlb) que opera de forma nativa en procesos x64
/// utilizando oleaut32.dll!LoadTypeLibEx con REGKIND_NONE, sin requerir registro COM previo
/// ni permisos de administrador.
/// </summary>
[SupportedOSPlatform("windows")]
public static class OcxTypeLibReader
{
    private const int REGKIND_NONE = 2;

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void LoadTypeLibEx(
        string szFile,
        int regkind,
        out ITypeLib pptlib);

    /// <summary>
    /// Intenta resolver la ruta de un archivo OCX/TLB a partir de su declaración en el proyecto VB6.
    /// </summary>
    public static string? ResolveOcxPath(string relativeOrDeclaredPath, string? projectDirectory = null, string? guid = null)
    {
        if (string.IsNullOrWhiteSpace(relativeOrDeclaredPath)) return null;

        // 1. Ruta absoluta existente
        if (Path.IsPathRooted(relativeOrDeclaredPath) && File.Exists(relativeOrDeclaredPath))
        {
            return Path.GetFullPath(relativeOrDeclaredPath);
        }

        // 2. Ruta relativa al directorio del proyecto
        if (!string.IsNullOrEmpty(projectDirectory))
        {
            string combined = Path.Combine(projectDirectory, relativeOrDeclaredPath);
            if (File.Exists(combined))
            {
                return Path.GetFullPath(combined);
            }
        }

        string fileName = Path.GetFileName(relativeOrDeclaredPath);

        // 3. Directorio de SysWOW64 (controles COM clásicos de 32 bits en Windows x64)
        string sysWow64 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64");
        string wow64Path = Path.Combine(sysWow64, fileName);
        if (File.Exists(wow64Path))
        {
            return wow64Path;
        }

        // 4. Directorio de System32
        string sys32 = Environment.SystemDirectory;
        string sys32Path = Path.Combine(sys32, fileName);
        if (File.Exists(sys32Path))
        {
            return sys32Path;
        }

        // 5. Búsqueda en el Registro de Windows bajo HKCR\TypeLib si se dispone del GUID
        if (!string.IsNullOrWhiteSpace(guid))
        {
            try
            {
                string cleanGuid = guid.Trim();
                if (!cleanGuid.StartsWith('{')) cleanGuid = "{" + cleanGuid + "}";

                using var tlibKey = Registry.ClassesRoot.OpenSubKey($@"TypeLib\{cleanGuid}");
                if (tlibKey != null)
                {
                    foreach (var verName in tlibKey.GetSubKeyNames())
                    {
                        using var verKey = tlibKey.OpenSubKey(verName);
                        if (verKey == null) continue;

                        using var win32Key = verKey.OpenSubKey(@"0\win32");
                        if (win32Key?.GetValue(null) is string regPath32 && File.Exists(regPath32))
                        {
                            return regPath32;
                        }

                        using var win64Key = verKey.OpenSubKey(@"0\win64");
                        if (win64Key?.GetValue(null) is string regPath64 && File.Exists(regPath64))
                        {
                            return regPath64;
                        }
                    }
                }
            }
            catch
            {
                // Ignorar excepciones de lectura de registro
            }
        }

        return null;
    }

    /// <summary>
    /// Intenta leer la Type Library de forma segura capturando posibles excepciones COM.
    /// </summary>
    public static bool TryReadTypeLib(string filePath, out OcxLibraryInfo? libInfo, out string? errorMessage)
    {
        libInfo = null;
        errorMessage = null;

        try
        {
            libInfo = ReadTypeLib(filePath);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return false;
        }
    }

    /// <summary>
    /// Carga la Type Library incrustada en un archivo .ocx, .dll o .tlb y extrae todos los controles,
    /// interfaces, métodos, propiedades, parámetros y enumeraciones.
    /// </summary>
    public static OcxLibraryInfo ReadTypeLib(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"El archivo OCX o TypeLib no existe: {filePath}", filePath);
        }

        LoadTypeLibEx(filePath, REGKIND_NONE, out ITypeLib typeLib);

        var libInfo = new OcxLibraryInfo
        {
            FilePath = filePath
        };

        try
        {
            // Atributos de la librería
            nint pLibAttr = nint.Zero;
            typeLib.GetLibAttr(out pLibAttr);
            if (pLibAttr != nint.Zero)
            {
                try
                {
                    TYPELIBATTR libAttr = Marshal.PtrToStructure<TYPELIBATTR>(pLibAttr);
                    libInfo.Guid = libAttr.guid.ToString("B").ToUpperInvariant();
                    libInfo.MajorVersion = libAttr.wMajorVerNum;
                    libInfo.MinorVersion = libAttr.wMinorVerNum;
                }
                finally
                {
                    typeLib.ReleaseTLibAttr(pLibAttr);
                }
            }

            typeLib.GetDocumentation(-1, out string libName, out string libDoc, out _, out _);
            libInfo.LibraryName = libName ?? Path.GetFileNameWithoutExtension(filePath);
            libInfo.Description = libDoc ?? string.Empty;

            int typeInfoCount = typeLib.GetTypeInfoCount();

            for (int i = 0; i < typeInfoCount; i++)
            {
                typeLib.GetTypeInfo(i, out ITypeInfo tinfo);
                nint pAttr = nint.Zero;
                tinfo.GetTypeAttr(out pAttr);
                if (pAttr == nint.Zero)
                {
                    Marshal.ReleaseComObject(tinfo);
                    continue;
                }

                try
                {
                    TYPEATTR attr = Marshal.PtrToStructure<TYPEATTR>(pAttr);
                    tinfo.GetDocumentation(-1, out string typeName, out string typeDoc, out _, out _);

                    switch (attr.typekind)
                    {
                        case TYPEKIND.TKIND_COCLASS:
                            bool isHiddenOrRestricted = (attr.wTypeFlags & (TYPEFLAGS.TYPEFLAG_FHIDDEN | TYPEFLAGS.TYPEFLAG_FRESTRICTED)) != 0;
                            var ctrl = ProcessCoClass(typeName, typeDoc, tinfo, attr, isHiddenOrRestricted);
                            libInfo.Controls.Add(ctrl);
                            break;

                        case TYPEKIND.TKIND_ENUM:
                            var enumInfo = ExtractEnumInfo(tinfo, attr, typeName, typeDoc);
                            libInfo.Enums.Add(enumInfo);
                            break;
                    }
                }
                finally
                {
                    tinfo.ReleaseTypeAttr(pAttr);
                    Marshal.ReleaseComObject(tinfo);
                }
            }

            // Si ningún CoClass declaró explícitamente TYPEFLAG_FCONTROL (común en algunos OCX antiguos),
            // marcar como controles aquellos CoClass que tengan miembros y no sean páginas de propiedades.
            if (libInfo.Controls.Count > 0 && !libInfo.Controls.Exists(c => c.IsControl))
            {
                foreach (var c in libInfo.Controls)
                {
                    if (c.Members.Count > 0 &&
                        !c.Name.EndsWith("PropPage", StringComparison.OrdinalIgnoreCase) &&
                        !c.Name.EndsWith("PropertyPage", StringComparison.OrdinalIgnoreCase))
                    {
                        c.IsControl = true;
                    }
                }
            }
        }
        finally
        {
            Marshal.ReleaseComObject(typeLib);
        }

        return libInfo;
    }

    private static bool IsRegisteredControlClsid(string clsid)
    {
        if (string.IsNullOrWhiteSpace(clsid)) return false;
        try
        {
            using var ctrlKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}\Control");
            if (ctrlKey != null) return true;

            using var bmpKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}\ToolboxBitmap32");
            if (bmpKey != null) return true;
        }
        catch
        {
            // Ignorar errores de acceso al registro
        }
        return false;
    }

    private static OcxControlInfo ProcessCoClass(
        string name,
        string doc,
        ITypeInfo coclassInfo,
        TYPEATTR coclassAttr,
        bool isHiddenOrRestricted)
    {
        string clsid = coclassAttr.guid.ToString("B").ToUpperInvariant();
        bool hasControlFlag = (coclassAttr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FCONTROL) != 0;
        bool isControl = !isHiddenOrRestricted && (hasControlFlag || IsRegisteredControlClsid(clsid));

        var ctrl = new OcxControlInfo
        {
            Name = name,
            Description = doc ?? string.Empty,
            Clsid = clsid,
            IsControl = isControl
        };

        var membersByName = new Dictionary<string, OcxMemberInfo>(StringComparer.OrdinalIgnoreCase);
        var visitedMemberInterfaces = new HashSet<Guid>();
        var visitedEventInterfaces = new HashSet<Guid>();

        for (int i = 0; i < coclassAttr.cImplTypes; i++)
        {
            coclassInfo.GetRefTypeOfImplType(i, out int hRefType);
            coclassInfo.GetRefTypeInfo(hRefType, out ITypeInfo implInfo);
            coclassInfo.GetImplTypeFlags(i, out IMPLTYPEFLAGS implFlags);

            nint pImplAttr = nint.Zero;
            implInfo.GetTypeAttr(out pImplAttr);
            if (pImplAttr == nint.Zero)
            {
                Marshal.ReleaseComObject(implInfo);
                continue;
            }

            try
            {
                TYPEATTR implAttr = Marshal.PtrToStructure<TYPEATTR>(pImplAttr);
                implInfo.GetDocumentation(-1, out string implName, out _, out _, out _);

                bool isSource = ((int)implFlags & (int)ComImplTypeFlags.Source) != 0;
                bool isDefault = ((int)implFlags & (int)ComImplTypeFlags.Default) != 0;

                if (isSource)
                {
                    // Interfaz de eventos (e.g. __axJColCombo)
                    if (string.IsNullOrEmpty(ctrl.EventsInterfaceName) || isDefault)
                    {
                        ctrl.EventsInterfaceName = implName;
                        if (implAttr.guid != Guid.Empty)
                        {
                            ctrl.EventsInterfaceIid = implAttr.guid.ToString("B").ToUpperInvariant();
                        }
                    }
                    ExtractMembersAndEvents(implInfo, implAttr, null, ctrl.Events, null, visitedEventInterfaces);
                }
                else
                {
                    // Interfaz principal de propiedades y métodos (e.g. _axJColCombo)
                    if (string.IsNullOrEmpty(ctrl.DefaultInterfaceName) || isDefault)
                    {
                        ctrl.DefaultInterfaceName = implName;
                        if (implAttr.guid != Guid.Empty)
                        {
                            ctrl.DefaultInterfaceIid = implAttr.guid.ToString("B").ToUpperInvariant();
                        }
                    }
                    ExtractMembersAndEvents(implInfo, implAttr, ctrl.Members, null, membersByName, visitedMemberInterfaces);
                }
            }
            finally
            {
                implInfo.ReleaseTypeAttr(pImplAttr);
                Marshal.ReleaseComObject(implInfo);
            }
        }

        return ctrl;
    }

    private static readonly Guid IID_IUnknown = new("00000000-0000-0000-C000-000000000046");
    private static readonly Guid IID_IDispatch = new("00020400-0000-0000-C000-000000000046");

    private static void ExtractMembersAndEvents(
        ITypeInfo tinfo,
        TYPEATTR attr,
        List<OcxMemberInfo>? membersList,
        List<OcxEventInfo>? eventsList,
        Dictionary<string, OcxMemberInfo>? sharedMembersByName = null,
        HashSet<Guid>? visitedInterfaces = null)
    {
        if (attr.guid == IID_IUnknown || attr.guid == IID_IDispatch)
        {
            return;
        }

        visitedInterfaces ??= new HashSet<Guid>();
        if (attr.guid != Guid.Empty && !visitedInterfaces.Add(attr.guid))
        {
            return;
        }

        var membersByName = sharedMembersByName ?? new Dictionary<string, OcxMemberInfo>(StringComparer.OrdinalIgnoreCase);
        var eventsByName = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (eventsList != null)
        {
            foreach (var existingEvt in eventsList)
            {
                eventsByName.Add(existingEvt.Name);
            }
        }

        // 1. Recorrer interfaces base / heredadas (ej. IDispatch dual o interfaz padre)
        for (int i = 0; i < attr.cImplTypes; i++)
        {
            try
            {
                tinfo.GetRefTypeOfImplType(i, out int hBaseRef);
                tinfo.GetRefTypeInfo(hBaseRef, out ITypeInfo baseInfo);
                nint pBaseAttr = nint.Zero;
                baseInfo.GetTypeAttr(out pBaseAttr);
                if (pBaseAttr != nint.Zero)
                {
                    try
                    {
                        TYPEATTR baseAttr = Marshal.PtrToStructure<TYPEATTR>(pBaseAttr);
                        ExtractMembersAndEvents(baseInfo, baseAttr, membersList, eventsList, membersByName, visitedInterfaces);
                    }
                    finally
                    {
                        baseInfo.ReleaseTypeAttr(pBaseAttr);
                    }
                }
                Marshal.ReleaseComObject(baseInfo);
            }
            catch
            {
                // Ignorar si una interfaz base no puede resolverse
            }
        }

        // Si es una dispinterface dual (TKIND_DISPATCH con TYPEFLAG_FDUAL), inspeccionar también su interfaz vtable asociada (-1)
        if (attr.typekind == TYPEKIND.TKIND_DISPATCH && (attr.wTypeFlags & TYPEFLAGS.TYPEFLAG_FDUAL) != 0)
        {
            try
            {
                tinfo.GetRefTypeOfImplType(-1, out int hPartnerRef);
                tinfo.GetRefTypeInfo(hPartnerRef, out ITypeInfo partnerInfo);
                nint pPartnerAttr = nint.Zero;
                partnerInfo.GetTypeAttr(out pPartnerAttr);
                if (pPartnerAttr != nint.Zero)
                {
                    try
                    {
                        TYPEATTR partnerAttr = Marshal.PtrToStructure<TYPEATTR>(pPartnerAttr);
                        ExtractMembersAndEvents(partnerInfo, partnerAttr, membersList, eventsList, membersByName, visitedInterfaces);
                    }
                    finally
                    {
                        partnerInfo.ReleaseTypeAttr(pPartnerAttr);
                    }
                }
                Marshal.ReleaseComObject(partnerInfo);
            }
            catch
            {
                // Ignorar si no tiene partner dual accesible
            }
        }

        // 2. Extraer funciones y propiedades de tabla de funciones (cFuncs)
        for (int f = 0; f < attr.cFuncs; f++)
        {
            nint pFuncDesc = nint.Zero;
            tinfo.GetFuncDesc(f, out pFuncDesc);
            if (pFuncDesc == nint.Zero) continue;

            try
            {
                FUNCDESC fd = Marshal.PtrToStructure<FUNCDESC>(pFuncDesc);

                // Ignorar funciones restringidas de bajo nivel IUnknown/IDispatch
                if ((fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FRESTRICTED) != 0) continue;

                bool isHidden = (fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FHIDDEN) != 0;
                bool isNonBrowsable = (fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FNONBROWSABLE) != 0;
                bool isBindable = (fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FBINDABLE) != 0;
                bool isDefaultBind = (fd.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FDEFAULTBIND) != 0 || fd.memid == 0;

                // Obtener nombres de función y parámetros
                string[] names = new string[Math.Max(1, fd.cParams + 1)];
                tinfo.GetNames(fd.memid, names, names.Length, out int cNames);

                string funcName = cNames > 0 ? names[0] : string.Empty;
                if (string.IsNullOrWhiteSpace(funcName))
                {
                    tinfo.GetDocumentation(fd.memid, out funcName, out _, out _, out _);
                }
                if (string.IsNullOrWhiteSpace(funcName)) continue;

                // Ignorar funciones IUnknown / IDispatch estándar
                if (IsStandardComMethod(funcName)) continue;

                tinfo.GetDocumentation(fd.memid, out _, out string funcDoc, out _, out _);

                // Si estamos extrayendo eventos
                if (eventsList != null)
                {
                    if (eventsByName.Add(funcName))
                    {
                        var evt = new OcxEventInfo
                        {
                            Name = funcName,
                            Description = funcDoc ?? string.Empty,
                            DispId = fd.memid
                        };

                        ExtractParameters(tinfo, fd, names, evt.Parameters, isPropertyPut: false);
                        eventsList.Add(evt);
                    }
                    continue;
                }

                // Si estamos extrayendo miembros (propiedades y métodos)
                if (membersList != null)
                {
                    bool isPropGet = fd.invkind == INVOKEKIND.INVOKE_PROPERTYGET;
                    bool isPropPut = fd.invkind is INVOKEKIND.INVOKE_PROPERTYPUT or INVOKEKIND.INVOKE_PROPERTYPUTREF;
                    bool isProperty = isPropGet || isPropPut;

                    string resolvedType = isPropPut
                        ? GetPropertyPutValueType(tinfo, fd)
                        : GetTypeNameFromElemDesc(tinfo, fd.elemdescFunc);

                    if (membersByName.TryGetValue(funcName, out var existing))
                    {
                        // Unificar Property Get y Let/Set
                        if (isPropGet)
                        {
                            existing.IsPropertyGet = true;
                            if (!string.IsNullOrEmpty(resolvedType) && resolvedType != "Void")
                            {
                                existing.ReturnType = resolvedType;
                            }
                        }
                        if (isPropPut)
                        {
                            existing.IsPropertyPut = true;
                            if ((string.IsNullOrEmpty(existing.ReturnType) || existing.ReturnType == "Void") &&
                                !string.IsNullOrEmpty(resolvedType) && resolvedType != "Void")
                            {
                                existing.ReturnType = resolvedType;
                            }
                        }
                        if (isHidden) existing.IsHidden = true;
                        if (isNonBrowsable) existing.IsBrowsable = false;
                        if (isBindable) existing.IsBindable = true;
                        if (isDefaultBind) existing.IsDefaultBind = true;
                        if (string.IsNullOrEmpty(existing.Description) && !string.IsNullOrEmpty(funcDoc))
                        {
                            existing.Description = funcDoc;
                        }
                    }
                    else
                    {
                        var member = new OcxMemberInfo
                        {
                            Name = funcName,
                            Kind = isProperty ? VbSymbolKind.Property : VbSymbolKind.Method,
                            ReturnType = resolvedType,
                            Description = funcDoc ?? string.Empty,
                            IsPropertyGet = isPropGet,
                            IsPropertyPut = isPropPut,
                            DispId = fd.memid,
                            IsHidden = isHidden,
                            IsBrowsable = !isNonBrowsable && !isHidden,
                            IsBindable = isBindable,
                            IsDefaultBind = isDefaultBind
                        };

                        ExtractParameters(tinfo, fd, names, member.Parameters, isPropPut);
                        membersByName[funcName] = member;
                        membersList.Add(member);
                    }
                }
            }
            finally
            {
                tinfo.ReleaseFuncDesc(pFuncDesc);
            }
        }

        // 3. Extraer propiedades definidas en dispinterfaces (cVars / GetVarDesc)
        // Muchos controles OCX en C++/MFC (COleControl), ATL y Delphi declaran sus propiedades en cVars dentro de su dispinterface.
        if (membersList != null && attr.cVars > 0)
        {
            for (int v = 0; v < attr.cVars; v++)
            {
                nint pVarDesc = nint.Zero;
                tinfo.GetVarDesc(v, out pVarDesc);
                if (pVarDesc == nint.Zero) continue;

                try
                {
                    VARDESC vd = Marshal.PtrToStructure<VARDESC>(pVarDesc);
                    if ((vd.wVarFlags & (short)VARFLAGS.VARFLAG_FRESTRICTED) != 0) continue;

                    bool isHidden = (vd.wVarFlags & (short)VARFLAGS.VARFLAG_FHIDDEN) != 0;
                    bool isNonBrowsable = (vd.wVarFlags & (short)VARFLAGS.VARFLAG_FNONBROWSABLE) != 0;
                    bool isBindable = (vd.wVarFlags & (short)VARFLAGS.VARFLAG_FBINDABLE) != 0;
                    bool isDefaultBind = (vd.wVarFlags & (short)VARFLAGS.VARFLAG_FDEFAULTBIND) != 0 || vd.memid == 0;

                    tinfo.GetDocumentation(vd.memid, out string propName, out string propDoc, out _, out _);
                    if (string.IsNullOrWhiteSpace(propName))
                    {
                        string[] varNames = new string[1];
                        tinfo.GetNames(vd.memid, varNames, 1, out int cVarNames);
                        if (cVarNames > 0) propName = varNames[0];
                    }

                    if (string.IsNullOrWhiteSpace(propName) || IsStandardComMethod(propName))
                        continue;

                    string propType = GetTypeNameFromElemDesc(tinfo, vd.elemdescVar);
                    bool isReadOnly = (vd.wVarFlags & (short)VARFLAGS.VARFLAG_FREADONLY) != 0;

                    if (membersByName.TryGetValue(propName, out var existingProp))
                    {
                        existingProp.IsPropertyGet = true;
                        if (!isReadOnly) existingProp.IsPropertyPut = true;
                        if (isHidden) existingProp.IsHidden = true;
                        if (isNonBrowsable) existingProp.IsBrowsable = false;
                        if (isBindable) existingProp.IsBindable = true;
                        if (isDefaultBind) existingProp.IsDefaultBind = true;
                        if ((string.IsNullOrEmpty(existingProp.ReturnType) || existingProp.ReturnType == "Void") &&
                            !string.IsNullOrEmpty(propType))
                        {
                            existingProp.ReturnType = propType;
                        }
                        if (string.IsNullOrEmpty(existingProp.Description) && !string.IsNullOrEmpty(propDoc))
                        {
                            existingProp.Description = propDoc;
                        }
                    }
                    else
                    {
                        var member = new OcxMemberInfo
                        {
                            Name = propName,
                            Kind = VbSymbolKind.Property,
                            ReturnType = propType,
                            Description = propDoc ?? string.Empty,
                            IsPropertyGet = true,
                            IsPropertyPut = !isReadOnly,
                            DispId = vd.memid,
                            IsHidden = isHidden,
                            IsBrowsable = !isNonBrowsable && !isHidden,
                            IsBindable = isBindable,
                            IsDefaultBind = isDefaultBind
                        };

                        membersByName[propName] = member;
                        membersList.Add(member);
                    }
                }
                finally
                {
                    tinfo.ReleaseVarDesc(pVarDesc);
                }
            }
        }
    }

    private static string GetPropertyPutValueType(ITypeInfo tinfo, FUNCDESC fd)
    {
        if (fd.cParams > 0 && fd.lprgelemdescParam != nint.Zero)
        {
            int elemDescSize = Marshal.SizeOf<ELEMDESC>();
            nint pLastElem = fd.lprgelemdescParam + ((fd.cParams - 1) * elemDescSize);
            ELEMDESC lastElem = Marshal.PtrToStructure<ELEMDESC>(pLastElem);
            string paramType = GetTypeNameFromElemDesc(tinfo, lastElem);
            if (!string.IsNullOrEmpty(paramType) && paramType != "Void")
            {
                return paramType;
            }
        }
        return "Variant";
    }

    private static void ExtractParameters(
        ITypeInfo tinfo,
        FUNCDESC fd,
        string[] names,
        List<OcxParameterInfo> parameters,
        bool isPropertyPut = false)
    {
        int paramCount = isPropertyPut ? Math.Max(0, fd.cParams - 1) : fd.cParams;
        if (paramCount <= 0 || fd.lprgelemdescParam == nint.Zero) return;

        int elemDescSize = Marshal.SizeOf<ELEMDESC>();

        for (int p = 0; p < paramCount; p++)
        {
            nint pElem = fd.lprgelemdescParam + (p * elemDescSize);
            ELEMDESC elem = Marshal.PtrToStructure<ELEMDESC>(pElem);

            string paramName = (p + 1 < names.Length && !string.IsNullOrEmpty(names[p + 1]))
                ? names[p + 1]
                : $"param{p + 1}";

            string paramType = GetTypeNameFromElemDesc(tinfo, elem);
            bool isOptional = (elem.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FOPT) != 0;
            bool isByRef = (VarEnum)elem.tdesc.vt == VarEnum.VT_PTR ||
                           (elem.desc.paramdesc.wParamFlags & PARAMFLAG.PARAMFLAG_FOUT) != 0;

            parameters.Add(new OcxParameterInfo
            {
                Name = paramName,
                Type = paramType,
                IsOptional = isOptional,
                IsByRef = isByRef
            });
        }
    }

    private static string GetTypeNameFromElemDesc(ITypeInfo tinfo, ELEMDESC elem)
    {
        return DecodeTypeDesc(tinfo, elem.tdesc);
    }

    private static string DecodeTypeDesc(ITypeInfo tinfo, TYPEDESC tdesc)
    {
        var vt = (VarEnum)tdesc.vt;

        if (vt == VarEnum.VT_PTR && tdesc.lpValue != nint.Zero)
        {
            TYPEDESC inner = Marshal.PtrToStructure<TYPEDESC>(tdesc.lpValue);
            return DecodeTypeDesc(tinfo, inner);
        }

        if (vt == VarEnum.VT_USERDEFINED)
        {
            try
            {
                int hRef = (int)tdesc.lpValue;
                tinfo.GetRefTypeInfo(hRef, out ITypeInfo refInfo);
                refInfo.GetDocumentation(-1, out string refName, out _, out _, out _);
                if (!string.IsNullOrEmpty(refName))
                {
                    return refName;
                }
            }
            catch
            {
                // Ignorar y retornar Object por defecto
            }
            return "Object";
        }

        return vt switch
        {
            VarEnum.VT_I2 => "Integer",
            VarEnum.VT_I4 => "Long",
            VarEnum.VT_R4 => "Single",
            VarEnum.VT_R8 => "Double",
            VarEnum.VT_CY => "Currency",
            VarEnum.VT_DATE => "Date",
            VarEnum.VT_BSTR => "String",
            VarEnum.VT_LPSTR => "String",
            VarEnum.VT_LPWSTR => "String",
            VarEnum.VT_DISPATCH => "Object",
            VarEnum.VT_UNKNOWN => "Object",
            VarEnum.VT_BOOL => "Boolean",
            VarEnum.VT_VARIANT => "Variant",
            VarEnum.VT_VOID => "Void",
            VarEnum.VT_EMPTY => "Void",
            VarEnum.VT_UI1 => "Byte",
            VarEnum.VT_UI2 => "Integer",
            VarEnum.VT_UI4 => "Long",
            _ => "Variant"
        };
    }

    private static OcxEnumInfo ExtractEnumInfo(ITypeInfo tinfo, TYPEATTR attr, string enumName, string enumDoc)
    {
        var enumInfo = new OcxEnumInfo
        {
            Name = enumName,
            Description = enumDoc ?? string.Empty
        };

        for (int v = 0; v < attr.cVars; v++)
        {
            nint pVarDesc = nint.Zero;
            tinfo.GetVarDesc(v, out pVarDesc);
            if (pVarDesc == nint.Zero) continue;

            try
            {
                VARDESC vd = Marshal.PtrToStructure<VARDESC>(pVarDesc);
                tinfo.GetDocumentation(vd.memid, out string varName, out string varDoc, out _, out _);

                long constValue = 0;
                if (vd.varkind == VARKIND.VAR_CONST && vd.desc.lpvarValue != nint.Zero)
                {
                    object? valObj = Marshal.GetObjectForNativeVariant(vd.desc.lpvarValue);
                    if (valObj != null)
                    {
                        try { constValue = Convert.ToInt64(valObj); } catch { }
                    }
                }

                enumInfo.Items.Add(new OcxEnumItemInfo
                {
                    Name = varName,
                    Value = constValue,
                    Description = varDoc ?? string.Empty
                });
            }
            finally
            {
                tinfo.ReleaseVarDesc(pVarDesc);
            }
        }

        return enumInfo;
    }

    private static bool IsStandardComMethod(string name)
    {
        return name.Equals("QueryInterface", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("AddRef", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("Release", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("GetTypeInfoCount", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("GetTypeInfo", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("GetIDsOfNames", StringComparison.OrdinalIgnoreCase) ||
               name.Equals("Invoke", StringComparison.OrdinalIgnoreCase);
    }
}
