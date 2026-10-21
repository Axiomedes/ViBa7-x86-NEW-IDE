using System;
using System.IO;
using System.Text;

namespace ViBa7.ProjectModel.Compilation;

/// <summary>
/// Modelado y serialización de las estructuras de cabeceras internas de Visual Basic 6.0
/// requeridas por el runtime oficial de Windows (MSVBVM60.DLL).
/// </summary>
public static class Vb6BinaryStructures
{
    // Firma fija de biblioteca de tipos OLE de VB6
    public const string DefaultOleAutGuid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation";

    /// <summary>
    /// Escribe la cabecera canónica VBHeader (firma "VB5!") de 108 bytes.
    /// </summary>
    public static void WriteVbHeader(
        BinaryWriter bw,
        uint lpProjInfoVa,
        uint lpSubMainVa = 0,
        uint dwLcid = 0x00000409,
        ushort formCount = 1,
        ushort externalComponentCount = 0)
    {
        // 1. Firma mágica: "VB5!" (4 bytes: 0x56, 0x42, 0x35, 0x21)
        bw.Write(Pe32Constants.Vb5Magic);

        // 2. wRuntimeBuild (2 bytes: versión de compilador, ej. 0x000F)
        bw.Write((ushort)0x000F);

        // 3. szLangDLL (14 bytes)
        byte[] langDll = new byte[14];
        Encoding.ASCII.GetBytes("*\\G{00020430-", 0, 13, langDll, 0);
        bw.Write(langDll);

        // 4. szSecLangDLL (14 bytes)
        bw.Write(new byte[14]);

        // 5. wRuntimeRevision (2 bytes)
        bw.Write((ushort)0x0001);

        // 6. dwLCID (4 bytes)
        bw.Write(dwLcid);

        // 7. szSecLangDLL2 (4 bytes)
        bw.Write((uint)0);

        // 8. lpSubMain (4 bytes) - VA de Sub Main (o 0 si inicia por Formulario)
        bw.Write(lpSubMainVa);

        // 9. lpProjInfo (4 bytes) - VA de la estructura ProjInfo
        bw.Write(lpProjInfoVa);

        // 10. fMode (4 bytes) - Flags de ejecución
        bw.Write((uint)0x00000000);

        // 11. dwThreadFlags (4 bytes)
        bw.Write((uint)0x00000000);

        // 12. dwThreadCount (4 bytes)
        bw.Write((uint)1);

        // 13. wFormCount (2 bytes)
        bw.Write(formCount);

        // 14. wExternalComponentCount (2 bytes)
        bw.Write(externalComponentCount);

        // 15. dwThunkCount (4 bytes)
        bw.Write((uint)0);

        // 16. lpGuiTable (4 bytes)
        bw.Write((uint)0);

        // 17. lpIdeData (4 bytes)
        bw.Write((uint)0);

        // 18. lpProjectData (4 bytes)
        bw.Write((uint)0);
    }

    /// <summary>
    /// Escribe la estructura ProjInfo (44 bytes).
    /// </summary>
    public static void WriteProjInfo(
        BinaryWriter bw,
        uint lpProjectDataVa,
        uint lpObjectTableVa,
        uint lpCodeStartVa,
        uint lpCodeEndVa,
        uint dwDataSize)
    {
        bw.Write((uint)0x000001F4); // dwVersion = 500 (VB5/VB6)
        bw.Write((uint)0);          // lpAppObject
        bw.Write(lpProjectDataVa);  // lpProjectData
        bw.Write((uint)0);          // dwTableFlags
        bw.Write((uint)0);          // dwReserved
        bw.Write(lpObjectTableVa);   // lpObjectTable
        bw.Write((uint)0);          // dwNull
        bw.Write(lpCodeStartVa);    // lpCodeStart
        bw.Write(lpCodeEndVa);      // lpCodeEnd
        bw.Write(dwDataSize);       // dwDataSize
        bw.Write((uint)0);          // lpThreadStart
    }

    /// <summary>
    /// Escribe la estructura ProjectData (8 bytes + nombre).
    /// </summary>
    public static void WriteProjectData(
        BinaryWriter bw,
        uint lpObjectTableVa,
        string projectName)
    {
        bw.Write((uint)0);          // dwVersion
        bw.Write(lpObjectTableVa);   // lpObjectTable

        // Nombre de proyecto terminado en nulo
        byte[] nameBytes = Encoding.ASCII.GetBytes(projectName + "\0");
        bw.Write(nameBytes);
    }

    /// <summary>
    /// Escribe la estructura ObjectTable (60 bytes).
    /// </summary>
    public static void WriteObjectTable(
        BinaryWriter bw,
        Guid projectGuid,
        ushort objectCount,
        uint lpObjectArrayVa,
        uint lpProjectTitleVa,
        uint lpProjectNameVa)
    {
        bw.Write((uint)0);                     // lpHeapLink
        bw.Write((uint)0);                     // lpExecProj
        bw.Write((uint)0);                     // lpProjInfo2
        bw.Write((uint)0);                     // dwReserved
        bw.Write((uint)0);                     // dwNull
        bw.Write((uint)0);                     // lpProjObject
        bw.Write(projectGuid.ToByteArray());   // uuidProject (16 bytes)
        bw.Write((ushort)0);                   // wCompileType (0 = P-Code, 1 = Native)
        bw.Write(objectCount);                 // cObjects
        bw.Write(objectCount);                 // cCompiledObjects
        bw.Write((ushort)0);                   // wReserved2
        bw.Write(lpObjectArrayVa);             // lpObjectArray
        bw.Write(lpProjectTitleVa);            // szProjectTitle
        bw.Write((uint)0);                     // dwHelpContext
        bw.Write((uint)0);                     // szHelpFile
        bw.Write(lpProjectNameVa);             // szProjectName
    }

    /// <summary>
    /// Escribe un descriptor básico de objeto ObjectInfo para un Formulario o Módulo (36 bytes).
    /// </summary>
    public static void WriteObjectInfo(
        BinaryWriter bw,
        uint lpObjectNameVa,
        ushort objectType = 1) // 1 = Formulario, 2 = Módulo
    {
        bw.Write((uint)0);              // lpHeapLink
        bw.Write((uint)0);              // lpObjectDescriptor
        bw.Write((uint)0);              // dwReserved
        bw.Write((uint)0);              // dwNull
        bw.Write((uint)0);              // lpControlArray
        bw.Write(lpObjectNameVa);       // lpObjectName
        bw.Write((uint)0);              // lpMethods
        bw.Write(objectType);           // wObjectType
        bw.Write((ushort)0);            // wReserved2
        bw.Write((uint)0);              // dwNull2
    }
}
