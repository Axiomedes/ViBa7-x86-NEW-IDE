using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.ProjectModel.Compilation;

/// <summary>
/// Opciones de configuración para la compilación de ejecutables nativos x32 de VB6.
/// </summary>
public record Vb6CompileOptions(
    string OutputExePath,
    string ProjectName = "Proyecto1",
    string ProjectTitle = "Proyecto1",
    string StartupObject = "Form1",
    int MajorVersion = 1,
    int MinorVersion = 0,
    int RevisionVersion = 0,
    string CompanyName = "",
    string FileDescription = "",
    string LegalCopyright = "",
    bool GenerateSxSManifest = true,
    IReadOnlyList<ViBa7.ProjectModel.ActiveX.OcxLibraryInfo>? ActiveXComponents = null
);

/// <summary>
/// Constructor de ejecutables binarios PE32 (x86) de Visual Basic 6 enlazados a MSVBVM60.DLL.
/// Genera el ejecutable nativo autónomo sin requerir VB6.EXE de Microsoft instalado.
/// </summary>
public class Msvbvm60PeBuilder
{
    private const uint DefaultImageBase = 0x00400000;
    private const uint SectionTextRva = 0x1000;
    private const uint SectionDataRva = 0x2000;

    /// <summary>
    /// Compila y genera el archivo ejecutable binario de 32 bits a partir de las opciones dadas.
    /// </summary>
    public byte[] BuildExecutable(Vb6CompileOptions options, IEnumerable<string>? formNames = null)
    {
        if (options == null) throw new ArgumentNullException(nameof(options));

        var peBuilder = new Pe32Builder
        {
            ImageBase = DefaultImageBase,
            Subsystem = Pe32Constants.SubsystemWindowsGui
        };

        // =========================================================================
        // 1. SECCIÓN .TEXT (Código ejecutable x86 + Tabla de Importaciones IAT)
        // =========================================================================
        var textSection = new Pe32Section(".text", Pe32Constants.SectionCharacteristicsText);
        var textMs = textSection.DataStream;
        var textBw = new BinaryWriter(textMs);

        // A. Entry Point x86 Stub (16 bytes):
        // push offset VBHeader (0x68, dd VBHeader_VA)
        // call dword ptr [IAT_ThunRTMain] (0xFF, 0x15, dd IAT_ThunRTMain_VA)
        // ret (0xC3) + NOP padding
        uint entryPointRva = SectionTextRva;
        uint vbHeaderVa = DefaultImageBase + SectionDataRva; // VBHeader estará al inicio de .data

        // Offset en .text donde colocaremos la IAT (ej. tras el stub, en offset 0x20)
        uint iatRva = SectionTextRva + 0x0020;
        uint iatThunRTMainVa = DefaultImageBase + iatRva;

        textBw.Write((byte)0x68);                 // PUSH imm32
        textBw.Write(vbHeaderVa);                 // [VA de VBHeader en .data]
        textBw.Write((byte)0xFF); textBw.Write((byte)0x15); // CALL dword ptr [...]
        textBw.Write(iatThunRTMainVa);            // [VA de IAT_ThunRTMain en .text]
        textBw.Write((byte)0xC3);                 // RET
        textBw.Write(new byte[5]);                // Padding NOPs/ceros hasta offset 16 (0x10)
        textBw.Write(new byte[16]);               // Padding hasta offset 32 (0x20)

        // B. IAT (Import Address Table) en offset 0x0020 de .text
        // Contiene 1 entrada: RVA de IMAGE_IMPORT_BY_NAME de ThunRTMain, seguida de null
        uint hintNameThunRTMainRva = SectionTextRva + 0x0070; // Colocaremos el Hint/Name más adelante
        textBw.Write(hintNameThunRTMainRva);      // IAT[0] -> RVA de "ThunRTMain"
        textBw.Write((uint)0);                    // IAT[1] -> Null terminator

        // C. Import Directory Table (IMAGE_IMPORT_DESCRIPTOR) en offset 0x0028 de .text
        uint importDirectoryRva = SectionTextRva + 0x0028;
        uint intThunRTMainRva = SectionTextRva + 0x0050; // INT (Import Name Table / OriginalFirstThunk)
        uint dllNameMsvbvmRva = SectionTextRva + 0x0060; // Cadena "MSVBVM60.DLL\0"

        textBw.Write(intThunRTMainRva);           // OriginalFirstThunk (RVA a INT)
        textBw.Write((uint)0);                    // TimeDateStamp
        textBw.Write((uint)0);                    // ForwarderChain
        textBw.Write(dllNameMsvbvmRva);           // Name (RVA a "MSVBVM60.DLL")
        textBw.Write(iatRva);                     // FirstThunk (RVA a IAT)

        // Descriptor nulo para terminar la Import Directory Table (20 bytes de cero)
        textBw.Write(new byte[20]);

        // D. INT (Import Name Table) en offset 0x0050 de .text (8 bytes)
        textBw.Write(hintNameThunRTMainRva);      // INT[0] -> RVA de "ThunRTMain"
        textBw.Write((uint)0);                    // INT[1] -> Null terminator
        textBw.Write(new byte[8]);                // Padding hasta offset 0x60

        // E. Cadena de DLL: "MSVBVM60.DLL\0" en offset 0x0060 de .text (16 bytes)
        byte[] dllNameBytes = Encoding.ASCII.GetBytes("MSVBVM60.DLL\0");
        textBw.Write(dllNameBytes);
        int padDll = 16 - dllNameBytes.Length;
        if (padDll > 0) textBw.Write(new byte[padDll]);

        // F. Hint/Name: [Hint: 0x0000 (2 bytes)] + ["ThunRTMain\0" (11 bytes)] en offset 0x0070 de .text
        textBw.Write((ushort)0);                  // Hint = 0
        byte[] funcNameBytes = Encoding.ASCII.GetBytes("ThunRTMain\0");
        textBw.Write(funcNameBytes);
        textSection.PadToAlignment(16);

        // =========================================================================
        // 2. SECCIÓN .DATA (Estructuras Nativas de VB6: VBHeader, ProjInfo, etc.)
        // =========================================================================
        var dataSection = new Pe32Section(".data", Pe32Constants.SectionCharacteristicsData);
        var dataMs = dataSection.DataStream;
        var dataBw = new BinaryWriter(dataMs);

        var formsList = new List<string>(formNames ?? new[] { options.StartupObject });
        if (formsList.Count == 0) formsList.Add("Form1");

        // Layout de .data:
        // Offset 0x0000: VBHeader (108 bytes)
        // Offset 0x0080: ProjInfo (44 bytes)
        // Offset 0x00C0: ProjectData (8 bytes + len(ProjectName) + 1)
        // Offset 0x0100: ObjectTable (60 bytes)
        // Offset 0x0140: Array de punteros a ObjectInfo (4 bytes * N)
        // Offset 0x0180+: ObjectInfo para cada objeto (36 bytes cada uno)
        // Cadenas ASCII al final

        uint projInfoVa = DefaultImageBase + SectionDataRva + 0x0080;
        uint projectDataVa = DefaultImageBase + SectionDataRva + 0x00C0;
        uint objectTableVa = DefaultImageBase + SectionDataRva + 0x0100;
        uint objectArrayVa = DefaultImageBase + SectionDataRva + 0x0140;

        // A. Escribir VBHeader (108 bytes) en offset 0
        ushort extCount = (ushort)(options.ActiveXComponents?.Count ?? 0);
        Vb6BinaryStructures.WriteVbHeader(
            dataBw,
            lpProjInfoVa: projInfoVa,
            lpSubMainVa: 0,
            dwLcid: 0x00000C0A, // Español (o 0x0409)
            formCount: (ushort)formsList.Count,
            externalComponentCount: extCount);

        // Padding hasta 0x0080
        dataSection.PadToAlignment(128);

        // B. Escribir ProjInfo (44 bytes) en offset 0x0080
        Vb6BinaryStructures.WriteProjInfo(
            dataBw,
            lpProjectDataVa: projectDataVa,
            lpObjectTableVa: objectTableVa,
            lpCodeStartVa: DefaultImageBase + SectionTextRva,
            lpCodeEndVa: DefaultImageBase + SectionTextRva + (uint)textMs.Length,
            dwDataSize: 0x1000);

        // Padding hasta 0x00C0
        dataSection.PadToAlignment(64);

        // C. Escribir ProjectData en offset 0x00C0
        Vb6BinaryStructures.WriteProjectData(
            dataBw,
            lpObjectTableVa: objectTableVa,
            projectName: options.ProjectName);

        // Padding hasta 0x0100
        dataSection.PadToAlignment(64);

        // Reservar espacio para ObjectTable (offset 0x0100)
        // Calcularemos las direcciones de cadenas y descriptores
        uint stringPoolOffset = 0x0200 + (uint)(formsList.Count * 64);
        uint projectNameVa = DefaultImageBase + SectionDataRva + stringPoolOffset;
        uint projectTitleVa = projectNameVa + (uint)options.ProjectName.Length + 1;

        // D. Escribir ObjectTable en offset 0x0100
        Vb6BinaryStructures.WriteObjectTable(
            dataBw,
            projectGuid: Guid.NewGuid(),
            objectCount: (ushort)formsList.Count,
            lpObjectArrayVa: objectArrayVa,
            lpProjectTitleVa: projectTitleVa,
            lpProjectNameVa: projectNameVa);

        // Padding hasta 0x0140
        dataSection.PadToAlignment(64);

        // E. Array de punteros a ObjectInfo en offset 0x0140
        uint objectInfoBaseOffset = 0x0160;
        for (int i = 0; i < formsList.Count; i++)
        {
            uint objInfoVa = DefaultImageBase + SectionDataRva + objectInfoBaseOffset + (uint)(i * 48);
            dataBw.Write(objInfoVa);
        }

        // Padding hasta 0x0160
        dataSection.PadToAlignment(32);

        // F. Descriptores ObjectInfo
        uint currentStrOffset = projectTitleVa + (uint)options.ProjectTitle.Length + 1;
        for (int i = 0; i < formsList.Count; i++)
        {
            string fName = formsList[i];
            uint nameVa = currentStrOffset;
            currentStrOffset += (uint)fName.Length + 1;

            Vb6BinaryStructures.WriteObjectInfo(
                dataBw,
                lpObjectNameVa: nameVa,
                objectType: 1); // 1 = Formulario
        }

        // G. Cadenas terminadas en nulo (String Pool)
        dataBw.Write(Encoding.ASCII.GetBytes(options.ProjectName + "\0"));
        dataBw.Write(Encoding.ASCII.GetBytes(options.ProjectTitle + "\0"));
        foreach (var fName in formsList)
        {
            dataBw.Write(Encoding.ASCII.GetBytes(fName + "\0"));
        }

        dataSection.PadToAlignment(16);

        // =========================================================================
        // 3. ENSAMBLAR ARCHIVO PE32
        // =========================================================================
        peBuilder.EntryPointRva = entryPointRva;
        peBuilder.ImportDirectoryRva = importDirectoryRva;
        peBuilder.ImportDirectorySize = 40; // 2 descriptores (20 bytes cada uno)
        peBuilder.IatRva = iatRva;
        peBuilder.IatSize = 8;  // 1 entrada + null

        peBuilder.AddSection(textSection);
        peBuilder.AddSection(dataSection);

        byte[] exeBytes = peBuilder.Build();

        // Escribir en disco si se proporcionó una ruta de salida
        if (!string.IsNullOrWhiteSpace(options.OutputExePath))
        {
            string? dir = Path.GetDirectoryName(options.OutputExePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllBytes(options.OutputExePath, exeBytes);
        }

        return exeBytes;
    }
}
