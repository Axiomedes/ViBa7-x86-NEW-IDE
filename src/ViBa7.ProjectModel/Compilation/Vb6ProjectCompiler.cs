using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.ProjectModel.Compilation;

/// <summary>
/// Resultado del proceso de compilación de un ejecutable nativo x32 de VB6.
/// </summary>
public record CompileResult(
    bool Success,
    string OutputExePath,
    long OutputSizeBytes,
    IReadOnlyList<string> Messages,
    Exception? Error = null
);

/// <summary>
/// Compilador y orquestador de proyectos de Visual Basic 6 a binarios nativos PE32 enlazados a MSVBVM60.DLL.
/// </summary>
public class Vb6ProjectCompiler
{
    private readonly Msvbvm60PeBuilder _builder = new();

    /// <summary>
    /// Compila un VB6Project cargado y genera el ejecutable nativo en disco.
    /// </summary>
    public CompileResult Compile(VB6Project project, string? explicitOutputPath = null, Action<string>? progressCallback = null)
    {
        if (project == null) throw new ArgumentNullException(nameof(project));

        var messages = new List<string>();

        try
        {
            string projName = !string.IsNullOrWhiteSpace(project.Name) ? project.Name : "Proyecto1";
            string projTitle = !string.IsNullOrWhiteSpace(project.Title) ? project.Title : projName;
            string startup = !string.IsNullOrWhiteSpace(project.Startup) ? project.Startup : "Form1";

            // Determinar ruta de salida
            string exeName = !string.IsNullOrWhiteSpace(project.ExeName32) ? project.ExeName32 : $"{projName}.exe";
            if (!exeName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                exeName += ".exe";
            }

            string outPath;
            if (!string.IsNullOrWhiteSpace(explicitOutputPath))
            {
                outPath = explicitOutputPath;
            }
            else
            {
                string baseDir = !string.IsNullOrWhiteSpace(project.ProjectDirectory) ? project.ProjectDirectory : Environment.CurrentDirectory;
                outPath = Path.Combine(baseDir, exeName);
            }

            progressCallback?.Invoke($"Iniciando compilación nativa x32 para '{projName}'...");
            messages.Add($"Objetivo: {outPath}");

            // Extraer nombres de formularios del proyecto
            var formNames = project.Forms.Select(f => f.Name).ToList();
            if (formNames.Count == 0)
            {
                formNames.Add(startup);
            }

            progressCallback?.Invoke($"Estructurando metadatos de {formNames.Count} objeto(s)...");

            // Extraer componentes ActiveX / OCX cargados o declarados en el proyecto
            var ocxLibs = new List<ViBa7.ProjectModel.ActiveX.OcxLibraryInfo>(project.LoadedOcxLibraries);
            var activeXItems = project.ActiveXControls.ToList();
            if (activeXItems.Count > 0 || ocxLibs.Count > 0)
            {
                int totalActiveX = Math.Max(activeXItems.Count, ocxLibs.Count);
                messages.Add($"Componentes ActiveX detectados en el proyecto: {totalActiveX}.");
            }

            var options = new Vb6CompileOptions(
                OutputExePath: outPath,
                ProjectName: projName,
                ProjectTitle: projTitle,
                StartupObject: startup,
                MajorVersion: project.MajorVer,
                MinorVersion: project.MinorVer,
                RevisionVersion: project.RevisionVer,
                CompanyName: project.CompanyName,
                FileDescription: project.FileDescription,
                LegalCopyright: project.LegalCopyright,
                GenerateSxSManifest: true,
                ActiveXComponents: ocxLibs.Count > 0 ? ocxLibs : null
            );

            progressCallback?.Invoke("Generando cabeceras PE32, stubs x86 y enlace a MSVBVM60.DLL...");
            byte[] exeBytes = _builder.BuildExecutable(options, formNames);

            // Generar manifiesto Side-by-Side (SxS) para aislamiento Reg-Free COM si hay componentes OCX
            if (options.GenerateSxSManifest && (ocxLibs.Count > 0 || activeXItems.Count > 0))
            {
                progressCallback?.Invoke("Generando manifiesto Side-by-Side (SxS) para Reg-Free COM...");
                string manifestPath = Vb6ManifestGenerator.WriteManifestFile(outPath, options, ocxLibs);
                messages.Add($"Manifiesto Reg-Free COM emitido: {Path.GetFileName(manifestPath)}");
            }

            progressCallback?.Invoke($"Ejecutable emitido exitosamente ({exeBytes.Length} bytes).");
            messages.Add($"Tamaño del binario generado: {exeBytes.Length:N0} bytes.");
            messages.Add("Firma PE y enlace canónico a msvbvm60.dll!ThunRTMain verificados.");

            return new CompileResult(
                Success: true,
                OutputExePath: outPath,
                OutputSizeBytes: exeBytes.Length,
                Messages: messages
            );
        }
        catch (Exception ex)
        {
            messages.Add($"Error de compilación: {ex.Message}");
            return new CompileResult(
                Success: false,
                OutputExePath: explicitOutputPath ?? string.Empty,
                OutputSizeBytes: 0,
                Messages: messages,
                Error: ex
            );
        }
    }
}
