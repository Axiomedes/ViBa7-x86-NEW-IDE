using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.ActiveX;

namespace ViBa7.ProjectModel.Compilation;

/// <summary>
/// Generador de manifiestos Side-by-Side (SxS) para habilitar despliegue y ejecución
/// aislada de controles ActiveX / OCX sin requerir registro previo en Windows (Reg-Free COM).
/// </summary>
public static class Vb6ManifestGenerator
{
    /// <summary>
    /// Genera el contenido XML del manifiesto de la aplicación para Reg-Free COM.
    /// </summary>
    public static string GenerateManifestXml(
        Vb6CompileOptions options,
        IEnumerable<OcxLibraryInfo>? ocxLibraries = null)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.AppendLine("<assembly xmlns=\"urn:schemas-microsoft-com:asm.v1\" manifestVersion=\"1.0\">");

        // Identidad de la aplicación
        string projName = !string.IsNullOrWhiteSpace(options.ProjectName) ? options.ProjectName : "Vb6Application";
        string version = $"{options.MajorVersion}.{options.MinorVersion}.{options.RevisionVersion}.0";
        sb.AppendLine($"  <assemblyIdentity type=\"win32\" name=\"{projName}\" version=\"{version}\" processorArchitecture=\"x86\" />");

        string desc = !string.IsNullOrWhiteSpace(options.FileDescription) ? options.FileDescription : projName;
        sb.AppendLine($"  <description>{desc}</description>");

        // Nivel de privilegios canónico de Windows
        sb.AppendLine("  <trustInfo xmlns=\"urn:schemas-microsoft-com:asm.v2\">");
        sb.AppendLine("    <security>");
        sb.AppendLine("      <requestedPrivileges>");
        sb.AppendLine("        <requestedExecutionLevel level=\"asInvoker\" uiAccess=\"false\" />");
        sb.AppendLine("      </requestedPrivileges>");
        sb.AppendLine("    </security>");
        sb.AppendLine("  </trustInfo>");

        // Entradas de aislamiento Reg-Free COM por cada librería OCX
        if (ocxLibraries != null)
        {
            foreach (var lib in ocxLibraries)
            {
                if (string.IsNullOrWhiteSpace(lib.FilePath)) continue;
                string fileName = Path.GetFileName(lib.FilePath);

                sb.AppendLine($"  <file name=\"{fileName}\">");

                // Declarar cada CoClass / Control OCX contenido en la TypeLib
                foreach (var ctrl in lib.Controls)
                {
                    string clsid = ctrl.Clsid;
                    if (!clsid.StartsWith('{')) clsid = "{" + clsid + "}";

                    string tlbid = lib.Guid;
                    if (!tlbid.StartsWith('{')) tlbid = "{" + tlbid + "}";

                    string progId = !string.IsNullOrWhiteSpace(lib.LibraryName) ? $"{lib.LibraryName}.{ctrl.Name}" : ctrl.Name;
                    string ctrlDesc = !string.IsNullOrWhiteSpace(ctrl.Description) ? ctrl.Description : ctrl.Name;

                    sb.Append("    <comClass");
                    sb.Append($" clsid=\"{clsid}\"");
                    sb.Append(" threadingModel=\"Apartment\"");
                    sb.Append($" tlbid=\"{tlbid}\"");
                    if (!string.IsNullOrWhiteSpace(progId))
                    {
                        sb.Append($" progid=\"{progId}\"");
                    }
                    if (!string.IsNullOrWhiteSpace(ctrlDesc))
                    {
                        sb.Append($" description=\"{ctrlDesc}\"");
                    }
                    sb.AppendLine(" />");
                }

                // Declarar la TypeLib asociada para marshaling e IDispatch
                string libTlbid = lib.Guid;
                if (!libTlbid.StartsWith('{')) libTlbid = "{" + libTlbid + "}";
                string tlbVer = $"{lib.MajorVersion}.{lib.MinorVersion}";

                sb.AppendLine($"    <typelib tlbid=\"{libTlbid}\" version=\"{tlbVer}\" helpdir=\"\" />");
                sb.AppendLine("  </file>");
            }
        }

        sb.AppendLine("</assembly>");
        return sb.ToString();
    }

    /// <summary>
    /// Escribe el archivo de manifiesto en disco junto al archivo ejecutable (.exe.manifest).
    /// </summary>
    public static string WriteManifestFile(
        string exePath,
        Vb6CompileOptions options,
        IEnumerable<OcxLibraryInfo>? ocxLibraries = null)
    {
        string manifestPath = $"{exePath}.manifest";
        string xml = GenerateManifestXml(options, ocxLibraries);
        File.WriteAllText(manifestPath, xml, Encoding.UTF8);
        return manifestPath;
    }
}
