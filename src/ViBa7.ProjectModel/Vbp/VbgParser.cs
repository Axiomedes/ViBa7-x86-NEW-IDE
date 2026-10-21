using System;
using System.IO;
using System.Text;
using ViBa7.ProjectModel.Documents;

namespace ViBa7.ProjectModel.Vbp;

public static class VbgParser
{
    public static VB6ProjectGroup Parse(string vbgFilePath, Encoding? fallbackEncoding = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(vbgFilePath);

        if (!File.Exists(vbgFilePath))
        {
            throw new FileNotFoundException($"El archivo de grupo de proyectos no existe: {vbgFilePath}", vbgFilePath);
        }

        string groupDir = Path.GetDirectoryName(vbgFilePath) ?? string.Empty;
        var doc = SafeFileStorage.LoadDocument(vbgFilePath, fallbackEncoding);

        var group = new VB6ProjectGroup
        {
            GroupFilePath = Path.GetFullPath(vbgFilePath),
            Name = Path.GetFileNameWithoutExtension(vbgFilePath)
        };

        using var reader = new StringReader(doc.Text);
        string? line;

        while ((line = reader.ReadLine()) != null)
        {
            string trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed)) continue;

            // Ignorar encabezado de grupo o comentarios
            if (trimmed.StartsWith("VBGROUP", StringComparison.OrdinalIgnoreCase)) continue;
            if (trimmed.StartsWith('\'') || trimmed.StartsWith(';')) continue;

            int eqIndex = trimmed.IndexOf('=');
            if (eqIndex <= 0) continue;

            string key = trimmed[..eqIndex].Trim();
            string value = trimmed[(eqIndex + 1)..].Trim().Trim('"');

            if (string.IsNullOrWhiteSpace(value)) continue;

            string fullVbpPath = VbpParser.ResolvePath(groupDir, value);

            switch (key.ToUpperInvariant())
            {
                case "STARTUPPROJECT":
                    group.StartupProjectPath = fullVbpPath;
                    LoadAndAddProject(group, fullVbpPath, fallbackEncoding);
                    break;

                case "PROJECT":
                    LoadAndAddProject(group, fullVbpPath, fallbackEncoding);
                    break;
            }
        }

        foreach (var p in group.Projects)
        {
            p.ParentGroup = group;
        }

        // Si no se especificó un StartupProject explícito en el .vbg, clasificar automáticamente:
        // En VB6 se prefiere el proyecto estándar ejecutable (Type=Exe); si no, el primer proyecto.
        if (string.IsNullOrWhiteSpace(group.StartupProjectPath) && group.Projects.Count > 0)
        {
            var exeProj = group.Projects.FirstOrDefault(p => p.ProjectType.Equals("Exe", StringComparison.OrdinalIgnoreCase));
            var primary = exeProj ?? group.Projects[0];
            group.StartupProjectPath = primary.ProjectFilePath;
        }

        return group;
    }

    private static void LoadAndAddProject(VB6ProjectGroup group, string fullVbpPath, Encoding? fallbackEncoding)
    {
        // Evitar agregar el mismo proyecto dos veces
        if (group.Projects.Exists(p => string.Equals(p.ProjectFilePath, fullVbpPath, StringComparison.OrdinalIgnoreCase)))
        {
            return;
        }

        if (File.Exists(fullVbpPath))
        {
            try
            {
                var project = VbpParser.Parse(fullVbpPath, fallbackEncoding);
                project.ParentGroup = group;
                group.Projects.Add(project);
            }
            catch
            {
                // Si falla el parseo de uno de los proyectos del grupo, añadir como proyecto no disponible
                var placeholder = new VB6Project
                {
                    ProjectFilePath = fullVbpPath,
                    Name = Path.GetFileNameWithoutExtension(fullVbpPath),
                    ParentGroup = group
                };
                group.Projects.Add(placeholder);
            }
        }
        else
        {
            // Proyecto referenciado no encontrado en disco
            var placeholder = new VB6Project
            {
                ProjectFilePath = fullVbpPath,
                Name = Path.GetFileNameWithoutExtension(fullVbpPath),
                ParentGroup = group
            };
            group.Projects.Add(placeholder);
        }
    }
}
