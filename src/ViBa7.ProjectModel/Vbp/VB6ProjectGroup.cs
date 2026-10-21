using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.ProjectModel.Vbp;

/// <summary>
/// Representa un Grupo de Proyectos de Visual Basic 6.0 (.vbg).
/// Gestiona la clasificación canónica entre el Proyecto Principal (proceso de inicio)
/// y los Proyectos Secundarios (que se ejecutan como subprocesos o bibliotecas dependientes).
/// </summary>
public class VB6ProjectGroup
{
    public string GroupFilePath { get; set; } = string.Empty;
    public string GroupDirectory => Path.GetDirectoryName(GroupFilePath) ?? string.Empty;
    public string Name { get; set; } = string.Empty;
    public string? StartupProjectPath { get; set; }
    public List<VB6Project> Projects { get; } = [];

    /// <summary>
    /// Proyecto Principal del grupo: aquel designado para iniciar la ejecución del proceso (StartupProject).
    /// </summary>
    public VB6Project? PrimaryProject
    {
        get => StartupProject;
        set
        {
            if (value != null)
            {
                SetPrimaryProject(value);
            }
        }
    }

    /// <summary>
    /// Proyecto de inicio canónico de VB6 (sinónimo canónico de PrimaryProject).
    /// Si no coincide por ruta exacta, prefiere el primer proyecto de tipo Exe.
    /// </summary>
    public VB6Project? StartupProject
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(StartupProjectPath))
            {
                var match = Projects.FirstOrDefault(p =>
                    string.Equals(p.ProjectFilePath, StartupProjectPath, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(Path.GetFileName(p.ProjectFilePath), Path.GetFileName(StartupProjectPath), StringComparison.OrdinalIgnoreCase));
                if (match != null) return match;
            }

            // Si no se especificó o no se encontró por ruta, preferir un ejecutable (.exe)
            var exeProj = Projects.FirstOrDefault(p => p.ProjectType.Equals("Exe", StringComparison.OrdinalIgnoreCase));
            return exeProj ?? Projects.FirstOrDefault();
        }
    }

    /// <summary>
    /// Lista de proyectos secundarios del grupo (aquellos que se ejecutan como subproceso / componente dependiente del principal).
    /// </summary>
    public IReadOnlyList<VB6Project> SecondaryProjects =>
        Projects.Where(p => !IsPrimaryProject(p)).ToList();

    /// <summary>
    /// Primer proyecto secundario del grupo (típico en la arquitectura de 2 proyectos VB6: Exe Principal + Control/DLL Secundario).
    /// </summary>
    public VB6Project? SecondaryProject => SecondaryProjects.FirstOrDefault();

    /// <summary>
    /// Determina si un proyecto específico es el proyecto principal (proceso de inicio).
    /// </summary>
    public bool IsPrimaryProject(VB6Project project)
    {
        if (project == null) return false;
        var primary = StartupProject;
        if (primary == null) return false;

        return ReferenceEquals(primary, project) ||
               string.Equals(primary.ProjectFilePath, project.ProjectFilePath, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Determina si un proyecto específico es un proyecto secundario (subproceso del principal).
    /// </summary>
    public bool IsSecondaryProject(VB6Project project)
    {
        if (project == null) return false;
        return Projects.Contains(project) && !IsPrimaryProject(project);
    }

    /// <summary>
    /// Establece un proyecto como el Proyecto Principal (proceso de inicio) del grupo.
    /// Actualiza la propiedad StartupProjectPath y los roles de todos los proyectos hijos.
    /// </summary>
    public void SetPrimaryProject(VB6Project project)
    {
        ArgumentNullException.ThrowIfNull(project);

        if (!Projects.Contains(project))
        {
            Projects.Add(project);
        }

        StartupProjectPath = project.ProjectFilePath;
        foreach (var p in Projects)
        {
            p.ParentGroup = this;
        }
    }

    /// <summary>
    /// Establece el proyecto principal por su ruta de archivo de proyecto (.vbp).
    /// </summary>
    public void SetPrimaryProjectByPath(string vbpPath)
    {
        if (string.IsNullOrWhiteSpace(vbpPath)) return;

        var found = Projects.FirstOrDefault(p =>
            string.Equals(p.ProjectFilePath, vbpPath, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetFileName(p.ProjectFilePath), Path.GetFileName(vbpPath), StringComparison.OrdinalIgnoreCase));

        if (found != null)
        {
            SetPrimaryProject(found);
        }
        else
        {
            StartupProjectPath = vbpPath;
        }
    }

    /// <summary>
    /// Busca el modelo de diseño (FormDesignModel) de un UserControl en todos los proyectos del grupo,
    /// consultando primero el proyecto principal y luego los proyectos secundarios (subprocesos de controles).
    /// </summary>
    public FormDesignModel? GetUserControlDesignModel(string typeOrName)
    {
        if (string.IsNullOrWhiteSpace(typeOrName)) return null;

        // 1. Buscar en el proyecto principal
        var primary = PrimaryProject;
        var model = primary?.GetUserControlDesignModel(typeOrName);
        if (model != null) return model;

        // 2. Buscar en los proyectos secundarios (subprocesos de controles)
        foreach (var sec in SecondaryProjects)
        {
            model = sec.GetUserControlDesignModel(typeOrName);
            if (model != null) return model;
        }

        return null;
    }

    /// <summary>
    /// Busca un ProjectItem de UserControl en todos los proyectos del grupo.
    /// </summary>
    public ProjectItem? FindUserControl(string typeOrName)
    {
        if (string.IsNullOrWhiteSpace(typeOrName)) return null;

        var primary = PrimaryProject;
        var item = primary?.FindUserControl(typeOrName);
        if (item != null) return item;

        foreach (var sec in SecondaryProjects)
        {
            item = sec.FindUserControl(typeOrName);
            if (item != null) return item;
        }

        return null;
    }

    /// <summary>
    /// Invalida la caché del modelo visual de controles de usuario en todos los proyectos del grupo.
    /// </summary>
    public void InvalidateUserControlCache(string? keyOrPath = null)
    {
        foreach (var p in Projects)
        {
            p.InvalidateUserControlCache(keyOrPath);
        }
    }

    /// <summary>
    /// Serializa y guarda el archivo de grupo de proyectos (.vbg) en disco con la sintaxis canónica de VB6.
    /// </summary>
    public void Save(string? targetFilePath = null)
    {
        string path = !string.IsNullOrWhiteSpace(targetFilePath) ? targetFilePath : GroupFilePath;
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("No se ha especificado la ruta para guardar el grupo de proyectos.");
        }

        string groupDir = Path.GetDirectoryName(path) ?? string.Empty;
        var sb = new StringBuilder();
        sb.AppendLine("VBGROUP 5.0");

        var primary = PrimaryProject;
        if (primary != null)
        {
            string relPrimary = Path.GetRelativePath(groupDir, primary.ProjectFilePath);
            sb.AppendLine($"StartupProject={relPrimary}");
        }

        foreach (var sec in SecondaryProjects)
        {
            string relSec = Path.GetRelativePath(groupDir, sec.ProjectFilePath);
            sb.AppendLine($"Project={relSec}");
        }

        File.WriteAllText(path, sb.ToString(), Encoding.GetEncoding(1252));
        GroupFilePath = Path.GetFullPath(path);
    }
}
