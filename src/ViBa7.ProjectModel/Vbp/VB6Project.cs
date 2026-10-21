using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Documents;

namespace ViBa7.ProjectModel.Vbp;

public class VB6Project
{
    public string ProjectFilePath { get; set; } = string.Empty;
    public string ProjectDirectory => Path.GetDirectoryName(ProjectFilePath) ?? string.Empty;

    public string Name { get; set; } = string.Empty;
    public string ProjectType { get; set; } = "Exe";
    public string Startup { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string ExeName32 { get; set; } = string.Empty;

    /// <summary>
    /// Grupo de proyectos (.vbg) al que pertenece este proyecto, si aplica.
    /// </summary>
    public VB6ProjectGroup? ParentGroup { get; set; }

    /// <summary>
    /// Indica si este proyecto es el proyecto principal (proceso de inicio) dentro del grupo.
    /// </summary>
    public bool IsPrimaryProject => ParentGroup != null ? ParentGroup.IsPrimaryProject(this) : true;

    /// <summary>
    /// Indica si este proyecto es un proyecto secundario (subproceso / componente dependiente) dentro del grupo.
    /// </summary>
    public bool IsSecondaryProject => ParentGroup != null && ParentGroup.IsSecondaryProject(this);

    /// <summary>
    /// Rol clasificado del proyecto dentro de su grupo o de forma individual.
    /// </summary>
    public ProjectGroupRole GroupRole
    {
        get
        {
            if (ParentGroup == null) return ProjectGroupRole.Standalone;
            return IsPrimaryProject ? ProjectGroupRole.Primary : ProjectGroupRole.Secondary;
        }
    }

    /// <summary>
    /// DescripciÃ³n textual en espaÃ±ol del rol del proyecto para visualizaciÃ³n en el IDE y barra de estado.
    /// </summary>
    public string RoleDescription
    {
        get
        {
            return GroupRole switch
            {
                ProjectGroupRole.Primary => "Proyecto Principal (Inicio / EjecuciÃ³n)",
                ProjectGroupRole.Secondary => $"Proyecto Secundario (Subproceso / {ProjectType})",
                _ => "Proyecto Individual"
            };
        }
    }

    public List<ProjectItem> Items { get; } = [];
    public List<ProjectReference> References { get; } = [];
    public List<string> RawLines { get; } = [];
    public Dictionary<string, string> Settings { get; } = new(StringComparer.OrdinalIgnoreCase);

    public int MajorVer
    {
        get => int.TryParse(Settings.TryGetValue("MajorVer", out var v) ? v : "1", out var n) ? n : 1;
        set => Settings["MajorVer"] = value.ToString();
    }

    public int MinorVer
    {
        get => int.TryParse(Settings.TryGetValue("MinorVer", out var v) ? v : "0", out var n) ? n : 0;
        set => Settings["MinorVer"] = value.ToString();
    }

    public int RevisionVer
    {
        get => int.TryParse(Settings.TryGetValue("RevisionVer", out var v) ? v : "0", out var n) ? n : 0;
        set => Settings["RevisionVer"] = value.ToString();
    }

    public bool AutoIncrementVer
    {
        get => Settings.TryGetValue("AutoIncrementVer", out var v) && (v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase));
        set => Settings["AutoIncrementVer"] = value ? "1" : "0";
    }

    public string CompanyName
    {
        get => Settings.TryGetValue("CompanyName", out var v) ? v : string.Empty;
        set => Settings["CompanyName"] = value;
    }

    public string FileDescription
    {
        get => Settings.TryGetValue("FileDescription", out var v) ? v : Title;
        set => Settings["FileDescription"] = value;
    }

    public string LegalCopyright
    {
        get => Settings.TryGetValue("LegalCopyright", out var v) ? v : string.Empty;
        set => Settings["LegalCopyright"] = value;
    }

    /// <summary>
    /// Dialecto de lenguaje seleccionado para el proyecto (VB6 Clásico, VB7 Office, o ViBa7 Moderno).
    /// Si no está presente en el .vbp, por defecto es VB6 Clásico Estándar (retrocompatibilidad total).
    /// </summary>
    public Execution.VbLanguageDialect Dialect
    {
        get
        {
            if (Settings.TryGetValue("LanguageDialect", out var v))
            {
                if (string.Equals(v, "Vb7Office", StringComparison.OrdinalIgnoreCase) || string.Equals(v, "VB7", StringComparison.OrdinalIgnoreCase))
                    return Execution.VbLanguageDialect.Vb7Office;
                if (string.Equals(v, "ViBaModern", StringComparison.OrdinalIgnoreCase) || string.Equals(v, "Modern", StringComparison.OrdinalIgnoreCase))
                    return Execution.VbLanguageDialect.ViBaModern;
            }
            return Execution.VbLanguageDialect.Vb6Classic;
        }
        set
        {
            Settings["LanguageDialect"] = value switch
            {
                Execution.VbLanguageDialect.Vb7Office => "Vb7Office",
                Execution.VbLanguageDialect.ViBaModern => "ViBaModern",
                _ => "VB6"
            };
        }
    }

    public IEnumerable<ProjectItem> Forms => Items.FindAll(i => i.Type == ProjectItemType.Form);
    public IEnumerable<ProjectItem> Modules => Items.FindAll(i => i.Type == ProjectItemType.Module);
    public IEnumerable<ProjectItem> Classes => Items.FindAll(i => i.Type == ProjectItemType.Class);
    public IEnumerable<ProjectItem> UserControls => Items.FindAll(i => i.Type == ProjectItemType.UserControl);
    public IEnumerable<ProjectItem> PropertyPages => Items.FindAll(i => i.Type == ProjectItemType.PropertyPage);
    public IEnumerable<ProjectItem> ActiveXControls => Items.FindAll(i => i.Type == ProjectItemType.ActiveXControl);

    /// <summary>
    /// Bibliotecas OCX / TypeLibs cargadas e integradas en este proyecto.
    /// </summary>
    public List<ViBa7.ProjectModel.ActiveX.OcxLibraryInfo> LoadedOcxLibraries { get; } = [];

    /// <summary>
    /// Registra un control ActiveX en el proyecto.
    /// </summary>
    public void AddActiveXControl(ViBa7.ProjectModel.ActiveX.OcxLibraryInfo libInfo, string relativeOrAbsolutePath)
    {
        ArgumentNullException.ThrowIfNull(libInfo);

        if (!LoadedOcxLibraries.Any(l => string.Equals(l.Guid, libInfo.Guid, StringComparison.OrdinalIgnoreCase) || string.Equals(l.FilePath, libInfo.FilePath, StringComparison.OrdinalIgnoreCase)))
        {
            LoadedOcxLibraries.Add(libInfo);
        }

        string fileName = Path.GetFileName(relativeOrAbsolutePath);
        if (!Items.Any(i => i.Type == ProjectItemType.ActiveXControl && (string.Equals(i.AbsolutePath, libInfo.FilePath, StringComparison.OrdinalIgnoreCase) || string.Equals(i.Name, libInfo.LibraryName, StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetFileName(i.RelativePath), fileName, StringComparison.OrdinalIgnoreCase))))
        {
            string guidFormatted = libInfo.Guid;
            if (!guidFormatted.StartsWith('{')) guidFormatted = "{" + guidFormatted + "}";
            string rawLine = $"Object={guidFormatted}#{libInfo.MajorVersion}.{libInfo.MinorVersion}#0; {fileName}";

            Items.Add(new ProjectItem
            {
                Type = ProjectItemType.ActiveXControl,
                Name = libInfo.LibraryName,
                RelativePath = relativeOrAbsolutePath,
                AbsolutePath = libInfo.FilePath,
                RawLine = rawLine,
                Guid = libInfo.Guid,
                MajorVersion = libInfo.MajorVersion.ToString(),
                MinorVersion = libInfo.MinorVersion.ToString(),
                FileExists = true
            });

            RawLines.Add(rawLine);
        }
    }

    /// <summary>
    /// Elimina un control ActiveX del proyecto en memoria y de la lista de librerÃ­as cargadas.
    /// </summary>
    public void RemoveActiveXControl(string libraryNameOrPathOrGuid)
    {
        if (string.IsNullOrWhiteSpace(libraryNameOrPathOrGuid)) return;

        string cleanGuid = libraryNameOrPathOrGuid.Trim('{', '}');
        string fileName = Path.GetFileName(libraryNameOrPathOrGuid);

        LoadedOcxLibraries.RemoveAll(l =>
            string.Equals(l.LibraryName, libraryNameOrPathOrGuid, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(l.FilePath, libraryNameOrPathOrGuid, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(l.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase));

        Items.RemoveAll(i => i.Type == ProjectItemType.ActiveXControl && (
            string.Equals(i.Name, libraryNameOrPathOrGuid, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(i.AbsolutePath, libraryNameOrPathOrGuid, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetFileName(i.RelativePath), fileName, StringComparison.OrdinalIgnoreCase) ||
            (!string.IsNullOrEmpty(i.Guid) && string.Equals(i.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase))));
    }

    /// <summary>
    /// Elimina una referencia COM / TypeLib del proyecto en memoria.
    /// </summary>
    public void RemoveReference(string guidOrPath)
    {
        if (string.IsNullOrWhiteSpace(guidOrPath)) return;

        string cleanGuid = guidOrPath.Trim('{', '}');
        References.RemoveAll(r =>
            (!string.IsNullOrEmpty(r.Guid) && string.Equals(r.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase)) ||
            string.Equals(r.FilePath, guidOrPath, StringComparison.OrdinalIgnoreCase));

        Items.RemoveAll(i => i.Type == ProjectItemType.Reference && (
            (!string.IsNullOrEmpty(i.Guid) && string.Equals(i.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase)) ||
            string.Equals(i.AbsolutePath, guidOrPath, StringComparison.OrdinalIgnoreCase)));
    }

    /// <summary>
    /// Agrega una referencia COM / TypeLib al proyecto si no existe previamente.
    /// </summary>
    public void AddReference(ProjectReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);

        string cleanGuid = reference.Guid.Trim('{', '}');
        if (!References.Any(r => (!string.IsNullOrEmpty(r.Guid) && string.Equals(r.Guid.Trim('{', '}'), cleanGuid, StringComparison.OrdinalIgnoreCase)) ||
                                 (!string.IsNullOrEmpty(r.FilePath) && string.Equals(r.FilePath, reference.FilePath, StringComparison.OrdinalIgnoreCase))))
        {
            if (string.IsNullOrEmpty(reference.RawLine))
            {
                string guidFormatted = reference.Guid;
                if (!guidFormatted.StartsWith('{')) guidFormatted = "{" + guidFormatted + "}";
                reference.RawLine = $"Reference=*\\G{guidFormatted}#{reference.MajorVersion}.{reference.MinorVersion}#{reference.Lcid}#{reference.FilePath}#{reference.Description}";
            }

            References.Add(reference);

            Items.Add(new ProjectItem
            {
                Type = ProjectItemType.Reference,
                Name = !string.IsNullOrWhiteSpace(reference.Description) ? reference.Description : Path.GetFileNameWithoutExtension(reference.FilePath),
                RelativePath = reference.FilePath,
                AbsolutePath = reference.FilePath,
                RawLine = reference.RawLine,
                Guid = reference.Guid,
                MajorVersion = reference.MajorVersion,
                MinorVersion = reference.MinorVersion,
                FileExists = File.Exists(reference.FilePath),
                ContainingProject = this
            });
        }
    }

    /// <summary>
    /// Reemplaza la lista completa de referencias del proyecto manteniendo el orden de prioridad especificado.
    /// </summary>
    public void SetReferences(IEnumerable<ProjectReference> newReferences)
    {
        ArgumentNullException.ThrowIfNull(newReferences);

        References.Clear();
        Items.RemoveAll(i => i.Type == ProjectItemType.Reference);

        foreach (var r in newReferences)
        {
            AddReference(r);
        }
    }

    private readonly Dictionary<string, FormDesignModel> _cachedUserControlModels = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Busca un ProjectItem de tipo UserControl por su nombre, tipo o nombre de archivo.
    /// </summary>
    public ProjectItem? FindUserControl(string typeOrName)
    {
        if (string.IsNullOrWhiteSpace(typeOrName)) return null;

        string cleanName = typeOrName.Trim();
        if (cleanName.StartsWith($"{Name}.", StringComparison.OrdinalIgnoreCase))
        {
            cleanName = cleanName[(Name.Length + 1)..];
        }

        return UserControls.FirstOrDefault(u =>
            string.Equals(u.Name, cleanName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetFileNameWithoutExtension(u.RelativePath), cleanName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetFileName(u.RelativePath), cleanName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals($"{Name}.{u.Name}", typeOrName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Obtiene y parsea el modelo de diseÃ±o (FormDesignModel) de un UserControl del proyecto.
    /// </summary>
    public FormDesignModel? GetUserControlDesignModel(string typeOrName)
    {
        var item = FindUserControl(typeOrName);
        if (item == null) return null;

        string key = item.AbsolutePath ?? item.Name;
        if (_cachedUserControlModels.TryGetValue(key, out var cached))
        {
            return cached;
        }

        if (!string.IsNullOrEmpty(item.AbsolutePath) && File.Exists(item.AbsolutePath))
        {
            try
            {
                var doc = SafeFileStorage.LoadDocument(item.AbsolutePath);
                if (!string.IsNullOrWhiteSpace(doc.Header))
                {
                    var model = FormDesignParser.Parse(doc.Header);
                    string? dir = Path.GetDirectoryName(item.AbsolutePath);
                    if (!string.IsNullOrEmpty(dir))
                    {
                        FormDesignParser.ResolveControlPictures(model, dir);
                    }

                    if (!string.IsNullOrWhiteSpace(doc.Code))
                    {
                        model.ExposedProperties = UserControlPropertyExtractor.ExtractProperties(doc.Code);
                    }

                    _cachedUserControlModels[key] = model;
                    return model;
                }
            }
            catch
            {
                // Fallback silencioso si no se puede leer
            }
        }

        return null;
    }

    /// <summary>
    /// Invalida la cachÃ© del modelo visual de un UserControl (por ejemplo cuando se guardan cambios en el editor).
    /// </summary>
    public void InvalidateUserControlCache(string? keyOrPath = null)
    {
        if (string.IsNullOrEmpty(keyOrPath))
        {
            _cachedUserControlModels.Clear();
        }
        else
        {
            _cachedUserControlModels.Remove(keyOrPath);
        }
    }
}

