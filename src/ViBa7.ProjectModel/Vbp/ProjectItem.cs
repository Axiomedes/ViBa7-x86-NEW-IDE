using System;
using System.IO;

namespace ViBa7.ProjectModel.Vbp;

public class ProjectItem
{
    public int LineNumber { get; set; }
    public string RawLine { get; set; } = string.Empty;
    public ProjectItemType Type { get; set; } = ProjectItemType.Unknown;
    public string Name { get; set; } = string.Empty;
    public string RelativePath { get; set; } = string.Empty;
    public string AbsolutePath { get; set; } = string.Empty;
    public bool FileExists { get; set; } = true;
    public string? Guid { get; set; }
    public string? MajorVersion { get; set; }
    public string? MinorVersion { get; set; }
    public VB6Project? ContainingProject { get; set; }

    public bool IsEditable => Type switch
    {
        ProjectItemType.Form => true,
        ProjectItemType.Module => true,
        ProjectItemType.Class => true,
        ProjectItemType.UserControl => true,
        ProjectItemType.PropertyPage => true,
        ProjectItemType.UserDocument => true,
        ProjectItemType.Designer => true,
        _ => false
    };

    public string DisplayIcon => Type switch
    {
        ProjectItemType.Form => "🗔",
        ProjectItemType.Module => "📜",
        ProjectItemType.Class => "🧩",
        ProjectItemType.UserControl => "🎛️",
        ProjectItemType.PropertyPage => "📑",
        ProjectItemType.Reference => "🔗",
        ProjectItemType.ActiveXControl => "📦",
        _ => "📄"
    };

    public string DisplayTitle => string.IsNullOrWhiteSpace(Name)
        ? (Path.GetFileName(RelativePath) is { Length: > 0 } fn ? fn : RawLine)
        : $"{Name} ({Path.GetFileName(RelativePath)})";
}
