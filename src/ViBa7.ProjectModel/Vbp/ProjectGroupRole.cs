namespace ViBa7.ProjectModel.Vbp;

/// <summary>
/// Rol de un proyecto dentro de un grupo de proyectos (.vbg) o de forma independiente.
/// </summary>
public enum ProjectGroupRole
{
    /// <summary>
    /// Proyecto individual fuera de un grupo.
    /// </summary>
    Standalone,

    /// <summary>
    /// Proyecto principal: el que se ejecuta como proceso de inicio del grupo.
    /// </summary>
    Primary,

    /// <summary>
    /// Proyecto secundario: se ejecuta como subproceso / componente dependiente del principal.
    /// </summary>
    Secondary
}
