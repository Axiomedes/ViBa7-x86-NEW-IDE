namespace ViBa7.IDE;

/// <summary>
/// Control de versiones centralizado de ViBa7-x86 IDE.
/// Estructura de 3 grupos: [Mayores].[Mejoras/Hitos].[Revisiones/Compilaciones] (00.000.0000).
/// </summary>
public static class AppVersion
{
    /// <summary>
    /// Versión final respecto de cambios grandes e importantes en el proyecto (2 dígitos).
    /// </summary>
    public const string Major = "00";

    /// <summary>
    /// Mejoras implementadas sin renovar demasiado el IDE o su funcionalidad (3 dígitos).
    /// </summary>
    public const string Minor = "001";

    /// <summary>
    /// Correcciones y compilaciones realizadas acumuladas desde el inicio (4 dígitos).
    /// </summary>
    public const string Revision = "0636";

    /// <summary>
    /// Número de versión completo formateado (ej. 00.001.0561).
    /// </summary>
    public static string VersionString => $"{Major}.{Minor}.{Revision}";

    /// <summary>
    /// Título oficial del proyecto.
    /// </summary>
    public static string AppTitle => "ViBa7-x86 IDE";

    /// <summary>
    /// Autoría oficial del proyecto.
    /// </summary>
    public static string Author => "David Rojas [Axio.UK]";

    /// <summary>
    /// Licencia oficial del proyecto.
    /// </summary>
    public const string License = "GNU General Public License v3.0 (GPL-3.0)";

    /// <summary>
    /// Declaración oficial y descargo de marcas registradas.
    /// </summary>
    public const string Disclaimer = "Este proyecto es un desarrollo independiente y no está afiliado, patrocinado ni respaldado por Microsoft. Las referencias a Visual Basic 6.0 y a otros productos de terceros se utilizan exclusivamente para identificar las tecnologías y los formatos con los que se pretende trabajar. Las marcas mencionadas pertenecen a sus respectivos titulares.";
}
