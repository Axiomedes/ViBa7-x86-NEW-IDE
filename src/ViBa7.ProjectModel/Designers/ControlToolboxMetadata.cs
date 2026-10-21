using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.ProjectModel.Designers;

/// <summary>
/// Metadatos y definición de un control disponible en la Caja de Herramientas (Toolbox) de VB6.
/// </summary>
public record ToolboxItemInfo(
    string TypeName,
    string DisplayName,
    string DefaultPrefix,
    string Glyph,
    string ToolTip,
    int DefaultWidthTwips,
    int DefaultHeightTwips,
    bool IsContainer = false,
    bool HasCaption = false,
    bool HasText = false,
    byte[]? IconData = null
);

/// <summary>
/// Registro central de herramientas intrínsecas de VB6 para la Caja de Herramientas y creación de controles.
/// </summary>
public static class ControlToolboxRegistry
{
    public const string PointerTool = "Pointer";

    private static readonly List<ToolboxItemInfo> _tools = new()
    {
        new ToolboxItemInfo(
            TypeName: PointerTool,
            DisplayName: "Puntero",
            DefaultPrefix: "",
            Glyph: "↖",
            ToolTip: "Puntero de selección e interactividad",
            DefaultWidthTwips: 0,
            DefaultHeightTwips: 0
        ),
        new ToolboxItemInfo(
            TypeName: "VB.PictureBox",
            DisplayName: "PictureBox",
            DefaultPrefix: "Picture",
            Glyph: "🖼",
            ToolTip: "Caja de imagen y lienzo contenedor (PictureBox)",
            DefaultWidthTwips: 1575,
            DefaultHeightTwips: 1215,
            IsContainer: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Label",
            DisplayName: "Label",
            DefaultPrefix: "Label",
            Glyph: "A",
            ToolTip: "Etiqueta de texto estático (Label)",
            DefaultWidthTwips: 1095,
            DefaultHeightTwips: 255,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.TextBox",
            DisplayName: "TextBox",
            DefaultPrefix: "Text",
            Glyph: "ab|",
            ToolTip: "Caja de edición de texto (TextBox)",
            DefaultWidthTwips: 1455,
            DefaultHeightTwips: 375,
            HasText: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Frame",
            DisplayName: "Frame",
            DefaultPrefix: "Frame",
            Glyph: "[...]",
            ToolTip: "Marco agrupador de controles (Frame)",
            DefaultWidthTwips: 2535,
            DefaultHeightTwips: 1695,
            IsContainer: true,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.CommandButton",
            DisplayName: "CommandButton",
            DefaultPrefix: "Command",
            Glyph: "🔲",
            ToolTip: "Botón de comando (CommandButton)",
            DefaultWidthTwips: 1335,
            DefaultHeightTwips: 495,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.CheckBox",
            DisplayName: "CheckBox",
            DefaultPrefix: "Check",
            Glyph: "☑",
            ToolTip: "Casilla de verificación (CheckBox)",
            DefaultWidthTwips: 1335,
            DefaultHeightTwips: 315,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.OptionButton",
            DisplayName: "OptionButton",
            DefaultPrefix: "Option",
            Glyph: "🔘",
            ToolTip: "Botón de opción radial (OptionButton)",
            DefaultWidthTwips: 1335,
            DefaultHeightTwips: 315,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.ComboBox",
            DisplayName: "ComboBox",
            DefaultPrefix: "Combo",
            Glyph: "▼",
            ToolTip: "Lista desplegable y cuadro combinado (ComboBox)",
            DefaultWidthTwips: 1575,
            DefaultHeightTwips: 315,
            HasText: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.ListBox",
            DisplayName: "ListBox",
            DefaultPrefix: "List",
            Glyph: "≡",
            ToolTip: "Cuadro de lista con opciones (ListBox)",
            DefaultWidthTwips: 1455,
            DefaultHeightTwips: 1410
        ),
        new ToolboxItemInfo(
            TypeName: "VB.HScrollBar",
            DisplayName: "HScrollBar",
            DefaultPrefix: "HScroll",
            Glyph: "◄▬►",
            ToolTip: "Barra de desplazamiento horizontal (HScrollBar)",
            DefaultWidthTwips: 1455,
            DefaultHeightTwips: 255
        ),
        new ToolboxItemInfo(
            TypeName: "VB.VScrollBar",
            DisplayName: "VScrollBar",
            DefaultPrefix: "VScroll",
            Glyph: "▲❚▼",
            ToolTip: "Barra de desplazamiento vertical (VScrollBar)",
            DefaultWidthTwips: 255,
            DefaultHeightTwips: 1455
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Timer",
            DisplayName: "Timer",
            DefaultPrefix: "Timer",
            Glyph: "⏱",
            ToolTip: "Temporizador periódico (Timer)",
            DefaultWidthTwips: 480,
            DefaultHeightTwips: 480
        ),
        new ToolboxItemInfo(
            TypeName: "VB.DriveListBox",
            DisplayName: "DriveListBox",
            DefaultPrefix: "Drive",
            Glyph: "🖴",
            ToolTip: "Lista de unidades de disco (DriveListBox)",
            DefaultWidthTwips: 1695,
            DefaultHeightTwips: 315
        ),
        new ToolboxItemInfo(
            TypeName: "VB.DirListBox",
            DisplayName: "DirListBox",
            DefaultPrefix: "Dir",
            Glyph: "📁",
            ToolTip: "Lista de directorios del sistema (DirListBox)",
            DefaultWidthTwips: 1695,
            DefaultHeightTwips: 1410
        ),
        new ToolboxItemInfo(
            TypeName: "VB.FileListBox",
            DisplayName: "FileListBox",
            DefaultPrefix: "File",
            Glyph: "📄",
            ToolTip: "Lista de archivos de carpeta (FileListBox)",
            DefaultWidthTwips: 1695,
            DefaultHeightTwips: 1410
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Shape",
            DisplayName: "Shape",
            DefaultPrefix: "Shape",
            Glyph: "◯",
            ToolTip: "Forma geométrica decorativa (Shape)",
            DefaultWidthTwips: 1335,
            DefaultHeightTwips: 855
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Line",
            DisplayName: "Line",
            DefaultPrefix: "Line",
            Glyph: "╱",
            ToolTip: "Línea gráfica divisoria (Line)",
            DefaultWidthTwips: 1455,
            DefaultHeightTwips: 0
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Image",
            DisplayName: "Image",
            DefaultPrefix: "Image",
            Glyph: "🎨",
            ToolTip: "Imagen gráfica ligera (Image)",
            DefaultWidthTwips: 1215,
            DefaultHeightTwips: 975
        ),
        new ToolboxItemInfo(
            TypeName: "VB.Data",
            DisplayName: "Data",
            DefaultPrefix: "Data",
            Glyph: "⛁",
            ToolTip: "Control de enlace de datos DAO/ODBC (Data)",
            DefaultWidthTwips: 1935,
            DefaultHeightTwips: 375,
            HasCaption: true
        ),
        new ToolboxItemInfo(
            TypeName: "VB.OLE",
            DisplayName: "OLE",
            DefaultPrefix: "OLE",
            Glyph: "◫",
            ToolTip: "Contenedor de objetos OLE incrustados (OLE)",
            DefaultWidthTwips: 1575,
            DefaultHeightTwips: 1215
        )
    };

    public static IReadOnlyList<ToolboxItemInfo> GetAllTools() => _tools;

    public static ToolboxItemInfo? Find(string typeName)
    {
        return _tools.FirstOrDefault(t =>
            string.Equals(t.TypeName, typeName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(t.DisplayName, typeName, StringComparison.OrdinalIgnoreCase)
        );
    }

    /// <summary>
    /// Registra una herramienta personalizada en la caja de herramientas.
    /// </summary>
    public static void RegisterCustomTool(ToolboxItemInfo tool)
    {
        if (tool == null) return;
        var existing = _tools.FirstOrDefault(t => string.Equals(t.TypeName, tool.TypeName, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            _tools.Remove(existing);
        }
        _tools.Add(tool);
    }

    /// <summary>
    /// Elimina herramientas que cumplan con la condición dada.
    /// </summary>
    public static void UnregisterCustomTools(Func<ToolboxItemInfo, bool> predicate)
    {
        _tools.RemoveAll(t => predicate(t));
    }

    /// <summary>
    /// Registra en la caja de herramientas todos los controles visuales expuestos por una biblioteca OCX.
    /// </summary>
    public static void RegisterOcxTools(ViBa7.ProjectModel.ActiveX.OcxLibraryInfo lib)
    {
        if (lib == null) return;

        foreach (var ctrl in lib.Controls.Where(c => c.IsControl))
        {
            string typeName = !string.IsNullOrEmpty(lib.LibraryName) ? $"{lib.LibraryName}.{ctrl.Name}" : ctrl.Name;
            string prefix = ctrl.Name;
            if (prefix.StartsWith("ax", StringComparison.OrdinalIgnoreCase))
            {
                prefix = prefix[2..];
            }

            byte[]? iconData = null;
            try
            {
                iconData = ViBa7.ProjectModel.ActiveX.OcxToolboxBitmapExtractor.GetToolboxBitmap(ctrl.Clsid, lib.FilePath);
            }
            catch
            {
                // Sin icono: se usará el glifo genérico
            }

            var tool = new ToolboxItemInfo(
                TypeName: typeName,
                DisplayName: ctrl.Name,
                DefaultPrefix: prefix,
                Glyph: "🧩",
                ToolTip: $"{ctrl.Name} ({(!string.IsNullOrEmpty(ctrl.Description) ? ctrl.Description : lib.LibraryName)})",
                DefaultWidthTwips: 1500,
                DefaultHeightTwips: 600,
                IsContainer: false,
                HasCaption: ctrl.Members.Any(m => m.Name.Equals("Caption", StringComparison.OrdinalIgnoreCase)),
                HasText: ctrl.Members.Any(m => m.Name.Equals("Text", StringComparison.OrdinalIgnoreCase)),
                IconData: iconData
            );

            RegisterCustomTool(tool);
        }
    }

    /// <summary>
    /// Elimina de la caja de herramientas todos los controles pertenecientes a una biblioteca OCX.
    /// </summary>
    public static void UnregisterOcxTools(string libraryName)
    {
        if (string.IsNullOrWhiteSpace(libraryName)) return;
        UnregisterCustomTools(t => t.TypeName.StartsWith($"{libraryName}.", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Registra en la caja de herramientas todos los UserControls (.ctl) pertenecientes al proyecto actual.
    /// </summary>
    public static void RegisterUserControlTools(VB6Project project, bool clearPrevious = true)
    {
        if (project == null) return;

        // Limpiar controles de usuario previos si se solicita
        if (clearPrevious)
        {
            UnregisterUserControlTools();
        }

        string projName = !string.IsNullOrWhiteSpace(project.Name) ? project.Name : "Proyecto1";

        foreach (var ucItem in project.UserControls)
        {
            string ucName = !string.IsNullOrWhiteSpace(ucItem.Name)
                ? ucItem.Name
                : Path.GetFileNameWithoutExtension(ucItem.RelativePath);

            string fullTypeName = $"{projName}.{ucName}";

            // Dimensiones por defecto leídas del archivo .ctl
            int defaultWidth = 2400;
            int defaultHeight = 1800;
            var ucModel = project.GetUserControlDesignModel(ucName);
            if (ucModel != null)
            {
                if (ucModel.ClientWidth > 0) defaultWidth = ucModel.ClientWidth;
                if (ucModel.ClientHeight > 0) defaultHeight = ucModel.ClientHeight;
            }

            var toolInfo = new ToolboxItemInfo(
                TypeName: fullTypeName,
                DisplayName: ucName,
                DefaultPrefix: ucName,
                Glyph: "🎛️",
                ToolTip: $"Control de Usuario: {ucName} ({projName})",
                DefaultWidthTwips: defaultWidth,
                DefaultHeightTwips: defaultHeight,
                IsContainer: ucModel?.ControlContainer ?? false,
                IconData: ucModel?.ToolboxBitmapData
            );

            RegisterCustomTool(toolInfo);
        }
    }

    /// <summary>
    /// Registra en la caja de herramientas todos los UserControls (.ctl) pertenecientes a un grupo de proyectos (.vbg).
    /// </summary>
    public static void RegisterUserControlTools(VB6ProjectGroup projectGroup)
    {
        if (projectGroup == null) return;

        UnregisterUserControlTools();

        foreach (var p in projectGroup.Projects)
        {
            RegisterUserControlTools(p, clearPrevious: false);
        }
    }

    /// <summary>
    /// Elimina de la caja de herramientas todos los UserControls personalizados registrados del proyecto.
    /// </summary>
    public static void UnregisterUserControlTools()
    {
        UnregisterCustomTools(t => t.Glyph == "🎛️" || t.ToolTip.StartsWith("Control de Usuario:", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Instancia un nuevo control con nombre autogenerado único, dimensiones estándar y textos por defecto.
    /// </summary>
    public static FormDesignerControl CreateControl(
        string typeName,
        FormDesignModel model,
        int leftTwips,
        int topTwips,
        int? widthTwips = null,
        int? heightTwips = null,
        FormDesignerControl? parentControl = null)
    {
        var info = Find(typeName);
        string prefix = info?.DefaultPrefix ?? "Control";
        string uniqueName = model.GenerateUniqueName(prefix);

        int w = widthTwips ?? info?.DefaultWidthTwips ?? 1200;
        int h = heightTwips ?? info?.DefaultHeightTwips ?? 400;

        // Asegurar que controles de 1 dimensión como Line mantengan dimensiones válidas
        if (w <= 0 && !string.Equals(typeName, "VB.Line", StringComparison.OrdinalIgnoreCase)) w = 300;
        if (h <= 0 && !string.Equals(typeName, "VB.Line", StringComparison.OrdinalIgnoreCase)) h = 300;

        string finalTypeName = typeName;
        bool isCustomUserControl = info != null && (info.Glyph == "🎛️" || info.ToolTip.StartsWith("Control de Usuario:", StringComparison.OrdinalIgnoreCase));
        if (!isCustomUserControl && !finalTypeName.StartsWith("VB.", StringComparison.OrdinalIgnoreCase) && !finalTypeName.Contains('.'))
        {
            finalTypeName = $"VB.{finalTypeName}";
        }

        var control = new FormDesignerControl
        {
            TypeName = finalTypeName,
            Name = uniqueName,
            Left = Math.Max(0, leftTwips),
            Top = Math.Max(0, topTwips),
            Width = w,
            Height = h,
            ParentControl = parentControl
        };

        if (info?.HasCaption == true)
        {
            control.Caption = uniqueName;
        }

        if (info?.HasText == true)
        {
            control.Text = uniqueName;
        }

        return control;
    }
}
