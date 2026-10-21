using System;
using System.Collections.Generic;
using System.Linq;

namespace ViBa7.ProjectModel.Designers;

/// <summary>
/// Modelo en memoria de la definición visual completa de un formulario VB6 (bloque Begin VB.Form ... End).
/// </summary>
public class FormDesignModel
{
    public const double TwipsPerPixel = 15.0;

    public string VersionHeader { get; set; } = "VERSION 5.00";
    public string RootType { get; set; } = "VB.Form";
    public string Name { get; set; } = "Form1";
    public string Caption { get; set; } = "Form1";

    public bool IsUserControl =>
        RootType.Equals("VB.UserControl", StringComparison.OrdinalIgnoreCase) ||
        RootType.Equals("UserControl", StringComparison.OrdinalIgnoreCase) ||
        RootType.EndsWith(".UserControl", StringComparison.OrdinalIgnoreCase);

    public bool IsPropertyPage =>
        RootType.Equals("VB.PropertyPage", StringComparison.OrdinalIgnoreCase) ||
        RootType.Equals("PropertyPage", StringComparison.OrdinalIgnoreCase);

    public bool HasTitleBar => !IsUserControl && !IsPropertyPage;

    // Geometría en Twips
    public int ClientLeft { get; set; } = 60;
    public int ClientTop { get; set; } = 345;
    public int ClientWidth { get; set; } = 4680;
    public int ClientHeight { get; set; } = 3195;

    public int? ScaleWidth { get; set; }
    public int? ScaleHeight { get; set; }
    public int? ScaleMode { get; set; } = 1; // 1 = Twip
    public int? StartUpPosition { get; set; } = 3; // 3 = Windows Default
    public int? BorderStyle { get; set; } = 2; // 2 = Sizable

    public string? BackColorHex { get; set; } = "&H8000000F&"; // ButtonFace
    public string? ForeColorHex { get; set; } = "&H80000012&"; // ButtonText
    public FormFontDescription? Font { get; set; }

    // Imagen de fondo del formulario
    public string? Picture { get; set; }
    public byte[]? PictureData { get; set; }

    // Icono de caja de herramientas (para UserControls)
    public string? ToolboxBitmap { get; set; }
    public byte[]? ToolboxBitmapData { get; set; }

    /// <summary>
    /// Indica si este UserControl puede alojar controles en tiempo de diseño (ControlContainer = -1).
    /// </summary>
    public bool ControlContainer { get; set; }

    /// <summary>
    /// Propiedades públicas expuestas por el código del UserControl (.ctl).
    /// </summary>
    public List<UserControlExposedProperty> ExposedProperties { get; set; } = new();

    // Declaraciones de bibliotecas ActiveX/OCX en la cabecera del archivo (ej. Object = "{GUID}#2.0#0"; "MSCOMCTL.OCX")
    public List<string> ObjectDeclarations { get; set; } = new();

    // Copia en memoria del archivo binario .frx / .ctx original para preservar offsets existentes sin pérdidas
    public byte[]? OriginalFrxBytes { get; set; }

    // Controles hijos de primer nivel
    public List<FormDesignerControl> Children { get; set; } = new();

    // Propiedades originales de la cabecera del formulario no mapeadas
    public List<FormPropertyItem> CustomProperties { get; set; } = new();

    // Atributos de módulo posteriores a la definición del formulario (Attribute VB_Name = ..., etc.)
    public List<string> Attributes { get; set; } = new();

    #region Conversión Twips ↔ Píxeles
    public double ClientWidthPixels
    {
        get => Math.Max(50, ClientWidth / TwipsPerPixel);
        set => ClientWidth = Math.Max(750, (int)Math.Round(value * TwipsPerPixel));
    }

    public double ClientHeightPixels
    {
        get => Math.Max(50, ClientHeight / TwipsPerPixel);
        set => ClientHeight = Math.Max(750, (int)Math.Round(value * TwipsPerPixel));
    }

    public void SetClientSizePixels(double width, double height)
    {
        ClientWidthPixels = width;
        ClientHeightPixels = height;
    }
    #endregion

    /// <summary>
    /// Retorna todos los controles del formulario de forma aplanada (incluyendo los anidados dentro de contenedores).
    /// </summary>
    public List<FormDesignerControl> GetAllControls()
    {
        var list = new List<FormDesignerControl>();
        CollectControlsRecursively(Children, list);
        return list;
    }

    private static void CollectControlsRecursively(IEnumerable<FormDesignerControl> controls, List<FormDesignerControl> accumulator)
    {
        foreach (var ctrl in controls)
        {
            accumulator.Add(ctrl);
            if (ctrl.Children.Count > 0)
            {
                CollectControlsRecursively(ctrl.Children, accumulator);
            }
        }
    }

    /// <summary>
    /// Busca un control por su nombre en cualquier nivel de la jerarquía.
    /// </summary>
    public FormDesignerControl? FindControl(string name)
    {
        return GetAllControls().FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Busca un control por su nombre y por su índice (para matrices de controles).
    /// </summary>
    public FormDesignerControl? FindControl(string name, int? index)
    {
        return GetAllControls().FirstOrDefault(c =>
            string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase) &&
            c.Index == index);
    }

    /// <summary>
    /// Retorna todos los controles que pertenecen a una matriz de controles con el nombre indicado, ordenados por su índice.
    /// </summary>
    public List<FormDesignerControl> GetControlArray(string name)
    {
        return GetAllControls()
            .Where(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
            .OrderBy(c => c.Index ?? 0)
            .ToList();
    }

    /// <summary>
    /// Elimina un control de la jerarquía (sea hijo directo del formulario o de un contenedor).
    /// </summary>
    public bool RemoveControl(FormDesignerControl control)
    {
        if (control.ParentControl != null)
        {
            return control.ParentControl.Children.Remove(control);
        }

        return Children.Remove(control);
    }

    /// <summary>
    /// Genera un nombre de control no repetido con prefijo estándar (ej. Command1, Command2).
    /// </summary>
    public string GenerateUniqueName(string prefix)
    {
        var allNames = new HashSet<string>(GetAllControls().Select(c => c.Name), StringComparer.OrdinalIgnoreCase);
        int index = 1;
        while (allNames.Contains($"{prefix}{index}"))
        {
            index++;
        }
        return $"{prefix}{index}";
    }

    #region Soporte para Menús (VB.Menu)
    /// <summary>
    /// Retorna los menús de nivel superior del formulario.
    /// </summary>
    public List<FormDesignerControl> TopLevelMenus => Children.Where(c => c.IsMenu).ToList();

    /// <summary>
    /// Retorna los controles del formulario que no son menús.
    /// </summary>
    public List<FormDesignerControl> NonMenuControls => Children.Where(c => !c.IsMenu).ToList();

    /// <summary>
    /// Retorna la lista aplanada de todos los elementos de menú con su respectivo nivel de sangría (0 para menú principal, 1 para submenú, etc.).
    /// </summary>
    public List<(FormDesignerControl Control, int Level)> GetFlattenedMenuItems()
    {
        var result = new List<(FormDesignerControl Control, int Level)>();
        foreach (var topMenu in TopLevelMenus)
        {
            CollectMenusFlattened(topMenu, 0, result);
        }
        return result;
    }

    private static void CollectMenusFlattened(FormDesignerControl current, int level, List<(FormDesignerControl Control, int Level)> list)
    {
        list.Add((current, level));
        foreach (var child in current.Children.Where(c => c.IsMenu))
        {
            CollectMenusFlattened(child, level + 1, list);
        }
    }

    /// <summary>
    /// Reemplaza la estructura completa de menús del formulario a partir de una lista aplanada con niveles de indentación.
    /// </summary>
    public void ReplaceMenus(IEnumerable<(FormDesignerControl Control, int Level)> flattenedList)
    {
        // 1. Quitar todos los menús existentes de Children
        Children.RemoveAll(c => c.IsMenu);

        var stack = new List<(FormDesignerControl Control, int Level)>();

        foreach (var (ctrl, level) in flattenedList)
        {
            ctrl.Children.Clear();
            ctrl.ParentControl = null;

            if (level == 0)
            {
                Children.Add(ctrl);
                stack.Clear();
                stack.Add((ctrl, 0));
            }
            else
            {
                // Buscar el padre directo en el stack (el último elemento con nivel = level - 1)
                while (stack.Count > 0 && stack[^1].Level >= level)
                {
                    stack.RemoveAt(stack.Count - 1);
                }

                if (stack.Count > 0)
                {
                    var parent = stack[^1].Control;
                    ctrl.ParentControl = parent;
                    parent.Children.Add(ctrl);
                }
                else
                {
                    // Fallback: si no hay padre disponible, agregar a nivel raíz
                    Children.Add(ctrl);
                }

                stack.Add((ctrl, level));
            }
        }
    }
    #endregion
}
