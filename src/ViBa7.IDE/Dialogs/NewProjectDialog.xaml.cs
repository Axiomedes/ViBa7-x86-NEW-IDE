using System;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using ViBa7.ProjectModel.Execution;
using ViBa7.ProjectModel.Vbp;

namespace ViBa7.IDE.Dialogs;

public enum NewProjectType
{
    StandardExe,
    ActiveXDll,
    ActiveXControl
}

public partial class NewProjectDialog : Window
{
    public NewProjectType SelectedProjectType { get; private set; } = NewProjectType.StandardExe;
    public VbLanguageDialect SelectedDialect { get; private set; } = VbLanguageDialect.Vb6Classic;
    public string ProjectName { get; private set; } = "Proyecto1";
    public string ProjectLocation { get; private set; } = string.Empty;
    public string ProjectFilePath => Path.Combine(ProjectLocation, $"{ProjectName}.vbp");

    public NewProjectDialog(string? defaultLocation = null)
    {
        InitializeComponent();

        ProjectLocation = !string.IsNullOrWhiteSpace(defaultLocation) && Directory.Exists(defaultLocation)
            ? defaultLocation
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        TxtProjectLocation.Text = ProjectLocation;
        TxtProjectName.Text = "Proyecto1";
        TxtProjectName.SelectAll();
        TxtProjectName.Focus();
    }

    private void BtnBrowseLocation_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog
        {
            Title = "Seleccionar carpeta para el nuevo proyecto",
            InitialDirectory = Directory.Exists(TxtProjectLocation.Text) ? TxtProjectLocation.Text : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (dlg.ShowDialog() == true)
        {
            TxtProjectLocation.Text = dlg.FolderName;
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        string name = TxtProjectName.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show("Por favor, ingresa un nombre para el proyecto.", "Nombre Requerido", MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtProjectName.Focus();
            return;
        }

        char[] invalidChars = Path.GetInvalidFileNameChars();
        if (name.IndexOfAny(invalidChars) >= 0)
        {
            MessageBox.Show("El nombre del proyecto contiene caracteres no válidos.", "Nombre Inválido", MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtProjectName.Focus();
            return;
        }

        string loc = TxtProjectLocation.Text.Trim();
        if (string.IsNullOrWhiteSpace(loc) || !Directory.Exists(loc))
        {
            try
            {
                Directory.CreateDirectory(loc);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"No se pudo crear o acceder al directorio especificado:\n{ex.Message}", "Directorio Inválido", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtProjectLocation.Focus();
                return;
            }
        }

        ProjectName = name;
        ProjectLocation = loc;

        // Tipo de proyecto
        if (RadioTypeDll.IsChecked == true) SelectedProjectType = NewProjectType.ActiveXDll;
        else if (RadioTypeControl.IsChecked == true) SelectedProjectType = NewProjectType.ActiveXControl;
        else SelectedProjectType = NewProjectType.StandardExe;

        // Dialecto de lenguaje
        if (RadioDialectVb7.IsChecked == true) SelectedDialect = VbLanguageDialect.Vb7Office;
        else if (RadioDialectModern.IsChecked == true) SelectedDialect = VbLanguageDialect.ViBaModern;
        else SelectedDialect = VbLanguageDialect.Vb6Classic;

        DialogResult = true;
        Close();
    }
}
