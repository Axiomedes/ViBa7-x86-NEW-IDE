using System;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class UserControlToolboxTests : IDisposable
{
    private readonly string _tempDir;

    static UserControlToolboxTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public UserControlToolboxTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "ViBa7_UcToolTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        ControlToolboxRegistry.UnregisterUserControlTools();
        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, true); } catch { }
        }
    }

    [Fact]
    public void RegisterUserControlTools_FromProject_RegistersToolboxItemCorrectly()
    {
        string ctlPath = Path.Combine(_tempDir, "MiBotonCustom.ctl");
        string ctlContent = @"VERSION 5.00
Begin VB.UserControl MiBotonCustom 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleWidth      =   2400
End
Attribute VB_Name = ""MiBotonCustom""
Option Explicit
";
        File.WriteAllText(ctlPath, ctlContent, Encoding.GetEncoding(1252));

        var project = new VB6Project
        {
            Name = "ProyectoPrueba",
            ProjectFilePath = Path.Combine(_tempDir, "ProyectoPrueba.vbp")
        };

        project.Items.Add(new ProjectItem
        {
            Name = "MiBotonCustom",
            RelativePath = "MiBotonCustom.ctl",
            AbsolutePath = ctlPath,
            Type = ProjectItemType.UserControl,
            FileExists = true,
            ContainingProject = project
        });

        ControlToolboxRegistry.RegisterUserControlTools(project);

        var tool = ControlToolboxRegistry.Find("ProyectoPrueba.MiBotonCustom");
        Assert.NotNull(tool);
        Assert.Equal("MiBotonCustom", tool.DisplayName);
        Assert.Equal("🎛️", tool.Glyph);
        Assert.Equal(2400, tool.DefaultWidthTwips);
        Assert.Equal(1200, tool.DefaultHeightTwips);

        // Crear una instancia en un formulario
        var formModel = new FormDesignModel { Name = "Form1" };
        var ctrl = ControlToolboxRegistry.CreateControl(tool.TypeName, formModel, 100, 200);

        Assert.NotNull(ctrl);
        Assert.Equal("ProyectoPrueba.MiBotonCustom", ctrl.TypeName);
        Assert.Equal(100, ctrl.Left);
        Assert.Equal(200, ctrl.Top);
        Assert.Equal(2400, ctrl.Width);
        Assert.Equal(1200, ctrl.Height);
    }

    [Fact]
    public void RegisterUserControlTools_FromProjectGroup_RegistersAllUserControls()
    {
        string ctlPath1 = Path.Combine(_tempDir, "UC1.ctl");
        File.WriteAllText(ctlPath1, "VERSION 5.00\r\nBegin VB.UserControl UC1\r\nClientHeight=1000\r\nClientWidth=1000\r\nEnd\r\n", Encoding.GetEncoding(1252));

        string ctlPath2 = Path.Combine(_tempDir, "UC2.ctl");
        File.WriteAllText(ctlPath2, "VERSION 5.00\r\nBegin VB.UserControl UC2\r\nClientHeight=2000\r\nClientWidth=2000\r\nEnd\r\n", Encoding.GetEncoding(1252));

        var group = new VB6ProjectGroup
        {
            Name = "GrupoControles",
            GroupFilePath = Path.Combine(_tempDir, "Grupo.vbg")
        };

        var proj1 = new VB6Project { Name = "AppPrincipal", ProjectFilePath = Path.Combine(_tempDir, "App.vbp") };
        proj1.Items.Add(new ProjectItem { Name = "UC1", RelativePath = "UC1.ctl", AbsolutePath = ctlPath1, Type = ProjectItemType.UserControl, FileExists = true });

        var proj2 = new VB6Project { Name = "BibliotecaControles", ProjectFilePath = Path.Combine(_tempDir, "Lib.vbp") };
        proj2.Items.Add(new ProjectItem { Name = "UC2", RelativePath = "UC2.ctl", AbsolutePath = ctlPath2, Type = ProjectItemType.UserControl, FileExists = true });

        group.Projects.Add(proj1);
        group.Projects.Add(proj2);

        ControlToolboxRegistry.RegisterUserControlTools(group);

        Assert.NotNull(ControlToolboxRegistry.Find("AppPrincipal.UC1"));
        Assert.NotNull(ControlToolboxRegistry.Find("BibliotecaControles.UC2"));

        // Limpiar controles de usuario no debe borrar herramientas intrínsecas
        ControlToolboxRegistry.UnregisterUserControlTools();
        Assert.Null(ControlToolboxRegistry.Find("AppPrincipal.UC1"));
        Assert.Null(ControlToolboxRegistry.Find("BibliotecaControles.UC2"));
        Assert.NotNull(ControlToolboxRegistry.Find("VB.CommandButton"));
    }
}
