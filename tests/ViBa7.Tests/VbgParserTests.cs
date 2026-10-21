using System;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class VbgParserTests : IDisposable
{
    private readonly string _tempDir;

    static VbgParserTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public VbgParserTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "ViBa7_VbgTests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try { Directory.Delete(_tempDir, true); } catch { }
        }
    }

    [Fact]
    public void Parse_ValidProjectGroup_LoadsAllProjectsAndStartupProject()
    {
        string vbgPath = Path.Combine(_tempDir, "Group1.vbg");
        string vbp1Path = Path.Combine(_tempDir, "MainApp.vbp");
        string vbp2Path = Path.Combine(_tempDir, "Component.vbp");

        File.WriteAllText(vbp1Path, "Type=Exe\r\nName=\"MainApp\"\r\nModule=Module1; Mod1.bas\r\n", Encoding.GetEncoding(1252));
        File.WriteAllText(vbp2Path, "Type=OleDll\r\nName=\"ComponentLib\"\r\nClass=Class1; Cls1.cls\r\n", Encoding.GetEncoding(1252));

        string vbgContent = @"VBGROUP 5.0
StartupProject=MainApp.vbp
Project=Component.vbp
";
        File.WriteAllText(vbgPath, vbgContent, Encoding.GetEncoding(1252));

        var group = VbgParser.Parse(vbgPath);

        Assert.Equal("Group1", group.Name);
        Assert.Equal(2, group.Projects.Count);
        Assert.NotNull(group.StartupProject);
        Assert.Equal("MainApp", group.StartupProject.Name);

        var compProj = group.Projects.FirstOrDefault(p => p.Name == "ComponentLib");
        Assert.NotNull(compProj);
        Assert.Single(compProj.Classes);
    }

    [Fact]
    public void Parse_ProjectGroup_CorrectlyClassifiesPrimaryAndSecondaryProjects()
    {
        string vbgPath = Path.Combine(_tempDir, "ClientServer.vbg");
        string vbp1Path = Path.Combine(_tempDir, "ClientExe.vbp");
        string vbp2Path = Path.Combine(_tempDir, "CustomControl.vbp");

        File.WriteAllText(vbp1Path, "Type=Exe\r\nName=\"ClientApp\"\r\nStartup=\"Form1\"\r\n", Encoding.GetEncoding(1252));
        File.WriteAllText(vbp2Path, "Type=Control\r\nName=\"MyControlsLib\"\r\n", Encoding.GetEncoding(1252));

        string vbgContent = @"VBGROUP 5.0
StartupProject=ClientExe.vbp
Project=CustomControl.vbp
";
        File.WriteAllText(vbgPath, vbgContent, Encoding.GetEncoding(1252));

        var group = VbgParser.Parse(vbgPath);

        // 1. Verificación de clasificación en el grupo
        Assert.NotNull(group.PrimaryProject);
        Assert.Equal("ClientApp", group.PrimaryProject.Name);
        Assert.Same(group.PrimaryProject, group.StartupProject);

        Assert.Single(group.SecondaryProjects);
        var secProj = group.SecondaryProject;
        Assert.NotNull(secProj);
        Assert.Equal("MyControlsLib", secProj.Name);

        // 2. Verificación de propiedades y roles en los proyectos individuales
        var primary = group.PrimaryProject;
        Assert.True(primary.IsPrimaryProject);
        Assert.False(primary.IsSecondaryProject);
        Assert.Equal(ProjectGroupRole.Primary, primary.GroupRole);
        Assert.Contains("Principal", primary.RoleDescription);

        Assert.False(secProj.IsPrimaryProject);
        Assert.True(secProj.IsSecondaryProject);
        Assert.Equal(ProjectGroupRole.Secondary, secProj.GroupRole);
        Assert.Contains("Secundario", secProj.RoleDescription);
        Assert.Contains("Subproceso", secProj.RoleDescription);

        // 3. Métodos booleanos de consulta
        Assert.True(group.IsPrimaryProject(primary));
        Assert.False(group.IsPrimaryProject(secProj));
        Assert.False(group.IsSecondaryProject(primary));
        Assert.True(group.IsSecondaryProject(secProj));
    }

    [Fact]
    public void Parse_ProjectGroupWithoutExplicitStartup_DefaultsToExeAsPrimary()
    {
        string vbgPath = Path.Combine(_tempDir, "NoStartup.vbg");
        string vbp1Path = Path.Combine(_tempDir, "Lib1.vbp");
        string vbp2Path = Path.Combine(_tempDir, "HostApp.vbp");

        File.WriteAllText(vbp1Path, "Type=OleDll\r\nName=\"MyDll\"\r\n", Encoding.GetEncoding(1252));
        File.WriteAllText(vbp2Path, "Type=Exe\r\nName=\"HostApp\"\r\n", Encoding.GetEncoding(1252));

        // En este VBG, no hay StartupProject=; solo Project=
        string vbgContent = @"VBGROUP 5.0
Project=Lib1.vbp
Project=HostApp.vbp
";
        File.WriteAllText(vbgPath, vbgContent, Encoding.GetEncoding(1252));

        var group = VbgParser.Parse(vbgPath);

        // Debe deducir que HostApp (Type=Exe) es el Proyecto Principal
        Assert.NotNull(group.PrimaryProject);
        Assert.Equal("HostApp", group.PrimaryProject.Name);
        Assert.True(group.PrimaryProject.IsPrimaryProject);

        var lib = group.Projects.First(p => p.Name == "MyDll");
        Assert.True(lib.IsSecondaryProject);
    }

    [Fact]
    public void SetPrimaryProject_SwitchesRolesAndSerializesAccurately()
    {
        string vbgPath = Path.Combine(_tempDir, "SwitchGroup.vbg");
        string vbp1Path = Path.Combine(_tempDir, "App1.vbp");
        string vbp2Path = Path.Combine(_tempDir, "App2.vbp");

        File.WriteAllText(vbp1Path, "Type=Exe\r\nName=\"App1\"\r\n", Encoding.GetEncoding(1252));
        File.WriteAllText(vbp2Path, "Type=Exe\r\nName=\"App2\"\r\n", Encoding.GetEncoding(1252));

        string vbgContent = @"VBGROUP 5.0
StartupProject=App1.vbp
Project=App2.vbp
";
        File.WriteAllText(vbgPath, vbgContent, Encoding.GetEncoding(1252));

        var group = VbgParser.Parse(vbgPath);
        var app1 = group.Projects.First(p => p.Name == "App1");
        var app2 = group.Projects.First(p => p.Name == "App2");

        Assert.True(app1.IsPrimaryProject);
        Assert.True(app2.IsSecondaryProject);

        // Cambiar App2 a Proyecto Principal
        group.SetPrimaryProject(app2);

        Assert.False(app1.IsPrimaryProject);
        Assert.True(app1.IsSecondaryProject);
        Assert.True(app2.IsPrimaryProject);
        Assert.False(app2.IsSecondaryProject);

        // Guardar y volver a cargar
        group.Save();
        var reloaded = VbgParser.Parse(vbgPath);

        Assert.Equal("App2", reloaded.PrimaryProject?.Name);
        Assert.True(reloaded.PrimaryProject?.IsPrimaryProject);
        Assert.True(reloaded.SecondaryProject?.IsSecondaryProject);
    }

    [Fact]
    public void GetUserControlDesignModel_ResolvesUserControlFromSecondaryProject()
    {
        string vbgPath = Path.Combine(_tempDir, "UcGroup.vbg");
        string vbpPrimary = Path.Combine(_tempDir, "MainExe.vbp");
        string vbpSecondary = Path.Combine(_tempDir, "ControlsLib.vbp");
        string ctlPath = Path.Combine(_tempDir, "SuperButton.ctl");

        string ctlContent = @"VERSION 5.00
Begin VB.UserControl SuperButton 
   ClientHeight    =   3600
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = ""SuperButton""
Public Property Get CustomColor() As Long
End Property
";
        File.WriteAllText(ctlPath, ctlContent, Encoding.GetEncoding(1252));

        File.WriteAllText(vbpPrimary, "Type=Exe\r\nName=\"MainExe\"\r\n", Encoding.GetEncoding(1252));
        File.WriteAllText(vbpSecondary, "Type=Control\r\nName=\"ControlsLib\"\r\nUserControl=SuperButton.ctl\r\n", Encoding.GetEncoding(1252));

        string vbgContent = @"VBGROUP 5.0
StartupProject=MainExe.vbp
Project=ControlsLib.vbp
";
        File.WriteAllText(vbgPath, vbgContent, Encoding.GetEncoding(1252));

        var group = VbgParser.Parse(vbgPath);

        // MainExe es el Principal, ControlsLib es el Secundario (Subproceso)
        Assert.Equal("MainExe", group.PrimaryProject?.Name);
        Assert.Equal("ControlsLib", group.SecondaryProject?.Name);

        // El grupo debe resolver el UserControl que reside en el proyecto secundario
        var ucModel = group.GetUserControlDesignModel("SuperButton");
        Assert.NotNull(ucModel);
        Assert.Equal("SuperButton", ucModel.Name);
        Assert.NotNull(ucModel.ExposedProperties);
        Assert.Contains(ucModel.ExposedProperties, p => p.Name == "CustomColor");

        var ucItem = group.FindUserControl("SuperButton");
        Assert.NotNull(ucItem);
        Assert.Equal("SuperButton", ucItem.Name);
        Assert.Equal("ControlsLib", ucItem.ContainingProject?.Name);
    }
}
