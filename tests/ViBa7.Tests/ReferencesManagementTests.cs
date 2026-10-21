using System;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class ReferencesManagementTests : IDisposable
{
    private readonly string _tempDir;

    static ReferencesManagementTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public ReferencesManagementTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "ViBa7_RefTests_" + Guid.NewGuid().ToString("N"));
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
    public void AddReference_AddsItemAndPreventsDuplicates()
    {
        var project = new VB6Project
        {
            Name = "RefTestProject",
            ProjectFilePath = Path.Combine(_tempDir, "Project1.vbp")
        };

        var ref1 = new ProjectReference
        {
            Guid = "{420B2830-E718-11CF-893D-00A0C9054228}",
            MajorVersion = "1",
            MinorVersion = "0",
            Lcid = "0",
            FilePath = @"C:\Windows\SysWOW64\scrrun.dll",
            Description = "Microsoft Scripting Runtime"
        };

        project.AddReference(ref1);

        Assert.Single(project.References);
        Assert.Equal("Microsoft Scripting Runtime", project.References[0].Description);
        Assert.Contains(project.Items, i => i.Type == ProjectItemType.Reference && i.Name == "Microsoft Scripting Runtime");

        // Intentar agregar la misma referencia por GUID duplicado
        var refDup = new ProjectReference
        {
            Guid = "{420B2830-E718-11CF-893D-00A0C9054228}",
            FilePath = @"C:\Other\scrrun.dll"
        };

        project.AddReference(refDup);

        // No debe haberse agregado duplicada
        Assert.Single(project.References);
    }

    [Fact]
    public void SetReferences_PreservesPriorityOrder()
    {
        var project = new VB6Project
        {
            Name = "RefOrderProject",
            ProjectFilePath = Path.Combine(_tempDir, "Project1.vbp")
        };

        var refA = new ProjectReference
        {
            Guid = "{00020430-0000-0000-C000-000000000046}",
            MajorVersion = "2",
            MinorVersion = "0",
            FilePath = @"C:\Windows\SysWOW64\stdole2.tlb",
            Description = "OLE Automation"
        };

        var refB = new ProjectReference
        {
            Guid = "{420B2830-E718-11CF-893D-00A0C9054228}",
            MajorVersion = "1",
            MinorVersion = "0",
            FilePath = @"C:\Windows\SysWOW64\scrrun.dll",
            Description = "Microsoft Scripting Runtime"
        };

        var refC = new ProjectReference
        {
            Guid = "{2A75196C-D9EB-4129-B803-931327F72D5C}",
            MajorVersion = "2",
            MinorVersion = "8",
            FilePath = @"C:\Program Files\Common Files\System\ado\msado15.dll",
            Description = "Microsoft ActiveX Data Objects 2.8 Library"
        };

        // Asignar en orden B, C, A
        project.SetReferences(new[] { refB, refC, refA });

        Assert.Equal(3, project.References.Count);
        Assert.Equal("Microsoft Scripting Runtime", project.References[0].Description);
        Assert.Equal("Microsoft ActiveX Data Objects 2.8 Library", project.References[1].Description);
        Assert.Equal("OLE Automation", project.References[2].Description);

        // Invertir orden a A, B, C
        project.SetReferences(new[] { refA, refB, refC });

        Assert.Equal(3, project.References.Count);
        Assert.Equal("OLE Automation", project.References[0].Description);
        Assert.Equal("Microsoft Scripting Runtime", project.References[1].Description);
        Assert.Equal("Microsoft ActiveX Data Objects 2.8 Library", project.References[2].Description);
    }

    [Fact]
    public void RemoveReference_RemovesCorrectItem()
    {
        var project = new VB6Project
        {
            Name = "RefRemoveProject",
            ProjectFilePath = Path.Combine(_tempDir, "Project1.vbp")
        };

        var ref1 = new ProjectReference
        {
            Guid = "{00020430-0000-0000-C000-000000000046}",
            FilePath = @"C:\Windows\SysWOW64\stdole2.tlb",
            Description = "OLE Automation"
        };

        var ref2 = new ProjectReference
        {
            Guid = "{420B2830-E718-11CF-893D-00A0C9054228}",
            FilePath = @"C:\Windows\SysWOW64\scrrun.dll",
            Description = "Microsoft Scripting Runtime"
        };

        project.AddReference(ref1);
        project.AddReference(ref2);
        Assert.Equal(2, project.References.Count);

        // Remover ref1 por GUID
        project.RemoveReference("{00020430-0000-0000-C000-000000000046}");

        Assert.Single(project.References);
        Assert.Equal("{420B2830-E718-11CF-893D-00A0C9054228}", project.References[0].Guid);
        Assert.DoesNotContain(project.Items, i => i.Type == ProjectItemType.Reference && i.Guid?.Contains("00020430") == true);
    }

    [Fact]
    public void VbpParser_ParsesMissingReferenceCorrectly()
    {
        string vbpPath = Path.Combine(_tempDir, "MissingRefApp.vbp");
        string fakePath = Path.Combine(_tempDir, "NoExiste_Inexistente.dll");

        string vbpContent = $@"Type=Exe
Reference=*\G{{11223344-5566-7788-99AA-BBCCDDEEFF00}}#1.0#0#{fakePath}#Librería Inexistente
Startup=""Sub Main""
Name=""MissingRefApp""
";
        File.WriteAllText(vbpPath, vbpContent, Encoding.GetEncoding(1252));

        var project = VbpParser.Parse(vbpPath);

        Assert.Single(project.References);
        var r = project.References[0];
        Assert.Equal("{11223344-5566-7788-99AA-BBCCDDEEFF00}", r.Guid);
        Assert.Equal("1", r.MajorVersion);
        Assert.Equal("0", r.MinorVersion);
        Assert.Equal("Librería Inexistente", r.Description);
        Assert.False(File.Exists(r.FilePath), "El archivo no debe existir en disco para simular [FALTA]");
    }
}
