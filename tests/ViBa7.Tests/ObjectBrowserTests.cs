using System;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.IntelliSense;
using ViBa7.ProjectModel.Navigation;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

public class ObjectBrowserTests : IDisposable
{
    private readonly string _tempDir;

    static ObjectBrowserTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public ObjectBrowserTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "ViBa7_ObTests_" + Guid.NewGuid().ToString("N"));
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
    public void VbControlCatalog_ExposesKnownIntrinsicControlTypes()
    {
        var controlTypes = VbControlMemberCatalog.GetKnownControlTypes().ToList();

        Assert.Contains("Form", controlTypes);
        Assert.Contains("CommandButton", controlTypes);
        Assert.Contains("TextBox", controlTypes);
        Assert.Contains("Label", controlTypes);
        Assert.Contains("ListBox", controlTypes);
        Assert.Contains("ComboBox", controlTypes);

        // Verificar que cada control tiene sus propiedades y eventos registrados
        var btnMembers = VbControlMemberCatalog.GetMembersForControl("CommandButton");
        Assert.Contains(btnMembers, m => m.Name == "Caption");
        Assert.Contains(btnMembers, m => m.Name == "Click");
        Assert.Contains(btnMembers, m => m.Name == "Enabled");
    }

    [Fact]
    public void VbRuntimeCatalog_ExposesVbaFunctionsAndConstants()
    {
        var msgBox = VbRuntimeCatalog.FindFunction("MsgBox");
        Assert.NotNull(msgBox);
        Assert.Equal("VbMsgBoxResult", msgBox.ReturnType);
        Assert.True(msgBox.Parameters.Count >= 3);

        var constants = VbRuntimeCatalog.GetAllConstants();
        Assert.Contains(constants, c => c.Name == "vbOKOnly");
        Assert.Contains(constants, c => c.Name == "vbCritical");
        Assert.Contains(constants, c => c.Name == "vbInformation");

        var enumConstants = VbRuntimeCatalog.GetConstantsForType("VbMsgBoxStyle");
        Assert.NotEmpty(enumConstants);
        Assert.Contains(enumConstants, c => c.Name == "vbYesNoCancel");
    }

    [Fact]
    public void ObjectBrowser_ParsesProjectClassesAndMembersForNavigation()
    {
        string modulePath = Path.Combine(_tempDir, "ModCalculos.bas");
        string moduleCode = @"Attribute VB_Name = ""ModCalculos""
Option Explicit

Public Const PI As Double = 3.1415926535

Public Function Sumar(ByVal a As Long, ByVal b As Long) As Long
    Sumar = a + b
End Function

Public Sub LimpiarBuffer()
    ' Código de limpieza
End Sub

Public Property Get VersionApp() As String
    VersionApp = ""1.0.0""
End Property
";
        File.WriteAllText(modulePath, moduleCode, Encoding.GetEncoding(1252));

        var members = CodeStructureParser.ParseMembers(moduleCode, modulePath);

        Assert.Equal(4, members.Count);

        var constMember = members.FirstOrDefault(m => m.Name == "PI");
        Assert.NotNull(constMember);
        Assert.Equal(CodeMemberKind.Constant, constMember.Kind);

        var funcMember = members.FirstOrDefault(m => m.Name == "Sumar");
        Assert.NotNull(funcMember);
        Assert.Equal(CodeMemberKind.Function, funcMember.Kind);
        Assert.Equal("Long", funcMember.TypeName);
        Assert.Equal("(ByVal a As Long, ByVal b As Long)", funcMember.Signature);

        var subMember = members.FirstOrDefault(m => m.Name == "LimpiarBuffer");
        Assert.NotNull(subMember);
        Assert.Equal(CodeMemberKind.Sub, subMember.Kind);

        var propMember = members.FirstOrDefault(m => m.Name == "VersionApp");
        Assert.NotNull(propMember);
        Assert.Equal(CodeMemberKind.PropertyGet, propMember.Kind);
        Assert.Equal("String", propMember.TypeName);
    }
}
