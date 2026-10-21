using System;
using System.Collections.Generic;
using System.Linq;
using ViBa7.ProjectModel.Designers;
using Xunit;

namespace ViBa7.Tests;

/// <summary>
/// Pruebas unitarias para la validación y tipado de IconCharCode, IconFont y eliminación de
/// propiedades Font fantasma/duplicadas en UserControls y controles ActiveX OCX.
/// </summary>
public class UserControlIconAndFontPropertyTests
{
    [Theory]
    [InlineData("EA01", "&HEA01")]
    [InlineData("ea01", "&HEA01")]
    [InlineData("&HEA01", "&HEA01")]
    [InlineData("&hea01&", "&HEA01")]
    [InlineData("U+EA01", "&HEA01")]
    [InlineData("u+ea01", "&HEA01")]
    [InlineData("0xEA01", "&HEA01")]
    [InlineData("0xea01", "&HEA01")]
    [InlineData("59905", "&HEA01")]
    [InlineData("0", "&H0")]
    [InlineData("", "")]
    public void NormalizeIconCharCode_NormalizesVariousFormatsToCanonicalHex(string input, string expected)
    {
        string actual = DesignerPropertyFactory.NormalizeIconCharCode(input);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void ParseCharCode_CorrectlyParsesHexAndDecimalValues()
    {
        Assert.Equal(0xEA01, DesignerPropertyFactory.ParseCharCode("&HEA01"));
        Assert.Equal(0xEA01, DesignerPropertyFactory.ParseCharCode("EA01"));
        Assert.Equal(0xEA01, DesignerPropertyFactory.ParseCharCode("U+EA01"));
        Assert.Equal(0xEA01, DesignerPropertyFactory.ParseCharCode("0xEA01"));
        Assert.Equal(59905, DesignerPropertyFactory.ParseCharCode("59905"));
        Assert.Null(DesignerPropertyFactory.ParseCharCode(null));
        Assert.Null(DesignerPropertyFactory.ParseCharCode(""));
    }

    [Fact]
    public void FormDesignerControl_NamedFontProperties_DoNotOverwriteMainFont()
    {
        var ctrl = new FormDesignerControl
        {
            TypeName = "MyProject.AxLabelPlus",
            Name = "AxLabelPlus1"
        };

        var mainFont = new FormFontDescription { Name = "Tahoma", Size = 8.25 };
        var iconFont = new FormFontDescription { Name = "IcoFont", Size = 14 };

        ctrl.Font = mainFont;
        ctrl.SetFontProperty("IconFont", iconFont);

        Assert.Equal("Tahoma", ctrl.Font.Name);
        Assert.Equal("Tahoma", ctrl.GetFontProperty("Font")?.Name);
        Assert.Equal("IcoFont", ctrl.GetFontProperty("IconFont")?.Name);
        Assert.Equal(14, ctrl.GetFontProperty("IconFont")?.Size);

        // Clonado preserva fuentes nombradas independientes
        var clone = ctrl.Clone();
        Assert.Equal("Tahoma", clone.Font?.Name);
        Assert.Equal("IcoFont", clone.GetFontProperty("IconFont")?.Name);

        // Modificar clone no afecta al original
        clone.SetFontProperty("IconFont", new FormFontDescription { Name = "FontAwesome", Size = 16 });
        Assert.Equal("FontAwesome", clone.GetFontProperty("IconFont")?.Name);
        Assert.Equal("IcoFont", ctrl.GetFontProperty("IconFont")?.Name);
    }

    [Fact]
    public void UserControl_DoesNotExposePhantomFontOrAppearance_WhenNotDeclaredInCode()
    {
        var ctrl = new FormDesignerControl
        {
            TypeName = "MyProject.AxLabelPlus",
            Name = "AxLabelPlus1",
            Left = 100,
            Top = 200,
            Width = 1500,
            Height = 400
        };

        // Simular que el UserControl solo expone IconFont e IconCharCode en su código .ctl
        DesignerPropertyFactory.UserControlPropertiesResolver = typeName =>
        {
            if (typeName.Contains("AxLabelPlus", StringComparison.OrdinalIgnoreCase))
            {
                return new List<UserControlExposedProperty>
                {
                    new UserControlExposedProperty
                    {
                        Name = "IconFont",
                        TypeName = "StdFont",
                        Category = "Font"
                    },
                    new UserControlExposedProperty
                    {
                        Name = "IconCharCode",
                        TypeName = "Variant",
                        DefaultValue = "EA01",
                        Category = "Appearance"
                    }
                };
            }
            return null;
        };

        try
        {
            var props = DesignerPropertyFactory.GetPropertiesForControl(ctrl, null, () => { });

            var propNames = props.Select(p => p.Name).ToList();

            // Extender properties esperadas
            Assert.Contains("Name", propNames);
            Assert.Contains("Index", propNames);
            Assert.Contains("Tag", propNames);
            Assert.Contains("Left", propNames);
            Assert.Contains("Top", propNames);
            Assert.Contains("Width", propNames);
            Assert.Contains("Height", propNames);
            Assert.Contains("Visible", propNames);
            Assert.Contains("Enabled", propNames);
            Assert.Contains("TabIndex", propNames);
            Assert.Contains("TabStop", propNames);

            // Propiedades legítimas del UserControl
            Assert.Contains("IconFont", propNames);
            Assert.Contains("IconCharCode", propNames);

            // NO debe tener propiedades intrínsecas no codificadas (Font, BackColor, ForeColor, Caption, etc.)
            Assert.DoesNotContain("Font", propNames);
            Assert.DoesNotContain("Caption", propNames);
            Assert.DoesNotContain("Text", propNames);
            Assert.DoesNotContain("BackColor", propNames);
            Assert.DoesNotContain("ForeColor", propNames);
            Assert.DoesNotContain("Alignment", propNames);
            Assert.DoesNotContain("MultiLine", propNames);
            Assert.DoesNotContain("Picture", propNames);

            // Tipos de las propiedades de icono
            var iconFontProp = props.First(p => p.Name == "IconFont");
            Assert.Equal(DesignerPropertyType.FontDescription, iconFontProp.Type);

            var iconCharCodeProp = props.First(p => p.Name == "IconCharCode");
            Assert.Equal(DesignerPropertyType.Text, iconCharCodeProp.Type);

            // Validador de IconCharCode en ValueSetter
            iconCharCodeProp.SetValue("ea02");
            Assert.Equal("&HEA02", ctrl.GetCustomProperty("IconCharCode"));
        }
        finally
        {
            DesignerPropertyFactory.UserControlPropertiesResolver = null;
        }
    }

    [Fact]
    public void UserControl_ExposesFont_WhenExplicitlyDeclaredInCode()
    {
        var ctrl = new FormDesignerControl
        {
            TypeName = "MyProject.LabelPlus",
            Name = "LabelPlus1"
        };

        DesignerPropertyFactory.UserControlPropertiesResolver = typeName =>
        {
            return new List<UserControlExposedProperty>
            {
                new UserControlExposedProperty { Name = "Font", TypeName = "StdFont", Category = "Font" },
                new UserControlExposedProperty { Name = "IconFont", TypeName = "StdFont", Category = "Font" },
                new UserControlExposedProperty { Name = "Caption", TypeName = "String", Category = "Appearance" }
            };
        };

        try
        {
            var props = DesignerPropertyFactory.GetPropertiesForControl(ctrl, null, () => { });
            var propNames = props.Select(p => p.Name).ToList();

            Assert.Contains("Font", propNames);
            Assert.Contains("IconFont", propNames);
            Assert.Contains("Caption", propNames);
            Assert.DoesNotContain("BackColor", propNames);
        }
        finally
        {
            DesignerPropertyFactory.UserControlPropertiesResolver = null;
        }
    }

    [Fact]
    public void FormDesignParser_And_Serializer_RoundTrip_NamedFonts()
    {
        string frmContent = @"VERSION 5.00
Begin VB.Form Form1 
   Caption         =   ""Test Form""
   ClientHeight    =   4000
   ClientWidth     =   6000
   Begin MyProject.AxLabelPlus AxLabelPlus1 
      Height          =   500
      Left            =   120
      Top             =   240
      Width           =   1800
      IconCharCode    =   ""&HEA01""
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ""IcoFont""
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
";

        var model = FormDesignParser.Parse(frmContent);
        Assert.Single(model.Children);

        var ctrl = model.Children[0];
        Assert.Equal("MyProject.AxLabelPlus", ctrl.TypeName);
        Assert.Equal("AxLabelPlus1", ctrl.Name);

        // Verificar que no se creó Font estándar fantasma
        Assert.Null(ctrl.Font);

        // Verificar que IconFont se parseó adecuadamente
        var iconFont = ctrl.GetFontProperty("IconFont");
        Assert.NotNull(iconFont);
        Assert.Equal("IcoFont", iconFont.Name);
        Assert.Equal(12, iconFont.Size);

        // Serializar de vuelta y verificar que solo emite BeginProperty IconFont
        string serialized = FormDesignSerializer.Serialize(model);
        Assert.Contains("BeginProperty IconFont", serialized);
        Assert.Contains("\"IcoFont\"", serialized);
        Assert.DoesNotContain("BeginProperty Font", serialized);
    }

    [Fact]
    public void UserControlPropertyExtractor_ExtractsIconCharCodeAsText_AndIconFontAsFont()
    {
        string ctlCode = @"VERSION 5.00
Begin VB.UserControl AxLabelPlus 
   ClientHeight    =   3600
   ClientWidth     =   4800
End
Attribute VB_Name = ""AxLabelPlus""

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(ByVal New_IconFont As StdFont)
    Set m_IconFont = New_IconFont
    PropertyChanged ""IconFont""
End Property

Public Property Get IconCharCode() As Variant
    IconCharCode = m_IconCharCode
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As Variant)
    m_IconCharCode = New_IconCharCode
    PropertyChanged ""IconCharCode""
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_IconFont = PropBag.ReadProperty(""IconFont"", Ambient.Font)
    m_IconCharCode = PropBag.ReadProperty(""IconCharCode"", 0)
End Sub
";

        var props = UserControlPropertyExtractor.ExtractProperties(ctlCode);
        Assert.NotNull(props);

        var iconFontProp = props.FirstOrDefault(p => p.Name == "IconFont");
        Assert.NotNull(iconFontProp);
        Assert.Equal(DesignerPropertyType.FontDescription, iconFontProp.InferredType);

        var iconCharCodeProp = props.FirstOrDefault(p => p.Name == "IconCharCode");
        Assert.NotNull(iconCharCodeProp);
        Assert.Equal(DesignerPropertyType.Text, iconCharCodeProp.InferredType);
    }
}
