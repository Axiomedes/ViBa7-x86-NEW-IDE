using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.IntelliSense;
using ViBa7.ProjectModel.Vbp;
using Xunit;

namespace ViBa7.Tests;

/// <summary>
/// Suite de pruebas automatizadas para certificar la carga y exposición funcional
/// de UserControls (.ctl) y controles ActiveX (.ocx) de terceros en el diseñador visual
/// bajo la arquitectura ViBa7-x86.
/// </summary>
public class ViBa7x86DesignSurfaceVerificationTests : IDisposable
{
    private readonly string _tempDir;

    static ViBa7x86DesignSurfaceVerificationTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public ViBa7x86DesignSurfaceVerificationTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "ViBa7x86_Verification_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Limpieza temporal
        }
    }

    [Fact]
    public void UserControl_FullLifecycle_Loading_PropertyExtraction_And_Serialization()
    {
        // 1. Crear código fuente de un UserControl (.ctl) con enumeraciones, propiedades públicas y persistencia
        string ctlContent = @"VERSION 5.00
Begin VB.UserControl CustomCard
   ClientHeight    =   3600
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.CommandButton btnAction
      Caption         =   ""Ejecutar""
      Height          =   495
      Left            =   1200
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label lblTitle
      Caption         =   ""Título de la Tarjeta""
      Height          =   375
      Left            =   240
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = ""CustomCard""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CardTheme
    [Classic Gray] = 0
    [Electric Blue] = 1
    [Emerald Dark] = 2
End Enum

Private m_Theme As CardTheme
Private m_CardHeader As String
Private m_AccentColor As OLE_COLOR
Private m_IsElevated As Boolean

Public Property Get Theme() As CardTheme
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal newTheme As CardTheme)
    m_Theme = newTheme
    PropertyChanged ""Theme""
End Property

Public Property Get CardHeader() As String
    CardHeader = m_CardHeader
End Property

Public Property Let CardHeader(ByVal newHeader As String)
    m_CardHeader = newHeader
    PropertyChanged ""CardHeader""
End Property

Public Property Get AccentColor() As OLE_COLOR
    AccentColor = m_AccentColor
End Property

Public Property Let AccentColor(ByVal newColor As OLE_COLOR)
    m_AccentColor = newColor
    PropertyChanged ""AccentColor""
End Property

Public Property Get IsElevated() As Boolean
    IsElevated = m_IsElevated
End Property

Public Property Let IsElevated(ByVal newVal As Boolean)
    m_IsElevated = newVal
    PropertyChanged ""IsElevated""
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Theme = PropBag.ReadProperty(""Theme"", 0)
    m_CardHeader = PropBag.ReadProperty(""CardHeader"", ""Encabezado Inicial"")
    m_AccentColor = PropBag.ReadProperty(""AccentColor"", &H00FF8000&)
    m_IsElevated = PropBag.ReadProperty(""IsElevated"", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(""Theme"", m_Theme, 0)
    Call PropBag.WriteProperty(""CardHeader"", m_CardHeader, ""Encabezado Inicial"")
    Call PropBag.WriteProperty(""AccentColor"", m_AccentColor, &H00FF8000&)
    Call PropBag.WriteProperty(""IsElevated"", m_IsElevated, True)
End Sub
";
        string ctlPath = Path.Combine(_tempDir, "CustomCard.ctl");
        File.WriteAllText(ctlPath, ctlContent, Encoding.Default);

        // 2. Crear proyecto .vbp que contiene el UserControl y un Formulario que lo consume
        string frmContent = @"VERSION 5.00
Begin VB.Form frmMain
   Caption         =   ""Ventana Principal""
   ClientHeight    =   6000
   ClientWidth     =   8000
   Begin MiProyecto.CustomCard Card1
      Height          =   3600
      Left            =   600
      Top             =   600
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
      Theme           =   1
      CardHeader      =   ""Facturación Mensual""
      AccentColor     =   &H000080FF&
      IsElevated      =   -1  'True
   End
End
Attribute VB_Name = ""frmMain""
Option Explicit
";
        string frmPath = Path.Combine(_tempDir, "frmMain.frm");
        File.WriteAllText(frmPath, frmContent, Encoding.Default);

        string vbpContent = @"Type=Exe
Form=frmMain.frm
UserControl=CustomCard.ctl
Name=""MiProyecto""
Startup=""frmMain""
";
        string vbpPath = Path.Combine(_tempDir, "MiProyecto.vbp");
        File.WriteAllText(vbpPath, vbpContent, Encoding.Default);

        // 3. Cargar el proyecto con VbpParser
        var project = VbpParser.Parse(vbpPath);
        Assert.NotNull(project);
        Assert.Single(project.UserControls);
        Assert.Single(project.Forms);

        // 4. Verificar resolución del modelo visual del UserControl
        var ucModel = project.GetUserControlDesignModel("CustomCard");
        Assert.NotNull(ucModel);
        Assert.Equal(2, ucModel.Children.Count); // btnAction y lblTitle
        Assert.Contains(ucModel.Children, c => c.Name == "btnAction");
        Assert.Contains(ucModel.Children, c => c.Name == "lblTitle");

        // 5. Configurar el resolver de propiedades del UserControl como hace el IDE
        DesignerPropertyFactory.UserControlPropertiesResolver = typeName => project.GetUserControlDesignModel(typeName)?.ExposedProperties;

        // 6. Parsear el formulario consumidor
        var formModel = FormDesignParser.Parse(frmContent);
        Assert.NotNull(formModel);
        Assert.Single(formModel.Children);

        var cardControl = formModel.Children[0];
        Assert.Equal("MiProyecto.CustomCard", cardControl.TypeName);
        Assert.Equal("Card1", cardControl.Name);

        // 7. Generar lista de propiedades funcionales para el PropertyGrid
        bool modelChangedCalled = false;
        var props = DesignerPropertyFactory.GetPropertiesForControl(
            cardControl,
            formModel,
            () => { modelChangedCalled = true; });

        Assert.NotEmpty(props);

        // Debe contener las propiedades estándar de identificación y geometría
        var propName = props.FirstOrDefault(p => p.Name == "Name");
        Assert.NotNull(propName);
        Assert.Equal("Card1", propName.CurrentValue);

        // Debe exponer la propiedad (Custom) para UserControls / Controles no intrínsecos
        var propCustom = props.FirstOrDefault(p => p.Name == "Custom");
        Assert.NotNull(propCustom);

        // 8. Verificar que las 4 propiedades públicas del UserControl fueron extraídas funcionalmente
        var propTheme = props.FirstOrDefault(p => p.Name == "Theme");
        Assert.NotNull(propTheme);
        Assert.Equal(DesignerPropertyType.EnumChoice, propTheme.Type);
        Assert.Equal(3, propTheme.Choices.Count);
        Assert.Contains(propTheme.Choices, c => c.DisplayText == "0 - Classic Gray");
        Assert.Contains(propTheme.Choices, c => c.DisplayText == "1 - Electric Blue");
        Assert.Contains(propTheme.Choices, c => c.DisplayText == "2 - Emerald Dark");
        Assert.Equal("1", propTheme.CurrentValue);

        var propHeader = props.FirstOrDefault(p => p.Name == "CardHeader");
        Assert.NotNull(propHeader);
        Assert.Equal("\"Facturación Mensual\"", propHeader.CurrentValue);

        var propAccent = props.FirstOrDefault(p => p.Name == "AccentColor");
        Assert.NotNull(propAccent);
        Assert.Equal(DesignerPropertyType.ColorHex, propAccent.Type);
        Assert.Equal("&H000080FF&", propAccent.CurrentValue);

        var propElevated = props.FirstOrDefault(p => p.Name == "IsElevated");
        Assert.NotNull(propElevated);
        Assert.Equal(DesignerPropertyType.Boolean, propElevated.Type);

        // 9. Modificar funcionalmente las propiedades vía ValueSetter (simulando interacción de usuario en el diseñador)
        modelChangedCalled = false;
        propTheme!.ValueSetter?.Invoke("2");
        Assert.True(modelChangedCalled);
        Assert.Equal("2", cardControl.GetCustomProperty("Theme"));

        modelChangedCalled = false;
        propHeader!.ValueSetter?.Invoke("\"Balance General 2026\"");
        Assert.True(modelChangedCalled);
        Assert.Equal("\"Balance General 2026\"", cardControl.GetCustomProperty("CardHeader"));

        modelChangedCalled = false;
        propAccent!.ValueSetter?.Invoke("&H0000FF00&");
        Assert.True(modelChangedCalled);
        Assert.Equal("&H0000FF00&", cardControl.GetCustomProperty("AccentColor"));

        // 10. Serializar y verificar que la persistencia en el .frm preserva los valores actualizados
        string serialized = FormDesignSerializer.Serialize(formModel);
        Assert.Contains("Theme           =   2", serialized);
        Assert.Contains("CardHeader      =   \"Balance General 2026\"", serialized);
        Assert.Contains("AccentColor     =   &H0000FF00&", serialized);
    }

    [Fact]
    public void ThirdPartyOcx_PropertyExposure_And_HiMetricConversion_Functional()
    {
        // 1. Simular formulario con un control de terceros tipo Grid / ListView con propiedades ricas y fuentes
        string frmContent = @"VERSION 5.00
Object = ""{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0""; ""MSFLXGRD.OCX""
Begin VB.Form frmGrid
   Caption         =   ""Visor de Datos""
   ClientHeight    =   6500
   ClientWidth     =   9000
   Begin MSFlexGridLib.MSFlexGrid GridData
      Height          =   4500
      Left            =   360
      Top             =   480
      Width           =   8200
      _ExtentX        =   14464
      _ExtentY        =   7938
      _Version        =   393216
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      AllowUserResizing=   3
      GridColor       =   &H00C0C0C0&
      FocusRect       =   2
      HighLight       =   1
      SelectionMode   =   1
      FormatString    =   ""Codigo|Descripcion|Precio|Stock""
   End
End
Attribute VB_Name = ""frmGrid""
";
        var formModel = FormDesignParser.Parse(frmContent);
        Assert.NotNull(formModel);
        Assert.Single(formModel.Children);

        var gridControl = formModel.Children[0];
        Assert.Equal("MSFlexGridLib.MSFlexGrid", gridControl.TypeName);
        Assert.Equal("GridData", gridControl.Name);

        // 2. Verificar Extents en HiMetric y conversión a Twips
        Assert.Equal(14464, gridControl.ExtentX);
        Assert.Equal(7938, gridControl.ExtentY);
        Assert.Equal(8200, gridControl.Width);
        Assert.Equal(4500, gridControl.Height);

        // 3. Obtener lista de propiedades expuestas por el diseñador
        bool modelChangedCalled = false;
        var props = DesignerPropertyFactory.GetPropertiesForControl(
            gridControl,
            formModel,
            () => { modelChangedCalled = true; });

        // Debe exponer (Custom) al inicio para controles OCX de terceros
        Assert.Equal("Custom", props[0].Name);

        // Debe exponer todas las propiedades personalizadas cargadas del .frm
        var propRows = props.FirstOrDefault(p => p.Name == "Rows");
        Assert.NotNull(propRows);
        Assert.Equal("50", propRows.CurrentValue);

        var propCols = props.FirstOrDefault(p => p.Name == "Cols");
        Assert.NotNull(propCols);
        Assert.Equal("10", propCols.CurrentValue);

        var propGridColor = props.FirstOrDefault(p => p.Name == "GridColor");
        Assert.NotNull(propGridColor);
        Assert.Equal(DesignerPropertyType.ColorHex, propGridColor.Type);
        Assert.Equal("&H00C0C0C0&", propGridColor.CurrentValue);

        var propFormat = props.FirstOrDefault(p => p.Name == "FormatString");
        Assert.NotNull(propFormat);
        Assert.Equal("\"Codigo|Descripcion|Precio|Stock\"", propFormat.CurrentValue);

        // 4. Modificar dimensiones del control y verificar sincronización automática de HiMetric ExtentX / ExtentY
        var propWidth = props.FirstOrDefault(p => p.Name == "Width");
        Assert.NotNull(propWidth);

        modelChangedCalled = false;
        propWidth!.ValueSetter?.Invoke("9000");
        Assert.True(modelChangedCalled);
        Assert.Equal(9000, gridControl.Width);
        // 9000 Twips * 2540 / 1440 = 15875 HiMetric
        Assert.Equal(15875, gridControl.ExtentX);

        var propHeight = props.FirstOrDefault(p => p.Name == "Height");
        Assert.NotNull(propHeight);

        modelChangedCalled = false;
        propHeight!.ValueSetter?.Invoke("5000");
        Assert.True(modelChangedCalled);
        Assert.Equal(5000, gridControl.Height);
        // 5000 Twips * 2540 / 1440 = 8819 HiMetric
        Assert.Equal(8819, gridControl.ExtentY);

        // 5. Modificar propiedad Rows
        modelChangedCalled = false;
        propRows!.ValueSetter?.Invoke("100");
        Assert.True(modelChangedCalled);
        Assert.Equal("100", gridControl.GetCustomProperty("Rows"));

        // 6. Serializar a .frm y validar que los Extents calculados y las propiedades se persisten fidedignamente
        string serialized = FormDesignSerializer.Serialize(formModel);
        Assert.Contains("_ExtentX        =   15875", serialized);
        Assert.Contains("_ExtentY        =   8819", serialized);
        Assert.Contains("Rows            =   100", serialized);
        Assert.Contains("Cols            =   10", serialized);
        Assert.Contains("FormatString    =   \"Codigo|Descripcion|Precio|Stock\"", serialized);
    }

    [Fact]
    public void TypeLibReader_RegisterAndExpose_ActiveXControls_To_Toolbox()
    {
        // 1. Simular biblioteca de tipos de un OCX con controles y enumeraciones
        var libInfo = new OcxLibraryInfo
        {
            LibraryName = "MiControlLib",
            FilePath = "C:\\Windows\\SysWOW64\\MiControl.ocx",
            Guid = "{12345678-1234-1234-1234-1234567890AB}",
            MajorVersion = 1,
            MinorVersion = 0,
            Description = "Biblioteca de Controles Avanzados"
        };

        var ctrlInfo = new OcxControlInfo
        {
            Name = "SuperGauge",
            Clsid = "{87654321-4321-4321-4321-BA0987654321}",
            Description = "Control de Manómetro Gráfico",
            Members = new List<OcxMemberInfo>
            {
                new OcxMemberInfo
                {
                    Name = "Value",
                    Kind = VbSymbolKind.Property,
                    ReturnType = "Double",
                    Description = "Valor actual del medidor",
                    IsPropertyGet = true,
                    IsPropertyPut = true
                },
                new OcxMemberInfo
                {
                    Name = "GaugeStyle",
                    Kind = VbSymbolKind.Property,
                    ReturnType = "GaugeStyles",
                    Description = "Estilo visual del medidor",
                    IsPropertyGet = true,
                    IsPropertyPut = true
                },
                new OcxMemberInfo
                {
                    Name = "NeedleColor",
                    Kind = VbSymbolKind.Property,
                    ReturnType = "OLE_COLOR",
                    Description = "Color de la aguja",
                    IsPropertyGet = true,
                    IsPropertyPut = true
                }
            }
        };
        libInfo.Controls.Add(ctrlInfo);

        // 2. Registrar en catálogo de miembros y en la Caja de Herramientas (Toolbox)
        VbControlMemberCatalog.RegisterOcxLibrary(libInfo);
        ControlToolboxRegistry.RegisterOcxTools(libInfo);

        // 3. Comprobar que el control está disponible en el catálogo de herramientas
        var tool = ControlToolboxRegistry.Find("MiControlLib.SuperGauge");
        Assert.NotNull(tool);
        Assert.Equal("SuperGauge", tool.DisplayName);
        Assert.Equal("MiControlLib.SuperGauge", tool.TypeName);

        // 4. Instanciar el control en un modelo de diseño
        var formModel = new FormDesignModel { Name = "Form1" };
        var created = ControlToolboxRegistry.CreateControl("MiControlLib.SuperGauge", formModel, 500, 500);
        formModel.Children.Add(created);

        Assert.NotNull(created);
        Assert.Equal("MiControlLib.SuperGauge", created.TypeName);
        Assert.StartsWith("SuperGauge", created.Name);

        // 5. Validar que DesignerPropertyFactory extrae los miembros del TypeLib para el PropertyGrid
        var props = DesignerPropertyFactory.GetPropertiesForControl(created, formModel, () => { });

        var propValue = props.FirstOrDefault(p => p.Name == "Value");
        Assert.NotNull(propValue);
        Assert.Equal("ActiveX / TypeLib", propValue.Category);

        var propColor = props.FirstOrDefault(p => p.Name == "NeedleColor");
        Assert.NotNull(propColor);
        Assert.Equal(DesignerPropertyType.ColorHex, propColor.Type);
        Assert.Equal("ActiveX / TypeLib", propColor.Category);

        // 6. Probar edición
        propValue!.ValueSetter?.Invoke("75.5");
        Assert.Equal("75.5", created.GetCustomProperty("Value"));

        propColor!.ValueSetter?.Invoke("&H00FF0000&");
        Assert.Equal("&H00FF0000&", created.GetCustomProperty("NeedleColor"));

        // Limpiar registro
        ControlToolboxRegistry.UnregisterOcxTools("MiControlLib");
        VbControlMemberCatalog.UnregisterOcxLibrary("MiControlLib");
    }
}
