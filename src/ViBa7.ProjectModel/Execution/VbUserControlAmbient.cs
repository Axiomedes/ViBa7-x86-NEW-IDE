using System;

namespace ViBa7.ProjectModel.Execution;

/// <summary>
/// Objeto ambiental AmbientProperties expuesto en tiempo de ejecución y diseño a los controles de usuario (UserControl).
/// Proporciona acceso a Ambient.UserMode, Ambient.DisplayName, Ambient.BackColor, etc.
/// </summary>
public sealed class VbUserControlAmbient : ILiveControl
{
    private readonly bool _userMode;
    private readonly string _displayName;
    private readonly long _backColor;
    private readonly long _foreColor;

    public string Name => "Ambient";
    public string TypeName => "AmbientProperties";

    public bool UserMode => _userMode;
    public string DisplayName => _displayName;

    public VbUserControlAmbient(
        bool isUserMode = true,
        string displayName = "UserControl1",
        long backColor = 0x8000000F, // &H8000000F& (vbButtonFace)
        long foreColor = 0x80000012) // &H80000012& (vbButtonText)
    {
        _userMode = isUserMode;
        _displayName = displayName;
        _backColor = backColor;
        _foreColor = foreColor;
    }

    public VbValue GetProperty(string propertyName)
    {
        switch (propertyName.ToUpperInvariant())
        {
            case "USERMODE":
                return VbValue.FromBoolean(_userMode);

            case "DISPLAYNAME":
                return VbValue.FromString(_displayName);

            case "BACKCOLOR":
                return VbValue.FromLong(unchecked((int)_backColor));

            case "FORECOLOR":
                return VbValue.FromLong(unchecked((int)_foreColor));

            case "LOCALEID":
                return VbValue.FromLong(0x00000409);

            case "SCALEUNITS":
                return VbValue.FromString("Twip");

            case "SHOWHINT":
                return VbValue.FromBoolean(true);

            case "SHOWGRABHANDLES":
                return VbValue.FromBoolean(!_userMode);

            case "UIDEAD":
                return VbValue.FromBoolean(false);

            case "RIGHTTOLEFT":
                return VbValue.FromBoolean(false);

            default:
                return VbValue.Empty;
        }
    }

    public void SetProperty(string propertyName, VbValue value)
    {
        // Las propiedades de Ambient son de solo lectura en VB6
    }

    public VbValue InvokeMethod(string methodName, VbValue[] args)
    {
        return VbValue.Empty;
    }
}
