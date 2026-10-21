using System;
using ViBa7.ProjectModel.Designers;

namespace ViBa7.ProjectModel.Execution;

/// <summary>
/// Objeto extendido VBControlExtender expuesto a controles UserControl (.ctl) en tiempo de ejecución.
/// Proporciona acceso a Extender.Name, Extender.Index, Extender.Left, Extender.Top, etc.
/// </summary>
public sealed class VbUserControlExtender : ILiveControl
{
    private readonly FormDesignerControl _control;

    public string Name => "Extender";
    public string TypeName => "VBControlExtender";

    public FormDesignerControl Control => _control;

    public VbUserControlExtender(FormDesignerControl control)
    {
        _control = control ?? throw new ArgumentNullException(nameof(control));
    }

    public VbValue GetProperty(string propertyName)
    {
        switch (propertyName.ToUpperInvariant())
        {
            case "NAME":
                return VbValue.FromString(_control.Name);

            case "INDEX":
                return _control.Index.HasValue ? VbValue.FromLong(_control.Index.Value) : VbValue.Empty;

            case "LEFT":
                return VbValue.FromDouble(_control.Left);

            case "TOP":
                return VbValue.FromDouble(_control.Top);

            case "WIDTH":
                return VbValue.FromDouble(_control.Width);

            case "HEIGHT":
                return VbValue.FromDouble(_control.Height);

            case "VISIBLE":
                return VbValue.FromBoolean(_control.Visible);

            case "ENABLED":
                return VbValue.FromBoolean(_control.Enabled);

            case "TABINDEX":
                return VbValue.FromLong(_control.TabIndex ?? 0);

            case "TABSTOP":
                return VbValue.FromBoolean(_control.TabStop ?? true);

            case "TAG":
                return VbValue.FromString(_control.GetCustomPropertyString("Tag") ?? string.Empty);

            case "TOOLTIPTEXT":
                return VbValue.FromString(_control.GetCustomPropertyString("ToolTipText") ?? string.Empty);

            case "HELPCONTEXTID":
                return VbValue.FromLong(_control.HelpContextID);

            default:
                var prop = _control.CustomProperties.Find(p => p.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
                return prop != null ? VbValue.FromString(prop.Value) : VbValue.Empty;
        }
    }

    public void SetProperty(string propertyName, VbValue value)
    {
        switch (propertyName.ToUpperInvariant())
        {
            case "LEFT":
                _control.Left = (int)value.AsDouble();
                break;

            case "TOP":
                _control.Top = (int)value.AsDouble();
                break;

            case "WIDTH":
                _control.Width = (int)value.AsDouble();
                break;

            case "HEIGHT":
                _control.Height = (int)value.AsDouble();
                break;

            case "VISIBLE":
                _control.Visible = value.AsBoolean();
                break;

            case "ENABLED":
                _control.Enabled = value.AsBoolean();
                break;

            case "TABINDEX":
                _control.TabIndex = (int)value.AsLong();
                break;

            case "TABSTOP":
                _control.TabStop = value.AsBoolean();
                break;

            case "TAG":
                _control.SetCustomProperty("Tag", value.AsString());
                break;

            case "TOOLTIPTEXT":
                _control.SetCustomProperty("ToolTipText", value.AsString());
                break;

            case "HELPCONTEXTID":
                _control.HelpContextID = (int)value.AsLong();
                break;

            default:
                _control.SetCustomProperty(propertyName, value.AsString());
                break;
        }
    }

    public VbValue InvokeMethod(string methodName, VbValue[] args)
    {
        switch (methodName.ToUpperInvariant())
        {
            case "MOVE":
                if (args.Length > 0) _control.Left = (int)args[0].AsDouble();
                if (args.Length > 1) _control.Top = (int)args[1].AsDouble();
                if (args.Length > 2) _control.Width = (int)args[2].AsDouble();
                if (args.Length > 3) _control.Height = (int)args[3].AsDouble();
                return VbValue.Empty;

            default:
                return VbValue.Empty;
        }
    }
}
