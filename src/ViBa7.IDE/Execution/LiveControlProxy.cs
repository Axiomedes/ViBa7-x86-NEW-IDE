using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using ViBa7.ActiveXHost;
using ViBa7.ProjectModel.Designers;
using ViBa7.ProjectModel.Execution;

namespace ViBa7.IDE.Execution;

/// <summary>
/// Proxy dinámico que envuelve un sub-objeto COM IDispatch retornado por un control ActiveX en tiempo de ejecución
/// (por ejemplo: TreeView.Nodes, ListView.ListItems, ColumnHeaders, Toolbar.Buttons, StatusBar.Panels, StdFont)
/// permitiendo encadenar llamadas y asignaciones desde el intérprete de VB6.
/// </summary>
public sealed class LiveComDispatchProxy : ILiveControl
{
    private readonly object _comObject;
    private readonly Dispatcher _dispatcher;

    public string Name { get; }
    public string TypeName => "COMDispatchObject";
    public object ComObject => _comObject;

    public LiveComDispatchProxy(object comObject, string name, Dispatcher dispatcher)
    {
        _comObject = comObject ?? throw new ArgumentNullException(nameof(comObject));
        Name = string.IsNullOrWhiteSpace(name) ? "COMObject" : name;
        _dispatcher = dispatcher ?? Dispatcher.CurrentDispatcher;
    }

    public VbValue GetProperty(string propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return VbValue.Empty;

        return _dispatcher.Invoke(() =>
        {
            object? res = DynamicActiveXControl.InvokeComDispatchMember(_comObject, propertyName);
            return LiveControlProxy.ConvertComResultToVbValue(res, propertyName, _dispatcher);
        });
    }

    public void SetProperty(string propertyName, VbValue value)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return;

        _dispatcher.Invoke(() =>
        {
            object? rawVal = LiveControlProxy.UnwrapVbValueForCom(value);
            DynamicActiveXControl.SetComDispatchProperty(_comObject, propertyName, rawVal);
        });
    }

    public void SetIndexedProperty(string propertyName, VbValue[] indices, VbValue value)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return;

        _dispatcher.Invoke(() =>
        {
            object? rawVal = LiveControlProxy.UnwrapVbValueForCom(value);
            object?[]? rawIndices = LiveControlProxy.UnwrapVbValuesForCom(indices);
            DynamicActiveXControl.SetComDispatchProperty(_comObject, propertyName, rawVal, rawIndices);
        });
    }

    public VbValue InvokeMethod(string methodName, VbValue[] args)
    {
        return _dispatcher.Invoke(() =>
        {
            object?[]? rawArgs = LiveControlProxy.UnwrapVbValuesForCom(args);
            object? res = DynamicActiveXControl.InvokeComDispatchMember(_comObject, methodName, rawArgs);
            return LiveControlProxy.ConvertComResultToVbValue(res, methodName, _dispatcher);
        });
    }
}

/// <summary>
/// Proxy interactivo en vivo de un control de VB6 o ActiveX (OCX) montado sobre un elemento visual WPF o HwndHost.
/// Permite leer y modificar propiedades escalares e indexadas en caliente (Text, Caption, TextMatrix, TabCaption, etc.),
/// navegar sub-objetos COM (Nodes, ListItems, Panels, Font) y ejecutar métodos (SetFocus, Move, AddItem, Clear, etc.).
/// </summary>
public class LiveControlProxy : ILiveControl
{
    private readonly FormDesignerControl _designerControl;
    private readonly FrameworkElement _visual;
    private readonly DispatcherTimer? _timer;
    private readonly Dictionary<string, VbValue> _dynamicProperties = new(StringComparer.OrdinalIgnoreCase);

    public string Name => _designerControl.Name;
    public string TypeName => _designerControl.TypeName;
    public FrameworkElement Visual => _visual;
    public FormDesignerControl DesignerControl => _designerControl;

    public LiveControlProxy(FormDesignerControl designerControl, FrameworkElement visual, DispatcherTimer? timer = null)
    {
        _designerControl = designerControl ?? throw new ArgumentNullException(nameof(designerControl));
        _visual = visual ?? throw new ArgumentNullException(nameof(visual));
        _timer = timer;
    }

    private DynamicActiveXControl? TryGetActiveXControl()
    {
        if (_visual is ActiveXRuntimeHost rtHost)
        {
            if (rtHost.AxControl.GetActiveXObject() == null && !rtHost.AxControl.HasInitializationError)
            {
                rtHost.EnsureControlCreatedEarly();
            }
            if (!rtHost.AxControl.HasInitializationError && rtHost.AxControl.GetActiveXObject() != null)
            {
                return rtHost.AxControl;
            }
        }
        else if (_visual is ActiveXInPlaceHost designHost)
        {
            if (!designHost.AxControl.HasInitializationError && designHost.AxControl.GetActiveXObject() != null)
            {
                return designHost.AxControl;
            }
        }

        return null;
    }

    public VbValue GetProperty(string propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return VbValue.Empty;

        return _visual.Dispatcher.Invoke(() =>
        {
            var axCtrl = TryGetActiveXControl();

            switch (propertyName.ToUpperInvariant())
            {
                case "NAME":
                    return VbValue.FromString(_designerControl.Name);

                case "INDEX":
                    return VbValue.FromLong(_designerControl.Index ?? 0);

                case "ENABLED":
                    if (axCtrl != null)
                    {
                        var enVal = axCtrl.GetPropertyValue("Enabled");
                        if (enVal is bool bEn) return VbValue.FromBoolean(bEn);
                    }
                    return VbValue.FromBoolean(_timer != null ? _timer.IsEnabled : _visual.IsEnabled);

                case "INTERVAL":
                    return VbValue.FromLong(_designerControl.Interval);

                case "VISIBLE":
                    return VbValue.FromBoolean(_visual.Visibility == Visibility.Visible);

                case "LEFT":
                    return VbValue.FromDouble(Canvas.GetLeft(_visual) * FormDesignModel.TwipsPerPixel);

                case "TOP":
                    return VbValue.FromDouble(Canvas.GetTop(_visual) * FormDesignModel.TwipsPerPixel);

                case "WIDTH":
                    return VbValue.FromDouble(_visual.Width * FormDesignModel.TwipsPerPixel);

                case "HEIGHT":
                    return VbValue.FromDouble(_visual.Height * FormDesignModel.TwipsPerPixel);

                case "TAG":
                    return _dynamicProperties.TryGetValue("Tag", out var tag)
                        ? tag
                        : VbValue.FromString(_designerControl.GetCustomPropertyString("Tag") ?? _visual.Tag?.ToString() ?? string.Empty);

                case "TOOLTIPTEXT":
                    return _dynamicProperties.TryGetValue("ToolTipText", out var tip)
                        ? tip
                        : VbValue.FromString(_designerControl.GetCustomPropertyString("ToolTipText") ?? string.Empty);

                case "HWND":
                    if (axCtrl != null)
                    {
                        IntPtr hwnd = axCtrl.GetControlWindowHandle();
                        if (hwnd != IntPtr.Zero) return VbValue.FromLong(unchecked((int)hwnd.ToInt64()));
                    }
                    return VbValue.FromLong(0);

                case "OBJECT":
                    if (axCtrl?.GetActiveXObject() is object rawOcx)
                    {
                        return VbValue.FromObject(new LiveComDispatchProxy(rawOcx, $"{_designerControl.Name}.Object", _visual.Dispatcher));
                    }
                    return VbValue.FromObject(this);

                case "TEXT":
                case "CAPTION":
                case "VALUE":
                    if (axCtrl != null)
                    {
                        var liveProp = axCtrl.GetPropertyValue(propertyName);
                        if (liveProp != null)
                        {
                            return ConvertComResultToVbValue(liveProp, propertyName, _visual.Dispatcher);
                        }
                    }
                    return GetControlTextOrValue();

                case "LISTCOUNT":
                    if (_visual is ListBox lb) return VbValue.FromLong(lb.Items.Count);
                    if (_visual is ComboBox cb) return VbValue.FromLong(cb.Items.Count);
                    if (axCtrl != null)
                    {
                        var lc = axCtrl.GetPropertyValue("ListCount");
                        if (lc != null) return ConvertComResultToVbValue(lc, "ListCount", _visual.Dispatcher);
                    }
                    return VbValue.FromLong(0);

                case "LISTINDEX":
                    if (_visual is ListBox lbIdx) return VbValue.FromLong(lbIdx.SelectedIndex);
                    if (_visual is ComboBox cbIdx) return VbValue.FromLong(cbIdx.SelectedIndex);
                    if (axCtrl != null)
                    {
                        var li = axCtrl.GetPropertyValue("ListIndex");
                        if (li != null) return ConvertComResultToVbValue(li, "ListIndex", _visual.Dispatcher);
                    }
                    return VbValue.FromLong(-1);

                default:
                    if (axCtrl != null)
                    {
                        var ocxVal = axCtrl.GetPropertyValue(propertyName);
                        if (ocxVal != null)
                        {
                            return ConvertComResultToVbValue(ocxVal, propertyName, _visual.Dispatcher);
                        }
                    }

                    if (_dynamicProperties.TryGetValue(propertyName, out var dynamicVal))
                    {
                        return dynamicVal;
                    }
                    var prop = _designerControl.CustomProperties.Find(p => p.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase));
                    return prop != null ? VbValue.FromString(prop.Value) : VbValue.Empty;
            }
        });
    }

    public void SetProperty(string propertyName, VbValue value)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return;

        _visual.Dispatcher.Invoke(() =>
        {
            var axCtrl = TryGetActiveXControl();

            switch (propertyName.ToUpperInvariant())
            {
                case "TEXT":
                case "CAPTION":
                    if (axCtrl != null)
                    {
                        axCtrl.SetPropertyValue(propertyName, UnwrapVbValueForCom(value));
                    }
                    SetControlText(value.AsString());
                    break;

                case "VALUE":
                    if (axCtrl != null)
                    {
                        axCtrl.SetPropertyValue(propertyName, UnwrapVbValueForCom(value));
                    }
                    SetControlValue(value);
                    break;

                case "ENABLED":
                    bool en = value.AsBoolean();
                    _visual.IsEnabled = en;
                    _designerControl.Enabled = en;
                    if (axCtrl != null)
                    {
                        axCtrl.Extender.Enabled = en;
                        axCtrl.SetPropertyValue("Enabled", en);
                    }
                    if (_timer != null)
                    {
                        _timer.IsEnabled = en;
                        if (en && _designerControl.Interval > 0)
                        {
                            if (!_timer.IsEnabled) _timer.Start();
                        }
                        else
                        {
                            _timer.Stop();
                        }
                    }
                    break;

                case "INTERVAL":
                    int ms = Math.Max(0, Math.Min(65535, (int)value.AsLong()));
                    _designerControl.Interval = ms;
                    if (_timer != null)
                    {
                        if (ms > 0)
                        {
                            _timer.Interval = TimeSpan.FromMilliseconds(ms);
                            if (_designerControl.Enabled)
                            {
                                _timer.Start();
                            }
                        }
                        else
                        {
                            _timer.Stop();
                        }
                    }
                    break;

                case "VISIBLE":
                    bool vis = value.AsBoolean();
                    _visual.Visibility = vis ? Visibility.Visible : Visibility.Hidden;
                    _designerControl.Visible = vis;
                    if (axCtrl != null)
                    {
                        axCtrl.Extender.Visible = vis;
                    }
                    break;

                case "LEFT":
                    Canvas.SetLeft(_visual, value.AsDouble() / FormDesignModel.TwipsPerPixel);
                    _designerControl.Left = (int)Math.Round(value.AsDouble());
                    if (axCtrl != null) axCtrl.Extender.Left = (float)value.AsDouble();
                    break;

                case "TOP":
                    Canvas.SetTop(_visual, value.AsDouble() / FormDesignModel.TwipsPerPixel);
                    _designerControl.Top = (int)Math.Round(value.AsDouble());
                    if (axCtrl != null) axCtrl.Extender.Top = (float)value.AsDouble();
                    break;

                case "WIDTH":
                    _visual.Width = Math.Max(1, value.AsDouble() / FormDesignModel.TwipsPerPixel);
                    _designerControl.Width = (int)Math.Round(value.AsDouble());
                    if (axCtrl != null) axCtrl.Extender.Width = (float)value.AsDouble();
                    break;

                case "HEIGHT":
                    _visual.Height = Math.Max(1, value.AsDouble() / FormDesignModel.TwipsPerPixel);
                    _designerControl.Height = (int)Math.Round(value.AsDouble());
                    if (axCtrl != null) axCtrl.Extender.Height = (float)value.AsDouble();
                    break;

                case "BACKCOLOR":
                    if (axCtrl != null)
                    {
                        axCtrl.SetPropertyValue("BackColor", (uint)value.AsLong());
                    }
                    ApplyBackground(value.AsLong());
                    break;

                case "FORECOLOR":
                    if (axCtrl != null)
                    {
                        axCtrl.SetPropertyValue("ForeColor", (uint)value.AsLong());
                    }
                    ApplyForeground(value.AsLong());
                    break;

                case "LISTINDEX":
                    int idx = (int)value.AsLong();
                    if (_visual is ListBox lb) lb.SelectedIndex = idx;
                    else if (_visual is ComboBox cb) cb.SelectedIndex = idx;
                    else if (axCtrl != null) axCtrl.SetPropertyValue("ListIndex", idx);
                    break;

                default:
                    _dynamicProperties[propertyName] = value;
                    if (axCtrl != null)
                    {
                        axCtrl.SetPropertyValue(propertyName, UnwrapVbValueForCom(value));
                    }
                    break;
            }
        });
    }

    public void SetIndexedProperty(string propertyName, VbValue[] indices, VbValue value)
    {
        if (string.IsNullOrWhiteSpace(propertyName)) return;

        _visual.Dispatcher.Invoke(() =>
        {
            string key = $"{propertyName}({string.Join(",", indices.Select(i => i.AsString()))})";
            _dynamicProperties[key] = value;

            var axCtrl = TryGetActiveXControl();
            if (axCtrl != null)
            {
                object?[]? rawIndices = UnwrapVbValuesForCom(indices);
                if (rawIndices is { Length: > 0 })
                {
                    axCtrl.SetIndexedPropertyValue(propertyName, UnwrapVbValueForCom(value), rawIndices!);
                }
            }
        });
    }

    public VbValue InvokeMethod(string methodName, VbValue[] args)
    {
        return _visual.Dispatcher.Invoke(() =>
        {
            var axCtrl = TryGetActiveXControl();

            switch (methodName.ToUpperInvariant())
            {
                case "SETFOCUS":
                    if (axCtrl != null && axCtrl.FocusControl())
                    {
                        return VbValue.Empty;
                    }
                    _visual.Focus();
                    return VbValue.Empty;

                case "MOVE":
                    if (args.Length >= 1 && !args[0].IsEmpty)
                    {
                        SetProperty("Left", args[0]);
                    }
                    if (args.Length >= 2 && !args[1].IsEmpty)
                    {
                        SetProperty("Top", args[1]);
                    }
                    if (args.Length >= 3 && !args[2].IsEmpty)
                    {
                        SetProperty("Width", args[2]);
                    }
                    if (args.Length >= 4 && !args[3].IsEmpty)
                    {
                        SetProperty("Height", args[3]);
                    }
                    return VbValue.Empty;

                case "CLEAR":
                    if (_visual is TextBox tb) { tb.Clear(); return VbValue.Empty; }
                    if (_visual is ListBox lb) { lb.Items.Clear(); return VbValue.Empty; }
                    if (_visual is ComboBox cb) { cb.Items.Clear(); return VbValue.Empty; }
                    if (axCtrl != null)
                    {
                        axCtrl.InvokeMethod("Clear");
                    }
                    return VbValue.Empty;

                case "ADDITEM":
                    if (args.Length > 0)
                    {
                        if (_visual is ListBox lbAdd) { lbAdd.Items.Add(args[0].AsString()); return VbValue.Empty; }
                        if (_visual is ComboBox cbAdd) { cbAdd.Items.Add(args[0].AsString()); return VbValue.Empty; }
                        if (axCtrl != null)
                        {
                            axCtrl.InvokeMethod("AddItem", UnwrapVbValuesForCom(args));
                        }
                    }
                    return VbValue.Empty;

                case "REMOVEITEM":
                    if (args.Length > 0)
                    {
                        int removeIdx = (int)args[0].AsLong();
                        if (_visual is ListBox lbRem && removeIdx >= 0 && removeIdx < lbRem.Items.Count) { lbRem.Items.RemoveAt(removeIdx); return VbValue.Empty; }
                        if (_visual is ComboBox cbRem && removeIdx >= 0 && removeIdx < cbRem.Items.Count) { cbRem.Items.RemoveAt(removeIdx); return VbValue.Empty; }
                        if (axCtrl != null)
                        {
                            axCtrl.InvokeMethod("RemoveItem", UnwrapVbValuesForCom(args));
                        }
                    }
                    return VbValue.Empty;

                case "REFRESH":
                    if (axCtrl != null)
                    {
                        axCtrl.InvokeMethod("Refresh");
                    }
                    _visual.UpdateLayout();
                    return VbValue.Empty;

                default:
                    if (axCtrl != null)
                    {
                        object? res = axCtrl.InvokeMethod(methodName, UnwrapVbValuesForCom(args));
                        if (res != null)
                        {
                            return ConvertComResultToVbValue(res, methodName, _visual.Dispatcher);
                        }
                    }

                    // Respaldo para lectura de propiedades indexadas cacheadas (ej. TextMatrix(r, c), TabCaption(i))
                    if (args.Length > 0)
                    {
                        string key = $"{methodName}({string.Join(",", args.Select(a => a.AsString()))})";
                        if (_dynamicProperties.TryGetValue(key, out var cachedIndexed))
                        {
                            return cachedIndexed;
                        }
                    }

                    return VbValue.Empty;
            }
        });
    }

    internal static object? UnwrapVbValueForCom(VbValue val)
    {
        if (val.IsEmpty) return Type.Missing;
        if (val.RawValue is LiveComDispatchProxy comProxy) return comProxy.ComObject;
        return val.RawValue;
    }

    internal static object?[]? UnwrapVbValuesForCom(VbValue[]? args)
    {
        if (args == null || args.Length == 0) return null;
        var raw = new object?[args.Length];
        for (int i = 0; i < args.Length; i++)
        {
            raw[i] = UnwrapVbValueForCom(args[i]);
        }
        return raw;
    }

    internal static VbValue ConvertComResultToVbValue(object? comResult, string memberName, Dispatcher dispatcher)
    {
        if (comResult == null) return VbValue.Empty;

        if (Marshal.IsComObject(comResult))
        {
            return VbValue.FromObject(new LiveComDispatchProxy(comResult, memberName, dispatcher));
        }

        return comResult switch
        {
            bool b => VbValue.FromBoolean(b),
            short s => VbValue.FromLong(s),
            ushort us => VbValue.FromLong(us),
            int i => VbValue.FromLong(i),
            uint ui => VbValue.FromLong(unchecked((int)ui)),
            long l => VbValue.FromLong(unchecked((int)l)),
            float f => VbValue.FromDouble(f),
            double d => VbValue.FromDouble(d),
            decimal dec => VbValue.FromDouble((double)dec),
            string str => VbValue.FromString(str),
            _ => VbValue.FromString(comResult.ToString() ?? string.Empty)
        };
    }

    private VbValue GetControlTextOrValue()
    {
        if (_visual is TextBox tb) return VbValue.FromString(tb.Text);
        if (_visual is Button btn) return VbValue.FromString(btn.Content?.ToString() ?? string.Empty);
        if (_visual is TextBlock txt) return VbValue.FromString(txt.Text);
        if (_visual is CheckBox chk) return VbValue.FromLong(chk.IsChecked == true ? 1 : 0);
        if (_visual is RadioButton rad) return VbValue.FromBoolean(rad.IsChecked == true);
        if (_visual is ComboBox cb) return VbValue.FromString(cb.SelectedItem?.ToString() ?? cb.Text);
        if (_visual is ListBox lb) return VbValue.FromString(lb.SelectedItem?.ToString() ?? string.Empty);

        return VbValue.FromString(_designerControl.Text.Length > 0 ? _designerControl.Text : _designerControl.Caption);
    }

    private void SetControlText(string text)
    {
        _designerControl.Text = text;
        _designerControl.Caption = text;

        if (_visual is TextBox tb) tb.Text = text;
        else if (_visual is Button btn) btn.Content = text;
        else if (_visual is TextBlock txt) txt.Text = text;
        else if (_visual is CheckBox chk) chk.Content = text;
        else if (_visual is RadioButton rad) rad.Content = text;
        else if (_visual is GroupBox gb) gb.Header = text;
        else if (_visual is ComboBox cb) cb.Text = text;
    }

    private void SetControlValue(VbValue value)
    {
        if (_visual is CheckBox chk)
        {
            chk.IsChecked = value.AsLong() != 0;
        }
        else if (_visual is RadioButton rad)
        {
            rad.IsChecked = value.AsBoolean();
        }
        else
        {
            SetControlText(value.AsString());
        }
    }

    private void ApplyBackground(long colorRef)
    {
        var brush = VbColorToBrush(colorRef);
        if (_visual is Control ctrl) ctrl.Background = brush;
        else if (_visual is Panel panel) panel.Background = brush;
        else if (_visual is Border border) border.Background = brush;
    }

    private void ApplyForeground(long colorRef)
    {
        var brush = VbColorToBrush(colorRef);
        if (_visual is Control ctrl) ctrl.Foreground = brush;
        else if (_visual is TextBlock tb) tb.Foreground = brush;
    }

    private static Brush VbColorToBrush(long colorRef)
    {
        // Si es color de sistema VB6 (&H800000xx) o RGB normal (0x00BBGGRR)
        if ((colorRef & 0x80000000) != 0)
        {
            int sysIdx = (int)(colorRef & 0xFF);
            return sysIdx switch
            {
                0x0F => SystemColors.ControlBrush,
                0x12 => SystemColors.ControlTextBrush,
                0x05 => SystemColors.WindowBrush,
                0x08 => SystemColors.WindowTextBrush,
                _ => SystemColors.ControlBrush
            };
        }

        byte r = (byte)(colorRef & 0xFF);
        byte g = (byte)((colorRef >> 8) & 0xFF);
        byte b = (byte)((colorRef >> 16) & 0xFF);
        return new SolidColorBrush(Color.FromRgb(r, g, b));
    }
}
