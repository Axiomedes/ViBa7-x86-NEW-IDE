using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ViBa7.ActiveXHost.Diagnostics;
using ViBa7.ActiveXHost.Native;
using ViBa7.ProjectModel.ActiveX;
using ViBa7.ProjectModel.Execution;

namespace ViBa7.ActiveXHost;

/// <summary>
/// Representa una invocación de evento COM disparada por un control ActiveX hacia el entorno anfitrión.
/// </summary>
public record ActiveXEventInvocation(
    int DispId,
    string EventName,
    VbValue[] Arguments,
    bool[] IsByRef,
    Action<int, VbValue> UpdateArgumentValue
);

/// <summary>
/// Sumidero de eventos COM (COM Event Sink) dinámico que implementa la interfaz de salida (Source Interface)
/// de un control ActiveX mediante IConnectionPointContainer e IDispatch::Invoke.
/// Desempaqueta argumentos DISPPARAMS respetando el orden inverso de la ABI COM de 32 bits,
/// propaga parámetros ByRef de vuelta al control nativo e inyecta Index As Integer para matrices de controles.
/// </summary>
[ComVisible(true)]
public sealed class ActiveXEventSink : ICustomQueryInterface, IDispatchNative, IDisposable
{
    private readonly Guid _eventIid;
    private readonly string _controlName;
    private readonly string _typeName;
    private readonly int? _controlIndex;
    private readonly Dictionary<int, string> _eventNameByDispId = new();
    private bool _disposed;

    public Guid EventIid => _eventIid;
    public string ControlName => _controlName;
    public string TypeName => _typeName;
    public int? ControlIndex => _controlIndex;

    /// <summary>
    /// Evento disparado cada vez que el control ActiveX emite una notificación a través de su interfaz Source.
    /// </summary>
    public event Action<ActiveXEventInvocation>? EventTriggered;

    public ActiveXEventSink(
        Guid eventIid,
        string controlName,
        string typeName,
        int? controlIndex = null,
        IEnumerable<OcxEventInfo>? knownEvents = null)
    {
        _eventIid = eventIid;
        _controlName = !string.IsNullOrWhiteSpace(controlName) ? controlName : "ActiveXControl1";
        _typeName = typeName ?? string.Empty;
        _controlIndex = controlIndex;

        // Registrar DISPIDs estándar de OLE / ActiveX
        _eventNameByDispId[Ole32Interop.DISPID_CLICK] = "Click";
        _eventNameByDispId[Ole32Interop.DISPID_DBLCLICK] = "DblClick";
        _eventNameByDispId[Ole32Interop.DISPID_KEYDOWN] = "KeyDown";
        _eventNameByDispId[Ole32Interop.DISPID_KEYPRESS] = "KeyPress";
        _eventNameByDispId[Ole32Interop.DISPID_KEYUP] = "KeyUp";
        _eventNameByDispId[Ole32Interop.DISPID_MOUSEDOWN] = "MouseDown";
        _eventNameByDispId[Ole32Interop.DISPID_MOUSEMOVE] = "MouseMove";
        _eventNameByDispId[Ole32Interop.DISPID_MOUSEUP] = "MouseUp";
        _eventNameByDispId[Ole32Interop.DISPID_ERROREVENT] = "Error";

        // Registrar eventos conocidos extraídos de la TypeLib del control
        if (knownEvents != null)
        {
            foreach (var ev in knownEvents)
            {
                if (!string.IsNullOrWhiteSpace(ev.Name))
                {
                    _eventNameByDispId[ev.DispId] = ev.Name;
                }
            }
        }
    }

    /// <summary>
    /// Permite responder afirmativamente al QueryInterface para la GUID exacta de la interfaz Source del OCX.
    /// </summary>
    public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
    {
        ppv = IntPtr.Zero;

        if (_disposed)
        {
            return CustomQueryInterfaceResult.Failed;
        }

        if (iid == _eventIid ||
            iid == Ole32Interop.IID_IDispatch ||
            iid == Ole32Interop.IID_IUnknown)
        {
            ppv = Marshal.GetComInterfaceForObject(this, typeof(IDispatchNative), CustomQueryInterfaceMode.Ignore);
            return CustomQueryInterfaceResult.Handled;
        }

        return CustomQueryInterfaceResult.NotHandled;
    }

    #region IDispatchNative Implementation

    public int GetTypeInfoCount(out uint pctinfo)
    {
        pctinfo = 0;
        return 0; // S_OK
    }

    public int GetTypeInfo(uint iTInfo, int lcid, out IntPtr ppTInfo)
    {
        ppTInfo = IntPtr.Zero;
        return unchecked((int)0x80004001); // E_NOTIMPL
    }

    public int GetIDsOfNames(
        ref Guid riid,
        string[] rgszNames,
        uint cNames,
        int lcid,
        int[] rgDispId)
    {
        return unchecked((int)0x80004001); // E_NOTIMPL
    }

    public int Invoke(
        int dispIdMember,
        ref Guid riid,
        int lcid,
        ushort wFlags,
        ref DISPPARAMS pDispParams,
        IntPtr pVarResult,
        IntPtr pExcepInfo,
        IntPtr puArgErr)
    {
        if (_disposed) return 0; // S_OK

        try
        {
            string eventName = ResolveEventName(dispIdMember);
            int argCount = (int)pDispParams.cArgs;

            var rawArgs = new object?[argCount];
            var isByRef = new bool[argCount];
            var targetPointers = new IntPtr[argCount];
            var varTypes = new ushort[argCount];

            // ═══════════════════════════════════════════════════════════════════
            // REGLA CRÍTICA COM/OLE: Los argumentos en DISPPARAMS.rgvarg vienen
            // en ORDEN INVERSO (el último argumento de la firma está en índice 0).
            // ═══════════════════════════════════════════════════════════════════
            if (pDispParams.rgvarg != IntPtr.Zero && argCount > 0)
            {
                for (int i = 0; i < argCount; i++)
                {
                    int revIndex = argCount - 1 - i;
                    IntPtr pVariant = pDispParams.rgvarg + (revIndex * 16); // 16 bytes por VARIANT en x86

                    ushort vt = (ushort)Marshal.ReadInt16(pVariant, 0);
                    varTypes[i] = vt;
                    bool byRef = (vt & (ushort)VarEnum.VT_BYREF) != 0;
                    isByRef[i] = byRef;

                    if (byRef)
                    {
                        targetPointers[i] = Marshal.ReadIntPtr(pVariant, 8);
                    }

                    try
                    {
                        rawArgs[i] = Marshal.GetObjectForNativeVariant(pVariant);
                    }
                    catch
                    {
                        rawArgs[i] = null;
                    }
                }
            }

            // Convertir argumentos COM a VbValue
            var vbArgs = new VbValue[argCount];
            for (int i = 0; i < argCount; i++)
            {
                vbArgs[i] = ConvertComValueToVbValue(rawArgs[i]);
            }

            // Delegado para propagar modificaciones ByRef de vuelta al control nativo
            void UpdateArgument(int paramIndex, VbValue newValue)
            {
                if (paramIndex < 0 || paramIndex >= argCount || !isByRef[paramIndex]) return;
                IntPtr pTarget = targetPointers[paramIndex];
                if (pTarget == IntPtr.Zero) return;

                ushort baseVt = (ushort)(varTypes[paramIndex] & ~(ushort)VarEnum.VT_BYREF);
                WriteBackByRefValue(pTarget, baseVt, newValue);
            }

            // Si el control pertenece a un arreglo de controles (Control Array), inyectar Index As Integer
            VbValue[] dispatchArgs;
            bool[] dispatchByRef;
            if (_controlIndex.HasValue)
            {
                dispatchArgs = new VbValue[argCount + 1];
                dispatchByRef = new bool[argCount + 1];
                dispatchArgs[0] = VbValue.FromLong(_controlIndex.Value);
                dispatchByRef[0] = false;
                Array.Copy(vbArgs, 0, dispatchArgs, 1, argCount);
                Array.Copy(isByRef, 0, dispatchByRef, 1, argCount);
            }
            else
            {
                dispatchArgs = vbArgs;
                dispatchByRef = isByRef;
            }

            var invocation = new ActiveXEventInvocation(
                dispIdMember,
                eventName,
                dispatchArgs,
                dispatchByRef,
                (idx, val) =>
                {
                    int realParamIdx = _controlIndex.HasValue ? idx - 1 : idx;
                    UpdateArgument(realParamIdx, val);
                });

            EventTriggered?.Invoke(invocation);
            ActiveXDiagnosticsLogger.LogSuccess("EventDispatch", _eventIid.ToString("B"), _controlName,
                $"Evento '{eventName}' (DISPID {dispIdMember}) despachado con {dispatchArgs.Length} argumentos.");
            return 0; // S_OK
        }
        catch (Exception ex)
        {
            ActiveXDiagnosticsLogger.LogError("EventInvokeException", _eventIid.ToString("B"), ex.HResult, _controlName, ex.Message);
            return 0; // Prevenir caídas no controladas en el hilo del control nativo
        }
    }

    #endregion

    private string ResolveEventName(int dispId)
    {
        if (_eventNameByDispId.TryGetValue(dispId, out string? name) && !string.IsNullOrWhiteSpace(name))
        {
            return name;
        }

        // Buscar en catálogo de TypeLib si se conoce el tipo de control
        if (!string.IsNullOrWhiteSpace(_typeName) &&
            ViBa7.ProjectModel.IntelliSense.VbControlMemberCatalog.TryResolveEventByDispId(_typeName, dispId, out var symbol) &&
            symbol != null && !string.IsNullOrWhiteSpace(symbol.Name))
        {
            _eventNameByDispId[dispId] = symbol.Name;
            return symbol.Name;
        }

        return dispId switch
        {
            Ole32Interop.DISPID_CLICK => "Click",
            Ole32Interop.DISPID_DBLCLICK => "DblClick",
            Ole32Interop.DISPID_KEYDOWN => "KeyDown",
            Ole32Interop.DISPID_KEYPRESS => "KeyPress",
            Ole32Interop.DISPID_KEYUP => "KeyUp",
            Ole32Interop.DISPID_MOUSEDOWN => "MouseDown",
            Ole32Interop.DISPID_MOUSEMOVE => "MouseMove",
            Ole32Interop.DISPID_MOUSEUP => "MouseUp",
            Ole32Interop.DISPID_ERROREVENT => "Error",
            _ => $"Event_{dispId}"
        };
    }

    private static VbValue ConvertComValueToVbValue(object? obj)
    {
        if (obj == null || obj is DBNull) return VbValue.Empty;
        if (obj is bool b) return VbValue.FromBoolean(b);
        if (obj is short s) return VbValue.FromLong(s);
        if (obj is ushort us) return VbValue.FromLong(us);
        if (obj is int i) return VbValue.FromLong(i);
        if (obj is uint ui) return VbValue.FromLong(unchecked((int)ui));
        if (obj is long l) return VbValue.FromLong(unchecked((int)l));
        if (obj is float f) return VbValue.FromDouble(f);
        if (obj is double d) return VbValue.FromDouble(d);
        if (obj is decimal m) return VbValue.FromDouble((double)m);
        if (obj is string str) return VbValue.FromString(str);
        return VbValue.FromObject(obj);
    }

    private static void WriteBackByRefValue(IntPtr pTarget, ushort baseVt, VbValue val)
    {
        try
        {
            switch ((VarEnum)baseVt)
            {
                case VarEnum.VT_I1:
                case VarEnum.VT_UI1:
                    Marshal.WriteByte(pTarget, (byte)val.AsLong());
                    break;

                case VarEnum.VT_I2:
                    Marshal.WriteInt16(pTarget, (short)val.AsLong());
                    break;

                case VarEnum.VT_UI2:
                    Marshal.WriteInt16(pTarget, unchecked((short)(ushort)val.AsLong()));
                    break;

                case VarEnum.VT_I4:
                case VarEnum.VT_INT:
                    Marshal.WriteInt32(pTarget, (int)val.AsLong());
                    break;

                case VarEnum.VT_UI4:
                case VarEnum.VT_UINT:
                    Marshal.WriteInt32(pTarget, unchecked((int)(uint)val.AsLong()));
                    break;

                case VarEnum.VT_BOOL:
                    // VARIANT_BOOL en COM: -1 para True (VARIANT_TRUE), 0 para False (VARIANT_FALSE)
                    Marshal.WriteInt16(pTarget, (short)(val.AsBoolean() ? -1 : 0));
                    break;

                case VarEnum.VT_R4:
                    byte[] r4Bytes = BitConverter.GetBytes((float)val.AsDouble());
                    Marshal.Copy(r4Bytes, 0, pTarget, 4);
                    break;

                case VarEnum.VT_R8:
                    byte[] r8Bytes = BitConverter.GetBytes(val.AsDouble());
                    Marshal.Copy(r8Bytes, 0, pTarget, 8);
                    break;

                case VarEnum.VT_BSTR:
                    IntPtr oldBstr = Marshal.ReadIntPtr(pTarget);
                    if (oldBstr != IntPtr.Zero)
                    {
                        Marshal.FreeBSTR(oldBstr);
                    }
                    Marshal.WriteIntPtr(pTarget, Marshal.StringToBSTR(val.AsString()));
                    break;

                case VarEnum.VT_VARIANT:
                    IntPtr tempVar = Marshal.AllocHGlobal(16);
                    try
                    {
                        Marshal.GetNativeVariantForObject(val.RawValue, tempVar);
                        byte[] varData = new byte[16];
                        Marshal.Copy(tempVar, varData, 0, 16);
                        Marshal.Copy(varData, 0, pTarget, 16);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(tempVar);
                    }
                    break;
            }
        }
        catch
        {
            // Ignorar fallos de escritura ByRef si el puntero ya no es accesible
        }
    }

    /// <summary>
    /// Intenta conectar el sumidero de eventos a un control ActiveX que implemente IConnectionPointContainer.
    /// Resuelve la interfaz Source a través del GUID preferido, IProvideClassInfo2 o enumeración de puntos de conexión.
    /// </summary>
    public static bool TryConnect(
        object ocx,
        string controlName,
        string typeName,
        int? controlIndex,
        Guid? preferredEventIid,
        IEnumerable<OcxEventInfo>? knownEvents,
        out ActiveXEventSink? sink,
        out IConnectionPoint? connectionPoint,
        out int cookie)
    {
        sink = null;
        connectionPoint = null;
        cookie = 0;

        if (ocx is not IConnectionPointContainer cpc)
        {
            return false;
        }

        Guid targetIid = Guid.Empty;

        // 1. Probar con GUID preferido de la TypeLib
        if (preferredEventIid.HasValue && preferredEventIid.Value != Guid.Empty)
        {
            Guid tryIid = preferredEventIid.Value;
            if (cpc.FindConnectionPoint(ref tryIid, out var cpPreferred) >= 0 && cpPreferred != null)
            {
                targetIid = tryIid;
                connectionPoint = cpPreferred;
            }
        }

        // 2. Probar con IProvideClassInfo2 (GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1)
        if (connectionPoint == null && ocx is IProvideClassInfo2 pci2)
        {
            try
            {
                if (pci2.GetGUID(Ole32Interop.GUIDKIND_DEFAULT_SOURCE_DISP_IID, out Guid defaultSourceIid) >= 0 &&
                    defaultSourceIid != Guid.Empty)
                {
                    if (cpc.FindConnectionPoint(ref defaultSourceIid, out var cpDefault) >= 0 && cpDefault != null)
                    {
                        targetIid = defaultSourceIid;
                        connectionPoint = cpDefault;
                    }
                }
            }
            catch { }
        }

        // 3. Probar enumerando puntos de conexión si aún no se resolvió
        if (connectionPoint == null)
        {
            try
            {
                if (cpc.EnumConnectionPoints(out IntPtr ppEnum) >= 0 && ppEnum != IntPtr.Zero)
                {
                    try
                    {
                        if (Marshal.GetObjectForIUnknown(ppEnum) is IEnumConnectionPoints enumCp)
                        {
                            var cpBuffer = new IConnectionPoint[1];
                            while (enumCp.Next(1, cpBuffer, out uint fetched) == 0 && fetched == 1)
                            {
                                var cpCandidate = cpBuffer[0];
                                if (cpCandidate != null)
                                {
                                    if (cpCandidate.GetConnectionInterface(out Guid cpIid) >= 0 &&
                                        cpIid != Ole32Interop.IID_IPropertyNotifySink &&
                                        cpIid != Guid.Empty)
                                    {
                                        targetIid = cpIid;
                                        connectionPoint = cpCandidate;
                                        break;
                                    }
                                    Marshal.ReleaseComObject(cpCandidate);
                                }
                            }
                        }
                    }
                    finally
                    {
                        Marshal.Release(ppEnum);
                    }
                }
            }
            catch { }
        }

        if (connectionPoint == null || targetIid == Guid.Empty)
        {
            return false;
        }

        sink = new ActiveXEventSink(targetIid, controlName, typeName, controlIndex, knownEvents);

        try
        {
            int hrAdvise = connectionPoint.Advise(sink, out int newCookie);
            if (hrAdvise >= 0 && newCookie != 0)
            {
                cookie = newCookie;
                return true;
            }
        }
        catch { }

        sink.Dispose();
        sink = null;
        if (Marshal.IsComObject(connectionPoint))
        {
            Marshal.ReleaseComObject(connectionPoint);
        }
        connectionPoint = null;
        return false;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        EventTriggered = null;
    }
}
