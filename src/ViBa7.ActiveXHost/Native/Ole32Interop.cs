using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ViBa7.ActiveXHost.Native;

[StructLayout(LayoutKind.Sequential)]
public struct CAUUID
{
    public int cElems;
    public IntPtr pElems;
}

[StructLayout(LayoutKind.Sequential)]
public struct COMRECT
{
    public int left;
    public int top;
    public int right;
    public int bottom;

    public COMRECT(int left, int top, int right, int bottom)
    {
        this.left = left;
        this.top = top;
        this.right = right;
        this.bottom = bottom;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct SIZEL
{
    public int cx;
    public int cy;

    public SIZEL(int cx, int cy)
    {
        this.cx = cx;
        this.cy = cy;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct OLEINPLACEFRAMEINFO
{
    public int cb;
    public int fMDIApp;
    public IntPtr hwndFrame;
    public IntPtr hAccel;
    public int cAccelEntries;
}

[StructLayout(LayoutKind.Sequential)]
public struct POINT
{
    public int x;
    public int y;

    public POINT(int x, int y)
    {
        this.x = x;
        this.y = y;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct POINTF
{
    public float x;
    public float y;

    public POINTF(float x, float y)
    {
        this.x = x;
        this.y = y;
    }
}

[StructLayout(LayoutKind.Sequential)]
public struct CONTROLINFO
{
    public int cb;
    public IntPtr hAccel;
    public ushort cAccel;
    public uint dwFlags;
}

[StructLayout(LayoutKind.Sequential)]
public struct MSG
{
    public IntPtr hwnd;
    public uint message;
    public IntPtr wParam;
    public IntPtr lParam;
    public uint time;
    public POINT pt;
}

[ComImport]
[Guid("00000115-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleInPlaceUIWindow
{
    [PreserveSig]
    int GetWindow(out IntPtr phwnd);

    [PreserveSig]
    int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);

    [PreserveSig]
    int GetBorder(ref COMRECT lprectBorder);

    [PreserveSig]
    int RequestBorderSpace(ref COMRECT pborderwidths);

    [PreserveSig]
    int SetBorderSpace(ref COMRECT pborderwidths);

    [PreserveSig]
    int SetActiveObject(
        [MarshalAs(UnmanagedType.IUnknown)] object? pActiveObj,
        [MarshalAs(UnmanagedType.LPWStr)] string? pszObjName);
}

[ComImport]
[Guid("00000116-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleInPlaceFrame : IOleInPlaceUIWindow
{
    [PreserveSig]
    new int GetWindow(out IntPtr phwnd);

    [PreserveSig]
    new int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);

    [PreserveSig]
    new int GetBorder(ref COMRECT lprectBorder);

    [PreserveSig]
    new int RequestBorderSpace(ref COMRECT pborderwidths);

    [PreserveSig]
    new int SetBorderSpace(ref COMRECT pborderwidths);

    [PreserveSig]
    new int SetActiveObject(
        [MarshalAs(UnmanagedType.IUnknown)] object? pActiveObj,
        [MarshalAs(UnmanagedType.LPWStr)] string? pszObjName);

    [PreserveSig]
    int InsertMenus(IntPtr hmenuShared, IntPtr lpMenuWidths);

    [PreserveSig]
    int SetMenu(IntPtr hmenuShared, IntPtr holemenu, IntPtr hwndActiveObject);

    [PreserveSig]
    int RemoveMenus(IntPtr hmenuShared);

    [PreserveSig]
    int SetStatusText([MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);

    [PreserveSig]
    int EnableModeless([MarshalAs(UnmanagedType.Bool)] bool fEnable);

    [PreserveSig]
    int TranslateAccelerator(ref MSG lpmsg, ushort wID);
}

[ComImport]
[Guid("B196B289-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleControlSite
{
    [PreserveSig]
    int OnControlInfoChanged();

    [PreserveSig]
    int LockInPlaceActive([MarshalAs(UnmanagedType.Bool)] bool fLock);

    [PreserveSig]
    int GetExtendedControl([MarshalAs(UnmanagedType.IDispatch)] out object? ppDisp);

    [PreserveSig]
    int TransformCoords(ref POINT pPtlHimetric, ref POINTF pPtfContainer, uint dwFlags);

    [PreserveSig]
    int TranslateAccelerator(ref MSG pMsg, uint grfModifiers);

    [PreserveSig]
    int OnFocus([MarshalAs(UnmanagedType.Bool)] bool fGotFocus);

    [PreserveSig]
    int ShowPropertyFrame();
}

[ComImport]
[Guid("00000118-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleClientSite
{
    [PreserveSig]
    int SaveObject();

    [PreserveSig]
    int GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk);

    [PreserveSig]
    int GetContainer(out IntPtr ppContainer);

    [PreserveSig]
    int ShowObject();

    [PreserveSig]
    int OnShowWindow([MarshalAs(UnmanagedType.Bool)] bool fShow);

    [PreserveSig]
    int RequestNewObjectLayout();
}

[ComImport]
[Guid("00000119-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleInPlaceSite
{
    [PreserveSig]
    int GetWindow(out IntPtr phwnd);

    [PreserveSig]
    int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);

    [PreserveSig]
    int CanInPlaceActivate();

    [PreserveSig]
    int OnInPlaceActivate();

    [PreserveSig]
    int OnUIActivate();

    [PreserveSig]
    int GetWindowContext(
        out IntPtr ppFrame,
        out IntPtr ppDoc,
        out COMRECT lprcPosRect,
        out COMRECT lprcClipRect,
        ref OLEINPLACEFRAMEINFO lpFrameInfo);

    [PreserveSig]
    int Scroll(SIZEL scrollExtant);

    [PreserveSig]
    int OnUIDeactivate([MarshalAs(UnmanagedType.Bool)] bool fUndoable);

    [PreserveSig]
    int OnInPlaceDeactivate();

    [PreserveSig]
    int DiscardUndoState();

    [PreserveSig]
    int DeactivateAndUndo();

    [PreserveSig]
    int OnPosRectChange(ref COMRECT lprcPosRect);
}

[ComImport]
[Guid("00000112-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleObject
{
    [PreserveSig]
    int SetClientSite([In, MarshalAs(UnmanagedType.Interface)] IOleClientSite? pClientSite);

    [PreserveSig]
    int GetClientSite(out IOleClientSite ppClientSite);

    [PreserveSig]
    int SetHostNames(
        [MarshalAs(UnmanagedType.LPWStr)] string szContainerApp,
        [MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);

    [PreserveSig]
    int Close(int dwSaveOption);

    [PreserveSig]
    int SetMoniker(int dwWhichMoniker, IntPtr pmk);

    [PreserveSig]
    int GetMoniker(int dwAssign, int dwWhichMoniker, out IntPtr ppmk);

    [PreserveSig]
    int InitFromData(IntPtr pDataObject, [MarshalAs(UnmanagedType.Bool)] bool fCreation, int dwReserved);

    [PreserveSig]
    int GetClipboardData(int dwReserved, out IntPtr ppDataObject);

    [PreserveSig]
    int DoVerb(
        int iVerb,
        IntPtr lpmsg,
        [In, MarshalAs(UnmanagedType.Interface)] IOleClientSite? pActiveSite,
        int lindex,
        IntPtr hwndParent,
        ref COMRECT lprcPosRect);

    [PreserveSig]
    int EnumVerbs(out IntPtr ppEnumOleVerb);

    [PreserveSig]
    int Update();

    [PreserveSig]
    int IsUpToDate();

    [PreserveSig]
    int GetUserClassID(out Guid pClsid);

    [PreserveSig]
    int GetUserType(int dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);

    [PreserveSig]
    int SetExtent(int dwDrawAspect, ref SIZEL psizel);

    [PreserveSig]
    int GetExtent(int dwDrawAspect, out SIZEL psizel);

    [PreserveSig]
    int Advise(IntPtr pAdvSink, out int pdwConnection);

    [PreserveSig]
    int Unadvise(int dwConnection);

    [PreserveSig]
    int EnumAdvise(out IntPtr ppenumAdvise);

    [PreserveSig]
    int GetMiscStatus(int dwAspect, out int pdwStatus);

    [PreserveSig]
    int SetColorScheme(IntPtr pLogpal);
}

[ComImport]
[Guid("00000113-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleInPlaceObject
{
    [PreserveSig]
    int GetWindow(out IntPtr phwnd);

    [PreserveSig]
    int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);

    [PreserveSig]
    int InPlaceDeactivate();

    [PreserveSig]
    int UIDeactivate();

    [PreserveSig]
    int SetObjectRects(ref COMRECT lprcPosRect, ref COMRECT lprcClipRect);

    [PreserveSig]
    int ReactivateAndUndo();
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct OLEVERB
{
    public int lVerb;
    [MarshalAs(UnmanagedType.LPWStr)]
    public string? lpszVerbName;
    public int fuFlags;
    public int grfAttribs;
}

[ComImport]
[Guid("00000104-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumOLEVERB
{
    [PreserveSig]
    int Next(uint celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] OLEVERB[] rgelt, out uint pceltFetched);

    [PreserveSig]
    int Skip(uint celt);

    [PreserveSig]
    int Reset();

    [PreserveSig]
    int Clone([MarshalAs(UnmanagedType.Interface)] out IEnumOLEVERB? ppenum);
}

[ComImport]
[Guid("00000117-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleInPlaceActiveObject
{
    [PreserveSig]
    int GetWindow(out IntPtr phwnd);

    [PreserveSig]
    int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);

    [PreserveSig]
    int TranslateAccelerator(ref MSG lpmsg);

    [PreserveSig]
    int OnFrameWindowActivate([MarshalAs(UnmanagedType.Bool)] bool fActivate);

    [PreserveSig]
    int OnDocWindowActivate([MarshalAs(UnmanagedType.Bool)] bool fActivate);

    [PreserveSig]
    int ResizeBorder(
        ref COMRECT prcBorder,
        [In, MarshalAs(UnmanagedType.Interface)] IOleInPlaceUIWindow? pUIWindow,
        [MarshalAs(UnmanagedType.Bool)] bool fFrameWindow);

    [PreserveSig]
    int EnableModeless([MarshalAs(UnmanagedType.Bool)] bool fEnable);
}

[ComImport]
[Guid("742B0E01-14E6-101B-914E-00AA00300CAB")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISimpleFrameSite
{
    [PreserveSig]
    int PreMessageFilter(
        IntPtr hWnd,
        uint msg,
        IntPtr wp,
        IntPtr lp,
        out IntPtr plResult,
        out uint pdwCookie);

    [PreserveSig]
    int PostMessageFilter(
        IntPtr hWnd,
        uint msg,
        IntPtr wp,
        IntPtr lp,
        out IntPtr plResult,
        uint dwCookie);
}

[StructLayout(LayoutKind.Sequential)]
public struct LICINFO
{
    public int cbLicInfo;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fRuntimeKeyAvail;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fLicVerified;
}

[ComImport]
[Guid("00000001-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IClassFactory
{
    [PreserveSig]
    int CreateInstance(
        IntPtr pUnkOuter,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

    [PreserveSig]
    int LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
}

[ComImport]
[Guid("B196B28F-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IClassFactory2 : IClassFactory
{
    [PreserveSig]
    new int CreateInstance(
        IntPtr pUnkOuter,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

    [PreserveSig]
    new int LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);

    [PreserveSig]
    int GetLicInfo(ref LICINFO pLicInfo);

    [PreserveSig]
    int RequestLicKey(
        uint dwReserved,
        [MarshalAs(UnmanagedType.BStr)] out string? pBstrKey);

    [PreserveSig]
    int CreateInstanceLic(
        IntPtr pUnkOuter,
        IntPtr pUnkReserved,
        ref Guid riid,
        [In, MarshalAs(UnmanagedType.BStr)] string bstrKey,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObj);
}

[ComImport]
[Guid("00000109-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPersistStream
{
    [PreserveSig]
    int GetClassID(out Guid pClassID);

    [PreserveSig]
    int IsDirty();

    [PreserveSig]
    int Load(IntPtr pStm);

    [PreserveSig]
    int Save(IntPtr pStm, [MarshalAs(UnmanagedType.Bool)] bool fClearDirty);

    [PreserveSig]
    int GetSizeMax(out long pcbSize);
}

[ComImport]
[Guid("7FD52380-4E07-101B-AE2D-08002B2EC713")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPersistStreamInit
{
    [PreserveSig]
    int GetClassID(out Guid pClassID);

    [PreserveSig]
    int IsDirty();

    [PreserveSig]
    int Load(IntPtr pStm);

    [PreserveSig]
    int Save(IntPtr pStm, [MarshalAs(UnmanagedType.Bool)] bool fClearDirty);

    [PreserveSig]
    int GetSizeMax(out long pcbSize);

    [PreserveSig]
    int InitNew();
}

[ComImport]
[Guid("0000010A-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPersistStorage
{
    [PreserveSig]
    int GetClassID(out Guid pClassID);

    [PreserveSig]
    int IsDirty();

    [PreserveSig]
    int InitNew(IntPtr pStg);

    [PreserveSig]
    int Load(IntPtr pStg);

    [PreserveSig]
    int Save(IntPtr pStgSave, [MarshalAs(UnmanagedType.Bool)] bool fSameAsLoad);

    [PreserveSig]
    int SaveCompleted(IntPtr pStgNew);

    [PreserveSig]
    int HandsOffStorage();
}

[ComImport]
[Guid("0000000c-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IComStream
{
    [PreserveSig]
    int Read(IntPtr pv, uint cb, out uint pcbRead);

    [PreserveSig]
    int Write(IntPtr pv, uint cb, out uint pcbWritten);

    [PreserveSig]
    int Seek(long dlibMove, uint dwOrigin, out ulong plibNewPosition);

    [PreserveSig]
    int SetSize(ulong libNewSize);
}

[ComImport]
[Guid("1CF2B120-547D-101B-8E65-08002B2BD119")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IErrorInfo
{
    [PreserveSig]
    int GetGUID(out Guid pGuid);

    [PreserveSig]
    int GetSource([MarshalAs(UnmanagedType.BStr)] out string? pBstrSource);

    [PreserveSig]
    int GetDescription([MarshalAs(UnmanagedType.BStr)] out string? pBstrDescription);

    [PreserveSig]
    int GetHelpFile([MarshalAs(UnmanagedType.BStr)] out string? pBstrHelpFile);

    [PreserveSig]
    int GetHelpContext(out uint pdwHelpContext);
}

[ComImport]
[Guid("B196B28D-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ISpecifyPropertyPages
{
    [PreserveSig]
    int GetPages(out CAUUID pPages);
}

[ComImport]
[Guid("3127CA40-446E-11CE-8135-00AA004BB851")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IErrorLog
{
    [PreserveSig]
    int AddError(
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
        IntPtr pExcepInfo);
}

[ComImport]
[Guid("55272A00-42CB-11CE-8135-00AA004BB851")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPropertyBag
{
    [PreserveSig]
    int Read(
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
        [In, Out, MarshalAs(UnmanagedType.Struct)] ref object pVar,
        [In, MarshalAs(UnmanagedType.Interface)] IErrorLog? pErrorLog);

    [PreserveSig]
    int Write(
        [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropName,
        [In, MarshalAs(UnmanagedType.Struct)] ref object pVar);
}

[ComImport]
[Guid("37D84F60-42CB-11CE-8135-00AA004BB851")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPersistPropertyBag
{
    [PreserveSig]
    int GetClassID(out Guid pClassID);

    [PreserveSig]
    int InitNew();

    [PreserveSig]
    int Load(
        [In, MarshalAs(UnmanagedType.Interface)] IPropertyBag pPropBag,
        [In, MarshalAs(UnmanagedType.Interface)] IErrorLog? pErrorLog);

    [PreserveSig]
    int Save(
        [In, MarshalAs(UnmanagedType.Interface)] IPropertyBag pPropBag,
        [In, MarshalAs(UnmanagedType.Bool)] bool fClearDirty,
        [In, MarshalAs(UnmanagedType.Bool)] bool fSaveAllProperties);
}

[ComImport]
[Guid("B196B288-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleControl
{
    [PreserveSig]
    int GetControlInfo(ref CONTROLINFO pCI);

    [PreserveSig]
    int OnMnemonic(ref MSG pMsg);

    [PreserveSig]
    int OnAmbientPropertyChange(int dispID);

    [PreserveSig]
    int FreezeEvents([MarshalAs(UnmanagedType.Bool)] bool bFreeze);
}

[ComImport]
[Guid("9BFBBC02-EFF1-101A-84ED-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPropertyNotifySink
{
    [PreserveSig]
    int OnChanged(int dispID);

    [PreserveSig]
    int OnRequestEdit(int dispID);
}

[ComImport]
[Guid("B196B286-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IConnectionPoint
{
    [PreserveSig]
    int GetConnectionInterface(out Guid pIID);

    [PreserveSig]
    int GetConnectionPointContainer([MarshalAs(UnmanagedType.Interface)] out IConnectionPointContainer? ppCPC);

    [PreserveSig]
    int Advise([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSink, out int pdwCookie);

    [PreserveSig]
    int Unadvise(int dwCookie);

    [PreserveSig]
    int EnumConnections(out IntPtr ppEnum);
}

[ComImport]
[Guid("B196B284-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IConnectionPointContainer
{
    [PreserveSig]
    int EnumConnectionPoints(out IntPtr ppEnum);

    [PreserveSig]
    int FindConnectionPoint(ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IConnectionPoint? ppCP);
}

[ComImport]
[Guid("B196B285-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IEnumConnectionPoints
{
    [PreserveSig]
    int Next(uint cConnections, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.Interface)] IConnectionPoint[] rgpcn, out uint pcFetched);

    [PreserveSig]
    int Skip(uint cConnections);

    [PreserveSig]
    int Reset();

    [PreserveSig]
    int Clone(out IEnumConnectionPoints ppEnum);
}

[ComImport]
[Guid("B196B283-BAB4-101A-B69C-00AA00341D07")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IProvideClassInfo
{
    [PreserveSig]
    int GetClassInfo(out IntPtr ppTI);
}

[ComImport]
[Guid("A6BC9AC0-DBAA-11CE-9DE3-00AA004BB851")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IProvideClassInfo2 : IProvideClassInfo
{
    [PreserveSig]
    new int GetClassInfo(out IntPtr ppTI);

    [PreserveSig]
    int GetGUID(int dwGuidKind, out Guid pGUID);
}

[StructLayout(LayoutKind.Sequential)]
public struct DISPPARAMS
{
    public IntPtr rgvarg;
    public IntPtr rgdispidNamedArgs;
    public uint cArgs;
    public uint cNamedArgs;
}

[ComImport]
[Guid("00020400-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDispatchNative
{
    [PreserveSig]
    int GetTypeInfoCount(out uint pctinfo);

    [PreserveSig]
    int GetTypeInfo(uint iTInfo, int lcid, out IntPtr ppTInfo);

    [PreserveSig]
    int GetIDsOfNames(
        ref Guid riid,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames,
        uint cNames,
        int lcid,
        [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

    [PreserveSig]
    int Invoke(
        int dispIdMember,
        ref Guid riid,
        int lcid,
        ushort wFlags,
        ref DISPPARAMS pDispParams,
        IntPtr pVarResult,
        IntPtr pExcepInfo,
        IntPtr puArgErr);
}

[ComImport]
[Guid("0000010F-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAdviseSink
{
    [PreserveSig]
    void OnDataChange(IntPtr pFormatetc, IntPtr pStgmed);

    [PreserveSig]
    void OnViewChange(int dwAspect, int lindex);

    [PreserveSig]
    void OnRename(IntPtr pmk);

    [PreserveSig]
    void OnSave();

    [PreserveSig]
    void OnClose();
}

[ComImport]
[Guid("0000010D-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IViewObject
{
    [PreserveSig]
    int Draw(
        int dwDrawAspect,
        int lindex,
        IntPtr pvAspect,
        IntPtr ptd,
        IntPtr hdcTargetDev,
        IntPtr hdcDraw,
        ref COMRECT lprcBounds,
        IntPtr lprcWBounds,
        IntPtr pfnContinue,
        int dwContinue);

    [PreserveSig]
    int GetColorSet(int dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hicTargetDev, out IntPtr ppColorSet);

    [PreserveSig]
    int Freeze(int dwDrawAspect, int lindex, IntPtr pvAspect, out int pdwFreeze);

    [PreserveSig]
    int Unfreeze(int dwFreeze);

    [PreserveSig]
    int SetAdvise(int aspects, int advf, [In, MarshalAs(UnmanagedType.Interface)] IAdviseSink? pAdvSink);

    [PreserveSig]
    int GetAdvise(out int pAspects, out int pAdvf, [MarshalAs(UnmanagedType.Interface)] out IAdviseSink? ppAdvSink);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct FONTDESC
{
    public int cbSizeofstruct;
    [MarshalAs(UnmanagedType.LPWStr)]
    public string lpstrName;
    public long cySize;
    public short sWeight;
    public short sCharset;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fItalic;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fUnderline;
    [MarshalAs(UnmanagedType.Bool)]
    public bool fStrikethrough;
}

[StructLayout(LayoutKind.Sequential)]
public struct PICTDESC_BMP
{
    public int cbSizeofstruct;
    public int picType;
    public IntPtr hbitmap;
    public IntPtr hpal;
    public int unused;
}

[StructLayout(LayoutKind.Sequential)]
public struct PICTDESC_ICON
{
    public int cbSizeofstruct;
    public int picType;
    public IntPtr hicon;
    public IntPtr unused1;
    public int unused2;
}

public delegate IntPtr SubclassProc(
    IntPtr hWnd,
    uint uMsg,
    IntPtr wParam,
    IntPtr lParam,
    UIntPtr uIdSubclass,
    UIntPtr dwRefData);

public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

public static class Ole32Interop
{
    public const uint WM_CLOSE = 0x0010;
    public const uint WM_ERASEBKGND = 0x0014;
    public const uint WM_SETCURSOR = 0x0020;
    public const uint WM_CONTEXTMENU = 0x007B;
    public const uint WM_KEYDOWN = 0x0100;
    public const uint WM_KEYUP = 0x0101;
    public const uint WM_CHAR = 0x0102;
    public const uint WM_SYSKEYDOWN = 0x0104;
    public const uint WM_SYSKEYUP = 0x0105;
    public const uint WM_SYSCHAR = 0x0106;
    public const uint WM_TIMER = 0x0113;
    public const uint WM_MOUSEMOVE = 0x0200;
    public const uint WM_LBUTTONDOWN = 0x0201;
    public const uint WM_LBUTTONUP = 0x0202;
    public const uint WM_LBUTTONDBLCLK = 0x0203;
    public const uint WM_RBUTTONDOWN = 0x0204;
    public const uint WM_RBUTTONUP = 0x0205;
    public const uint WM_PRINT = 0x0317;

    public const int PRF_CHECKVISIBLE = 0x01;
    public const int PRF_NONCLIENT = 0x02;
    public const int PRF_CLIENT = 0x04;
    public const int PRF_ERASEBKGND = 0x08;
    public const int PRF_CHILDREN = 0x10;

    public const int DVASPECT_CONTENT = 1;
    public const int OLEIVERB_PRIMARY = 0;
    public const int OLEIVERB_SHOW = -1;
    public const int OLEIVERB_OPEN = -2;
    public const int OLEIVERB_HIDE = -3;
    public const int OLEIVERB_UIACTIVATE = -4;
    public const int OLEIVERB_INPLACEACTIVATE = -5;
    public const int OLEIVERB_PROPERTIES = -7;
    public const int OLECLOSE_NOSAVE = 1;
    public const uint CLSCTX_INPROC_SERVER = 0x1;
    public const uint CLSCTX_LOCAL_SERVER = 0x4;

    public const int OLEMISC_INVISIBLEATRUNTIME = 0x00000400;
    public const int OLEMISC_ACTSLIKEBUTTON = 0x00001000;
    public const int OLEMISC_ACTSLIKELABEL = 0x00002000;
    public const int OLEMISC_SIMPLEFRAME = 0x00010000;

    public const uint XFORMCOORDS_POSITION = 0x1;
    public const uint XFORMCOORDS_SIZE = 0x2;
    public const uint XFORMCOORDS_HIMETRICTOCONTAINER = 0x4;
    public const uint XFORMCOORDS_CONTAINERTOHIMETRIC = 0x8;

    public const int PICTYPE_BITMAP = 1;
    public const int PICTYPE_ICON = 3;

    public const int CLASS_E_NOTLICENSED = unchecked((int)0x80040112);

    public const uint GMEM_MOVEABLE = 0x0002;
    public const uint STREAM_SEEK_SET = 0;
    public const uint STREAM_SEEK_CUR = 1;
    public const uint STREAM_SEEK_END = 2;

    public const uint WS_POPUP = 0x80000000;
    public const uint WS_VISIBLE = 0x10000000;
    public const uint WS_CLIPCHILDREN = 0x02000000;
    public const uint WS_CLIPSIBLINGS = 0x04000000;
    public const uint WS_TABSTOP = 0x00010000;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_NOACTIVATE = 0x08000000;

    public static readonly Guid IID_IUnknown = new("00000000-0000-0000-C000-000000000046");
    public static readonly Guid IID_IDispatch = new("00020400-0000-0000-C000-000000000046");
    public static readonly Guid IID_IClassFactory = new("00000001-0000-0000-C000-000000000046");
    public static readonly Guid IID_IClassFactory2 = new("B196B28F-BAB4-101A-B69C-00AA00341D07");
    public static readonly Guid IID_IPropertyNotifySink = new("9BFBBC02-EFF1-101A-84ED-00AA00341D07");
    public static readonly Guid IID_IFontDisp = new("BEF6E003-A874-101A-8BBA-00AA00300CAB");
    public static readonly Guid IID_IPictureDisp = new("7BF80981-BF32-101A-8BBB-00AA00300CAB");

    public const int GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1;

    public const int DISPID_CLICK = -600;
    public const int DISPID_DBLCLICK = -601;
    public const int DISPID_KEYDOWN = -602;
    public const int DISPID_KEYPRESS = -603;
    public const int DISPID_KEYUP = -604;
    public const int DISPID_MOUSEDOWN = -605;
    public const int DISPID_MOUSEMOVE = -606;
    public const int DISPID_MOUSEUP = -607;
    public const int DISPID_ERROREVENT = -608;

    public const int WH_CBT = 5;
    public const int HCBT_CREATEWND = 3;
    public const int HCBT_ACTIVATE = 5;

    private const uint _MCW_EM = 0x0008001F;

    [DllImport("msvcrt.dll", CallingConvention = CallingConvention.Cdecl)]
    private static extern uint _control87(uint newCtrl, uint mask);

    /// <summary>
    /// Enmascara todas las excepciones de punto flotante x87/SSE (división por cero, overflow, underflow)
    /// para evitar que controles OCX nativos de 32 bits generen fallos de hardware en modo diseño.
    /// </summary>
    public static void MaskFpuExceptions()
    {
        try
        {
            _control87(_MCW_EM, _MCW_EM);
        }
        catch
        {
            // Ignorar si no aplica
        }
    }

    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int CoCreateInstance(
        ref Guid rclsid,
        IntPtr pUnkOuter,
        uint dwClsContext,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppv);

    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int CoGetClassObject(
        ref Guid rclsid,
        uint dwClsContext,
        IntPtr pvReserved,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppv);

    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int CreateStreamOnHGlobal(
        IntPtr hGlobal,
        [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease,
        out IntPtr ppstm);

    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int GetHGlobalFromStream(
        IntPtr pstm,
        out IntPtr phglobal);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GlobalUnlock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern UIntPtr GlobalSize(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalFree(IntPtr hMem);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    public static extern int GetErrorInfo(
        uint dwReserved,
        [MarshalAs(UnmanagedType.Interface)] out IErrorInfo? pperrinfo);

    /// <summary>
    /// Crea un IStream nativo de COM respaldado por memoria global (HGlobal) inicializado con los bytes dados.
    /// </summary>
    public static IntPtr CreateComStreamFromBytes(byte[]? data)
    {
        IntPtr hGlobal = IntPtr.Zero;
        if (data != null && data.Length > 0)
        {
            hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)data.Length);
            if (hGlobal == IntPtr.Zero) return IntPtr.Zero;

            IntPtr ptr = GlobalLock(hGlobal);
            if (ptr == IntPtr.Zero)
            {
                GlobalFree(hGlobal);
                return IntPtr.Zero;
            }

            try
            {
                Marshal.Copy(data, 0, ptr, data.Length);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }
        }

        int hr = CreateStreamOnHGlobal(hGlobal, fDeleteOnRelease: true, out IntPtr pStream);
        if (hr < 0 || pStream == IntPtr.Zero)
        {
            if (hGlobal != IntPtr.Zero)
            {
                GlobalFree(hGlobal);
            }
            return IntPtr.Zero;
        }

        if (data != null && data.Length > 0)
        {
            try
            {
                if (Marshal.GetObjectForIUnknown(pStream) is IComStream comStm)
                {
                    comStm.SetSize((ulong)data.Length);
                    comStm.Seek(0, STREAM_SEEK_SET, out _);
                }
            }
            catch { }
        }

        return pStream;
    }

    /// <summary>
    /// Extrae el arreglo de bytes escrito en un IStream nativo creado con CreateStreamOnHGlobal.
    /// </summary>
    public static byte[]? ReadBytesFromComStream(IntPtr pStream)
    {
        if (pStream == IntPtr.Zero) return null;

        int length = -1;
        try
        {
            if (Marshal.GetObjectForIUnknown(pStream) is IComStream comStm)
            {
                if (comStm.Seek(0, STREAM_SEEK_CUR, out ulong pos) == 0 && pos > 0 && pos <= int.MaxValue)
                {
                    length = (int)pos;
                }
                else if (comStm.Seek(0, STREAM_SEEK_END, out ulong endPos) == 0 && endPos <= int.MaxValue)
                {
                    length = (int)endPos;
                }
            }
        }
        catch { }

        if (GetHGlobalFromStream(pStream, out IntPtr hGlobal) < 0 || hGlobal == IntPtr.Zero)
        {
            return null;
        }

        int globalLen = (int)GlobalSize(hGlobal);
        if (length < 0 || length > globalLen)
        {
            length = globalLen;
        }

        if (length <= 0)
        {
            return Array.Empty<byte>();
        }

        IntPtr ptr = GlobalLock(hGlobal);
        if (ptr == IntPtr.Zero) return null;

        try
        {
            byte[] result = new byte[length];
            Marshal.Copy(ptr, result, 0, length);
            return result;
        }
        finally
        {
            GlobalUnlock(hGlobal);
        }
    }

    /// <summary>
    /// Obtiene la descripción textual del último error COM (IErrorInfo) en el hilo actual si existe.
    /// </summary>
    public static bool TryGetComErrorDescription(out string? description, out string? source)
    {
        description = null;
        source = null;
        try
        {
            if (GetErrorInfo(0, out var errInfo) == 0 && errInfo != null)
            {
                errInfo.GetDescription(out description);
                errInfo.GetSource(out source);
                Marshal.ReleaseComObject(errInfo);
                return !string.IsNullOrWhiteSpace(description) || !string.IsNullOrWhiteSpace(source);
            }
        }
        catch { }
        return false;
    }

    [DllImport("ole32.dll", PreserveSig = true)]
    public static extern int OleDraw(
        [MarshalAs(UnmanagedType.IUnknown)] object pUnk,
        int dwAspect,
        IntPtr hdcDraw,
        ref COMRECT lprcBounds);

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    public static extern int OleCreatePropertyFrame(
        IntPtr hwndOwner,
        int x,
        int y,
        string lpszCaption,
        int cObjects,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.IUnknown)]
        object[] lplpUnk,
        int cPages,
        IntPtr lpPageClsID,
        int lcid,
        int dwReserved,
        IntPtr lpvReserved);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    public static extern int OleCreateFontIndirect(
        ref FONTDESC lpFontDesc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.IDispatch)] out object? lplpvObj);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    public static extern int OleCreatePictureIndirect(
        ref PICTDESC_BMP lpPictDesc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Bool)] bool fOwn,
        [MarshalAs(UnmanagedType.IDispatch)] out object? lplpvObj);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    public static extern int OleCreatePictureIndirect(
        ref PICTDESC_ICON lpPictDesc,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Bool)] bool fOwn,
        [MarshalAs(UnmanagedType.IDispatch)] out object? lplpvObj);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    public static extern int OleTranslateColor(
        uint clr,
        IntPtr hpal,
        out uint lpcolorref);

    /// <summary>
    /// Crea un objeto COM StdFont (IFontDisp) nativo mediante OleCreateFontIndirect.
    /// </summary>
    public static object? CreateStdFont(
        string fontName = "MS Sans Serif",
        float sizePoints = 8.25f,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        bool strikethrough = false,
        short charset = 0)
    {
        try
        {
            var fd = new FONTDESC
            {
                cbSizeofstruct = Marshal.SizeOf<FONTDESC>(),
                lpstrName = string.IsNullOrWhiteSpace(fontName) ? "MS Sans Serif" : fontName,
                // En FONTDESC, cySize es un entero CURRENCY escalado por 10000 (ej. 8.25pt = 82500)
                cySize = (long)Math.Round(Math.Max(1.0f, sizePoints) * 10000.0f),
                sWeight = (short)(bold ? 700 : 400),
                sCharset = charset,
                fItalic = italic,
                fUnderline = underline,
                fStrikethrough = strikethrough
            };

            Guid iidFontDisp = IID_IFontDisp;
            int hr = OleCreateFontIndirect(ref fd, ref iidFontDisp, out object? fontObj);
            return hr >= 0 ? fontObj : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Crea un objeto COM StdPicture (IPictureDisp) nativo a partir de los bytes de una imagen o icono.
    /// </summary>
    public static object? CreateStdPictureFromBytes(byte[]? imageBytes)
    {
        if (imageBytes == null || imageBytes.Length < 4) return null;

        try
        {
            using var ms = new System.IO.MemoryStream(imageBytes);
            // Detectar si es un archivo ICO (00 00 01 00)
            if (imageBytes[0] == 0 && imageBytes[1] == 0 && imageBytes[2] == 1 && imageBytes[3] == 0)
            {
                using var ico = new System.Drawing.Icon(ms);
                IntPtr hIcon = ico.ToBitmap().GetHicon();
                var pdIcon = new PICTDESC_ICON
                {
                    cbSizeofstruct = Marshal.SizeOf<PICTDESC_ICON>(),
                    picType = PICTYPE_ICON,
                    hicon = hIcon
                };
                Guid iidPict = IID_IPictureDisp;
                int hr = OleCreatePictureIndirect(ref pdIcon, ref iidPict, fOwn: true, out object? pictIcon);
                return hr >= 0 ? pictIcon : null;
            }
            else
            {
                using var bmp = new System.Drawing.Bitmap(ms);
                IntPtr hBmp = bmp.GetHbitmap();
                var pdBmp = new PICTDESC_BMP
                {
                    cbSizeofstruct = Marshal.SizeOf<PICTDESC_BMP>(),
                    picType = PICTYPE_BITMAP,
                    hbitmap = hBmp,
                    hpal = IntPtr.Zero
                };
                Guid iidPict = IID_IPictureDisp;
                int hr = OleCreatePictureIndirect(ref pdBmp, ref iidPict, fOwn: true, out object? pictBmp);
                if (hr < 0)
                {
                    DeleteObject(hBmp);
                    return null;
                }
                return pictBmp;
            }
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Convierte un valor OLE_COLOR (ej. 0x8000000F o 0x00BBGGRR) a System.Drawing.Color usando OleTranslateColor.
    /// </summary>
    public static System.Drawing.Color TranslateOleColor(uint oleColor)
    {
        if (OleTranslateColor(oleColor, IntPtr.Zero, out uint colorRef) >= 0)
        {
            int r = (int)(colorRef & 0xFF);
            int g = (int)((colorRef >> 8) & 0xFF);
            int b = (int)((colorRef >> 16) & 0xFF);
            return System.Drawing.Color.FromArgb(255, r, g, b);
        }
        return System.Drawing.Color.FromArgb(255, 240, 240, 240);
    }

    [DllImport("ole32.dll")]
    public static extern void CoTaskMemFree(IntPtr pv);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteObject(IntPtr hObject);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateWindowExW(
        uint dwExStyle,
        string lpClassName,
        string lpWindowName,
        uint dwStyle,
        int x,
        int y,
        int nWidth,
        int nHeight,
        IntPtr hWndParent,
        IntPtr hMenu,
        IntPtr hInstance,
        IntPtr lpParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public const uint SMTO_NORMAL = 0x0000;
    public const uint SMTO_BLOCK = 0x0001;
    public const uint SMTO_ABORTIFHUNG = 0x0002;
    public const uint SMTO_NOTIMEOUTIFNOTHUNG = 0x0008;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeoutW(
        IntPtr hWnd,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam,
        uint fuFlags,
        uint uTimeout,
        out IntPtr lpdwResult);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);

    [DllImport("comctl32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowSubclass(
        IntPtr hWnd,
        SubclassProc pfnSubclass,
        UIntPtr uIdSubclass,
        UIntPtr dwRefData);

    [DllImport("comctl32.dll")]
    public static extern IntPtr DefSubclassProc(
        IntPtr hWnd,
        uint uMsg,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport("comctl32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RemoveWindowSubclass(
        IntPtr hWnd,
        SubclassProc pfnSubclass,
        UIntPtr uIdSubclass);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(
        IntPtr hWndParent,
        EnumWindowsProc lpEnumFunc,
        IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr SetCursor(IntPtr hCursor);

    [DllImport("user32.dll")]
    public static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SetWindowsHookExW(
        int idHook,
        HookProc lpfn,
        IntPtr hMod,
        uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    public const int GWL_STYLE = -16;
    public const uint WS_CHILD = 0x40000000;
    public const uint WM_COMMAND = 0x0111;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    public static extern int GetWindowLongW(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EndDialog(IntPtr hDlg, IntPtr nResult);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, [MarshalAs(UnmanagedType.Bool)] bool bRepaint);

    [DllImport("user32.dll")]
    public static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    public const int SW_SHOW = 5;
    public const int IDC_ARROW = 32512;

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetGuiResources(IntPtr hProcess, uint uiFlags);

    public const uint GR_GDIOBJECTS = 0;
    public const uint GR_USEROBJECTS = 1;
    public const uint GR_GDIOBJECTS_PEAK = 2;
    public const uint GR_USEROBJECTS_PEAK = 4;
}

/// <summary>
/// Guardián de hilo local (WH_CBT) que intercepta y suprime cualquier cuadro de diálogo de error modal (#32770)
/// disparado por msvbvm60.dll o controles OCX de terceros durante la instanciación o redimensionamiento en modo diseño
/// (por ejemplo, "Run-time error '11': Division by zero" o alertas de licencia en controles no inicializados).
/// Nunca aborta la creación de ventanas hijas de clase #32770 usadas por controles compuestos de MFC/Delphi.
/// </summary>
public sealed class NativeDialogGuard : IDisposable
{
    private IntPtr _hookHandle;
    private readonly HookProc _hookProc;

    public bool DialogSuppressed { get; private set; }

    public NativeDialogGuard()
    {
        Ole32Interop.MaskFpuExceptions();
        _hookProc = CbtHookCallback;
        uint threadId = Ole32Interop.GetCurrentThreadId();
        _hookHandle = Ole32Interop.SetWindowsHookExW(Ole32Interop.WH_CBT, _hookProc, IntPtr.Zero, threadId);
    }

    private IntPtr CbtHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode == Ole32Interop.HCBT_ACTIVATE && wParam != IntPtr.Zero)
        {
            var sb = new StringBuilder(64);
            int len = Ole32Interop.GetClassNameW(wParam, sb, sb.Capacity);
            if (len > 0 && sb.ToString() == "#32770")
            {
                int style = Ole32Interop.GetWindowLongW(wParam, Ole32Interop.GWL_STYLE);
                // Solo suprimir si es un cuadro de diálogo modal superior (no una ventana hija integrada en el control)
                if ((style & (int)Ole32Interop.WS_CHILD) == 0)
                {
                    DialogSuppressed = true;
                    Ole32Interop.EndDialog(wParam, (IntPtr)1);
                    Ole32Interop.PostMessageW(wParam, Ole32Interop.WM_COMMAND, (IntPtr)1, IntPtr.Zero);
                    Ole32Interop.PostMessageW(wParam, Ole32Interop.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    return (IntPtr)1;
                }
            }
        }

        return Ole32Interop.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    public void Dispose()
    {
        if (_hookHandle != IntPtr.Zero)
        {
            Ole32Interop.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
        }
        Ole32Interop.MaskFpuExceptions();
    }
}
