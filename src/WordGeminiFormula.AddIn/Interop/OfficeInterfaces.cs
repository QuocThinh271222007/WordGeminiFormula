using System;
using System.Runtime.InteropServices;

namespace WordGeminiFormula.AddIn.Interop
{
    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ExtConnectMode : uint
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3,
        Solution = 4,
        UISetup = 5
    }

    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ExtDisconnectMode : uint
    {
        HostShutdown = 0,
        UserClosed = 1,
        UISetupComplete = 2,
        SolutionClosed = 3
    }

    // Managed declaration of Extensibility.IDTExtensibility2 that is exported by
    // the CCW. The IID and marshaling match the Office type library, while
    // InterfaceIsIDispatch ensures Office can invoke the callbacks correctly.
    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [TypeLibType((short)4160)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] ExtConnectMode connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ExtDisconnectMode removeMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    // Managed declaration of Microsoft.Office.Core.IRibbonExtensibility exported
    // by the same CCW for Ribbon discovery.
    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [TypeLibType((short)0x1040)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
}
