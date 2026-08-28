using System;
using System.Runtime.InteropServices;

namespace WordGeminiFormula.AddIn.Interop
{
    public enum ExtConnectMode
    {
        AfterStartup = 0,
        Startup = 1,
        External = 2,
        CommandLine = 3,
        Solution = 4,
        UISetup = 5
    }

    public enum ExtDisconnectMode
    {
        HostShutdown = 0,
        UserClosed = 1,
        UISetupComplete = 2,
        SolutionClosed = 3
    }

    [ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] ExtConnectMode connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            [In] ref Array custom);

        [DispId(2)]
        void OnDisconnection([In] ExtDisconnectMode removeMode, [In] ref Array custom);
        [DispId(3)]
        void OnAddInsUpdate([In] ref Array custom);
        [DispId(4)]
        void OnStartupComplete([In] ref Array custom);
        [DispId(5)]
        void OnBeginShutdown([In] ref Array custom);
    }

    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
}
