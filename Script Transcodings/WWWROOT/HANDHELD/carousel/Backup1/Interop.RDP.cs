using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {

  public static class RDP {

    public static string GetRDPClientName() {
      const int WTS_CURRENT_SERVER_HANDLE = -1;
      IntPtr buffer = IntPtr.Zero;
      uint bytesReturned;
      string retval = string.Empty;
      try {
        WTSQuerySessionInformation(IntPtr.Zero, WTS_CURRENT_SERVER_HANDLE, WTS_INFO_CLASS.WTSClientName, out buffer, out bytesReturned);
        retval = Marshal.PtrToStringAnsi(buffer);
      }
      finally {
        WTSFreeMemory(buffer);
        buffer = IntPtr.Zero;
      }
      return retval.ToUpper();
    }

    public static bool HookWindowsMessageQueue(IntPtr handle) {
      return WTSRegisterSessionNotification(handle, (int)WTSSessionNotification.NotifyForThisSession);
    }

    public static bool UnhookWindowsMessageQueue(IntPtr handle) {
      return WTSUnRegisterSessionNotification(handle);
    }

    private enum WTS_INFO_CLASS {
      WTSInitialProgram,
      WTSApplicationName,
      WTSWorkingDirectory,
      WTSOEMId,
      WTSSessionId,
      WTSUserName,
      WTSWinStationName,
      WTSDomainName,
      WTSConnectState,
      WTSClientBuildNumber,
      WTSClientName,
      WTSClientDirectory,
      WTSClientProductId,
      WTSClientHardwareId,
      WTSClientAddress,
      WTSClientDisplay,
      WTSClientProtocolType
    };

    public enum WindowsMessage : int {
      WTSSessionChange = 689,    // WM_WTSSESSION_CHANGE    = 0x2b1
    };
    private enum WTSSessionNotification : int {
      NotifyForThisSession = 0,  // NOTIFY_FOR_THIS_SESSION = 0x0
      NotifyForAllSession = 1,   // NOTIFY_FOR_ALL_SESSIONS = 0x1
    };
    public enum WTSMessage : int {
      ConsoleConnect = 1,        // WTS_CONSOLE_CONNECT        = 0x1
      ConsoleDisconnect = 2,     // WTS_CONSOLE_DISCONNECT     = 0x2
      RemoteConnect = 3,         // WTS_REMOTE_CONNECT         = 0x3
      RemoteDisconnect = 4,      // WTS_REMOTE_DISCONNECT      = 0x4
      SessionLogon = 5,          // WTS_SESSION_LOGON          = 0x5
      SessionLogoff = 6,         // WTS_SESSION_LOGOFF         = 0x6
      SessionLock = 7,           // WTS_SESSION_LOCK           = 0x7
      SessionUnlock = 8,         // WTS_SESSION_UNLOCK         = 0x8
      SessionRemoteControl = 9,  // WTS_SESSION_REMOTE_CONTROL = 0x9
    };

    [DllImport("wtsapi32.dll")]
    private static extern bool WTSQuerySessionInformation(IntPtr hServer, int sessionId, WTS_INFO_CLASS wtsInfoClass, out IntPtr ppBuffer, out uint pBytesReturned);
    [DllImport("wtsapi32.dll", ExactSpelling = true, SetLastError = false)]
    private static extern void WTSFreeMemory(IntPtr memory);

    [DllImport("WtsApi32.dll")]
    private static extern bool WTSRegisterSessionNotification(IntPtr hWnd, [MarshalAs(UnmanagedType.U4)] int dwFlags);
    [DllImport("WtsApi32.dll")]
    private static extern bool WTSUnRegisterSessionNotification(IntPtr hWnd);

  }

}