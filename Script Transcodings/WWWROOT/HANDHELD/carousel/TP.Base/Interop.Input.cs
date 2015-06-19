using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {

  public static class Input {

    public static long GetLastInputTime(int lastScanTick) {
      LASTINPUTINFO lii = new LASTINPUTINFO();
      lii.cbSize = Marshal.SizeOf(lii);
      lii.dwTime = 0;
      bool retval = GetLastInputInfo(ref lii);
      if (retval) {
        long idle = Environment.TickCount - Math.Max(lii.dwTime, (uint)lastScanTick);
        //System.Diagnostics.Debug.WriteLine("idle for " + idle.ToString());
        return idle;
      }
      else {
        return 0;
      }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    private struct LASTINPUTINFO {
      [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)]
      public int cbSize;
      [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)]
      public UInt32 dwTime;
    }

  }
}
