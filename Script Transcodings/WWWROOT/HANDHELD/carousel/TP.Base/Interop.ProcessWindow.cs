using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {
  public static class ProcessWindow {

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    public static int? GetForegroundWindowProcessId() {
      IntPtr activatedHandle = GetForegroundWindow();
      if (activatedHandle == IntPtr.Zero) {
        return null; // not sure what can cause this, but possible according to docs
      }
      int activeProcID;
      GetWindowThreadProcessId(activatedHandle, out activeProcID);
      return activeProcID;
    }

    public static string GetFocusedWindowText() {
      IntPtr activeWindowHandle = GetForegroundWindow();

      int length = GetWindowTextLength(activeWindowHandle);

      StringBuilder buf = new StringBuilder(length + 1);
      int ret = GetWindowText(activeWindowHandle, buf, buf.Capacity);
      if (ret > 0) {
        return buf.ToString();
      }
      else {
        return null;
      }
    }

    public static bool SetForegroundForm(System.Windows.Forms.Form form) {
      return SetForegroundWindow(form.Handle);
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int GetWindowThreadProcessId(IntPtr handle, out int processId);

    [DllImport("user32.dll")]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr hWnd);

  }
}
