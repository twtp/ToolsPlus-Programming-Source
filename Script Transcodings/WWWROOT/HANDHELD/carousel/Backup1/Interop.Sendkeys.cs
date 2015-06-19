using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {
  public static class Sendkeys {

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint numberOfInputs, [MarshalAs(UnmanagedType.LPArray, SizeConst = 1)] INPUT[] pInputs, int structSize);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")]
    static extern uint MapVirtualKey(uint uCode, uint uMapType);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern short VkKeyScan(char key);

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT {
      internal uint type;
      internal InputUnion U;
      internal static int Size { get { return Marshal.SizeOf(typeof(INPUT)); } }
    }
    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion {
      [FieldOffset(0)]
      internal MOUSEINPUT mi;
      [FieldOffset(0)]
      internal KEYBDINPUT ki;
      [FieldOffset(0)]
      internal HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT {
      public ushort wVk;
      public ushort wScan;
      public uint dwFlags;
      public uint time;
      public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT {
      public int dx;
      public int dy;
      public uint mouseData;
      public uint dwFlags;
      public uint time;
      public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct HARDWAREINPUT {
      public uint uMsg;
      public ushort wParamL;
      public ushort wParamH;
    }

    private const int INPUT_KEYBOARD = 1;
    private const int KEY_UP = 0x0002;
    private const int KEY_SCANCODE = 0x0004;

    public static bool Send(string s) {
      if (s.Length == 0) {
        return true;
      }

      INPUT[] input = new INPUT[2 * s.Length];
      int i = 0;
      foreach (char c in s.ToCharArray()) {
        input[i] = new INPUT();
        input[i].type = INPUT_KEYBOARD;
        input[i].U.ki = new KEYBDINPUT();
        input[i].U.ki.dwFlags = KEY_SCANCODE;
        input[i].U.ki.wScan = (ushort)(c & 0xFF);
        i++;

        input[i] = new INPUT();
        input[i].type = INPUT_KEYBOARD;
        input[i].U.ki = new KEYBDINPUT();
        input[i].U.ki.dwFlags = KEY_SCANCODE;
        input[i].U.ki.dwFlags |= KEY_UP;
        input[i].U.ki.wScan = (ushort)(c & 0xFF);
        i++;
      }

      uint result = SendInput((uint)input.Length, input, Marshal.SizeOf(input[0]));
      if (result == 0) {
        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error().ToString());
      }
      else {
        return true;
      }
    }

    public static bool SendVirtualKey(byte vk) {
      byte sc = (byte)MapVirtualKey((uint)vk, 0);
      keybd_event(vk, sc, 0, (UIntPtr)0);
      keybd_event(vk, sc, 2, (UIntPtr)0);
      return true;
    }

    public static bool SendAltCombo(char otherKey) {
      byte keycode = (System.Text.Encoding.ASCII.GetBytes(new char[] { otherKey }))[0];
      keybd_event(0x12, 0xB8, 0, (UIntPtr)0);
      keybd_event(keycode, 0x92, 0, (UIntPtr)0);
      keybd_event(keycode, 0x92, 2, (UIntPtr)0);
      keybd_event(0x12, 0xB8, 2, (UIntPtr)0);
      return true;
    }

    public static bool SendAltE() {
      return SendAltCombo('E');
    }

    public static bool SendEnter() {
      return SendVirtualKey(0x0D);
    }

    public static bool SendTab() {
      return SendVirtualKey(0x09);
    }

    public static bool SendDelete() {
      return SendVirtualKey(0x2E);
    }

    public static short GetKeyForCharacter(char c) {
      return VkKeyScan(c);
    }
  }
}
