using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {
  public static class Audio {

    public static bool PlaySoundByBytes(byte[] data, bool playAsync) {
      if (playAsync) {
        return PlaySound(data, IntPtr.Zero, SND_ASYNC | SND_MEMORY);
      }
      else {
        return PlaySound(data, IntPtr.Zero, SND_MEMORY);
      }
    }

    [DllImport("Winmm.dll")]
    private static extern bool PlaySound(byte[] data, IntPtr hMod, UInt32 dwFlags);
    private const UInt32 SND_SYNC = 0;
    private const UInt32 SND_ASYNC = 1;
    private const UInt32 SND_MEMORY = 4;

  }
}
