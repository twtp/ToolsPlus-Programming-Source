using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {
  public static class IPC {

    public static Microsoft.Win32.SafeHandles.SafeFileHandle CreateReadablePipeFile(string pipeName, uint bufferSize) {
      return CreateNamedPipe(
          pipeName,
          PIPE_ACCESS_INBOUND | FILE_FLAG_OVERLAPPED,
          0,
          255,
          bufferSize,
          bufferSize,
          0,
          IntPtr.Zero
        );
    }

    public static int WaitForPipeConnection(Microsoft.Win32.SafeHandles.SafeFileHandle pipeHandle) {
      return ConnectNamedPipe(
          pipeHandle,
          IntPtr.Zero
        );
    }

    public static Microsoft.Win32.SafeHandles.SafeFileHandle CreateWritablePipeFile(string pipeName) {
      return CreateFile(
          pipeName,
          GENERIC_WRITE,
          0,
          IntPtr.Zero,
          OPEN_EXISTING,
          FILE_FLAG_OVERLAPPED,
          IntPtr.Zero
        );
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateNamedPipe(
        String pipeName,
        uint dwOpenMode,
        uint dwPipeMode,
        uint nMaxInstances,
        uint nOutBufferSize,
        uint nInBufferSize,
        uint nDefaultTimeOut,
        IntPtr lpSecurityAttributes
      );
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern int ConnectNamedPipe(
        Microsoft.Win32.SafeHandles.SafeFileHandle hNamedPipe,
        IntPtr lpOverlapped
      );
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateFile(
        String pipeName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplate
      );

    private const uint PIPE_ACCESS_DUPLEX = (0x00000003);
    private const uint PIPE_ACCESS_INBOUND = (0x00000001);
    private const uint PIPE_ACCESS_OUTBOUND = (0x00000002);

    private const uint FILE_FLAG_OVERLAPPED = (0x40000000);

    private const uint GENERIC_READ = (0x80000000);
    private const uint GENERIC_WRITE = (0x40000000);
    private const uint OPEN_EXISTING = 3;

  }
}
