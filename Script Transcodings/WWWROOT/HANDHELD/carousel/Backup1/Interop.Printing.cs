using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace TP.Base.Interop {
  public static class Printing {

    public static bool PrintRAWString(string szPrinterName, string szString, string documentName) {
      IntPtr pBytes;
      Int32 dwCount;
      // How many characters are in the string?
      //dwCount = szString.Length;
      dwCount = (szString.Length + 1) * Marshal.SystemMaxDBCSCharSize; // fix to add null char and character width
      // Assume that the printer is expecting ANSI text, and then convert
      // the string to ANSI text.
      pBytes = Marshal.StringToCoTaskMemAnsi(szString);
      // Send the converted ANSI string to the printer.
      printBytes(szPrinterName, pBytes, dwCount, documentName);
      Marshal.FreeCoTaskMem(pBytes);
      return true;
    }

    public static bool PrintRAWBytes(string szPrinterName, byte[] pBytes, string documentName) {
      IntPtr p = Marshal.AllocHGlobal(pBytes.Length);
      Marshal.Copy(pBytes, 0, p, pBytes.Length);
      printBytes(szPrinterName, p, pBytes.Length, documentName);
      Marshal.FreeHGlobal(p);
      return true;
    }


    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    private class DOCINFOA {
      [MarshalAs(UnmanagedType.LPStr)]
      public string pDocName;
      [MarshalAs(UnmanagedType.LPStr)]
      public string pOutputFile;
      [MarshalAs(UnmanagedType.LPStr)]
      public string pDataType;
    }

    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool OpenPrinter([MarshalAs(UnmanagedType.LPStr)] string szPrinter, out IntPtr hPrinter, IntPtr pd);
    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool ClosePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool StartDocPrinter(IntPtr hPrinter, Int32 level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);
    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool EndDocPrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool StartPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool EndPagePrinter(IntPtr hPrinter);
    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    private static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, Int32 dwCount, out Int32 dwWritten);


    // SendBytesToPrinter()
    // When the function is given a printer name and an unmanaged array
    // of bytes, the function sends those bytes to the print queue.
    // Returns true on success, false on failure.
    private static bool printBytes(string szPrinterName, IntPtr pBytes, Int32 dwCount, string documentName) {
      Int32 dwError = 0, dwWritten = 0;
      IntPtr hPrinter = new IntPtr(0);
      DOCINFOA di = new DOCINFOA();
      bool bSuccess = false; // Assume failure unless you specifically succeed.

      di.pDocName = documentName;
      di.pDataType = "RAW";

      // Open the printer.
      if (OpenPrinter(szPrinterName.Normalize(), out hPrinter, IntPtr.Zero)) {
        // Start a document.
        if (StartDocPrinter(hPrinter, 1, di)) {
          // Start a page.
          if (StartPagePrinter(hPrinter)) {
            // Write your bytes.
            bSuccess = WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
            EndPagePrinter(hPrinter);
          }
          EndDocPrinter(hPrinter);
        }
        ClosePrinter(hPrinter);
      }
      // If you did not succeed, GetLastError may give more information
      // about why not.
      if (bSuccess == false) {
        dwError = Marshal.GetLastWin32Error();
      }
      return bSuccess;
    }

  }

}
