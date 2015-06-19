using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.IPC {

  public delegate void IPCMessageReceivedEventHandler(object sender, IPCMessageEventArgs e);

  public class IPCMessageEventArgs : EventArgs {
    private string message;
    public IPCMessageEventArgs(byte[] buffer) : this(new ASCIIEncoding().GetString(buffer)) {
      // nothing
    }
    public IPCMessageEventArgs(string message) {
      this.message = message.TrimEnd('\0');
    }
    public string Message { get { return this.message; } }
  }

  public static class NamedPipes {

    public static event IPCMessageReceivedEventHandler IPCMessageReceived;

    public const uint BUFFER_SIZE = 4096;

    private static System.Threading.Thread listenerThread;

    private static string sid = System.Diagnostics.Process.GetCurrentProcess().SessionId.ToString();
    private static string buildPipeName() {
      return "\\\\.\\pipe\\" + "whsepipe-" + sid;
    }

    public static bool Write(string message) {
      try {
        using (Microsoft.Win32.SafeHandles.SafeFileHandle pipeHandle = TP.Base.Interop.IPC.CreateWritablePipeFile(buildPipeName())) {

          if (pipeHandle.IsInvalid) {
            return false;
          }

          using (System.IO.FileStream fs = new System.IO.FileStream(pipeHandle, System.IO.FileAccess.Write, (int)BUFFER_SIZE)) {
            byte[] towrite = new ASCIIEncoding().GetBytes(message);
            fs.Write(towrite, 0, towrite.Length);
            fs.Flush();
            fs.Close();
            return true;
          }
        }
      }
      catch (Exception ex) {
        TP.Base.Logger.Log("----> EXCEPTION IN IPC.NamedPipes.Write(message:=" + message + "): " + ex.Message);
        return false;
      }
    }

    public static void ListenStart() {
      if (IsListening()) {
        ListenStop();
      }
      listenerThread = new System.Threading.Thread(new System.Threading.ThreadStart(listener));
      listenerThread.IsBackground = true;
      listenerThread.Start();
    }
    public static void ListenStop() {
      // this doesn't work at all, the thread gets blocked on ConnectNamedPipe
      // and won't respond to anything else. is there an async version or something
      // to use?
      // anyway, only called when program is dying, so we can force kill it there instead. fuck.
      if (listenerThread != null) {
        listenerThread = null;
      }
    }
    public static bool IsListening() {
      return listenerThread != null && listenerThread.IsAlive;
    }
    public static void ListenSync() {
      listener();
    }

    private static void listener() {
      try {
        Microsoft.Win32.SafeHandles.SafeFileHandle clientPipeHandle;
        while (true) {
          clientPipeHandle = TP.Base.Interop.IPC.CreateReadablePipeFile(buildPipeName(), BUFFER_SIZE);
          
          if (clientPipeHandle.IsInvalid) {
            TP.Base.Logger.Log("IPC thread invalid pipe?");
            clientPipeHandle.Dispose();
            break;
          }

          TP.Base.Logger.Log("IPC thread waiting for connection");
          int success = TP.Base.Interop.IPC.WaitForPipeConnection(clientPipeHandle);

          if (success != 1) {
            TP.Base.Logger.Log("IPC thread bad connection?");
            clientPipeHandle.Dispose();
            break;
          }

          // client connection successful
          // handle client communication

          TP.Base.Logger.Log("IPC thread received connection, reading");
          StringBuilder sb = new StringBuilder();
          while (clientPipeHandle.IsClosed == false) {
            using (System.IO.FileStream fs = new System.IO.FileStream(clientPipeHandle, System.IO.FileAccess.Read, (int)BUFFER_SIZE)) {
              byte[] buffer = new byte[BUFFER_SIZE];
              int bytesRead = fs.Read(buffer, 0, (int)BUFFER_SIZE);
              if (bytesRead == 0) {
                break;
              }
              sb.Append(System.Text.Encoding.ASCII.GetString(buffer));
            }
          }
          string msg = sb.ToString();
          if (!string.IsNullOrEmpty(msg)) {
            if (IPCMessageReceived != null) {
              IPCMessageReceived(null, new IPCMessageEventArgs(msg));
            }
          }
          clientPipeHandle.Dispose();
        }
      }
      catch (Exception ex) {
        TP.Base.Logger.Log(" --> thread for IPC listener() died! " + ex.Message + ex.StackTrace);
        // silently throw away the error
        // i guess people will complain eventually if this dies, and i can check the logs?
        //throw;
      }
    }

  }
}
