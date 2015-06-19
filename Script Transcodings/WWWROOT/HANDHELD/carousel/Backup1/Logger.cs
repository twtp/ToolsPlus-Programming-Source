using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {
  public static class Logger {

    public delegate bool LoggingRedirect(string message);

    private static readonly object lockObj = new object();
    private static string fileName = null;

    private static bool enabled = true;
    public static bool Enabled { get { return enabled; } set { enabled = value; } }

    private static LoggingRedirect redirectLogs = null;
    public static LoggingRedirect RedirectLogsTo { get { return redirectLogs; } set { redirectLogs = value; } }

    public static bool Log(string message) {
      if (enabled) {
        lock (lockObj) {
          if (redirectLogs != null) {
            return redirectLogs(message);
          }

          if (fileName == null) {
            doInitialization();
          }
          System.IO.File.AppendAllText(
              fileName,
              string.Format("{0:M/d/yyyy h:mm:ss tt}: {1}{2}", DateTime.Now, message, System.Environment.NewLine)
            );
        }
      }
      return true;
    }

    private static bool doInitialization() {
      DateTime now = DateTime.Now;

      string logPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "log");
      System.IO.Directory.CreateDirectory(logPath);
      foreach (System.IO.FileInfo fi in new System.IO.DirectoryInfo(logPath).GetFiles("*.log")) {
        TimeSpan ts = now.Subtract(fi.LastWriteTime);
        if (ts.Days > 10) {
          fi.Delete();
        }
      }

      fileName = string.Format(".\\log\\{0:yyyyMMddHHmmss}.log", now);

      return true;
    }

    // what a terrible failure
    public static bool WTF(string message) {
      return WTF(message, "warehousing error message");
    }
    public static bool WTF(string message, string subject) {
      string local = Environment.GetEnvironmentVariable("CLIENTNAME");
      if (local == null) {
        local = string.Empty;
      }
      else {
        local = "local: " + local + Environment.NewLine;
      }
      message = "client: " + Environment.MachineName + Environment.NewLine +
                local +
                "version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + Environment.NewLine +
                "message: " + message;
      return TP.Base.SMTP.SendEmail(
          "no-reply@tools-plus.com",
          "brian@tools-plus.com",
          subject,
          message,
          false
        );
    }

    // even more terrible, email the entire log file
    public static bool SuperWTF() {
      string message = System.IO.File.ReadAllText(fileName);
      return TP.Base.SMTP.SendEmail(
          "no-reply@tools-plus.com",
          "brian@tools-plus.com",
          "warehousing error logfile",
          message,
          false
        );
    }
  }
}
