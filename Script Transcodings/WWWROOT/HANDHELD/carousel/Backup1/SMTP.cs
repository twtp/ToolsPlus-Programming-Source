using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base {
  public static class SMTP {

    private static string exchangeServer = "toolsplus06";
    private static int exchangePort = 25;
    private static string exchangeUsername = "masemail";
    private static string exchangePassword = "masemail";

    public static bool SendEmail(string fromEmail, string toEmail, string subject, string message, bool async) {
      return SendEmail(fromEmail, new string[] { toEmail, }, subject, message, async);
    }

    public static bool SendEmail(string fromEmail, List<string> toEmails, string subject, string message, bool async) {
      return SendEmail(fromEmail, toEmails.ToArray(), subject, message, async);
    }

    public static bool SendEmail(string fromEmail, string[] toEmails, string subject, string message, bool async) {
      System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
      msg.From = new System.Net.Mail.MailAddress(fromEmail);
      foreach (string toEmail in toEmails) {
        msg.To.Add(toEmail);
      }
      msg.Subject = subject;
      msg.Body = message;
      if (message.StartsWith("<html", StringComparison.CurrentCultureIgnoreCase) ||
          message.StartsWith("<!DOCTYPE", StringComparison.CurrentCultureIgnoreCase)) {
        msg.IsBodyHtml = true;
      }
      return SendEmail(msg, async);
    }

    public static bool SendEmail(System.Net.Mail.MailMessage msg, bool async) {
      // this is public, so if we wanted to we could build a mailmessage somewhere
      // else, possibly specifying multipart/alternative pieces.
      // or should an override handle textpart + htmlpart as params?
      if (async) {
        System.Threading.ThreadPool.QueueUserWorkItem(sendEmail, msg);
      }
      else {
        sendEmail(msg);
      }
      return true;
    }

    // System.Threading.WaitCallback signature
    private static void sendEmail(object stateInfo) {
      try {
        System.Net.Mail.MailMessage msg = (System.Net.Mail.MailMessage)stateInfo;
        System.Net.Mail.SmtpClient c = new System.Net.Mail.SmtpClient(exchangeServer, exchangePort);
        c.Credentials = new System.Net.NetworkCredential(exchangeUsername, exchangePassword);
        c.Send(msg);
      }
      catch (System.Net.Mail.SmtpException /*smtpex*/) {
        //TP.Base.Logger.Log(" --> thread for sendEmail() died! " + smtpex.Message + smtpex.StackTrace);
        throw;
      }
      catch (Exception /*ex*/) {
        //TP.Base.Logger.Log(" --> thread for sendEmail() died! " + ex.Message + ex.StackTrace);
        throw;
      }
    }

    /*public static bool SendEmail(string fromEmail, string toEmail, string subject, string message, bool async) {
      List<string> temp = new List<string>();
      temp.Add(toEmail);
      if (!async) {
        return sendEmail(fromEmail, temp, subject, message);
      }
      else {
        System.ComponentModel.BackgroundWorker bw = new System.ComponentModel.BackgroundWorker();
        bw.DoWork += delegate(object sender, System.ComponentModel.DoWorkEventArgs e) {
            sendEmail(fromEmail, temp, subject, message);
          };
        bw.RunWorkerAsync();
        return true;
      }
    }
    public static bool SendEmail(string fromEmail, string[] toEmails, string subject, string message, bool async) {
      List<string> temp = new List<string>();
      foreach (string e in toEmails) {
        temp.Add(e);
      }
      if (!async) {
        return sendEmail(fromEmail, temp, subject, message);
      }
      else {
        System.ComponentModel.BackgroundWorker bw = new System.ComponentModel.BackgroundWorker();
        bw.DoWork += delegate(object sender, System.ComponentModel.DoWorkEventArgs e) {
            sendEmail(fromEmail, temp, subject, message);
          };
        bw.RunWorkerAsync();
        return true;
      }
    }
    public static bool SendEmail(string fromEmail, List<string> toEmails, string subject, string message, bool async) {
      if (!async) {
        return sendEmail(fromEmail, toEmails, subject, message);
      }
      else {
        System.ComponentModel.BackgroundWorker bw = new System.ComponentModel.BackgroundWorker();
        bw.DoWork += delegate(object sender, System.ComponentModel.DoWorkEventArgs e) {
          sendEmail(fromEmail, toEmails, subject, message);
        };
        bw.RunWorkerAsync();
        return true;
      }
    }
    
    private static bool sendEmail(string fromEmail, List<string> toEmails, string subject, string message) {
      System.Net.Sockets.TcpClient smtp = new System.Net.Sockets.TcpClient();
      try {
        smtp.Connect(exchangeServer, exchangePort);
      }
      catch {
        return false;
      }
      System.Net.Sockets.NetworkStream nst = smtp.GetStream();
      System.IO.StreamWriter stw = new System.IO.StreamWriter(nst);
      System.IO.StreamReader str = new System.IO.StreamReader(nst);

      try {
        // greeting
        waitForResponse("220", nst, str);
        // return greeting
        stw.WriteLine("EHLO internalmx.tools-plus.com");
        stw.Flush();
        // greeting ok message
        waitForResponse("250", nst, str);
        // login
        stw.WriteLine("AUTH LOGIN");
        stw.Flush();
        waitForResponse("334", nst, str);
        stw.WriteLine(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(exchangeUsername)));
        stw.Flush();
        waitForResponse("334", nst, str);
        stw.WriteLine(Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(exchangePassword)));
        stw.Flush();
        // login ok message
        waitForResponse("235", nst, str);
        // send from email
        stw.WriteLine("MAIL FROM: " + fromEmail);
        stw.Flush();
        // from email ok
        waitForResponse("250", nst, str);
        // send to email
        foreach (string toEmail in toEmails) {
          stw.WriteLine("RCPT TO: " + toEmail);
          stw.Flush();
          // to email ok
          waitForResponse("250", nst, str);
        }
        // prepare data
        stw.WriteLine("DATA");
        stw.Flush();
        // prepare data ok
        waitForResponse("354", nst, str);
        // send data
        string smtpData = "";
        smtpData += "From: " + fromEmail + Environment.NewLine;
        smtpData += "To: " + string.Join("; ", toEmails.ToArray()) + Environment.NewLine;
        smtpData += "Subject: " + subject + Environment.NewLine;
        smtpData += Environment.NewLine;
        smtpData += message + Environment.NewLine;
        smtpData += "." + Environment.NewLine;
        stw.Write(smtpData);
        stw.Flush();
        // send data ok
        waitForResponse("250", nst, str);
        // send quit
        stw.WriteLine("QUIT");
        stw.Flush();
        // quit acknowledged
        waitForResponse("221", nst, str);
        smtp.Close();
        return true;
      }
      catch (Exception) {
        smtp.Close();
        throw;
      }
    }

    private static bool waitForResponse(string responseCode, System.Net.Sockets.NetworkStream nst, System.IO.StreamReader str) {
      return waitForResponse(new string[] { responseCode }, nst, str);
    }

    private static bool waitForResponse(string[] responseCodes, System.Net.Sockets.NetworkStream nst, System.IO.StreamReader str) {
      DateTime starttime = DateTime.Now;
      TimeSpan tsp;
      do {
        System.Threading.Thread.Sleep(50);
        if (nst.DataAvailable) {
          string incoming = str.ReadLine();
          foreach (string rc in responseCodes) {
            if (incoming.Substring(0, rc.Length) == rc) {
              return true;
            }
          }
        }
        tsp = DateTime.Now - starttime;
      } while (tsp.Minutes * 60 + tsp.Seconds < smtpTimeOut);

      throw new Exception("Timeout waiting for MTA response " + responseCodes[0]);
    }*/


  }
}
