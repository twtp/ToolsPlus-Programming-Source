using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface {
  internal static class CIC332 {

    // serial port access
    private static TP.Base.Port.SerialPort port = null;
    private static readonly object writeLock = new object();
    private static List<string> chunker(ref string incoming) {
      List<string> retval = new List<string>();

      bool done;
      do {
        if (incoming.Length == 0) {
          done = true;
        }
        else if (re_ACK.IsMatch(incoming)) {
          // ACK
          int pos = incoming.IndexOf('\x06');
          retval.Add(incoming.Substring(0, pos + 1));
          incoming = incoming.Substring(pos + 1);
          done = false;
        }
        else {
          // other messages end with ETX and checksum
          int pos = incoming.IndexOf('\x03');
          if (pos == -1 || incoming.Length == pos + 1) {
            // partial transmission
            done = true;
          }
          else {
            retval.Add(incoming.Substring(0, pos + 2));
            incoming = incoming.Substring(pos + 2);
            done = false;
          }
        }
      } while (!done);

      return retval;
    }

    // message passing stuff, a buffer to hold partial messages, a queue
    // for outgoing commands, and a flag and lock for incoming ACKs
    private static TP.Base.DataStructure.BlockingQueue<Message.OutgoingMessage> messagesToSend = new TP.Base.DataStructure.BlockingQueue<Message.OutgoingMessage>();
    private static bool ackFlag = false;
    private static readonly object ackLock = new object();

    // regular expressions to identify and handle incoming messages
    private static System.Text.RegularExpressions.Regex re_ACK = new System.Text.RegularExpressions.Regex("^\x00*\x06");

    // published events, probably to be handled by TP.Warehousing.White.Pod.Controller
    public static event CIC332StatusReceivedEventHandler StatusReceived;
    public static event CIC332FaultsReceivedEventHandler FaultsReceived;
    public static event CIC332TextMessageReceivedEventHandler TextReceived;
    public static event LEDIlluminatedEventHandler LEDIlluminated;
    public static event LEDClearedEventHandler LEDCleared;
        

    // static initializer
    static CIC332() {
      /*Dictionary<string, string> cfg = TP.Warehousing.Config.For("pod")[0];
      string comPort = "COM" + cfg["com"];*/
      string comPort = "COM" + TP.Base.Config.Pod.COMPort.ToString();

      port = new TP.Base.Port.SerialPort(chunker, comPort, 9600, System.IO.Ports.Parity.Even, 7, System.IO.Ports.StopBits.One);
      port.IncomingData += new TP.Base.Port.IncomingDataEventHandler(port_IncomingData);

      System.Threading.Thread qpThread = new System.Threading.Thread(new System.Threading.ThreadStart(QueueProcessor));
      qpThread.IsBackground = true;
      qpThread.Start();
    }

    public static string FormatAddress(string address) {
      return address.PadLeft(3, '0');
    }

    public static void WaitForQueueCompletion() {
      TP.Base.Logger.Log("CIC332.WaitForQueueCompletion(): blocking");
      while (messagesToSend.Count > 0) {
        //System.Threading.Monitor.Wait(messagesToSend, 2000);
        System.Threading.Thread.Sleep(100);
      }
      TP.Base.Logger.Log("CIC332.WaitForQueueCompletion(): queue completed");
      return;
    }

    private static void port_IncomingData(object sender, TP.Base.Port.IncomingDataEventArgs e) {
      identifyAndRaise(e.Message);
    }

    // enqueues a command to send to the CIC332
    public static bool SendCommand(Message.OutgoingMessage message) {
      //TP.Base.Logger.Log("CIC332 Queueing message '" + message + "'");
      messagesToSend.Enqueue(message);
      return true;
    }

    // this is a separate thread that runs to process outgoing messages from
    // the queue
    private static void QueueProcessor() {
      try {
        Message.OutgoingMessage msg;
        //TP.Base.Logger.Log("CIC332 QueueProcessor() started");
        while (messagesToSend.TryDequeue(out msg)) {
          TP.Base.Logger.Log("CIC332.QueueProcessor(): dequeueing message '" + msg.ToString() + "', have " + messagesToSend.Count + " remaining");
          int retryCount = 0;
          do {
            int start = System.Environment.TickCount;
            if (writeString(msg.ToSendable())) {
              while (ackFlag == false && System.Environment.TickCount - start < 1500) {
                System.Threading.Thread.Sleep(10);
              }
              lock (ackLock) {
                if (ackFlag == true) {
                  ackFlag = false;
                  TP.Base.Logger.Log("CIC332.QueueProcessor(): ack received");
                  break;
                }
                else {
                  // timed out, resend
                  TP.Base.Logger.Log("CIC332.QueueProcessor(): timed out waiting for ack");
                  retryCount++;
                }
              }
            }
            else {
              TP.Base.Logger.Log("CIC332.QueueProcessor(): error calling port.WriteString()");
              System.Threading.Thread.Sleep(1500);
              retryCount++;
            }
          } while (retryCount < 5);
          if (retryCount == 5) {
            // no ack received, should i throw an error?
            TP.Base.Logger.Log("CIC332.QueueProcessor(): NO ACK RECEIVED FOR MESSAGE '" + msg.ToString() + "', FAILING SILENTLY");
          }
          else {
            // successful, what to do here?

            // here we need to set up the event for the lights, since there's
            // no asynchronous "finished" notice. this isn't very pretty, but
            // it's a nice way to signal a queued event finishing. we'll
            // shovel this off to a separate thread, since it might take a
            // while to complete? what could the program be doing with this?
            Message.ILightMessage ilm = msg as Message.ILightMessage;
            if (ilm != null) {
              System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(lightSignaler), (object)ilm);
            }
          }
        }
        //TP.Base.Logger.Log("CIC332 QueueProcessor() ending");
        return;
      }
      catch (Exception ex) {
        TP.Base.Logger.Log(" --> thread for CIC332 QueueProcessor() died! " + ex.Message + ex.StackTrace);
        throw;
      }
    }

    // System.Threading.WaitCallback signature
    // called by QueueProcessor thread when a light message is
    // successfully sent to the CIC
    private static void lightSignaler(object stateInfo) {
      try {
        // stateInfo is an ILightMessage
        Message.Outgoing.LEDIlluminate ledi = stateInfo as Message.Outgoing.LEDIlluminate;
        if (ledi != null) {
          if (LEDIlluminated != null) {
            //TP.Base.Logger.Log("creating return event for led illumination");
            LEDIlluminated(null, new LEDIlluminatedEventArgs(ledi));
          }
          return;
        }
        Message.Outgoing.LEDClearAll ledca = stateInfo as Message.Outgoing.LEDClearAll;
        if (ledca != null) {
          if (LEDCleared != null) {
            //TP.Base.Logger.Log("creating return event for led clear all");
            LEDCleared(null, new LEDClearedEventArgs(ledca));
          }
          return;
        }
        Message.Outgoing.LEDClearSingle ledcs = stateInfo as Message.Outgoing.LEDClearSingle;
        if (ledcs != null) {
          if (LEDCleared != null) {
            //TP.Base.Logger.Log("creating return event for led clear single");
            LEDCleared(null, new LEDClearedEventArgs(ledcs));
          }
          return;
        }
      }
      catch (Exception ex) {
        TP.Base.Logger.Log(" --> thread for CIC332 lightSignaler() died! " + ex.Message + ex.StackTrace);
        throw;
      }
    }

    // low-level call to write a message to the serial port
    private static bool writeString(string message) {
      //TP.Base.Logger.Log("CIC332 Preparing to send message \"" + message + "\"");
      lock (writeLock) {
        try {
          return port.Write(message);
        }
        catch (Exception ex) {
          TP.Base.Logger.Log("CIC332.writeString(): Exception sending message \"" + message + "\": " + ex.Message + ex.StackTrace);
          return false;
        }
      }
    }

    // low-level identification of CIC332 message types
    private static void identifyAndRaise(string message) {
      //TP.Base.Logger.Log("attempting to identify message: " + getPrintable(message));
      if (re_ACK.IsMatch(message)) {
        lock (ackLock) {
          ackFlag = true;
          //TP.Base.Logger.Log("message is ACK, flag set");
        }
        return;
      }

      TP.Base.Logger.Log("CIC332.identifyAndRaise(): attempting to identify message: " + message);

      Message.IncomingMessage im;

      im = Message.Incoming.CarouselStatusReply.Parse(message);
      if (im != null) {
        TP.Base.Logger.Log("CIC332.identifyAndRaise(): message is SS, ACK'ing and raising");
        acknowledge();
        if (StatusReceived != null) {
          StatusReceived(null, new StatusMessageEventArgs((Message.Incoming.CarouselStatusReply)im));
        }
        return;
      }

      im = Message.Incoming.CarouselFaultsReply.Parse(message);
      if (im != null) {
        TP.Base.Logger.Log("CIC332.identifyAndRaise(): message is SF, ACK'ing and raising");
        acknowledge();
        if (FaultsReceived != null) {
          FaultsReceived(null, new FaultsMessageEventArgs((Message.Incoming.CarouselFaultsReply)im));
        }
        return;
      }

      im = Message.Incoming.ASCIIText.Parse(message);
      if (im != null) {
        TP.Base.Logger.Log("CIC332.identifyAndRaise(): message is TX, ACK'ing and raising");
        acknowledge();
        if (TextReceived != null) {
          TextReceived(null, new TextMessageEventArgs((Message.Incoming.ASCIIText)im));
        }
        return;
      }

      // unknown message, ignore?
      TP.Base.Logger.Log("CIC332.identifyAndRaise(): Unknown message received: " + message);
      return;
    }

    // acknowledgement of receipt of an incoming message
    private static bool acknowledge() {
      return writeString("\x06");
    }

    private static string getPrintable(string message) {
      string retval = string.Empty;
      foreach (byte b in System.Text.Encoding.ASCII.GetBytes(message)) {
        retval += " " + b.ToString();
      }
      if (retval.Length > 0) {
        return retval.Substring(1);
      }
      else {
        return string.Empty;
      }
    }
  }
}
