using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Port {

  public delegate void IncomingDataEventHandler(object sender, IncomingDataEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class IncomingDataEventArgs : EventArgs {
    private string message;
    internal IncomingDataEventArgs(string message) {
      this.message = message;
    }
    public string Message { get { return this.message; } }
  }

  public class PortMissingException : System.IO.IOException {
    public PortMissingException(string message, System.IO.IOException innerException) : base(message, innerException) { /*nothing*/ }
  }
  public class DeviceErrorException : System.IO.IOException {
    public DeviceErrorException(string message, System.IO.IOException innerException) : base(message, innerException) { /*nothing*/ }
  }
  public class PortInUseException : System.UnauthorizedAccessException {
    public PortInUseException(string message, System.UnauthorizedAccessException innerException) : base(message, innerException) { /*nothing*/ }
  }

  public class SerialPort {
    // message raiser
    public event IncomingDataEventHandler IncomingData;
    // buffer chunker
    public delegate List<string> BufferChunker(ref string incoming);

    // actual .net serial port access
    private System.IO.Ports.SerialPort port = null;
    // buffer to hold partial messages
    private string messageBuffer = string.Empty;
    // buffer chunking algorithm
    private BufferChunker chunkingAlgorithm;

    public SerialPort(BufferChunker chunkingAlgorithm, string comPort, int baudRate, System.IO.Ports.Parity parity, int dataBits, System.IO.Ports.StopBits stopBits) {
      TP.Base.Logger.Log("TP.Ports.SerialPort CTOR(): starting");
      this.chunkingAlgorithm = chunkingAlgorithm;
      this.setupPort(comPort, baudRate, parity, dataBits, stopBits);
      TP.Base.Logger.Log("TP.Ports.SerialPort CTOR(): finished");
    }
    public SerialPort(BufferChunker chunkingAlgorithm, string connString) {
      this.chunkingAlgorithm = chunkingAlgorithm;
      // expects a string like "COM1,9600,8,N,1"
      string[] pieces = connString.Split(',');
      this.setupPort(
          pieces[0],
          Convert.ToInt32(pieces[1]),
          toParity(pieces[2]),
          Convert.ToInt32(pieces[3]),
          toStopBits(pieces[4])
        );
    }

    public bool Kill() {
      TP.Base.Logger.Log("TP.Ports.SerialPort.Kill(): starting");
      try {
        TP.Base.Logger.Log("TP.Ports.SerialPort.Kill(): disassociating data received event");
        this.port.DataReceived -= new System.IO.Ports.SerialDataReceivedEventHandler(this.port_DataReceived);
        TP.Base.Logger.Log("TP.Ports.SerialPort.Kill(): closing port");
        this.port.Close();
        TP.Base.Logger.Log("TP.Ports.SerialPort.Kill(): nulling port object");
        this.port = null;
        TP.Base.Logger.Log("TP.Ports.SerialPort.Kill(): returning true");
        return true;
      }
      catch (System.IO.IOException ioex) {
        // The port is in an invalid state. - or - An attempt to set the state of the underlying port
        // failed. For example, the parameters passed from this SerialPort object were invalid.
        TP.Base.Logger.Log("--> IOException occurred closing RS232 port: " + ioex.Message);
        this.port = null;
        return false;
      }
      catch (Exception ex) {
        // non-documented error?
        TP.Base.Logger.Log("--> Exception occurred closing RS232 port: " + ex.Message);
        this.port = null;
        return false;
      }
    }

    public bool Write(string outgoing) {
      //TP.Base.Logger.Log("--> port write: " + outgoing);
      try {
        this.port.Write(outgoing);
        return true;
      }
      catch (Exception ex) {
        TP.Base.Logger.Log("--> Exception occurred writing to RS232 port: " + ex.Message + ex.StackTrace);
        return false;
      }
    }

    public string COMPort { get { return this.port.PortName; } }
    public int BaudRate { get { return this.port.BaudRate; } }
    public int DataBits { get { return this.port.DataBits; } }
    public System.IO.Ports.Parity Parity { get { return this.port.Parity; } }
    public System.IO.Ports.StopBits StopBits { get { return this.port.StopBits; } }

    private void port_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e) {
      try {
        lock (this.messageBuffer) {
          string incoming = this.port.ReadExisting();
          //TP.Base.Logger.Log("--> port read: " + incoming);
          if (incoming != string.Empty) {
            this.messageBuffer += incoming;
            List<string> todo = this.chunkingAlgorithm(ref this.messageBuffer);
            //TP.Base.Logger.Log("--> port chunker found " + todo.Count + " messages, buffer is now " + (this.messageBuffer == string.Empty ? "<empty>" : this.messageBuffer));
            if (todo.Count != 0) {
              if (this.IncomingData != null) {
                // raising this event can be long...db lookup for a barcode, etc.
                // so, fire it off to a threadpool worker who can handle it, and
                // then we can get back to just reading and chunking
                foreach (string s in todo) {
                  // just fire these off, the chunking algorithm is responsible
                  // for sorting out duplicates and other weirdness?
                  System.Threading.ThreadPool.QueueUserWorkItem(new System.Threading.WaitCallback(messageRaiser), (object)s);
                  //System.Threading.Thread.Sleep(500);
                }
              }
            }
          }
        }
      }
      catch (Exception ex) {
        TP.Base.Logger.Log(" --> thread for TP.Ports.SerialPort.port_DataReceived() died! " + ex.Message + ex.StackTrace);
        throw;
      }
    }

    // System.Threading.WaitCallback signature
    private void messageRaiser(object stateInfo) {
      try {
        this.IncomingData(this, new IncomingDataEventArgs((string)stateInfo));
      }
      catch (Exception ex) {
        TP.Base.Logger.Log(" --> thread for TP.Ports.SerialPort.messageRaiser() died! " + ex.Message + ex.StackTrace);
        throw;
      }
    }


    private static System.IO.Ports.Parity toParity(string parity) {
      switch (parity) {
        case "O":
          return System.IO.Ports.Parity.Odd;
        case "E":
          return System.IO.Ports.Parity.Even;
        case "N":
          return System.IO.Ports.Parity.None;
        case "M":
          return System.IO.Ports.Parity.Mark;
        case "S":
          return System.IO.Ports.Parity.Space;
        default:
          throw new ArgumentException("invalid parity char '" + parity + "'");
      }
    }

    private static System.IO.Ports.StopBits toStopBits(string stopBits) {
      switch (stopBits) {
        case "0":
          return System.IO.Ports.StopBits.None;
        case "1":
          return System.IO.Ports.StopBits.One;
        case "2":
          return System.IO.Ports.StopBits.Two;
        default:
          throw new ArgumentException("invalid stop bit char '" + stopBits + "'");
      }
    }

    private void setupPort(string comPort, int baudRate, System.IO.Ports.Parity parity, int dataBits, System.IO.Ports.StopBits stopBits) {
      TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: starting");
      if (this.port != null) {
        TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: port is not null!");
        if (!this.Kill()) {
          // wtf, can't kill the port? guess just stopping here is the right thing to do?
          return;
        }
      }
      TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: creating new basic port object");
      this.port = new System.IO.Ports.SerialPort(comPort, baudRate, parity, dataBits, stopBits);
      TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: associating basic port data received event");
      this.port.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.port_DataReceived);
      try {
        TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: opening basic port");
        this.port.Open();
        TP.Base.Logger.Log("TP.Ports.SerialPort.setupPort: finished");
      }
      catch (System.IO.IOException ioex) {
        // The port is in an invalid state. - or - An attempt to set the state of the underlying port
        // failed. For example, the parameters passed from this System.IO.Ports.SerialPort object were
        // invalid.
        TP.Base.Logger.Log("--> IOException occurred opening RS232 port " + comPort + ": " + ioex.Message);
        if (ioex.Message.Contains("device attached to the system is not functioning")) {
          // usb unplug/plug fixes this
          throw new DeviceErrorException("Error starting device! Unplug the USB adapter, wait 5 seconds, plug it back in, wait another 5 seconds, then restart this program.", ioex);
        }
        else if (ioex.Message.Contains("does not exist")) {
          // incorrect config, or usb unplugged
          throw new PortMissingException("Error starting device! Either Brian messed up the configuration, or your USB adapter isn't plugged in!", ioex);
        }
        else {
          throw;
        }
      }
      catch (ArgumentOutOfRangeException aoorex) {
        // One or more of the properties for this instance are invalid.
        TP.Base.Logger.Log("--> ArgumentOutOfRangeException occurred opening RS232 port " + comPort + ": " + aoorex.Message);
        throw;
      }
      catch (UnauthorizedAccessException uaex) {
        // Access is denied to the port.
        TP.Base.Logger.Log("--> UnauthorizedAccessException occurred opening RS232 port " + comPort + ": " + uaex.Message);
        if (uaex.Message.Contains("is denied")) {
          throw new PortInUseException("Error starting device! You probably already have a program running that is using the device!", uaex);
        }
        else {
          throw;
        }
      }
      catch (InvalidOperationException ioex) {
        // The specified port is open.
        TP.Base.Logger.Log("--> InvalidOperationException occurred opening RS232 port " + comPort + ": " + ioex.Message);
        throw;
      }
      catch (ArgumentException aex) {
        // The port name does not begin with "COM". - or -The file type of the port is not supported.
        TP.Base.Logger.Log("--> ArgumentException occurred opening RS232 port " + comPort + ": " + aex.Message);
        throw;
      }
      catch (Exception ex) {
        // non-documented error?
        TP.Base.Logger.Log("--> Exception occurred opening RS232 port " + comPort + ": " + ex.Message);
        throw;
      }
    }
  }
}
