using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message {
  internal abstract class Base {

    protected const string STX_CHAR = "\x02";
    protected const string ETX_CHAR = "\x03";

    protected string fromAddress;
    public string FromAddress { get { return this.fromAddress; } }
    protected string fromSubchannel;
    public string FromSubchannel { get { return this.fromSubchannel; } }
    protected string toAddress;
    public string ToAddress { get { return this.toAddress; } }
    protected string toSubchannel;
    public string ToSubchannel { get { return this.toSubchannel; } }

    public string wrapMessage(string interior) {
      string msg = STX_CHAR + this.toAddress + this.toSubchannel + this.fromAddress + this.fromSubchannel + interior + ETX_CHAR;
      msg += calculateChecksum(msg);
      return msg;
    }

    // checks the message body against the given checksum
    protected static bool checksumIsValid(string message, string letter) {
      return letter == calculateChecksum(message.Substring(0, message.Length - 1));
    }

    // checksum is the sum of the ASCII values of the message, from STX to ETX inclusive.
    // this is then mod-divided by 128, bitwise and'ed with 127, and bitwise or'ed with 64
    protected static string calculateChecksum(string message) {
      int total = 0;
      foreach (byte c in System.Text.Encoding.ASCII.GetBytes(message)) {
        total += c;
      }
      total = (((total % 128) & 127) | 64);
      return Convert.ToChar(total).ToString();
    }
  }

  internal abstract class OutgoingMessage : Base {
    public abstract string ToSendable();

    public bool Send() {
      return Interface.CIC332.SendCommand(this);
    }
  }

  internal abstract class IncomingMessage : Base {
    // nothing here yet?
    // each should implement a factory method -- public *Message Parse(string incoming) -- or similar
  }

  internal interface ILightMessage {
    // TODO
  }
}
