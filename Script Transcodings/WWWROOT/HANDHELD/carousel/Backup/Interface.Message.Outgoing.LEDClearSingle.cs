using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Outgoing {
  internal class LEDClearSingle : OutgoingMessage, ILightMessage {

    protected int lightNumber;
    public int LightNumber { get { return this.lightNumber; } }
    protected int numChars;
    // no getter
    protected string forCarouselAddr;
    public string ForCarouselAddr { get { return this.forCarouselAddr; } }

    public LEDClearSingle(string lightAddr, string lightSubchannel, int lightNumber, int numChars, string forCarouselAddr) {
      this.fromAddress = "000";
      this.fromSubchannel = "0";
      this.toAddress = lightAddr;
      this.toSubchannel = lightSubchannel;
      this.lightNumber = lightNumber;
      this.numChars = numChars;
      this.forCarouselAddr = forCarouselAddr;
    }

    public override string ToSendable() {
      return base.wrapMessage("{" + this.lightNumber + new string(' ', numChars) + "}");
    }

    public override string ToString() {
      return string.Format("Light {0} Clear Light {1} Request", this.toAddress, this.lightNumber);
    }
  }
}
