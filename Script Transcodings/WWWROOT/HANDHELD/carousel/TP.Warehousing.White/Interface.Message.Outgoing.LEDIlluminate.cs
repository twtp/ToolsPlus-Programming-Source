using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Outgoing {
  internal class LEDIlluminate : OutgoingMessage, ILightMessage {

    private int lightNumber;
    public int LightNumber { get { return this.lightNumber; } }
    private string message;
    public string Message { get { return this.message; } }
    private string forCarouselAddr;
    public string ForCarouselAddr { get { return this.forCarouselAddr; } }

    public LEDIlluminate(string lightAddr, string lightSubchannel, int lightNumber, string message, string forCarouselAddr) {
      this.fromAddress = "000";
      this.fromSubchannel = "0";
      this.toAddress = lightAddr;
      this.toSubchannel = lightSubchannel;
      this.lightNumber = lightNumber;
      this.message = message;
      this.forCarouselAddr = forCarouselAddr;
    }

    public override string ToSendable() {
      return base.wrapMessage("{" + this.lightNumber + this.message + "}");
    }

    public override string ToString() {
      return string.Format("Light {0} Illuminate {1} With '{2}' Request", this.toAddress, this.lightNumber, this.message);
    }
    
  }
}
