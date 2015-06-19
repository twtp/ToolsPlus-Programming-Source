using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Incoming {
  internal class LEDActivated : IncomingMessage {

    protected int lightNumber;
    public int LightNumber { get { return this.lightNumber; } }
    protected string message;
    public string Message { get { return this.message; } }
    protected string forCarouselAddr;
    public string ForCarouselAddr { get { return this.forCarouselAddr; } }

    protected LEDActivated(string lightAddr, string lightSubchannel, string toAddress, string toSubchannel, int lightNumber, string message, string forCarouselAddr) {
      this.fromAddress = lightAddr;
      this.fromSubchannel = lightSubchannel;
      this.toAddress = toAddress;
      this.toSubchannel = toSubchannel;
      this.lightNumber = lightNumber;
      this.message = message;
      this.forCarouselAddr = forCarouselAddr;
    }

    public LEDActivated Parse(Outgoing.LEDIlluminate outgoingMessage) {
      return new LEDActivated(
          outgoingMessage.ToAddress,
          outgoingMessage.ToSubchannel,
          outgoingMessage.FromAddress,
          outgoingMessage.FromSubchannel,
          outgoingMessage.LightNumber,
          outgoingMessage.Message,
          outgoingMessage.ForCarouselAddr
        );
    }

  }
}
