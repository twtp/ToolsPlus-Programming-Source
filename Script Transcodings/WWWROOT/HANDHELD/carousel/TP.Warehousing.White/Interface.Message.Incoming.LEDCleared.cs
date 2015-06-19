using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Incoming {
  internal class LEDCleared : IncomingMessage {

    protected int? lightNumber;
    public int? LightNumber { get { return this.lightNumber; } }
    protected string forCarouselAddr;
    public string ForCarouselAddr { get { return this.forCarouselAddr; } }

    public LEDCleared(string lightAddr, string lightSubchannel, string toAddress, string toSubchannel, int? lightNumber, string forCarouselAddr) {
      this.fromAddress = lightAddr;
      this.fromSubchannel = lightSubchannel;
      this.toAddress = toAddress;
      this.toSubchannel = toSubchannel;
      this.lightNumber = lightNumber;
      this.forCarouselAddr = forCarouselAddr;
    }

    public LEDCleared Parse(Outgoing.LEDClearAll outgoingMessage) {
      return new LEDCleared(
          outgoingMessage.ToAddress,
          outgoingMessage.ToSubchannel,
          outgoingMessage.FromAddress,
          outgoingMessage.FromSubchannel,
          null,
          outgoingMessage.ForCarouselAddr
        );
    }

    public LEDCleared Parse(Outgoing.LEDClearSingle outgoingMessage) {
      return new LEDCleared(
          outgoingMessage.ToAddress,
          outgoingMessage.ToSubchannel,
          outgoingMessage.FromAddress,
          outgoingMessage.FromSubchannel,
          outgoingMessage.LightNumber,
          outgoingMessage.ForCarouselAddr
        );
    }

  }
}
