using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Outgoing {
  internal class CarouselMove : OutgoingMessage {

    protected string targetBin;

    public CarouselMove(string carouselAddr, string carouselSubchannel, string targetBin) {
      this.fromAddress = "000";
      this.fromSubchannel = "0";
      this.toAddress = carouselAddr;
      this.toSubchannel = carouselSubchannel;
      this.targetBin = targetBin;
    }

    public override string ToSendable() {
      return base.wrapMessage("G" + this.targetBin.PadLeft(3, '0'));
    }

    public override string ToString() {
      return string.Format("Carousel {0} Move to {1} Request", this.toAddress, this.targetBin);
    }
  }
}
