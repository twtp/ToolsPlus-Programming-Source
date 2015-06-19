using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Outgoing {
  internal class CarouselStatusQuery : OutgoingMessage {

    public CarouselStatusQuery(string carouselAddr, string carouselSubchannel) {
      this.fromAddress = "000";
      this.fromSubchannel = "0";
      this.toAddress = carouselAddr;
      this.toSubchannel = carouselSubchannel;
    }

    public override string ToSendable() {
      return base.wrapMessage("SS");
    }

    public override string ToString() {
      return string.Format("Carousel {0} Status Request", this.toAddress);
    }
  }
}
