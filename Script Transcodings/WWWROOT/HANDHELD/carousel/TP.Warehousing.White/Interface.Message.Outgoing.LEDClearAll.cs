using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Outgoing {
  internal class LEDClearAll : OutgoingMessage, ILightMessage {

    protected LEDClearStyle ledClearStyle;
    protected string forCarouselAddr;
    public string ForCarouselAddr { get { return this.forCarouselAddr; } }

    public LEDClearAll(string lightAddr, string lightSubchannel, LEDClearStyle ledClearStyle, string forCarouselAddr) {
      this.fromAddress = "000";
      this.fromSubchannel = "0";
      this.toAddress = lightAddr;
      this.toSubchannel = lightSubchannel;
      this.ledClearStyle = ledClearStyle;
      this.forCarouselAddr = forCarouselAddr;
    }

    public override string ToSendable() {
      switch (this.ledClearStyle) {
        case LEDClearStyle.LightTree:
          return base.wrapMessage("{W}");
        case LEDClearStyle.OmniCell:
          return base.wrapMessage("{cl}");
        default:
          throw new ApplicationException("missing LEDClearStyle in switch");
      }
    }

    public override string ToString() {
      return string.Format("Light {0} Clear All Request", this.toAddress);
    }
  }
}
