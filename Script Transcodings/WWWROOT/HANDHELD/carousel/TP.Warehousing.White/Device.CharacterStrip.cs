using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class CharacterStrip : LightDevice, ILightsQuantities {

    public CharacterStrip(int address, int numberOfDisplays, int charsPerDisplay) : this(address.ToString(), numberOfDisplays, charsPerDisplay) { }
    public CharacterStrip(string address, int numberOfDisplays, int charsPerDisplay) : base(address, "1", numberOfDisplays, charsPerDisplay) {
      // nothing?
    }

    public bool Display(int index, char message, string forCarouselAddr) {
      // omnicell strips are indexed from 1, wtf?
      return base.Display(index + 1, message.ToString(), forCarouselAddr);
    }

    public bool LightQuantity(string div1, string div2, int quantity, LightTreeDirections direction, string forCarouselAddr) {
      // arrow pointing to div1's bin
      // this is pretty carousel-layout dependent, we light 2 lights for each bin
      // bin number (1-indexed) doubled (because 2 lights per bin), subtract 2 (that gives you the end, go back to beginning)
      // BIN BIN BIN
      // 0 1 2 3 4 5
      int first;
      if (Int32.TryParse(div1, out first)) {
        first *= 2;
        first -= 2;
        return this.Display(first, '^', forCarouselAddr) && this.Display(first + 1, '^', forCarouselAddr);
      }
      return true; // true?
    }

    public override bool Clear(string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearAll ledca = new Interface.Message.Outgoing.LEDClearAll(this.address, this.subchannel, Interface.LEDClearStyle.OmniCell, forCarouselAddr);
      return ledca.Send();
    }
    public override bool Clear(int index, string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearSingle ledcs = new Interface.Message.Outgoing.LEDClearSingle(this.address, this.subchannel, index, this.charsPerDisplay, forCarouselAddr);
      return ledcs.Send();
    }

    public override DeviceType DeviceType { get { return DeviceType.CharacterStrip; } }

    public override bool Initialize() {
      return this.Clear(null);
    }
  }
}
