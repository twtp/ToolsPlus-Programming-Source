using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class SurfaceDisplay : LightDevice, ILightsQuantities {

    public SurfaceDisplay(int address, int numberOfDisplays, int charsPerDisplay) : this(address.ToString(), numberOfDisplays, charsPerDisplay) { }
    public SurfaceDisplay(string address, int numberOfDisplays, int charsPerDisplay) : base(address, "2", numberOfDisplays, charsPerDisplay) {
      // nothing?
    }

    public bool Display(string message, string forCarouselAddr) {
      return base.Display(0, message, forCarouselAddr);
    }

    public bool LightQuantity(string div1, string div2, int quantity, LightTreeDirections direction, string forCarouselAddr) {
      return this.Display(div1 + " " + div2 + " => " + quantity.ToString(), forCarouselAddr);
    }

    public override bool Clear(string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearAll ledca = new Interface.Message.Outgoing.LEDClearAll(this.address, this.subchannel, Interface.LEDClearStyle.LightTree, forCarouselAddr);
      return ledca.Send();
    }
    public override bool Clear(int index, string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearSingle ledcs = new Interface.Message.Outgoing.LEDClearSingle(this.address, this.subchannel, index, this.charsPerDisplay, forCarouselAddr);
      return ledcs.Send();
    }

    public override DeviceType DeviceType { get { return DeviceType.SurfaceDisplay; } }

    public override bool Initialize() {
      return this.Clear(null);
    }
  }
}

