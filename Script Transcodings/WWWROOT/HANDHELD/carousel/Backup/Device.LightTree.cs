using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class LightTree : LightDevice, ILightsQuantities {

    public LightTree(int address, int numberOfDisplays, int charsPerDisplay) : this(address.ToString(), numberOfDisplays, charsPerDisplay) { }
    public LightTree(string address, int numberOfDisplays, int charsPerDisplay) : base(address, "1", numberOfDisplays, charsPerDisplay) {
      // ?
    }

    public bool LightQuantity(string div1, string div2, int quantity, LightTreeDirections direction, string forCarouselAddr) {
      if (!this.Display(7, new string(DirectionToArrow(direction), this.charsPerDisplay), forCarouselAddr)) {
        return false;
      }
      if (!this.Display(6, div1 + " " + div2, forCarouselAddr)) {
        return false;
      }
      if (!this.Display(5, quantity.ToString(), forCarouselAddr)) {
        return false;
      }
      if (!this.Display(4, new string(DirectionToArrow(direction), this.charsPerDisplay), forCarouselAddr)) {
        return false;
      }
      return true;
    }

    public override bool Clear(string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearAll ledca = new Interface.Message.Outgoing.LEDClearAll(this.address, this.subchannel, Interface.LEDClearStyle.LightTree, forCarouselAddr);
      return ledca.Send();
    }
    public override bool Clear(int index, string forCarouselAddr) {
      if (index > this.numberOfDisplays) {
        throw new ArgumentException("invalid light number");
      }
      Interface.Message.Outgoing.LEDClearSingle ledcs = new Interface.Message.Outgoing.LEDClearSingle(this.address, this.subchannel, index, this.charsPerDisplay, forCarouselAddr);
      return ledcs.Send();
    }

    private static char DirectionToArrow(LightTreeDirections direction) {
      switch (direction) {
        case LightTreeDirections.Left:
          return '<';
        case LightTreeDirections.Right:
          return '>';
      }
      return ' ';
    }

    public static LightTreeDirections ArrowToDirection(string message) {
      if (message.StartsWith("<")) {
        return LightTreeDirections.Left;
      }
      else if (message.EndsWith(">")) {
        return LightTreeDirections.Right;
      }
      else {
        return LightTreeDirections.Unknown;
      }
    }

    public override DeviceType DeviceType { get { return DeviceType.LightTree; } }

    public override bool Initialize() {
      return this.Clear(null);
    }
  }
}
