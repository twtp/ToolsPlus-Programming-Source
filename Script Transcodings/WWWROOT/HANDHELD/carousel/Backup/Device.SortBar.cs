using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class SortBar : LightDevice {

    public SortBar(int address, int numberOfDisplays, int charsPerDisplay) : this(address.ToString(), numberOfDisplays, charsPerDisplay) { }
    public SortBar(string address, int numberOfDisplays, int charsPerDisplay) : base(address, "1", numberOfDisplays, charsPerDisplay) {
      // nothing?
    }

    public bool Arrow(int index, int pickQuantity, SortBarDirections direction) {
      // using "v" and "^" for arrows looks dumb, just do the quantities w/o arrow
      if (pickQuantity.ToString().Length > this.charsPerDisplay) {
        throw new ArgumentException("holy crap, we actually got an order for this many? TODO!");
      }
      string msg = pickQuantity.ToString();
      if (msg.Length % 2 == 1) {
        msg = msg + ' ';
      }
      while (msg.Length < this.charsPerDisplay) {
        msg = ' ' + msg + ' ';
      }
      return this.Display(index, msg, null);
    }
    public bool Arrow(int index, SortBarDirections direction) {
      return this.Display(index, new String(DirectionToArrow(direction), this.charsPerDisplay), null);
    }
    public bool Arrow(int index, int pickQuantity) {
      return Arrow(index, pickQuantity, 0);
    }
    public bool Arrow(int index) {
      return Arrow(index, 0);
    }

    public override bool Clear(string forCarouselAddr) {
      Interface.Message.Outgoing.LEDClearAll ledca = new Interface.Message.Outgoing.LEDClearAll(this.address, this.subchannel, Interface.LEDClearStyle.LightTree, null);
      return ledca.Send();
    }
    public override bool Clear(int index, string forCarouselAddr) {
      if (index > this.numberOfDisplays) {
        throw new ArgumentException("invalid light number");
      }
      Interface.Message.Outgoing.LEDClearSingle ledcs = new Interface.Message.Outgoing.LEDClearSingle(this.address, this.subchannel, index, this.charsPerDisplay, null);
      return ledcs.Send();
    }

    public new bool Display(int index, string message, string forCarouselAddr) {
      return base.Display(index, message, forCarouselAddr);
    }

    private char DirectionToArrow(SortBarDirections direction) {
      switch (direction) {
        case SortBarDirections.Down:
          return 'v';
        case SortBarDirections.Up:
          return '^';
      }
      return ' ';
    }

    public override DeviceType DeviceType { get { return DeviceType.SortBar; } }

    public override bool Initialize() {
      return this.Clear(null);
    }
  }
}
