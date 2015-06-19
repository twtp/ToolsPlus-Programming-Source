using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal abstract class LightDevice : BaseDevice {
    protected int numberOfDisplays;
    protected int charsPerDisplay;

    //
    // LightTreeDirections was originally intended to be from the POV of the
    // carousel, so left means the lightree is to the left of the carousel.
    // it gets reused here, so the directions are backwards!
    //

    protected struct CarouselLinkage {
      public Carousel Carousel;
      public LightTreeDirections Direction;
      public CarouselLinkage(Carousel carousel, LightTreeDirections direction) {
        this.Carousel = carousel;
        this.Direction = direction;
      }
    }

    protected List<CarouselLinkage> carousels;

    protected LightDevice(string address, string subchannel, int numberOfDisplays, int charsPerDisplay) : base(address, subchannel) {
      this.numberOfDisplays = numberOfDisplays;
      this.charsPerDisplay = charsPerDisplay;
      this.carousels = new List<CarouselLinkage>();
    }

    protected bool AttachCarousel(Carousel carousel) {
      return this.AttachCarousel(carousel, LightTreeDirections.Unknown);
    }
    protected bool AttachCarousel(Carousel carousel, LightTreeDirections direction) {
      this.carousels.Add(new CarouselLinkage(carousel, direction));
      return true;
    }

    protected virtual bool Display(int index, string message, string forCarouselAddr) {
      if (index > this.numberOfDisplays) {
        throw new ArgumentException("invalid light number");
      }
      if (message.Length > this.charsPerDisplay) {
        message = message.Substring(0, this.charsPerDisplay);
      }
      if (message.Length < this.charsPerDisplay) {
        message = message.PadRight(this.charsPerDisplay, ' ');
      }

      Interface.Message.Outgoing.LEDIlluminate ledi = new Interface.Message.Outgoing.LEDIlluminate(this.address, this.subchannel, index, message, forCarouselAddr);
      return ledi.Send();
    }

    public abstract bool Clear(string forCarouselAddr);
    public abstract bool Clear(int index, string forCarouselAddr);

    public bool DisplayArbitraryMessage(int lightIndex, string message, string forCarouselAddr) {
      return this.Display(lightIndex, message, forCarouselAddr);
    }

    public int NumberOfDisplays { get { return this.numberOfDisplays; } }
  }
}
