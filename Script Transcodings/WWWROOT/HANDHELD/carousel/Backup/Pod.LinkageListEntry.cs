using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Pod {
  internal class LinkageListEntry {
    private Device.Carousel carousel;
    private Device.LightDevice lightDevice;
    private Device.LightTreeDirections direction;

    public LinkageListEntry(Device.Carousel carousel, Device.LightDevice lightDevice, Device.LightTreeDirections direction) {
      this.carousel = carousel;
      this.lightDevice = lightDevice;
      this.direction = direction;
    }

    public Device.Carousel Carousel { get { return this.carousel; } }
    public Device.LightDevice LightDevice { get { return this.lightDevice; } }
    public Device.LightTreeDirections Direction { get { return this.direction; } }

  }
}
