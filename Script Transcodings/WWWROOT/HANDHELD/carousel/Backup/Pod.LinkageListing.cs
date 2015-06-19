using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Pod {
  internal class LinkageListing {
    public LinkageListing() {
      this.linkages = new List<LinkageListEntry>();
    }

    private List<LinkageListEntry> linkages;

    /*public bool AddLinkage(Device.Carousel carousel, Device.LightDevice lightDevice, string direction) {*/
    public bool AddLinkage(Device.Carousel carousel, Device.LightDevice lightDevice, TP.Base.ConfigParam.Direction direction) {
      Device.LightTreeDirections dir;
      switch (direction) {
        /*case "left":
          dir = Device.LightTreeDirections.Left;
          break;
        case "right":
          dir = Device.LightTreeDirections.Right;
          break;
        default:
          dir = Device.LightTreeDirections.Unknown;
          break;*/
        case TP.Base.ConfigParam.Direction.Left:
          dir = Device.LightTreeDirections.Left;
          break;
        case TP.Base.ConfigParam.Direction.Right:
          dir = Device.LightTreeDirections.Right;
          break;
        case TP.Base.ConfigParam.Direction.Unknown:
          dir = Device.LightTreeDirections.Unknown;
          break;
        default:
          throw new ApplicationException("oops");
      }
      return this.AddLinkage(carousel, lightDevice, dir);
    }
    public bool AddLinkage(Device.Carousel carousel, Device.LightDevice lightDevice, Device.LightTreeDirections direction) {
      foreach (LinkageListEntry lle in this.linkages) {
        if (lle.Carousel == carousel && lle.Direction == direction) {
          // conflict, can't have same carousel and direction pointing to a different light
          return false;
        }
        if (lle.LightDevice == lightDevice && lle.Direction == direction) {
          // conflict, can't have a light and direction pointing to multiple carousels
          return false;
        }
      }
      // no conflicts, yay
      this.linkages.Add(new LinkageListEntry(carousel, lightDevice, direction));
      return true;
    }

    public Device.Carousel GetCarousel(Device.LightDevice lightDevice, Device.LightTreeDirections direction) {
      foreach (LinkageListEntry lle in this.linkages) {
        if (lle.LightDevice == lightDevice && lle.Direction == direction) {
          return lle.Carousel;
        }
      }
      return null;
    }

    public Device.LightDevice GetLight(Device.Carousel carousel, Device.LightTreeDirections direction) {
      foreach (LinkageListEntry lle in this.linkages) {
        if (lle.Carousel == carousel && lle.Direction == direction) {
          return lle.LightDevice;
        }
      }
      return null;
    }

    public List<LinkageListEntry> GetLights(Device.Carousel carousel) {
      List<LinkageListEntry> retval = new List<LinkageListEntry>();
      foreach (LinkageListEntry lle in this.linkages) {
        if (lle.Carousel == carousel) {
          retval.Add(lle);
        }
      }
      return retval;
    }

    public List<LinkageListEntry> GetLights(Device.Carousel carousel, DeviceType deviceType) {
      List<LinkageListEntry> retval = new List<LinkageListEntry>();
      foreach (LinkageListEntry lle in this.linkages) {
        if (lle.Carousel == carousel && lle.LightDevice.DeviceType == deviceType) {
          retval.Add(lle);
        }
      }
      return retval;
    }
  }
}
