using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {

  internal interface ILightsQuantities  {
    bool LightQuantity(string div1, string div2, int quantity, LightTreeDirections direction, string forCarouselAddr);
  }

}
