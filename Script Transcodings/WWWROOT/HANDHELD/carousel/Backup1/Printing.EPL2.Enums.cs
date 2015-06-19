using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Printing.EPL2 {
  
  public enum EPL2PaperSize {
    ShippingLabel,
    BarcodeLabel,
  };

  public enum EPL2Rotation {
    Standard = 0,
    Clockwise = 1,
    UpsideDown = 2,
    CounterClockwise = 3,
  };

  public enum EPL2Font {
    CPI20_08x12 = 1,
    CPI17_10x16 = 2,
    CPI15_12x20 = 3,
    CPI13_14x24 = 4,
    CPI06_32x48 = 5,
  };

}
