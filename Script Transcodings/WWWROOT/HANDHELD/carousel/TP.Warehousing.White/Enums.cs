using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White {

  public enum DeviceType {
    Carousel,
    LightTree,
    SortBar,
    CharacterStrip,
    SurfaceDisplay,
    Operator, // only exists for the blockFlags
  };

  public enum CarouselMode {
    Automatic,
    Manual,
    Unknown,
  };

  public enum CarouselStatus {
    Success,
    Failure,
    InProgress,
    Braking,
    Uninitialized,
    Unknown,
  };

  public enum KeypadState {
    Manual,
    Host,
    Unknown, // unused, source is a bit field
  };

  public enum CarouselOrientation {
    Horizontal,
    Vertical,
  };

}
