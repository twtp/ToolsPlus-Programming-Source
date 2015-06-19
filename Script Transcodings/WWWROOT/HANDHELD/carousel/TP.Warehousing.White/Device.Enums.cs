using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {

  internal enum LockType {
    Exclusive,
    //OpenOnly,
    //CloseOnly,
    Cooperative,
  };

  internal enum LightTreeDirections {
    Left,
    Right,
    Unknown,
  };

  internal enum SortBarDirections {
    Down,
    Up,
    Unknown,
  };

}
