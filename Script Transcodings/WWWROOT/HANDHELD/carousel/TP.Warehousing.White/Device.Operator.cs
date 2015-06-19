using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  // ick, operator doesn't meet most of the BaseDevice criteria, just needs the block flags
  // maybe the block flags should be an interface that everything implements?
  internal class Operator : BaseDevice {
    public Operator() : base("0000", "0") {
      // nothing?
    }

    public override DeviceType DeviceType { get { return DeviceType.Operator; } }

    public override bool Initialize() {
      // nothing
      return true;
    }
  }
}
