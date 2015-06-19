using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class DeviceDefinition {
    private DeviceType device;
    private string address;
    private LockType lockType;
    private int coopID;
    private bool coopStartpoint;
    private bool coopEndpoint;

    public DeviceDefinition(DeviceType device, string address, LockType lockType) {
      this.device = device;
      this.address = Interface.CIC332.FormatAddress(address);
      this.lockType = lockType;
    }
    public DeviceDefinition(DeviceType device, string address, LockType lockType, int coopID, bool coopStartpoint, bool coopEndpoint) : this(device, address, lockType) {
      if (LockType == LockType.Cooperative) {
        this.coopID = coopID;
        this.coopStartpoint = coopStartpoint;
        this.coopEndpoint = coopEndpoint;
      }
      else {
        throw new ArgumentException("argument coopID is invalid for this lockType");
      }
    }

    public DeviceType DeviceType { get { return this.device; } }
    public string Address { get { return this.address; } }
    public LockType LockType { get { return this.lockType; } }
    public int CooperativeActionID { get { return this.coopID; } }
    public bool CooperativeStartpoint { get { return this.coopStartpoint; } }
    public bool CooperativeEndpoint { get { return this.coopEndpoint; } }
  }
}
