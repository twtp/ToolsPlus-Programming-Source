using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal abstract class BaseDevice {
    protected string address;
    protected string subchannel;
    protected BlockInfo blockInfo;

    protected class BlockInfo {
      private bool isBlocked = false;
      private int blockingAction = -1;

      public BlockInfo() {
        // nothing?
      }

      public bool IsBlocked { get { return this.isBlocked; } set { this.isBlocked = value; } }
      public int BlockingAction { get { return this.blockingAction; } set { this.blockingAction = value; } }
    }

    protected BaseDevice(string address, string subchannel) {
      this.address = Interface.CIC332.FormatAddress(address);
      this.subchannel = subchannel;
      this.blockInfo = new BlockInfo();
    }

    public string Address { get { return this.address; } }
    public string Subchannel { get { return this.subchannel; } }

    public static string FormatAddress(string address) {
      return Interface.CIC332.FormatAddress(address);
    }

    public bool IsLocked { get { return this.blockInfo.IsBlocked; } }
    public int LockedByActionID { get { return this.blockInfo.BlockingAction; } }
    
    public bool LockDevice(int byActionID) {
      lock (this.blockInfo) {
        if (this.blockInfo.IsBlocked) {
          TP.Base.Logger.Log("FAILURE ATTEMPTING TO lock device " + this.DeviceType.ToString() + " address " + this.Address);
          return false;
        }
        else {
          this.blockInfo.IsBlocked = true;
          this.blockInfo.BlockingAction = byActionID;
          TP.Base.Logger.Log("locking device " + this.DeviceType.ToString() + " address " + this.Address + " for " + byActionID.ToString());
          return true;
        }
      }
    }

    public bool UnlockDevice() {
      lock (this.blockInfo) {
        if (this.blockInfo.IsBlocked) {
          this.blockInfo.IsBlocked = false;
          this.blockInfo.BlockingAction = -1;
          TP.Base.Logger.Log("unlocking device " + this.DeviceType.ToString() + " address " + this.Address);
          return true;
        }
        else {
          TP.Base.Logger.Log("DEVICE " + this.DeviceType.ToString() + " address " + this.Address + " IS NOT LOCKED, NO UNLOCK REQUIRED");
          return true;
        }
      }
    }

    public abstract DeviceType DeviceType { get; }
    public abstract bool Initialize();

  }
}
