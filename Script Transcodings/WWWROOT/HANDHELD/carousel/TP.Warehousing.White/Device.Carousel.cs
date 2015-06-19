using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Device {
  internal class Carousel : BaseDevice, IEquatable<Carousel> {
    protected int numberOfBins;
    protected int? currentBinPosition;
    protected CarouselStatus currentStatus;

    public Carousel(int address, int numberOfBins) : this(address.ToString(), numberOfBins) { }
    public Carousel(string address, int numberOfBins) : base(address, "3") {
      this.numberOfBins = numberOfBins;
      this.currentBinPosition = null;
      this.currentStatus = CarouselStatus.Unknown;

      Interface.CIC332.StatusReceived += new Interface.CIC332StatusReceivedEventHandler(this.carousel_CIC332StatusReceivedEventHandler);
    }

    public bool MoveToBin(int binNumber) {
      if (binNumber > numberOfBins) {
        throw new ArgumentException("invalid bin number");
      }
      Interface.Message.Outgoing.CarouselMove cm = new Interface.Message.Outgoing.CarouselMove(this.address, this.subchannel, binNumber.ToString());
      return cm.Send();
    }

    public bool Halt() {
      TP.Base.Logger.Log("Device.Carousel.Halt(): queueing halt message for carousel " + this.address);
      Interface.Message.Outgoing.CarouselHalt ch = new Interface.Message.Outgoing.CarouselHalt(this.address, this.subchannel);
      return ch.Send();
    }

    public bool AskForStatus() {
      TP.Base.Logger.Log("Device.Carousel.AskForStatus(): queueing status message for carousel " + this.address);
      Interface.Message.Outgoing.CarouselStatusQuery csq = new Interface.Message.Outgoing.CarouselStatusQuery(this.address, this.subchannel);
      return csq.Send();
    }

    public bool AskForFaults() {
      TP.Base.Logger.Log("Device.Carousel.AskForFaults(): queueing faults message for carousel " + this.address);
      Interface.Message.Outgoing.CarouselFaultsQuery cfq = new Interface.Message.Outgoing.CarouselFaultsQuery(this.address, this.subchannel);
      return cfq.Send();
    }

    public int? CurrentBinPosition { get { return this.currentBinPosition; } }
    public CarouselStatus CurrentStatus { get { return this.currentStatus; } }

    public override DeviceType DeviceType { get { return DeviceType.Carousel; } }

    public override bool Initialize() {
      if (this.AskForStatus()) {
        // status receipt should complete initialization (homing if required), controlled by pod
        return true;
      }
      return false;
    }

    private void carousel_CIC332StatusReceivedEventHandler(object sender, Interface.StatusMessageEventArgs e) {
      if (e.CarouselAddress == this.Address) {
        this.currentBinPosition = e.CurrentPosition;
        this.currentStatus = e.CarouselStatus;
        TP.Base.Logger.Log("Device.Carousel.carousel_SSReceived(): carousel " + this.address + " is " + this.currentStatus.ToString() + ", at bin " + this.currentBinPosition.ToString());
      }
    }

    #region IEquatable<Carousel> Members

    public override bool Equals(object obj) {
      return base.Equals(obj);
    }

    public override int GetHashCode() {
      return base.GetHashCode();
    }

    public bool Equals(Carousel other) {
      if ((object)other == null) {
        return false;
      }
      return this.Address == other.Address;
    }

    public static bool operator ==(Carousel c1, Carousel c2) {
      if ((object)c1 == null) {
        return (object)c2 == null;
      }
      return c1.Equals(c2);
    }
    public static bool operator !=(Carousel c1, Carousel c2) {
      if ((object)c1 == null) {
        return (object)c2 != null;
      }
      return !c1.Equals(c2);
    }

    #endregion
  }

}
