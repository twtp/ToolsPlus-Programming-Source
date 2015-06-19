using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Pod {
  public delegate void MotorFaultEventHandler(object sender, CarouselFaultEventArgs e);
  public delegate void PhotoEyeFaultEventHandler(object sender, CarouselFaultEventArgs e);
  public delegate void KeypadModeFaultEventHandler(object sender, CarouselFaultEventArgs e);
  public delegate void EmergencyStopEventHandler(object sender, CarouselFaultEventArgs e);
  public delegate void CarouselOffEventHandler(object sender, CarouselFaultEventArgs e);
  public delegate void NoFaultsEventHandler(object sender, CarouselFaultEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class CarouselFaultEventArgs : EventArgs {
    private string carouselAddr;

    internal CarouselFaultEventArgs(string carouselAddr) {
      this.carouselAddr = carouselAddr;
    }

    public string CarouselAddress {
      get { return this.carouselAddr; }
    }
  }


  public delegate void CarouselStatusUpdateEventHandler(object sender, CarouselStatusUpdateEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class CarouselStatusUpdateEventArgs : EventArgs {
    private string carouselAddr;
    private CarouselMode mode;
    private int currentPos;
    private CarouselStatus status;

    internal CarouselStatusUpdateEventArgs(string carouselAddr, CarouselMode mode, int currentPos, CarouselStatus status) {
      this.carouselAddr = carouselAddr;
      this.mode = mode;
      this.currentPos = currentPos;
      this.status = status;
    }

    public string CarouselAddress {
      get { return this.carouselAddr; }
    }
    public CarouselMode CarouselMode {
      get { return this.mode; }
    }
    public int CurrentPosition {
      get { return this.currentPos; }
    }
    public CarouselStatus CurrentStatus {
      get { return this.status; }
    }
  }


  public delegate void LightActivationEventHandler(object sender, LightActivationEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class LightActivationEventArgs : EventArgs {
    private string associatedCarouselAddr;
    private string lightTreeAddress;
    private int lightNumber;
    private string text;

    internal LightActivationEventArgs(Interface.LEDIlluminatedEventArgs subevent) {
      this.associatedCarouselAddr = subevent.CarouselAddress;
      this.lightTreeAddress = subevent.LightTreeAddress;
      this.lightNumber = subevent.LightNumber;
      this.text = subevent.Text;
    }

    public string AssociatedCarouselAddr { get { return this.associatedCarouselAddr; } }
    public string LightTreeAddress { get { return this.lightTreeAddress; } }
    public int LightNumber { get { return this.lightNumber; } }
    public string Text { get { return this.text; } }
  }

  public delegate void LightDeactivationEventHandler(object sender, LightDeactivationEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class LightDeactivationEventArgs : EventArgs {
    private string associatedCarouselAddr;
    private string lightTreeAddress;
    private int? lightNumber;

    internal LightDeactivationEventArgs(Interface.LEDClearedEventArgs subevent) {
      //TP.Base.Logger.Log("creating LightDeactivationEventArgs object");
      this.associatedCarouselAddr = subevent.CarouselAddress;
      this.lightTreeAddress = subevent.LightTreeAddress;
      this.lightNumber = subevent.LightNumber;
    }

    public string AssociatedCarouselAddr { get { return this.associatedCarouselAddr; } }
    public string LightTreeAddress { get { return this.lightTreeAddress; } }
    public int? LightNumber { get { return this.lightNumber; } }
  }


  public delegate void LightreeButtonEventHandler(object sender, TextMessageEventArgs e);
  public delegate void OtherTextMessageEventHandler(object sender, TextMessageEventArgs e);

  [System.Diagnostics.DebuggerStepThroughAttribute()]
  public class TextMessageEventArgs : EventArgs {
    private string lightreeAddr;

    internal TextMessageEventArgs(string lightreeAddr) {
      this.lightreeAddr = lightreeAddr;
    }

    public string LightreeAddress {
      get { return this.lightreeAddr; }
    }
  }
}
