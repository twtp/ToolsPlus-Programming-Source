using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface {

  internal delegate void CIC332StatusReceivedEventHandler(object sender, StatusMessageEventArgs e);
  internal class StatusMessageEventArgs : EventArgs {
    private Message.Incoming.CarouselStatusReply message;
    public string CarouselAddress { get { return this.message.FromAddress; } }
    public CarouselMode CarouselMode { get { return this.message.CarouselMode; } }
    public CarouselStatus CarouselStatus { get { return this.message.CarouselStatus; } }
    public int BinCount { get { return this.message.CarouselLength; } }
    public int CurrentPosition { get { return this.message.CarouselPosition; } }

    public StatusMessageEventArgs(Message.Incoming.CarouselStatusReply csr) {
      this.message = csr;
    }
  }

  internal delegate void CIC332FaultsReceivedEventHandler(object sender, FaultsMessageEventArgs e);
  internal class FaultsMessageEventArgs : EventArgs {
    private Message.Incoming.CarouselFaultsReply message;
    public string CarouselAddress { get { return this.message.FromAddress; } }
    public bool CarouselError { get { return this.message.CarouselError; } }
    public bool MotorError { get { return this.message.MotorError; } }
    public KeypadState KeypadState { get { return this.message.KeypadState; } }
    public bool EmergencyStopEngaged { get { return this.message.EmergencyStopEngaged; } }
    public bool PhotoeyeRightError { get { return this.message.PhotoeyeRightError; } }
    public bool PhotoeyeLeftError { get { return this.message.PhotoeyeLeftError; } }
    public bool PhotoeyeAnyError { get { return this.message.PhotoeyeAnyError; } }

    public FaultsMessageEventArgs(Message.Incoming.CarouselFaultsReply cfr) {
      this.message = cfr;
    }
  }

  internal delegate void CIC332TextMessageReceivedEventHandler(object sender, TextMessageEventArgs e);
  internal class TextMessageEventArgs : EventArgs {
    private Message.Incoming.ASCIIText message;
    public string MessageText { get { return this.message.MessageText; } }
    public string FromLightTree { get { return this.message.FromAddress; } }

    public TextMessageEventArgs(Message.Incoming.ASCIIText t) {
      this.message = t;
    }
  }

  internal delegate void LEDIlluminatedEventHandler(object sender, LEDIlluminatedEventArgs e);
  internal class LEDIlluminatedEventArgs : EventArgs {
    private Message.Outgoing.LEDIlluminate message;
    public string LightTreeAddress { get { return this.message.ToAddress; } }
    public string CarouselAddress { get { return this.message.ForCarouselAddr; } }
    public int LightNumber { get { return this.message.LightNumber; } }
    public string Text { get { return this.message.Message; } }

    public LEDIlluminatedEventArgs(Message.Outgoing.LEDIlluminate ledi) {
      this.message = ledi;
    }
  }

  internal delegate void LEDClearedEventHandler(object sender, LEDClearedEventArgs e);
  internal class LEDClearedEventArgs : EventArgs {
    private string carouselAddress;
    public string CarouselAddress { get { return this.carouselAddress; } }
    private string lightTreeAddress;
    public string LightTreeAddress { get { return this.lightTreeAddress; } }
    private int? lightNumber;
    public int? LightNumber { get { return this.lightNumber; } }

    public LEDClearedEventArgs(Message.Outgoing.LEDClearAll ledca) {
      //TP.Base.Logger.Log("creating LEDClearedEventArgs object");
      this.carouselAddress = ledca.ForCarouselAddr;
      this.lightTreeAddress = ledca.ToAddress;
      this.lightNumber = null;
    }
    public LEDClearedEventArgs(Message.Outgoing.LEDClearSingle ledcs) {
      //TP.Base.Logger.Log("creating LEDClearedEventArgs object");
      this.carouselAddress = ledcs.ForCarouselAddr;
      this.lightTreeAddress = ledcs.ToAddress;
      this.lightNumber = ledcs.LightNumber;
    }
  }

}
