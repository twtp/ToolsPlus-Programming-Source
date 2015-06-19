using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Incoming {
  internal class CarouselStatusReply : IncomingMessage {

    // message format is <stx>0000ccc3mrlllbbb<etx>s
    //    - to address and subchannel always "0000"
    //    - sender address "ccc"
    //    - sender subchannel always "3"
    //    - carousel mode "m"
    //       - "A" means automatic
    //       - "M" means manual
    //    - carousel status "r"
    //       - "Y" means success
    //       - "F" means failure
    //       - "I" means in-progress
    //       - "B" means braking
    //       - "U" means uninitialized
    //    - carousel bin count "lll"
    //    - carousel's current position "bbb"
    //    - checksum "s"
    protected static System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(
        @"^" + STX_CHAR +
        @"(?<toaddr>\d{3})" +
        @"(?<tosub>\d{1})" +
        @"(?<fromaddr>\d{3})" +
        @"(?<fromsub>\d{1})" +
        @"(?<mode>[AM])" +
        @"(?<status>[YFIBU])" +
        @"(?<count>\d{3})" +
        @"(?<pos>\d{3})" +
        ETX_CHAR + @"(?<checksum>.)$"
      );

    protected CarouselMode carouselMode;
    public CarouselMode CarouselMode { get { return this.carouselMode; } }
    protected CarouselStatus carouselStatus;
    public CarouselStatus CarouselStatus { get { return this.carouselStatus; } }
    protected int carouselLength;
    public int CarouselLength { get { return this.carouselLength; } }
    protected int carouselPosition;
    public int CarouselPosition { get { return this.carouselPosition; } }

    protected CarouselStatusReply(string fromAddress, string fromSubchannel, string toAddress, string toSubchannel, CarouselMode carouselMode, CarouselStatus carouselStatus, int carouselLength, int carouselPosition) {
      this.fromAddress = fromAddress;
      this.fromSubchannel = fromSubchannel;
      this.toAddress = toAddress;
      this.toSubchannel = toSubchannel;
      this.carouselMode = carouselMode;
      this.carouselStatus = carouselStatus;
      this.carouselLength = carouselLength;
      this.carouselPosition = carouselPosition;
    }

    public static CarouselStatusReply Parse(string incoming) {
      System.Text.RegularExpressions.Match m = re.Match(incoming);
      if (!m.Success) {
        return null;
      }
      if (!checksumIsValid(incoming, m.Groups["checksum"].Value)) {
        return null;
      }

      return new CarouselStatusReply(
          m.Groups["fromaddr"].Value,
          m.Groups["fromsub"].Value,
          m.Groups["toaddr"].Value,
          m.Groups["tosub"].Value,
          modeToEnum(m.Groups["mode"].Value),
          statusToEnum(m.Groups["status"].Value),
          Convert.ToInt32(m.Groups["count"].Value),
          Convert.ToInt32(m.Groups["pos"].Value)
        );
    }

    protected static CarouselMode modeToEnum(string mode) {
      switch (mode) {
        case "A":
          return CarouselMode.Automatic;
        case "M":
          return CarouselMode.Manual;
        default:
          return CarouselMode.Unknown;
      }                                                          
    }

    protected static CarouselStatus statusToEnum(string status) {
      switch (status) {
        case "Y":
          return CarouselStatus.Success;
        case "F":
          return CarouselStatus.Failure;
        case "I":
          return CarouselStatus.InProgress;
        case "B":
          return CarouselStatus.Braking;
        case "U":
          return CarouselStatus.Uninitialized;
        default:
          return CarouselStatus.Unknown;
      }
    }

  }
}
