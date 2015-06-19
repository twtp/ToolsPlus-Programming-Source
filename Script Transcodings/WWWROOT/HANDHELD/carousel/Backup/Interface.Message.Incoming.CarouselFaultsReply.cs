using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Incoming {
  internal class CarouselFaultsReply : IncomingMessage {

    // message format is <stx>0000ccc3r<etx>s
    //    - to address and subchannel always "0000"
    //    - sender address "ccc"
    //    - sender subchannel always "3"
    //    - fault status bitfield "r"
    //       - 0 = carousel error
    //       - 1 = motor error
    //       - 2 = keypad manual
    //       - 3 = estop engaged
    //       - 4 = right photoeye blocked
    //       - 5 = left photoeye blocked
    //       - 6 = on
    //       - 7 = off
    //    - checksum "s"
    protected static System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(
        @"^" + STX_CHAR +
        @"(?<toaddr>\d{3})" +
        @"(?<tosub>\d{1})" +
        @"(?<fromaddr>\d{3})" +
        @"(?<fromsub>\d{1})" +
        @"(?<char>.)" +
        ETX_CHAR + @"(?<checksum>.)$"
      );

    protected bool carouselError;
    public bool CarouselError { get { return this.carouselError; } }
    protected bool motorError;
    public bool MotorError { get { return this.motorError; } }
    protected KeypadState keypadState;
    public KeypadState KeypadState { get { return this.keypadState; } }
    protected bool emergencyStopEngaged;
    public bool EmergencyStopEngaged { get { return this.emergencyStopEngaged; } }
    protected bool photoeyeRightError;
    public bool PhotoeyeRightError { get { return this.photoeyeRightError; } }
    protected bool photoeyeLeftError;
    public bool PhotoeyeLeftError { get { return this.photoeyeLeftError; } }
    public bool PhotoeyeAnyError { get { return this.photoeyeRightError || this.photoeyeLeftError; } }

    protected CarouselFaultsReply(string fromAddress, string fromSubchannel, string toAddress, string toSubchannel, bool carouselError, bool motorError, KeypadState keypadState, bool emergencyStopEngaged, bool photoeyeRightError, bool photoeyeLeftError) {
      this.fromAddress = fromAddress;
      this.fromSubchannel = fromSubchannel;
      this.toAddress = toAddress;
      this.toSubchannel = toSubchannel;
      this.carouselError = carouselError;
      this.motorError = motorError;
      this.keypadState = keypadState;
      this.emergencyStopEngaged = emergencyStopEngaged;
      this.photoeyeRightError = photoeyeRightError;
      this.photoeyeLeftError = photoeyeLeftError;
    }

    public static CarouselFaultsReply Parse(string incoming) {
      System.Text.RegularExpressions.Match m = re.Match(incoming);
      if (!m.Success) {
        return null;
      }
      if (!checksumIsValid(incoming, m.Groups["checksum"].Value)) {
        return null;
      }

      byte faultByte = System.Text.Encoding.ASCII.GetBytes(m.Groups["char"].Value)[0];

      return new CarouselFaultsReply(
          m.Groups["fromaddr"].Value,
          m.Groups["fromsub"].Value,
          m.Groups["toaddr"].Value,
          m.Groups["tosub"].Value,
          byteToCarouselError(faultByte),
          byteToMotorError(faultByte),
          byteToKeypadState(faultByte),
          byteToEmergencyStopEngaged(faultByte),
          byteToPhotoeyeRightError(faultByte),
          byteToPhotoeyeLeftError(faultByte)
        );
    }

    protected static bool byteToCarouselError(byte faultByte) {
      return (faultByte & (byte)1) == (byte)1;
    }

    protected static bool byteToMotorError(byte faultByte) {
      return (faultByte & (byte)2) == (byte)2;
    }

    protected static KeypadState byteToKeypadState(byte faultByte) {
      return (faultByte & (byte)4) == (byte)4 ? KeypadState.Manual : KeypadState.Host;
    }

    protected static bool byteToEmergencyStopEngaged(byte faultByte) {
      return (faultByte & (byte)8) == (byte)8;
    }

    protected static bool byteToPhotoeyeRightError(byte faultByte) {
      return (faultByte & (byte)16) == (byte)16;
    }

    protected static bool byteToPhotoeyeLeftError(byte faultByte) {
      return (faultByte & (byte)32) == (byte)32;
    }

  }
}
