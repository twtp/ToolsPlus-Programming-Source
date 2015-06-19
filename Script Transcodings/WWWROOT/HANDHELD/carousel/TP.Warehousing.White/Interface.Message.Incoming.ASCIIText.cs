using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Warehousing.White.Interface.Message.Incoming {
  internal class ASCIIText : IncomingMessage {

    // message format is <stx>0000cccpt...t<etx>s
    //    - to address and subchannel always "0000"
    //    - sender address "ccc"
    //    - sender subchannel "p"
    //    - textual message received "t...t"
    //    - checksum "s"
    protected static System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex(
        @"^" + STX_CHAR +
        @"(?<toaddr>\d{3})" +
        @"(?<tosub>\d{1})" +
        @"(?<fromaddr>\d{3})" +
        @"(?<fromsub>\d{1})" +
        @"(?<msg>[^" + ETX_CHAR + "]+)" +
        ETX_CHAR + @"(?<checksum>.)$"
      );

    protected string messageText;
    public string MessageText { get { return this.messageText; } }

    protected ASCIIText(string fromAddress, string fromSubchannel, string toAddress, string toSubchannel, string messageText) {
      this.fromAddress = fromAddress;
      this.fromSubchannel = fromSubchannel;
      this.toAddress = toAddress;
      this.toSubchannel = toSubchannel;
      this.messageText = messageText;
    }

    public static ASCIIText Parse(string incoming) {
      System.Text.RegularExpressions.Match m = re.Match(incoming);
      if (!m.Success) {
        return null;
      }
      if (!checksumIsValid(incoming, m.Groups["checksum"].Value)) {
        return null;
      }

      return new ASCIIText(
          m.Groups["fromaddr"].Value,
          m.Groups["fromsub"].Value,
          m.Groups["toaddr"].Value,
          m.Groups["tosub"].Value,
          m.Groups["msg"].Value
        );
    }

  }
}
