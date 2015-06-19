using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Printing.EPL2 {

  public abstract class EPL2Document {
    private int paperWidth;
    private int paperHeight;
    private List<IEPL2Object> elements;

    // XXX: everything but the paperSize ctor should probably be internal, don't want to muck
    // with paper sizes everywhere
    protected EPL2Document() : this(456, 152) {
      // nothing
    }
    protected EPL2Document(float width, float height) : this(inchesToDots(width), inchesToDots(height)) {
      // nothing
    }
    protected EPL2Document(int width, int height) {
      this.paperWidth = width;
      this.paperHeight = height;
      this.elements = new List<IEPL2Object>();
    }
    public EPL2Document(EPL2PaperSize paperSize) {
      switch (paperSize) {
        case EPL2PaperSize.BarcodeLabel:
          this.paperWidth = 456;
          this.paperHeight = 152;
          break;
        case EPL2PaperSize.ShippingLabel:
          this.paperWidth = 812;
          this.paperHeight = 1700;
          break;
        default:
          throw new ApplicationException("missing EPL2PaperSize in switch()!");
      }
      this.elements = new List<IEPL2Object>();
    }

    public bool AddBarcode(float x_offset, float y_offset, float height, int barWidth, string barcode, bool readable) {
      return this.AddBarcode(inchesToDots(x_offset), inchesToDots(y_offset), inchesToDots(height), barWidth, barcode, readable);
    }
    public bool AddBarcode(int x_offset, int y_offset, int height, int barWidth, string barcode, bool readable) {
      EPL2Barcode bc = new EPL2Barcode(
          x_offset,
          y_offset,
          height,
          barWidth,
          barcode,
          readable
        );
      this.elements.Add(bc);
      return true;
    }

    public bool AddText(int x_offset, int y_offset, int h_stretch, int v_stretch, string text) {
      EPL2String s = new EPL2String(
          x_offset,
          y_offset,
          EPL2Font.CPI15_12x20,
          h_stretch,
          v_stretch,
          text
        );
      this.elements.Add(s);
      return true;
    }
    public bool AddText(int x_offset, int y_offset, EPL2Font font, int h_stretch, int v_stretch, string text) {
      EPL2String s = new EPL2String(
          x_offset,
          y_offset,
          font,
          h_stretch,
          v_stretch,
          text
        );
      this.elements.Add(s);
      return true;
    }
    public bool AddText(int x_offset, int y_offset, int h_stretch, int v_stretch, string text, EPL2Rotation rotation) {
      EPL2String s = new EPL2String(
          x_offset,
          y_offset,
          EPL2Font.CPI15_12x20,
          h_stretch,
          v_stretch,
          text,
          rotation
        );
      this.elements.Add(s);
      return true;
    }

    public bool AddImage(int x_offset, int y_offset, System.Drawing.Bitmap bmp) {
      EPL2Image i = new EPL2Image(
          x_offset,
          y_offset,
          bmp
        );
      this.elements.Add(i);
      return true;
    }
    public bool AddImage(int x_offset, int y_offset, string resourceName) {
      EPL2Image i = new EPL2Image(
          x_offset,
          y_offset,
          resourceName
        );
      this.elements.Add(i);
      return true;
    }

    public bool AddLine(int x_offset, int y_offset, int x_end, int y_end, int thickness) {
      EPL2Line l = new EPL2Line(
          x_offset,
          y_offset,
          x_end,
          y_end,
          thickness
        );
      this.elements.Add(l);
      return true;
    }

    public bool Print(string printerName) {
      return this.Print(printerName, 1);
    }
    public bool Print(string printerName, int copies) {
      /*string epl2 = string.Format(
          "EPL2\r\nq{0}\r\nQ{1},24+0\r\nS4\r\nUN\r\nWN\r\nZT\r\nN\r\n",
          this.paperWidth,
          this.paperHeight
        );
      foreach (IEPL2Object o in this.elements) {
        epl2 += o.Print();
      }
      epl2 += string.Format("P{0}\r\n", copies);

      //System.Diagnostics.Debug.Print(epl2);
      return TP.Base.Interop.Printing.PrintRAWString(printerName, epl2, "EPL2 Document");*/

      bool retval;
      using (System.IO.MemoryStream ms = new System.IO.MemoryStream()) {
        using (System.IO.BinaryWriter bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.ASCII)) {
          bw.Write(System.Text.Encoding.ASCII.GetBytes(string.Format(
              "EPL2\r\nq{0}\r\nQ{1},24+0\r\nS4\r\nUN\r\nWN\r\nZT\r\nN\r\n",
              this.paperWidth,
              this.paperHeight
            )));
          foreach (IEPL2Object o in this.elements) {
            bw.Write(o.Bytes());
          }
          bw.Write(System.Text.Encoding.ASCII.GetBytes(string.Format("P{0}\r\n", copies)));

          bw.Flush();
        }

        retval = TP.Base.Interop.Printing.PrintRAWBytes(printerName, ms.ToArray(), "EPL2 Document");
      }
      return retval;
    }

    private static int inchesToDots(float inches) {
      return (int)Math.Floor(inches * 203); // DPI
    }
  }

  public class EPL2BarcodeDocument : EPL2Document {
    public EPL2BarcodeDocument(string barcode, string readable) : base(EPL2PaperSize.BarcodeLabel) {
      this.AddBarcode(40, 10, 100, 2, barcode, false);
      this.AddText(80, 120, 1, 1, readable);
    }
  }

  public class EPL2SerialNumberLabelDocument : EPL2Document {
    public EPL2SerialNumberLabelDocument(string modelNo, string serialNo, string purchaseOrderNo, string dateOfReceipt) : base(EPL2PaperSize.ShippingLabel) {
      this.AddText(760, 1160, 4, 4, "Model #:", EPL2Rotation.UpsideDown);
      this.AddText(700, 1060, 4, 4, modelNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 940, 4, 4, "Serial #:", EPL2Rotation.UpsideDown);
      this.AddText(700, 840, 4, 4, serialNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 720, 4, 4, "P/O #:", EPL2Rotation.UpsideDown);
      this.AddText(700, 620, 4, 4, purchaseOrderNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 500, 4, 4, "Received:", EPL2Rotation.UpsideDown);
      this.AddText(700, 400, 4, 4, dateOfReceipt, EPL2Rotation.UpsideDown);
      this.AddText(760, 280, 4, 4, "S/O #:", EPL2Rotation.UpsideDown);
      this.AddText(700, 180, 4, 4, "_______", EPL2Rotation.UpsideDown);
      this.AddText(760, 1640, 2, 2, "Model #:  " + modelNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 1580, 2, 2, "Serial #: " + serialNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 1520, 2, 2, "P/O #:    " + purchaseOrderNo, EPL2Rotation.UpsideDown);
      this.AddText(760, 1460, 2, 2, "Received: " + dateOfReceipt, EPL2Rotation.UpsideDown);
      this.AddText(760, 1400, 2, 2, "S/O #:    " + "_______", EPL2Rotation.UpsideDown);
    }
  }

  public class EPL2CustomerAccountDocument : EPL2Document {
    public EPL2CustomerAccountDocument(string accountNo, string customerName) : base(EPL2PaperSize.BarcodeLabel) {
      this.AddText(80, 10, 2, 1, "TOOLS PLUS");
      this.AddBarcode(50, 40, 70, 2, "CACCT" + accountNo.Replace("-", ""), false);
      this.AddText(60, 120, 1, 1, customerName.Length > 28 ? customerName.Substring(0, 28) : customerName);
    }
  }

  public class EPL2ThisSideUpLabelDocument : EPL2Document {
    public EPL2ThisSideUpLabelDocument() : base(EPL2PaperSize.ShippingLabel) {
      this.AddText(250, 150, 6, 6, "THIS");
      this.AddText(100, 300, 6, 6, "SIDE UP");

      int x_offset = 200;
      int y_offset = 550;

      int overall_width = 400;
      int overall_height = 400;
      int thickness = 8;

      int current_width = 0;
      for (int i = y_offset; i < y_offset + (overall_height / 2); i += thickness) {
        current_width += thickness * 2;
        int this_offset = (overall_width - current_width) / 2;
        this.AddLine(x_offset + this_offset, i, x_offset + this_offset + current_width, i, thickness);
      }

      int bar_width = overall_width / 2;
      int bar_offset = overall_width / 4;
      for (int i = y_offset + (overall_height / 2); i < y_offset + overall_height; i += thickness) {
        this.AddLine(x_offset + bar_offset, i, x_offset + bar_offset + bar_width, i, thickness);
      }
    }
  }

  public class EPL2HorizontalCarouselBinLabelDocument : EPL2Document {
    public EPL2HorizontalCarouselBinLabelDocument(string carouselNo, string binNo, string shelfNo, string cellNo, int locationID) : base(EPL2PaperSize.ShippingLabel) {
      this.AddText(20, 10, EPL2Font.CPI13_14x24, 2, 2, "Carousel " + carouselNo + " Bin " + binNo);
      this.AddText(150, 60, EPL2Font.CPI06_32x48, 2, 2, shelfNo + " " + cellNo);
      this.AddBarcode(30, 174, 66, 3, string.Format("LOC{0:000000000}", locationID), false);
      this.AddLine(609, 0, 609, 254, 1);
      this.AddLine(0, 254, 609, 254, 1);

      this.AddText(20, 10 + 256, EPL2Font.CPI13_14x24, 2, 2, "Carousel " + carouselNo + " Bin " + binNo);
      this.AddText(150, 60 + 256, EPL2Font.CPI06_32x48, 2, 2, shelfNo + " " + cellNo);
      this.AddBarcode(30, 174 + 256, 66, 3, string.Format("LOC{0:000000000}", locationID), false);
      this.AddLine(609, 0 + 256, 609, 254 + 256, 1);
      this.AddLine(0, 254 + 256, 609, 254 + 256, 1);
    }
  }

}
