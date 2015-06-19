using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Printing {
  internal class PrintPrimitiveBarcode : IPrintPrimitive {
    private TP.Base.Barcode.IBarcode[] bcs;
    public PrintPrimitiveBarcode(TP.Base.Barcode.IBarcode barcode) {
      this.bcs = new TP.Base.Barcode.IBarcode[1];
      this.bcs[0] = barcode;
    }
    public PrintPrimitiveBarcode(TP.Base.Barcode.IBarcode[] barcodes) {
      this.bcs = barcodes;
    }
    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      return this.bcs[0].Height; // assume all equal
    }
    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle elementBounds) {
      int leftOffset = 0;
      foreach (TP.Base.Barcode.IBarcode bc in this.bcs) {
        bc.AsPrintDocumentElement(graphics, new Point(elementBounds.Left + leftOffset, elementBounds.Top));
        leftOffset += bc.Width + 150;
      }
    }
  }
}
