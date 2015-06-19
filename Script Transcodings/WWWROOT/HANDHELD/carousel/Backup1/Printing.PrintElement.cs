using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;
using System.Drawing.Drawing2D;

namespace TP.Base.Printing {
  public class PrintElement {
    private List<IPrintPrimitive> printPrimitives = new List<IPrintPrimitive>();
    private IPrintable printObject;

    private Rectangle? overridePosition = null;
    public Rectangle? OverridePosition { get { return this.overridePosition; } set { this.overridePosition = value; } }
    private Font overrideFont = null;
    public Font OverrideFont { get { return this.overrideFont; } set { this.overrideFont = value; } }

    public PrintElement(IPrintable printObject) {
      this.printObject = printObject;
    }

    public void Print() {
      this.printObject.Print(this);
    }

    public void AddText(string buf) {
      AddPrimitive(new PrintPrimitiveText(buf, null));
    }
    public void AddText(string buf, Font font) {
      AddPrimitive(new PrintPrimitiveText(buf, font));
    }
    public void AddText(string buf, float fontSize) {
      AddPrimitive(new PrintPrimitiveText(buf, new Font(FontFamily.GenericSansSerif, fontSize)));
    }

    public void AddTabularText(List<TabularDataPair> columns) {
      AddPrimitive(new PrintPrimitiveTableRow(columns, null));
    }
    public void AddTabularText(List<TabularDataPair> columns, Font font) {
      AddPrimitive(new PrintPrimitiveTableRow(columns, font));
    }
    public void AddTabularText(List<TabularDataPair> columns, float fontSize) {
      AddPrimitive(new PrintPrimitiveTableRow(columns, new Font(FontFamily.GenericSansSerif, fontSize)));
    }

    public void AddTitledText(TabularDataPair title, TabularDataPair content) {
      AddPrimitive(new PrintPrimitivePair(title, content, null));
    }
    public void AddTitledText(TabularDataPair title, TabularDataPair content, Font font) {
      AddPrimitive(new PrintPrimitivePair(title, content, font));
    }

    internal void AddPrimitive(IPrintPrimitive primitive) {
      this.printPrimitives.Add(primitive);
    }

    public void AddData(string dataName, string dataValue) {
      AddText(dataName + ": " + dataValue);
    }

    public void AddHorizontalRule() {
      AddHorizontalRule(DashStyle.Solid);
    }

    public void AddHorizontalRule(DashStyle dashStyle) {
      AddPrimitive(new PrintPrimitiveRule(dashStyle));
    }

    public void AddBarcode(TP.Base.Barcode.IBarcode barcode) {
      AddPrimitive(new PrintPrimitiveBarcode(barcode));
    }
    public void AddBarcode(params TP.Base.Barcode.IBarcode[] barcodes) {
      AddPrimitive(new PrintPrimitiveBarcode(barcodes));
    }

    public void AddBlankLine() {
      AddText(string.Empty, null);
    }
    public void AddBlankLine(Font font) {
      AddText(string.Empty, font);
    }
    public void AddBlankLine(float fontSize) {
      AddText(string.Empty, new Font(FontFamily.GenericSansSerif, fontSize));
    }

    public void AddHeader(string buf) {
      AddText(buf);
      AddHorizontalRule();
    }

    private float? cachedHeight = null;
    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      if (this.cachedHeight == null) {
        float height = 0;
        foreach (IPrintPrimitive primitive in this.printPrimitives) {
          height += primitive.CalculateHeight(engine, graphics);
        }
        this.cachedHeight = height;
      }
      return this.cachedHeight.Value;
    }

    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle pageBounds) {
      System.Drawing.Rectangle elementBounds;
      if (this.overridePosition == null) {
        float height = CalculateHeight(engine, graphics);
        elementBounds = new Rectangle(
            pageBounds.Left,
            (int)yPos,
            pageBounds.Right - pageBounds.Left,
            (int)height
          );
      }
      else {
        //elementBounds = (System.Drawing.Rectangle)overridePosition;
        elementBounds = new Rectangle(
            pageBounds.Left + this.overridePosition.Value.Left,
            pageBounds.Top + this.overridePosition.Value.Top,
            pageBounds.Left + this.overridePosition.Value.Right,
            pageBounds.Top + this.overridePosition.Value.Bottom
          );
      }

      foreach (IPrintPrimitive primitive in this.printPrimitives) {
        if (this.overridePosition == null) {
          primitive.Draw(engine, yPos, graphics, elementBounds);
          yPos += primitive.CalculateHeight(engine, graphics);
        }
        else {
          primitive.Draw(engine, elementBounds.Top, graphics, elementBounds);
        }
      }
    }
  }
}
