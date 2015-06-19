using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Printing {
  internal class PrintPrimitiveTableRow : IPrintPrimitive {
    private List<TabularDataPair> columns;
    private Font font;

    public PrintPrimitiveTableRow(List<TabularDataPair> columns, Font font) {
      this.columns = columns;
      this.font = font;
    }
    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      return this.font == null ? engine.PrintFont.GetHeight(graphics) : this.font.GetHeight(graphics);
    }
    public float CalculateWidth(PrintEngine engine, Graphics graphics, string text) {
      return graphics.MeasureString(text, this.font == null ? engine.PrintFont : this.font).Width;
    }
    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle elementBounds) {
      float height = this.CalculateHeight(engine, graphics);
      foreach (TabularDataPair tdp in this.columns) {
        string actualText = engine.ReplaceTokens(tdp.Text);
        RectangleF rect;
        float actualWidth;
        switch (tdp.Justify) {
          case Justification.Left:
            rect = new RectangleF(elementBounds.Left + tdp.Offset, yPos, tdp.MaxWidth, height);
            break;
          case Justification.Right:
            actualWidth = this.CalculateWidth(engine, graphics, actualText);
            while (actualWidth > tdp.MaxWidth) {
              actualText = actualText.Substring(1); // chop char from front and recalc? this seems bad all around
              actualWidth = this.CalculateWidth(engine, graphics, actualText);
            }
            rect = new RectangleF(elementBounds.Left + tdp.Offset + tdp.MaxWidth - actualWidth, yPos, actualWidth, height);
            break;
          case Justification.Center:
            actualWidth = this.CalculateWidth(engine, graphics, actualText);
            while (actualWidth > tdp.MaxWidth) {
              actualText = actualText.Substring(actualText.Length); // chop char from back, i guess
              actualWidth = this.CalculateWidth(engine, graphics, actualText);
            }
            rect = new RectangleF(elementBounds.Left + tdp.Offset + ((tdp.MaxWidth - actualWidth) / 2), yPos, actualWidth, height);
            break;
          default:
            throw new ApplicationException("invalid justification");
        }

        graphics.DrawString(
            actualText,
            tdp.Font != null ? tdp.Font : this.font != null ? this.font : engine.PrintFont,
            engine.PrintBrush,
            rect,
            new StringFormat()
          );
      }
    }
  }
}
