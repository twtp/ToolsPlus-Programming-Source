using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Printing {
  internal class PrintPrimitiveText : IPrintPrimitive {
    private string text;
    private Font font;

    public PrintPrimitiveText(string buf, Font font) {
      this.text = buf;
      this.font = font;
    }
    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      return font == null ? engine.PrintFont.GetHeight(graphics) : font.GetHeight(graphics);
    }
    public float CalculateWidth(PrintEngine engine, Graphics graphics, string text) {
      return graphics.MeasureString(text, font == null ? engine.PrintFont : font).Width;
    }
    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle elementBounds) {
      graphics.DrawString(engine.ReplaceTokens(this.text), font == null ? engine.PrintFont : font, engine.PrintBrush, elementBounds.Left, yPos, new StringFormat());
    }
  }
}
