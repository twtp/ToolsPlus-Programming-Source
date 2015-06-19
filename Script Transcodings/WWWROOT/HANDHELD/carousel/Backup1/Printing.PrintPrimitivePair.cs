using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Printing {
  internal class PrintPrimitivePair : IPrintPrimitive {
    private TabularDataPair first;
    private TabularDataPair second;
    private bool multiLineProcessed = false;

    public PrintPrimitivePair(TabularDataPair first, TabularDataPair second, Font font) {
      this.first = first;
      this.second = second;

      if (font == null) {
        font = new Font(FontFamily.GenericSansSerif, 10);
      }
      this.first.Font = new Font(font, FontStyle.Bold);
      this.second.Font = font;
    }

    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      if (this.second.Text.Contains("\n") && this.multiLineProcessed == false) {
        // multi-line, need to check word wrapping for proper height
        // TODO: pull this out so other text handlers can use it
        List<string> newLines = new List<string>();
        string[] origlines = this.second.Text.Replace("\r", "").Split('\n');
        foreach (string ln in origlines) {
          float origWidth = graphics.MeasureString(ln, this.second.Font).Width;
          if (origWidth <= this.second.MaxWidth) {
            newLines.Add(ln);
          }
          else {
            string[] words = ln.Split(' ');
            string checkLine = string.Empty;
            foreach (string word in words) {
              if (checkLine == string.Empty) {
                checkLine = word;
              }
              else {
                if (graphics.MeasureString(checkLine + " " + word, this.second.Font).Width <= this.second.MaxWidth) {
                  checkLine += " " + word;
                }
                else {
                  newLines.Add(checkLine);
                  checkLine = word;
                }
              }
            }
            if (checkLine != string.Empty) {
              newLines.Add(checkLine);
            }
          }
          this.second.Text = string.Join("\r\n", newLines.ToArray());
        }
        this.multiLineProcessed = true;
      }
      return Math.Max(graphics.MeasureString(this.first.Text, this.first.Font).Height, graphics.MeasureString(this.second.Text, this.second.Font).Height);
    }
    
    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle elementBounds) {
      float height = this.CalculateHeight(engine, graphics);
      RectangleF rect;

      rect = new RectangleF(elementBounds.Left + this.first.Offset, yPos, this.first.MaxWidth, height);
      graphics.DrawString(
          this.first.Text,
          this.first.Font,
          engine.PrintBrush,
          rect,
          new StringFormat()
        );

      rect = new RectangleF(elementBounds.Left + this.first.Offset + this.first.MaxWidth, yPos, this.second.MaxWidth, height);
      graphics.DrawString(
          this.second.Text,
          this.second.Font,
          engine.PrintBrush,
          rect,
          new StringFormat()
        );

    }
  }
}
