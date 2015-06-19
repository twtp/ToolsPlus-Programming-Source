using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;
using System.Drawing.Drawing2D;

namespace TP.Base.Printing {
  internal class PrintPrimitiveRule : IPrintPrimitive {
    private DashStyle dashStyle = DashStyle.Solid;

    public PrintPrimitiveRule() { }
    public PrintPrimitiveRule(DashStyle dashStyle) {
      this.dashStyle = dashStyle;
    }

    public float CalculateHeight(PrintEngine engine, Graphics graphics) {
      return 5;
    }
    public void Draw(PrintEngine engine, float yPos, Graphics graphics, Rectangle elementBounds) {
      Pen pen = new Pen(engine.PrintBrush, 1);
      pen.DashStyle = this.dashStyle;
      graphics.DrawLine(pen, elementBounds.Left, yPos + 2, elementBounds.Right, yPos + 2);
    }
  }

  

  

  

  

  
}
