using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Printing {

  public interface IPrintable {
    void Print(PrintElement element);
  }

  internal interface IPrintPrimitive {
    float CalculateHeight(PrintEngine engine, System.Drawing.Graphics graphics);
    void Draw(PrintEngine engine, float yPos, System.Drawing.Graphics graphics, System.Drawing.Rectangle elementBounds);
  }

}
