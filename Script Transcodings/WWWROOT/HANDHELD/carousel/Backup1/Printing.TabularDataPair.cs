using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Printing {
  public class TabularDataPair {
    private string text;
    private int offset;
    private int maxWidth;
    private Justification justify = Justification.Left;
    private System.Drawing.Font font = null;

    public TabularDataPair(string text, int offset, int maxWidth) {
      this.text = text;
      this.offset = offset;
      this.maxWidth = maxWidth;
    }
    public TabularDataPair(string text, int offset, int maxWidth, Justification justify)
      : this(text, offset, maxWidth) {
      this.justify = justify;
    }

    public string Text { get { return this.text; } internal set { this.text = value; } }
    public int Offset { get { return this.offset; } }
    public int MaxWidth { get { return this.maxWidth; } }
    public Justification Justify { get { return this.justify; } }
    public System.Drawing.Font Font { get { return this.font; } set { this.font = value; } }
  }
}
