using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Barcode {

  public interface IBarcode {
    void AsPrintDocumentElement(Graphics g, Point pt);

    Bitmap AsBitmap();
    Bitmap AsBitmap(Bitmap bmp, Point pt);

    string AsLineString();

    int Width { get; }
    int Height { get; }
  }


  public class Code39 : IBarcode {
    private int _lineWidth = 1;
    private int _lineRatio = 3;
    private int _height = 30;

    private string barcode = string.Empty;

    public Code39(string barcode) {
      this.barcode = "*" + barcode + "*";
      //this.barcode = "*ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789*";
    }

    public void AsPrintDocumentElement(Graphics g, Point pt) {
      this.asGraphicalElement(g, pt);
    }

    public Bitmap AsBitmap() {
      return this.AsBitmap(new Bitmap(this.Width, this.Height), new Point(0, 0));
    }
    public Bitmap AsBitmap(Bitmap bmp, Point pt) {
      Graphics g = Graphics.FromImage(bmp);
      this.asGraphicalElement(g, pt);
      return bmp;
    }

    public string AsLineString() {
      return this.convertBarcodeToLineString();
    }

    private int asGraphicalElement(Graphics g, Point pt) {
      Brush brush = new SolidBrush(Color.Black);
      int leftPoint = pt.X;
      string lines = this.convertBarcodeToLineString();
      bool color = true;

      foreach (char c in lines) {
        int numUnits = 0;
        switch (c) {
          case '0':
            numUnits = this._lineWidth;
            break;
          case '1':
            numUnits = this._lineWidth * this._lineRatio;
            break;
        }
        if (color == true) {
          g.FillRectangle(brush, leftPoint, pt.Y, numUnits, this._height);
        }
        leftPoint += numUnits;
        color = !color;
      }

      return leftPoint - pt.X;
    }

    private string convertBarcodeToLineString() {
      string retval = string.Empty;
      char[] chars = this.barcode.ToCharArray();
      for (int i = 0; i < chars.Length; i++) {
        retval += this.convertCharToLineString(chars[i]);
        if (i < chars.Length - 1) {
          retval += "0";
        }
      }
      return retval;
    }

    private string convertCharToLineString(char c) {
      switch (c) {
        case '0': return "000110100";
        case '1': return "100100001";
        case '2': return "001100001";
        case '3': return "101100000";
        case '4': return "000110001";
        case '5': return "100110000";
        case '6': return "001110000";
        case '7': return "000100101";
        case '8': return "100100100";
        case '9': return "001100100";
        case 'A': return "100001001";
        case 'B': return "001001001";
        case 'C': return "101001000";
        case 'D': return "000011001";
        case 'E': return "100011000";
        case 'F': return "001011000";
        case 'G': return "000001101";
        case 'H': return "100001100";
        case 'I': return "001001100";
        case 'J': return "000011100";
        case 'K': return "100000011";
        case 'L': return "001000011";
        case 'M': return "101000010";
        case 'N': return "000010011";
        case 'O': return "100010010";
        case 'P': return "001010010";
        case 'Q': return "000000111";
        case 'R': return "100000110";
        case 'S': return "001000110";
        case 'T': return "000010110";
        case 'U': return "110000001";
        case 'V': return "011000001";
        case 'W': return "111000000";
        case 'X': return "010010001";
        case 'Y': return "110010000";
        case 'Z': return "011010000";
        case '-': return "010000101";
        case '*': return "010010100";
        case '+': return "010001010";
        case '$': return "010101000";
        case '%': return "000101010";
        case '/': return "010100010";
        case '.': return "110000100";
        case ' ': return "011000100";
        default: return "";
      }
    }

    public int Width {
      get {
        return this.barcode.Length * ((3 * this._lineWidth * this._lineRatio) + // three wide
                                      (6 * this._lineWidth) +                   // six narrow
                                      (1 * this._lineWidth)                     // between spacer
                                     ) - this._lineWidth;                       // minus the last spacer
      }
    }
    public int Height { get { return this._height; } }
  }
}
