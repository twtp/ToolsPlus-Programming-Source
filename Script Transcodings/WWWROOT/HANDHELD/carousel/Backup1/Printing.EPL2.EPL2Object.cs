using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Printing.EPL2 {

  internal interface IEPL2Object {
    string Print();
    byte[] Bytes();
  }

  internal class EPL2String : IEPL2Object {
    private int x_offset;
    private int y_offset;
    private int rotation;
    private int font;
    private int h_stretch;
    private int v_stretch;
    private bool reverse;
    private string text;

    internal EPL2String(int x_offset, int y_offset, EPL2Font font, int h_stretch, int v_stretch, string text) {
      this.x_offset = x_offset;
      this.y_offset = y_offset;
      this.rotation = 0;
      this.font = (int)font;
      this.h_stretch = h_stretch;
      this.v_stretch = v_stretch;
      this.reverse = false;
      this.text = text;
    }
    internal EPL2String(int x_offset, int y_offset, EPL2Font font, int h_stretch, int v_stretch, string text, EPL2Rotation rotation) : this(x_offset, y_offset, font, h_stretch, v_stretch, text) {
      this.rotation = (int)rotation;
    }

    public string Print() {
      return string.Format(
          "A{0},{1},{2},{3},{4},{5},{6},{7}\r\n",
          this.x_offset,
          this.y_offset,
          this.rotation,
          this.font,
          this.h_stretch,
          this.v_stretch,
          (this.reverse ? "R" : "N"),
          ("\"" + this.text.Replace("\"", "\\\"") + "\"")
        );
    }

    public byte[] Bytes() {
      return System.Text.Encoding.ASCII.GetBytes(this.Print());
    }
  }

  internal class EPL2Barcode : IEPL2Object {
    private int x_offset;
    private int y_offset;
    private int rotation;
    private int format;
    private int width1;
    private int width2;
    private int height;
    private bool readable;
    private string barcode;

    internal EPL2Barcode(int x_offset, int y_offset, int height, int width1, string barcode, bool readable) {
      this.x_offset = x_offset;
      this.y_offset = y_offset;
      this.rotation = (int)EPL2Rotation.Standard;
      this.format = 1;
      this.width1 = width1;
      this.width2 = width1 * 2;
      this.height = height;
      this.readable = readable;
      this.barcode = barcode;
    }
    internal EPL2Barcode(int x_offset, int y_offset, int height, int width1, string barcode, bool readable, EPL2Rotation rotation) : this(x_offset, y_offset, height, width1, barcode, readable) {
      this.rotation = (int)rotation;
    }

    public string Print() {
      return string.Format(
          "B{0},{1},{2},{3},{4},{5},{6},{7},{8}\r\n",
          this.x_offset,
          this.y_offset,
          this.rotation,
          this.format,
          this.width1,
          this.width2,
          this.height,
          (this.readable ? "Y" : "N"),
          ("\"" + this.barcode.Replace("\"", "\\\"") + "\"")
        );
    }

    public byte[] Bytes() {
      return System.Text.Encoding.ASCII.GetBytes(this.Print());
    }
  }

  internal class EPL2Line : IEPL2Object {
    private int x_offset;
    private int y_offset;
    private int x_end;
    private int y_end;
    private int thickness;

    public EPL2Line(int x_offset, int y_offset, int x_end, int y_end, int thickness) {
      this.x_offset = x_offset;
      this.y_offset = y_offset;
      this.x_end = x_end;
      this.y_end = y_end;
      this.thickness = thickness;
    }

    public string Print() {
      return string.Format(
          "LS{0},{1},{2},{3},{4}\r\n",
          this.x_offset,
          this.y_offset,
          this.thickness,
          this.x_end,
          this.y_end
        );
    }

    public byte[] Bytes() {
      return System.Text.Encoding.ASCII.GetBytes(this.Print());
    }
  }

  internal class EPL2Image : IEPL2Object {
    private int x_offset;
    private int y_offset;
    private byte[] bytes;

    private static Dictionary<string, byte[]> cache = new Dictionary<string, byte[]>();

    public EPL2Image(int x_offset, int y_offset, System.Drawing.Bitmap bmp) {
      this.x_offset = x_offset;
      this.y_offset = y_offset;
      // bitmap -> byte[] caching? GetHashCode returns different for same bmp
      // for now, if we're doing the same bmp all the time, add as a resource and use that?
      this.bytes = this.convertImageToEPL2Data(bmp);
    }
    public EPL2Image(int x_offset, int y_offset, string resourceName) {
      this.x_offset = x_offset;
      this.y_offset = y_offset;
      if (!cache.ContainsKey(resourceName)) {
        cache.Add(resourceName, this.convertImageToEPL2Data((System.Drawing.Bitmap)Properties.Resources.ResourceManager.GetObject(resourceName)));
      }
      this.bytes = cache[resourceName];
    }

    private byte[] convertImageToEPL2Data(System.Drawing.Bitmap bmp) {
      using (System.IO.MemoryStream ms = new System.IO.MemoryStream()) {
        using (System.IO.BinaryWriter bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.ASCII)) {
          bw.Write(System.Text.Encoding.ASCII.GetBytes(string.Format(
              "GW{0},{1},{2},{3}\r\n", // comma or CRLF between p4 and data, not in EPL2 manual per https://support.zebra.com/cpws/docs/eltron/gw_command.htm
              this.x_offset,
              this.y_offset,
              Math.Ceiling((decimal)bmp.Width / 8),
              bmp.Height
            )));

          // foreach pixel, use the standard luminance algorithm to check whether it
          // should be black or white. using a 50/50 split right now. should that be
          // different and/or configurable? also, we might want a expansion param
          // (similar to the text expansion natively supported), so we don't need
          // giant images due to the dpi, e.g., 600px == 3in, we'd be able to save
          // space in the resources file. 2x or 3x would probably not lose too much
          // quality.
          for (int y = 0; y < bmp.Height; y++) {
            byte packed = 0;
            int bitCount = 0;

            for (int x = 0; x < bmp.Width; x++) {
              System.Drawing.Color color = bmp.GetPixel(x, y);
              int luminance = (int)((color.R * 0.3) + (color.G * 0.59) + (color.B * 0.11));
              bool dot = luminance < 127;

              // zero means burn a dot, one means skip? ass-backwards.
              packed |= (byte)((dot ? 0 : 1) << (7 - bitCount));
              bitCount++;

              if (bitCount == 8) {
                bw.Write(packed);
                packed = 0;
                bitCount = 0;
              }
            }

            // pad to full byte if leftover
            if (bitCount > 0) {
              while (bitCount < 8) {
                packed |= (byte)((1) << (7 - bitCount));
                bitCount++;
              }
              bw.Write(packed);
            }
          }

          bw.Write(System.Text.Encoding.ASCII.GetBytes("\r\n"));
          bw.Flush();
          return ms.ToArray();
        }
      }
    }

    public string Print() {
      return System.Text.Encoding.ASCII.GetString(this.bytes);
    }

    public byte[] Bytes() {
      return this.bytes;
    }
  }

}
