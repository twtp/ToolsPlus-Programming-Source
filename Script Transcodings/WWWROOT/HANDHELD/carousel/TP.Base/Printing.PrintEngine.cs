using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;

namespace TP.Base.Printing {
  public class PrintEngine : System.Drawing.Printing.PrintDocument {

    private List<IPrintable> printObjects = new List<IPrintable>();
    private List<PositionedObject> positionedObjects = new List<PositionedObject>();
    private Font printFont = new Font("Arial", 12);
    private Brush printBrush = Brushes.Black;
    private PrintElement header = null;
    private PrintElement footer = null;
    private List<PrintElement> printElements;
    private List<PrintElement> positionedElements;
    private int printIndex = 0;
    private int pageNum = 0;
    private int printObjectStart = 0;
    private int totalPages = 0;

    public PrintEngine() {
      // default to 1/2" margins on l/r
      this.DefaultPageSettings.Margins.Left = 50;
      this.DefaultPageSettings.Margins.Right = 50;
    }

    public void SetFontSize(float points) {
      this.printFont = new Font(this.PrintFont.FontFamily, points);
    }

    public void SetMargins(int topBottom, int leftRight) {
      this.SetMargins(topBottom, leftRight, topBottom, leftRight);
    }
    public void SetMargins(int top, int left, int bottom, int right) {
      this.DefaultPageSettings.Margins.Top = top;
      this.DefaultPageSettings.Margins.Left = left;
      this.DefaultPageSettings.Margins.Bottom = bottom;
      this.DefaultPageSettings.Margins.Right = right;
    }

    public void AddHeaderText(string buf) {
      this.header = new PrintElement(null);
      this.header.AddText(buf);
      this.header.AddHorizontalRule();
      this.header.AddBlankLine();
    }

    public void AddHeader(IPrintable printObject) {
      this.header = new PrintElement(printObject);
      this.header.Print();
    }

    public void AddFooterText(string buf) {
      this.footer = new PrintElement(null);
      this.footer.AddHorizontalRule();
      this.footer.AddText(buf);
    }

    public void AddFooter(IPrintable printObject) {
      this.footer = new PrintElement(printObject);
      this.footer.Print();
    }

    public void AddPositionedElement(Font baseFont, string text, Rectangle position) {
      this.positionedObjects.Add(new PositionedObject(baseFont, text, position));
    }
    private class PositionedObject : IPrintable {
      private Font font;
      private string text;
      private Rectangle position;
      public PositionedObject(Font font, string text, Rectangle position) {
        this.font = font;
        this.text = text;
        this.position = position;
      }
      public Font Font { get { return this.font; } }
      public string Text { get { return this.text; } }
      public Rectangle Position { get { return this.position; } }

      public void Print(TP.Base.Printing.PrintElement element) {
        element.AddText(this.text, this.font);
      }
    }

    public void AddPrintObject(IPrintable printObject) {
      this.printObjects.Add(printObject);
    }

    public void SetPrintObjectStart(int yPos) {
      this.printObjectStart = yPos;
    }

    public string ReplaceTokens(string buf) {
      buf = buf.Replace("[pagenum]", this.pageNum.ToString());
      buf = buf.Replace("[pagetotal]", this.totalPages.ToString());
      return buf;
    }

    protected override void OnBeginPrint(System.Drawing.Printing.PrintEventArgs e) {
      this.positionedElements = new List<PrintElement>();
      this.printElements = new List<PrintElement>();
      this.pageNum = 0;
      this.printIndex = 0;

      foreach (PositionedObject positionedObject in this.positionedObjects) {
        PrintElement element = new PrintElement(positionedObject);
        element.OverrideFont = positionedObject.Font;
        element.OverridePosition = positionedObject.Position;
        this.positionedElements.Add(element);
        element.Print();
      }

      foreach (IPrintable printObject in this.printObjects) {
        PrintElement element = new PrintElement(printObject);
        this.printElements.Add(element);
        element.Print();
      }
    }

    private int calculateTotalNumberOfPages(Graphics g, int marginHeight) {
      int pagesSoFar = 1;
      // TODO: calculate page size, determine elements to put on page
      // eventually figure out number of pages?
      float headerHeight = this.header != null ? this.header.CalculateHeight(this, g) : 0;
      float footerHeight = this.footer != null ? this.footer.CalculateHeight(this, g) : 0;
      //int pageTop = (int)(e.MarginBounds.Top + headerHeight);
      int contentHeight = (int)(marginHeight - footerHeight - headerHeight);
      float yPos = this.printObjectStart; // first page, can have static elements
      foreach (PrintElement pe in this.printElements) {
        float peHeight = pe.CalculateHeight(this, g);
        if (yPos + peHeight > contentHeight) {
          pagesSoFar++;
          yPos = peHeight;
        }
        else {
          yPos += peHeight;
        }
      }
      return pagesSoFar;
    }

    protected override void OnPrintPage(System.Drawing.Printing.PrintPageEventArgs e) {
      this.pageNum++;

      if (this.totalPages == 0) {
        this.totalPages = this.calculateTotalNumberOfPages(e.Graphics, e.MarginBounds.Height);
      }

      float headerHeight = 0;
      if (this.header != null) {
        headerHeight = this.header.CalculateHeight(this, e.Graphics);
        this.header.Draw(this, e.MarginBounds.Top, e.Graphics, e.MarginBounds);
      }

      float footerHeight = 0;
      if (this.footer != null) {
        footerHeight = this.footer.CalculateHeight(this, e.Graphics);
        this.footer.Draw(this, e.MarginBounds.Bottom - footerHeight, e.Graphics, e.MarginBounds);
      }

      System.Drawing.Rectangle pageBounds = new Rectangle(
          e.MarginBounds.Left,
          (int)(e.MarginBounds.Top + headerHeight),
          e.MarginBounds.Width,
          (int)(e.MarginBounds.Height - footerHeight - headerHeight)
        );
      float yPos = this.printObjectStart == 0 || this.pageNum > 1 ? pageBounds.Top : (this.printObjectStart + pageBounds.Top);

      if (this.pageNum == 1 && this.positionedElements.Count > 0) {
        foreach (PrintElement pe in this.positionedElements) {
          pe.Draw(this, yPos, e.Graphics, pageBounds);
        }
      }

      bool morePages = false;
      int elementsOnPage = 0;
      while (this.printIndex < this.printElements.Count) {
        PrintElement element = this.printElements[this.printIndex];
        float height = element.CalculateHeight(this, e.Graphics);
        if (yPos + height > pageBounds.Bottom) {
          if (elementsOnPage != 0) {
            morePages = true;
            break;
          }
        }
        element.Draw(this, yPos, e.Graphics, pageBounds);
        yPos += height;
        this.printIndex++;
        elementsOnPage++;
      }

      e.HasMorePages = morePages;
    }

    public Font PrintFont { get { return this.printFont; } }
    public Brush PrintBrush { get { return this.printBrush; } }
    public PrintElement Header { get { return this.header; } }
    public PrintElement Footer { get { return this.footer; } }

    public List<IPrintable> PrintObjects { get { return this.printObjects; } }
  }
}
