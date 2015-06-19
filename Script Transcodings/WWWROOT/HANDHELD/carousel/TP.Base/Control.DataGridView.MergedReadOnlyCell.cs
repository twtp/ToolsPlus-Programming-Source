using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;
using System.Windows.Forms;

namespace TP.Base.Control.DataGridView {
  public class MergedReadOnlyCell : DataGridViewTextBoxCell {

    private int leftColumn = 0;
    private int rightColumn = 0;
    private StringAlignment alignment = StringAlignment.Near;

    public int LeftColumn { get { return this.leftColumn; } set { this.leftColumn = value; } }
    public int RightColumn { get { return this.rightColumn; } set { this.rightColumn = value; } }
    public StringAlignment Alignment { get { return this.alignment; } set { this.alignment = value; } }

    protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex, DataGridViewElementStates cellState, object value, object formattedValue, string errorText, DataGridViewCellStyle cellStyle, DataGridViewAdvancedBorderStyle advancedBorderStyle, DataGridViewPaintParts paintParts) {
      // background
      graphics.FillRectangle(SystemBrushes.Control, cellBounds);
      // bottom line
      graphics.DrawLine(SystemPens.ControlDark, cellBounds.Left, cellBounds.Bottom - 1, cellBounds.Right, cellBounds.Bottom - 1);
      // right line if last column in merge
      if (this.ColumnIndex == this.rightColumn) {
        graphics.DrawLine(SystemPens.ControlDark, cellBounds.Right - 1, cellBounds.Top, cellBounds.Right - 1, cellBounds.Bottom);
      }

      // width of text area, width of all merged cells
      int width = 0;
      for (int i = this.leftColumn; i <= this.rightColumn; i++) {
        width += this.OwningRow.Cells[i].Size.Width;
      }

      // x-offset of text area, width of previous cells
      int offset = 0;
      for (int i = this.leftColumn; i < this.ColumnIndex; i++) {
        offset += this.OwningRow.Cells[i].Size.Width;
      }

      // formatting for string
      StringFormat sf = new StringFormat();
      sf.Alignment = this.alignment;
      sf.LineAlignment = StringAlignment.Center;
      sf.Trimming = StringTrimming.EllipsisCharacter;

      graphics.DrawString(
          this.OwningRow.Cells[this.leftColumn].Value.ToString(), // always grab first cell's text
          this.DataGridView.Font,
          new SolidBrush(this.DataGridView.ForeColor),
          new RectangleF(
              cellBounds.Left - offset,
              cellBounds.Top,
              width,
              cellBounds.Height
            ),
          sf
        );
    }

  }
}
