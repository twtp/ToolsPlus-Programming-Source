using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Control {
  public class TPListView : System.Windows.Forms.ListView {

    protected TPListView() : base() {
      this.FullRowSelect = true;
      this.GridLines = true;
      this.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
      this.HideSelection = false;
      this.MultiSelect = false;
      this.ShowItemToolTips = true;
      this.UseCompatibleStateImageBehavior = false;
      this.View = System.Windows.Forms.View.Details;
    }

    protected System.Windows.Forms.ColumnHeader addColumn(string text, int width) {
      return this.addColumn(text, width, System.Windows.Forms.HorizontalAlignment.Left);
    }
    protected System.Windows.Forms.ColumnHeader addColumn(string text, int width, System.Windows.Forms.HorizontalAlignment textAlign) {
      return this.Columns.Add(text, width, textAlign);
    }

    protected System.Windows.Forms.ListViewGroup addGroup(string text) {
      return this.Groups.Add(text, text);
    }

  }
}
