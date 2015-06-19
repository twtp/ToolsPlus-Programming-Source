using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Control {
  public class ListViewSelectDeselectAllCheckBox : System.Windows.Forms.CheckBox {

    public ListViewSelectDeselectAllCheckBox()
      : base() {
      // nothing, for designer support
    }
    public ListViewSelectDeselectAllCheckBox(System.Windows.Forms.ListView linkedLV)
      : base() {
      this.LinkedListView = linkedLV; // handles event association too
    }

    private System.Windows.Forms.ListView linkedLV = null;
    public System.Windows.Forms.ListView LinkedListView {
      get {
        return this.linkedLV;
      }
      set {
        if (this.linkedLV != null) {
          this.linkedLV.ItemChecked -= new System.Windows.Forms.ItemCheckedEventHandler(this.linkedLV_ItemChecked);
          this.linkedLV = null;
        }
        this.linkedLV = value;
        if (this.linkedLV != null) {
          this.linkedLV.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.linkedLV_ItemChecked);
        }
      }
    }

    private int stopEvents = 0;

    protected override void OnCheckStateChanged(EventArgs e) {
      base.OnCheckStateChanged(e);

      if (this.stopEvents > 0) {
        return;
      }

      this.stopEvents++;
      switch (this.CheckState) {
        case System.Windows.Forms.CheckState.Checked:
          // select all
          foreach (System.Windows.Forms.ListViewItem lvi in this.linkedLV.Items) {
            if (lvi.Checked == false) {
              lvi.Checked = true;
            }
          }
          break;
        case System.Windows.Forms.CheckState.Unchecked:
          // deselect all
          foreach (System.Windows.Forms.ListViewItem lvi in this.linkedLV.Items) {
            if (lvi.Checked == true) {
              lvi.Checked = false;
            }
          }
          break;
        default:
          // indeterminate, should never happen
          break;
      }
      this.stopEvents--;
    }

    private void linkedLV_ItemChecked(object sender, System.Windows.Forms.ItemCheckedEventArgs e) {
      if (this.stopEvents > 0) {
        return;
      }

      this.stopEvents++;
      if (this.linkedLV.CheckedItems.Count == 0) {
        if (this.CheckState != System.Windows.Forms.CheckState.Unchecked) {
          this.CheckState = System.Windows.Forms.CheckState.Unchecked;
        }
      }
      else if (this.linkedLV.Items.Count == this.linkedLV.CheckedItems.Count) {
        if (this.CheckState != System.Windows.Forms.CheckState.Checked) {
          this.CheckState = System.Windows.Forms.CheckState.Checked;
        }
      }
      else {
        if (this.CheckState != System.Windows.Forms.CheckState.Indeterminate) {
          this.CheckState = System.Windows.Forms.CheckState.Indeterminate;
        }
      }
      this.stopEvents--;
    }
  }
}
