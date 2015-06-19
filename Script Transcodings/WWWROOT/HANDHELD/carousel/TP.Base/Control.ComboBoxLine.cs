using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Control {
  public class ComboBoxLine : IComparable {
    private string display;
    private object data;
    private System.Collections.IComparer comparator;

    private static DefaultComboBoxLineComparer defaultComparator = new DefaultComboBoxLineComparer();

    public ComboBoxLine(string display, object data) : this(display, data, defaultComparator) { }
    public ComboBoxLine(string display, object data, System.Collections.IComparer comparator) {
      this.display = display;
      this.data = data;
      this.comparator = comparator;
    }
    public string Display {
      get { return this.display; }
      set { this.display = value; }
    }
    public object Data {
      get { return this.data; }
      set { this.data = value; }
    }
    public System.Collections.IComparer Comparator {
      get { return this.comparator; }
      set { this.comparator = value; }
    }
    public override string ToString() {
      return this.Display;
    }

    #region IComparable Members

    public int CompareTo(object obj) {
      return this.comparator.Compare(this, obj);
    }

    #endregion
  }

  internal class DefaultComboBoxLineComparer : System.Collections.IComparer {
    public int Compare(object x, object y) {
      return String.Compare(
          ((ComboBoxLine)x).Display,
          ((ComboBoxLine)y).Display,
          true
        );
    }
  }
}
