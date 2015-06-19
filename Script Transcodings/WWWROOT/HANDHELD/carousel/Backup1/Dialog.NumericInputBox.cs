using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TP.Base.Dialog {
  public partial class NumericInputBox : Form {

    private decimal? returnValue = null;
    private decimal? ReturnValue { get { return this.returnValue; } }

    public static decimal? Show(string prompt) {
      return Show(prompt, prompt, 0, 0, 0, 99999);
    }
    public static decimal? Show(string prompt, string title) {
      return Show(prompt, title, 0, 0, 0, 99999);
    }
    public static decimal? Show(string prompt, string title, decimal defaultValue) {
      return Show(prompt, title, defaultValue, 0, 0, 99999);
    }
    public static decimal? Show(string prompt, string title, decimal defaultValue, int decimalPlaces, decimal minValue, decimal maxValue) {
      NumericInputBox ib = new NumericInputBox(prompt, title, defaultValue, decimalPlaces, minValue, maxValue);
      ib.ShowDialog();
      decimal? retval = ib.ReturnValue;
      ib.Dispose();
      return retval;
    }

    private NumericInputBox(string prompt, string title, decimal defaultValue, int decimalPlaces, decimal minValue, decimal maxValue) {
      InitializeComponent();

      this.PromptLabel.Text = prompt;
      this.Text = title;

      this.InputField.DecimalPlaces = decimalPlaces;
      this.InputField.Minimum = minValue;
      this.InputField.Maximum = maxValue;
      this.InputField.Value = defaultValue;
    }

    private void OKButton_Click(object sender, EventArgs e) {
      this.returnValue = this.InputField.Value;
      this.Close();
    }

    private void CancelInputButton_Click(object sender, EventArgs e) {
      this.returnValue = null;
      this.Close();
    }
  }
}