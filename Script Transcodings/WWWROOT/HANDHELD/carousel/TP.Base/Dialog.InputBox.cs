using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TP.Base.Dialog {
  public partial class InputBox : Form {

    private string returnValue = string.Empty;
    private string ReturnValue { get { return this.returnValue; } }

    public static string Show(string prompt) {
      return Show(prompt, prompt, string.Empty);
    }
    public static string Show(string prompt, string title) {
      return Show(prompt, title, string.Empty);
    }
    public static string Show(string prompt, string title, string defaultValue) {
      InputBox ib = new InputBox(prompt, title, defaultValue);
      ib.ShowDialog();
      string retval = ib.ReturnValue;
      ib.Dispose();
      return retval;
    }

    private InputBox(string prompt, string title, string defaultValue) {
      InitializeComponent();

      this.PromptLabel.Text = prompt;
      this.Text = title;
      this.InputField.Text = defaultValue;
    }

    private void OKButton_Click(object sender, EventArgs e) {
      this.returnValue = this.InputField.Text;
      this.Close();
    }

    private void CancelInputButton_Click(object sender, EventArgs e) {
      this.returnValue = string.Empty;
      this.Close();
    }

  }
}