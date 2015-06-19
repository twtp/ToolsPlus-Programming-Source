namespace TP.Base.Dialog {
  partial class InputBox {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing) {
      if (disposing && (components != null)) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
      this.PromptLabel = new System.Windows.Forms.Label();
      this.InputField = new System.Windows.Forms.TextBox();
      this.OKButton = new System.Windows.Forms.Button();
      this.CancelInputButton = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // PromptLabel
      // 
      this.PromptLabel.AutoSize = true;
      this.PromptLabel.Location = new System.Drawing.Point(12, 9);
      this.PromptLabel.Name = "PromptLabel";
      this.PromptLabel.Size = new System.Drawing.Size(119, 13);
      this.PromptLabel.TabIndex = 0;
      this.PromptLabel.Text = "PROMPT GOES HERE";
      // 
      // InputField
      // 
      this.InputField.Location = new System.Drawing.Point(12, 70);
      this.InputField.Name = "InputField";
      this.InputField.Size = new System.Drawing.Size(521, 20);
      this.InputField.TabIndex = 1;
      this.InputField.Text = "DEFAULT GOES HERE";
      // 
      // OKButton
      // 
      this.OKButton.Location = new System.Drawing.Point(458, 12);
      this.OKButton.Name = "OKButton";
      this.OKButton.Size = new System.Drawing.Size(75, 23);
      this.OKButton.TabIndex = 2;
      this.OKButton.Text = "OK";
      this.OKButton.UseVisualStyleBackColor = true;
      this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
      // 
      // CancelInputButton
      // 
      this.CancelInputButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
      this.CancelInputButton.Location = new System.Drawing.Point(458, 41);
      this.CancelInputButton.Name = "CancelInputButton";
      this.CancelInputButton.Size = new System.Drawing.Size(75, 23);
      this.CancelInputButton.TabIndex = 3;
      this.CancelInputButton.Text = "Cancel";
      this.CancelInputButton.UseVisualStyleBackColor = true;
      this.CancelInputButton.Click += new System.EventHandler(this.CancelInputButton_Click);
      // 
      // InputBox
      // 
      this.AcceptButton = this.OKButton;
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.CancelButton = this.CancelInputButton;
      this.ClientSize = new System.Drawing.Size(545, 102);
      this.Controls.Add(this.CancelInputButton);
      this.Controls.Add(this.OKButton);
      this.Controls.Add(this.InputField);
      this.Controls.Add(this.PromptLabel);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.MaximizeBox = false;
      this.MinimizeBox = false;
      this.Name = "InputBox";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "InputBox";
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Label PromptLabel;
    private System.Windows.Forms.TextBox InputField;
    private System.Windows.Forms.Button OKButton;
    private System.Windows.Forms.Button CancelInputButton;
  }
}