namespace TP.Base.Dialog {
  partial class NumericInputBox {
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
      this.CancelInputButton = new System.Windows.Forms.Button();
      this.OKButton = new System.Windows.Forms.Button();
      this.PromptLabel = new System.Windows.Forms.Label();
      this.InputField = new System.Windows.Forms.NumericUpDown();
      ((System.ComponentModel.ISupportInitialize)(this.InputField)).BeginInit();
      this.SuspendLayout();
      // 
      // CancelInputButton
      // 
      this.CancelInputButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
      this.CancelInputButton.Location = new System.Drawing.Point(458, 41);
      this.CancelInputButton.Name = "CancelInputButton";
      this.CancelInputButton.Size = new System.Drawing.Size(75, 23);
      this.CancelInputButton.TabIndex = 6;
      this.CancelInputButton.Text = "Cancel";
      this.CancelInputButton.UseVisualStyleBackColor = true;
      this.CancelInputButton.Click += new System.EventHandler(this.CancelInputButton_Click);
      // 
      // OKButton
      // 
      this.OKButton.Location = new System.Drawing.Point(458, 12);
      this.OKButton.Name = "OKButton";
      this.OKButton.Size = new System.Drawing.Size(75, 23);
      this.OKButton.TabIndex = 5;
      this.OKButton.Text = "OK";
      this.OKButton.UseVisualStyleBackColor = true;
      this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
      // 
      // PromptLabel
      // 
      this.PromptLabel.AutoSize = true;
      this.PromptLabel.Location = new System.Drawing.Point(12, 9);
      this.PromptLabel.Name = "PromptLabel";
      this.PromptLabel.Size = new System.Drawing.Size(119, 13);
      this.PromptLabel.TabIndex = 4;
      this.PromptLabel.Text = "PROMPT GOES HERE";
      // 
      // InputField
      // 
      this.InputField.Location = new System.Drawing.Point(12, 68);
      this.InputField.Name = "InputField";
      this.InputField.Size = new System.Drawing.Size(134, 20);
      this.InputField.TabIndex = 7;
      // 
      // NumericInputBox
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(543, 100);
      this.Controls.Add(this.InputField);
      this.Controls.Add(this.CancelInputButton);
      this.Controls.Add(this.OKButton);
      this.Controls.Add(this.PromptLabel);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.MaximizeBox = false;
      this.MinimizeBox = false;
      this.Name = "NumericInputBox";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "NumericInputBox";
      ((System.ComponentModel.ISupportInitialize)(this.InputField)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Button CancelInputButton;
    private System.Windows.Forms.Button OKButton;
    private System.Windows.Forms.Label PromptLabel;
    private System.Windows.Forms.NumericUpDown InputField;
  }
}