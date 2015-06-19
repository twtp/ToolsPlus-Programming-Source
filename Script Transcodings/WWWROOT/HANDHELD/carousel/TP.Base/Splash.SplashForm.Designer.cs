namespace TP.Base.Splash {
  partial class SplashForm {
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
      this.LoadingText = new System.Windows.Forms.Label();
      this.SuspendLayout();
      //
      // LoadingText
      //
      this.LoadingText.AutoSize = true;
      this.LoadingText.BackColor = System.Drawing.Color.Transparent;
      this.LoadingText.Font = new System.Drawing.Font(this.LoadingText.Font, System.Drawing.FontStyle.Bold);
      this.LoadingText.Location = new System.Drawing.Point(0, 0);
      this.LoadingText.Name = "LoadingText";
      this.LoadingText.Visible = false;
      //
      // SplashForm
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.BackgroundImage = global::TP.Base.Properties.Resources.Splash_DefaultSplash;
      this.ClientSize = new System.Drawing.Size(476, 125);
      this.ControlBox = false;
      this.Controls.Add(this.LoadingText);
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.Name = "SplashForm";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "Starting up...";
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.Label LoadingText;
  }
}