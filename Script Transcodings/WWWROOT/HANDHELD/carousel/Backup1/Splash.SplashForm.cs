using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TP.Base.Splash {
  public partial class SplashForm : Form {

    private static SplashForm splashForm = null;

    public static void ShowSplash() {
      ShowSplash(Properties.Resources.Splash_DefaultSplash);
    }
    public static void ShowSplash(int width, int height) {
      if (splashForm == null) {
        splashForm = new SplashForm();
        splashForm.Width = width;
        splashForm.Height = height;
        splashForm.Show();
      }
    }
    public static void ShowSplash(System.Drawing.Image pic) {
      if (splashForm == null) {
        splashForm = new SplashForm();
        splashForm.BackgroundImage = pic;
        splashForm.ClientSize = new System.Drawing.Size(pic.Width, pic.Height);
        splashForm.Show();
      }
    }
    public static void ShowSplash(int width, int height, System.Drawing.Image pic) {
      if (splashForm == null) {
        splashForm = new SplashForm();
        splashForm.Width = width;
        splashForm.Height = height;
        splashForm.BackgroundImage = pic;
        splashForm.Show();
      }
    }

    public static void ShowLoadingText(string text) {
      splashForm.SetLoadingText(text);
    }
    public static void ShowLoadingText(string text, System.Drawing.Color color, System.Drawing.Point location) {
      splashForm.SetLoadingText(text, color, location);
    }

    public static void ClearSplash() {
      if (splashForm != null) {
        splashForm.Close();
        splashForm = null;
      }
    }





    public SplashForm() {
      InitializeComponent();
    }

    public void SetLoadingText(string text) {
      this.SetLoadingText(text, this.LoadingText.ForeColor, this.LoadingText.Location);
    }
    public void SetLoadingText(string text, Color color, Point location) {
      if (text == null) {
        this.LoadingText.Visible = false;
      }
      else {
        this.LoadingText.Text = text;
        this.LoadingText.ForeColor = color;
        this.LoadingText.Location = location;
        this.LoadingText.Visible = true;
      }
      this.Invalidate();
      Application.DoEvents();
    }
  }
}