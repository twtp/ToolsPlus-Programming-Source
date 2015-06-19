namespace TPTickerCtl
{
    partial class TPTicker
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.headerBackgroundPic = new System.Windows.Forms.PictureBox();
            this.headerTextLbl = new System.Windows.Forms.Label();
            this.tickerTextLbl = new System.Windows.Forms.Label();
            this.tickerTmr = new System.Windows.Forms.Timer(this.components);
            this.tickerBackgroundPic = new System.Windows.Forms.PictureBox();
            this.tickerBackDropColorLbl = new System.Windows.Forms.Label();
            this.tickerEndPic = new System.Windows.Forms.PictureBox();
            this.headerSubLbl = new System.Windows.Forms.Label();
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.headerBackgroundPic)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tickerBackgroundPic)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tickerEndPic)).BeginInit();
            this.SuspendLayout();
            // 
            // headerBackgroundPic
            // 
            this.headerBackgroundPic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.headerBackgroundPic.Location = new System.Drawing.Point(0, 0);
            this.headerBackgroundPic.Name = "headerBackgroundPic";
            this.headerBackgroundPic.Size = new System.Drawing.Size(319, 97);
            this.headerBackgroundPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.headerBackgroundPic.TabIndex = 0;
            this.headerBackgroundPic.TabStop = false;
            // 
            // headerTextLbl
            // 
            this.headerTextLbl.BackColor = System.Drawing.Color.Transparent;
            this.headerTextLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headerTextLbl.Location = new System.Drawing.Point(17, 21);
            this.headerTextLbl.Name = "headerTextLbl";
            this.headerTextLbl.Size = new System.Drawing.Size(284, 55);
            this.headerTextLbl.TabIndex = 1;
            this.headerTextLbl.Text = "DEFAULT";
            this.headerTextLbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tickerTextLbl
            // 
            this.tickerTextLbl.BackColor = System.Drawing.Color.Transparent;
            this.tickerTextLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tickerTextLbl.Location = new System.Drawing.Point(325, 0);
            this.tickerTextLbl.Name = "tickerTextLbl";
            this.tickerTextLbl.Size = new System.Drawing.Size(587, 97);
            this.tickerTextLbl.TabIndex = 2;
            this.tickerTextLbl.Text = "DEFAULT";
            this.tickerTextLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tickerTextLbl.Click += new System.EventHandler(this.tickerTextLbl_Click);
            // 
            // tickerTmr
            // 
            this.tickerTmr.Interval = 15;
            this.tickerTmr.Tick += new System.EventHandler(this.tickerTmr_Tick);
            // 
            // tickerBackgroundPic
            // 
            this.tickerBackgroundPic.Location = new System.Drawing.Point(297, 0);
            this.tickerBackgroundPic.Name = "tickerBackgroundPic";
            this.tickerBackgroundPic.Size = new System.Drawing.Size(319, 97);
            this.tickerBackgroundPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.tickerBackgroundPic.TabIndex = 3;
            this.tickerBackgroundPic.TabStop = false;
            // 
            // tickerBackDropColorLbl
            // 
            this.tickerBackDropColorLbl.BackColor = System.Drawing.Color.Transparent;
            this.tickerBackDropColorLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tickerBackDropColorLbl.Location = new System.Drawing.Point(245, 0);
            this.tickerBackDropColorLbl.Name = "tickerBackDropColorLbl";
            this.tickerBackDropColorLbl.Size = new System.Drawing.Size(587, 97);
            this.tickerBackDropColorLbl.TabIndex = 4;
            this.tickerBackDropColorLbl.Text = "DEFAULT";
            this.tickerBackDropColorLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.tickerBackDropColorLbl.Click += new System.EventHandler(this.tickerBackDropColorLbl_Click);
            // 
            // tickerEndPic
            // 
            this.tickerEndPic.Location = new System.Drawing.Point(589, 3);
            this.tickerEndPic.Name = "tickerEndPic";
            this.tickerEndPic.Size = new System.Drawing.Size(47, 93);
            this.tickerEndPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.tickerEndPic.TabIndex = 5;
            this.tickerEndPic.TabStop = false;
            // 
            // headerSubLbl
            // 
            this.headerSubLbl.BackColor = System.Drawing.Color.Transparent;
            this.headerSubLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headerSubLbl.Location = new System.Drawing.Point(50, 21);
            this.headerSubLbl.Name = "headerSubLbl";
            this.headerSubLbl.Size = new System.Drawing.Size(284, 55);
            this.headerSubLbl.TabIndex = 6;
            this.headerSubLbl.Text = "sublabel";
            // 
            // updateTmr
            // 
            this.updateTmr.Interval = 1000;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // TPTicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.Controls.Add(this.headerSubLbl);
            this.Controls.Add(this.tickerEndPic);
            this.Controls.Add(this.tickerBackDropColorLbl);
            this.Controls.Add(this.tickerBackgroundPic);
            this.Controls.Add(this.tickerTextLbl);
            this.Controls.Add(this.headerTextLbl);
            this.Controls.Add(this.headerBackgroundPic);
            this.Name = "TPTicker";
            this.Size = new System.Drawing.Size(912, 98);
            this.Load += new System.EventHandler(this.TPTicker_Load);
            ((System.ComponentModel.ISupportInitialize)(this.headerBackgroundPic)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tickerBackgroundPic)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tickerEndPic)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox headerBackgroundPic;
        private System.Windows.Forms.Label headerTextLbl;
        private System.Windows.Forms.Label tickerTextLbl;
        private System.Windows.Forms.Timer tickerTmr;
        private System.Windows.Forms.PictureBox tickerBackgroundPic;
        private System.Windows.Forms.Label tickerBackDropColorLbl;
        private System.Windows.Forms.PictureBox tickerEndPic;
        private System.Windows.Forms.Label headerSubLbl;
        private System.Windows.Forms.Timer updateTmr;
    }
}
