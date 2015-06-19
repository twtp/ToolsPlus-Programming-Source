namespace TPClock
{
    partial class TPClock
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
            this.timeLbl = new System.Windows.Forms.Label();
            this.timeTmr = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // timeLbl
            // 
            this.timeLbl.BackColor = System.Drawing.Color.Transparent;
            this.timeLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeLbl.Location = new System.Drawing.Point(3, 0);
            this.timeLbl.Name = "timeLbl";
            this.timeLbl.Size = new System.Drawing.Size(424, 166);
            this.timeLbl.TabIndex = 0;
            this.timeLbl.Text = "0:00";
            this.timeLbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.timeLbl.Click += new System.EventHandler(this.timeLbl_Click);
            // 
            // timeTmr
            // 
            this.timeTmr.Enabled = true;
            this.timeTmr.Interval = 1000;
            this.timeTmr.Tick += new System.EventHandler(this.timeTmr_Tick);
            // 
            // UserControl1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.timeLbl);
            this.Name = "UserControl1";
            this.Size = new System.Drawing.Size(430, 166);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label timeLbl;
        private System.Windows.Forms.Timer timeTmr;
    }
}
