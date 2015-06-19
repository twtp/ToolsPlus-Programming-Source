namespace TPWillCall
{
    partial class TPWillCall
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
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            this.strobeTmr = new System.Windows.Forms.Timer(this.components);
            this.WillCallHeader = new System.Windows.Forms.Label();
            this.detailsLbl = new System.Windows.Forms.Label();
            this.subDetailsLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // updateTmr
            // 
            this.updateTmr.Enabled = true;
            this.updateTmr.Interval = 60000;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // strobeTmr
            // 
            this.strobeTmr.Tick += new System.EventHandler(this.strobeTmr_Tick);
            // 
            // WillCallHeader
            // 
            this.WillCallHeader.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.WillCallHeader.Location = new System.Drawing.Point(0, 0);
            this.WillCallHeader.Name = "WillCallHeader";
            this.WillCallHeader.Size = new System.Drawing.Size(590, 73);
            this.WillCallHeader.TabIndex = 0;
            this.WillCallHeader.Text = "Will Call";
            this.WillCallHeader.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // detailsLbl
            // 
            this.detailsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.detailsLbl.Location = new System.Drawing.Point(0, 73);
            this.detailsLbl.Name = "detailsLbl";
            this.detailsLbl.Size = new System.Drawing.Size(590, 37);
            this.detailsLbl.TabIndex = 1;
            this.detailsLbl.Text = "details";
            this.detailsLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // subDetailsLbl
            // 
            this.subDetailsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.subDetailsLbl.Location = new System.Drawing.Point(0, 120);
            this.subDetailsLbl.Name = "subDetailsLbl";
            this.subDetailsLbl.Size = new System.Drawing.Size(590, 37);
            this.subDetailsLbl.TabIndex = 2;
            this.subDetailsLbl.Text = "sub-details";
            this.subDetailsLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.subDetailsLbl.Click += new System.EventHandler(this.subDetailsLbl_Click);
            // 
            // TPWillCall
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.Controls.Add(this.subDetailsLbl);
            this.Controls.Add(this.detailsLbl);
            this.Controls.Add(this.WillCallHeader);
            this.Name = "TPWillCall";
            this.Size = new System.Drawing.Size(593, 163);
            this.Load += new System.EventHandler(this.UserControl1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer updateTmr;
        private System.Windows.Forms.Timer strobeTmr;
        private System.Windows.Forms.Label WillCallHeader;
        private System.Windows.Forms.Label detailsLbl;
        private System.Windows.Forms.Label subDetailsLbl;
    }
}
