namespace TPTransferAlert
{
    partial class TPTransferAlert
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
            this.transferDetailsLbl = new System.Windows.Forms.Label();
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            this.transferHeaderLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // transferDetailsLbl
            // 
            this.transferDetailsLbl.AutoSize = true;
            this.transferDetailsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.transferDetailsLbl.Location = new System.Drawing.Point(173, 0);
            this.transferDetailsLbl.Name = "transferDetailsLbl";
            this.transferDetailsLbl.Size = new System.Drawing.Size(96, 29);
            this.transferDetailsLbl.TabIndex = 0;
            this.transferDetailsLbl.Text = "details..";
            this.transferDetailsLbl.Click += new System.EventHandler(this.transferHeaderLbl_Click);
            // 
            // updateTmr
            // 
            this.updateTmr.Enabled = true;
            this.updateTmr.Interval = 54848;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // transferHeaderLbl
            // 
            this.transferHeaderLbl.AutoSize = true;
            this.transferHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.transferHeaderLbl.Location = new System.Drawing.Point(3, 0);
            this.transferHeaderLbl.Name = "transferHeaderLbl";
            this.transferHeaderLbl.Size = new System.Drawing.Size(216, 55);
            this.transferHeaderLbl.TabIndex = 1;
            this.transferHeaderLbl.Text = "Transfer:";
            this.transferHeaderLbl.Click += new System.EventHandler(this.transferHeaderLbl_Click_1);
            // 
            // TPTransferAlert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.Controls.Add(this.transferHeaderLbl);
            this.Controls.Add(this.transferDetailsLbl);
            this.Name = "TPTransferAlert";
            this.Size = new System.Drawing.Size(772, 59);
            this.Load += new System.EventHandler(this.TPTransferAlert_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label transferDetailsLbl;
        private System.Windows.Forms.Timer updateTmr;
        private System.Windows.Forms.Label transferHeaderLbl;
    }
}
