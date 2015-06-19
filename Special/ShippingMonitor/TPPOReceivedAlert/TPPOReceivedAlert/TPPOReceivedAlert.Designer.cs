namespace TPPOReceivedAlert
{
    partial class TPPOReceivedAlert
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
            this.poReceivedHeaderLbl = new System.Windows.Forms.Label();
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            this.poHeaderLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // poReceivedHeaderLbl
            // 
            this.poReceivedHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.poReceivedHeaderLbl.Location = new System.Drawing.Point(520, 15);
            this.poReceivedHeaderLbl.Name = "poReceivedHeaderLbl";
            this.poReceivedHeaderLbl.Size = new System.Drawing.Size(280, 37);
            this.poReceivedHeaderLbl.TabIndex = 0;
            this.poReceivedHeaderLbl.Text = "PO Received:";
            // 
            // updateTmr
            // 
            this.updateTmr.Enabled = true;
            this.updateTmr.Interval = 58765;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // poHeaderLbl
            // 
            this.poHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.poHeaderLbl.Location = new System.Drawing.Point(3, 0);
            this.poHeaderLbl.Name = "poHeaderLbl";
            this.poHeaderLbl.Size = new System.Drawing.Size(324, 67);
            this.poHeaderLbl.TabIndex = 1;
            this.poHeaderLbl.Text = "PO Received:";
            // 
            // TPPOReceivedAlert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.Controls.Add(this.poHeaderLbl);
            this.Controls.Add(this.poReceivedHeaderLbl);
            this.Name = "TPPOReceivedAlert";
            this.Size = new System.Drawing.Size(784, 67);
            this.Load += new System.EventHandler(this.TPPOReceivedAlert_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label poReceivedHeaderLbl;
        private System.Windows.Forms.Timer updateTmr;
        private System.Windows.Forms.Label poHeaderLbl;
    }
}
