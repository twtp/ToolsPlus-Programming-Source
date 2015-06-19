namespace AlertStandalone
{
    partial class settingsFrm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.titleLbl = new System.Windows.Forms.Label();
            this.rightWrench = new System.Windows.Forms.PictureBox();
            this.leftWrench = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.rightWrench)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftWrench)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(438, 24);
            this.menuStrip1.TabIndex = 13;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // titleLbl
            // 
            this.titleLbl.AutoSize = true;
            this.titleLbl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.titleLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLbl.Location = new System.Drawing.Point(165, 2);
            this.titleLbl.Name = "titleLbl";
            this.titleLbl.Size = new System.Drawing.Size(97, 20);
            this.titleLbl.TabIndex = 14;
            this.titleLbl.Text = "SETTINGS";
            // 
            // rightWrench
            // 
            this.rightWrench.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.rightWrench.Image = global::AlertStandalone.Properties.Resources.wrench2;
            this.rightWrench.Location = new System.Drawing.Point(257, 2);
            this.rightWrench.Name = "rightWrench";
            this.rightWrench.Size = new System.Drawing.Size(22, 21);
            this.rightWrench.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.rightWrench.TabIndex = 15;
            this.rightWrench.TabStop = false;
            // 
            // leftWrench
            // 
            this.leftWrench.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.leftWrench.Image = global::AlertStandalone.Properties.Resources.wrench;
            this.leftWrench.Location = new System.Drawing.Point(147, 2);
            this.leftWrench.Name = "leftWrench";
            this.leftWrench.Size = new System.Drawing.Size(22, 21);
            this.leftWrench.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.leftWrench.TabIndex = 16;
            this.leftWrench.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox2.Image = global::AlertStandalone.Properties.Resources.Exit_transparent;
            this.pictureBox2.Location = new System.Drawing.Point(408, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(30, 24);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 12;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // settingsFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(438, 311);
            this.Controls.Add(this.rightWrench);
            this.Controls.Add(this.leftWrench);
            this.Controls.Add(this.titleLbl);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "settingsFrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "settingsFrm";
            this.Load += new System.EventHandler(this.settingsFrm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rightWrench)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftWrench)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.Label titleLbl;
        private System.Windows.Forms.PictureBox rightWrench;
        private System.Windows.Forms.PictureBox leftWrench;
    }
}