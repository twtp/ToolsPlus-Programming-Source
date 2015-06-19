namespace FeedGenerator
{
    partial class ModifyLine
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
            this.modifyOnceBtn = new System.Windows.Forms.Button();
            this.itemPanel = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // modifyOnceBtn
            // 
            this.modifyOnceBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.modifyOnceBtn.Location = new System.Drawing.Point(392, 107);
            this.modifyOnceBtn.Name = "modifyOnceBtn";
            this.modifyOnceBtn.Size = new System.Drawing.Size(93, 28);
            this.modifyOnceBtn.TabIndex = 0;
            this.modifyOnceBtn.Text = "Modify Once";
            this.modifyOnceBtn.UseVisualStyleBackColor = false;
            this.modifyOnceBtn.Click += new System.EventHandler(this.modifyOnceBtn_Click);
            // 
            // itemPanel
            // 
            this.itemPanel.AutoScroll = true;
            this.itemPanel.BackColor = System.Drawing.Color.Gray;
            this.itemPanel.Location = new System.Drawing.Point(1, 1);
            this.itemPanel.Name = "itemPanel";
            this.itemPanel.Size = new System.Drawing.Size(876, 100);
            this.itemPanel.TabIndex = 1;
            // 
            // ModifyLine
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(877, 140);
            this.Controls.Add(this.itemPanel);
            this.Controls.Add(this.modifyOnceBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ModifyLine";
            this.Text = "ModifyLine";
            this.Load += new System.EventHandler(this.ModifyLine_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button modifyOnceBtn;
        private System.Windows.Forms.Panel itemPanel;
    }
}