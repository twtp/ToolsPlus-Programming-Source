namespace WindowsFormsApplication1
{
    partial class Form1
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
            this.scriptViewerLV = new System.Windows.Forms.ListView();
            this.label1 = new System.Windows.Forms.Label();
            this.testBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // scriptViewerLV
            // 
            this.scriptViewerLV.CheckBoxes = true;
            this.scriptViewerLV.FullRowSelect = true;
            this.scriptViewerLV.Location = new System.Drawing.Point(12, 21);
            this.scriptViewerLV.Name = "scriptViewerLV";
            this.scriptViewerLV.Size = new System.Drawing.Size(786, 404);
            this.scriptViewerLV.TabIndex = 0;
            this.scriptViewerLV.UseCompatibleStateImageBehavior = false;
            this.scriptViewerLV.View = System.Windows.Forms.View.Details;
            this.scriptViewerLV.SelectedIndexChanged += new System.EventHandler(this.scriptViewerLV_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(301, 454);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(294, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "For demo purposes, highlight a script and click \'Test\' to run it.";
            // 
            // testBtn
            // 
            this.testBtn.BackColor = System.Drawing.Color.IndianRed;
            this.testBtn.Enabled = false;
            this.testBtn.Location = new System.Drawing.Point(629, 438);
            this.testBtn.Name = "testBtn";
            this.testBtn.Size = new System.Drawing.Size(169, 44);
            this.testBtn.TabIndex = 2;
            this.testBtn.Text = "Test";
            this.testBtn.UseVisualStyleBackColor = false;
            this.testBtn.Click += new System.EventHandler(this.testBtn_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(815, 522);
            this.Controls.Add(this.testBtn);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.scriptViewerLV);
            this.Name = "Form1";
            this.Text = "Calling External Code in C#";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView scriptViewerLV;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button testBtn;

    }
}

