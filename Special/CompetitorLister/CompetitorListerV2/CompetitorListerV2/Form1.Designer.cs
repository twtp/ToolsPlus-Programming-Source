namespace CompetitorListerV2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.enabledLst = new System.Windows.Forms.ListBox();
            this.disabledLst = new System.Windows.Forms.ListBox();
            this.moveToDisabledBtn = new System.Windows.Forms.Button();
            this.moveToEnabledBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // enabledLst
            // 
            this.enabledLst.FormattingEnabled = true;
            this.enabledLst.Location = new System.Drawing.Point(8, 24);
            this.enabledLst.Name = "enabledLst";
            this.enabledLst.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.enabledLst.Size = new System.Drawing.Size(176, 238);
            this.enabledLst.TabIndex = 0;
            this.enabledLst.DoubleClick += new System.EventHandler(this.enabledLst_DoubleClick);
            // 
            // disabledLst
            // 
            this.disabledLst.FormattingEnabled = true;
            this.disabledLst.Location = new System.Drawing.Point(239, 24);
            this.disabledLst.Name = "disabledLst";
            this.disabledLst.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.disabledLst.Size = new System.Drawing.Size(176, 238);
            this.disabledLst.TabIndex = 1;
            this.disabledLst.DoubleClick += new System.EventHandler(this.disabledLst_DoubleClick);
            // 
            // moveToDisabledBtn
            // 
            this.moveToDisabledBtn.Location = new System.Drawing.Point(190, 71);
            this.moveToDisabledBtn.Name = "moveToDisabledBtn";
            this.moveToDisabledBtn.Size = new System.Drawing.Size(43, 30);
            this.moveToDisabledBtn.TabIndex = 2;
            this.moveToDisabledBtn.Text = ">>";
            this.moveToDisabledBtn.UseVisualStyleBackColor = true;
            this.moveToDisabledBtn.Click += new System.EventHandler(this.moveToDisabledBtn_Click);
            // 
            // moveToEnabledBtn
            // 
            this.moveToEnabledBtn.Location = new System.Drawing.Point(190, 153);
            this.moveToEnabledBtn.Name = "moveToEnabledBtn";
            this.moveToEnabledBtn.Size = new System.Drawing.Size(43, 30);
            this.moveToEnabledBtn.TabIndex = 3;
            this.moveToEnabledBtn.Text = "<<";
            this.moveToEnabledBtn.UseVisualStyleBackColor = true;
            this.moveToEnabledBtn.Click += new System.EventHandler(this.moveToEnabledBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(66, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Active";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(301, 8);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Blacklisted";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(427, 271);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.moveToEnabledBtn);
            this.Controls.Add(this.moveToDisabledBtn);
            this.Controls.Add(this.disabledLst);
            this.Controls.Add(this.enabledLst);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Competitor Lister";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox enabledLst;
        private System.Windows.Forms.ListBox disabledLst;
        private System.Windows.Forms.Button moveToDisabledBtn;
        private System.Windows.Forms.Button moveToEnabledBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

