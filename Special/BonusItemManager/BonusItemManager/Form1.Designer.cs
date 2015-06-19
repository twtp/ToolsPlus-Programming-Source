namespace BonusItemManager
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
            this.bodyTxt = new System.Windows.Forms.TextBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.headerTxt = new System.Windows.Forms.TextBox();
            this.footerTxt = new System.Windows.Forms.TextBox();
            this.dbugTxt = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.perlBodyTxt = new System.Windows.Forms.TextBox();
            this.perlHeaderTxt = new System.Windows.Forms.TextBox();
            this.perlFooterTxt = new System.Windows.Forms.TextBox();
            this.listView2 = new System.Windows.Forms.ListView();
            this.verifyBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bodyTxt
            // 
            this.bodyTxt.Location = new System.Drawing.Point(895, 90);
            this.bodyTxt.Multiline = true;
            this.bodyTxt.Name = "bodyTxt";
            this.bodyTxt.Size = new System.Drawing.Size(323, 255);
            this.bodyTxt.TabIndex = 0;
            this.bodyTxt.TextChanged += new System.EventHandler(this.bodyTxt_TextChanged);
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(8, 12);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(866, 265);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // headerTxt
            // 
            this.headerTxt.Location = new System.Drawing.Point(895, 12);
            this.headerTxt.Multiline = true;
            this.headerTxt.Name = "headerTxt";
            this.headerTxt.Size = new System.Drawing.Size(323, 72);
            this.headerTxt.TabIndex = 2;
            // 
            // footerTxt
            // 
            this.footerTxt.Location = new System.Drawing.Point(895, 351);
            this.footerTxt.Multiline = true;
            this.footerTxt.Name = "footerTxt";
            this.footerTxt.Size = new System.Drawing.Size(323, 72);
            this.footerTxt.TabIndex = 3;
            // 
            // dbugTxt
            // 
            this.dbugTxt.Location = new System.Drawing.Point(895, 429);
            this.dbugTxt.Multiline = true;
            this.dbugTxt.Name = "dbugTxt";
            this.dbugTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dbugTxt.Size = new System.Drawing.Size(323, 100);
            this.dbugTxt.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(257, 283);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(78, 38);
            this.button1.TabIndex = 5;
            this.button1.Text = "Add";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(427, 283);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(78, 38);
            this.button2.TabIndex = 6;
            this.button2.Text = "Remove";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(343, 283);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(78, 38);
            this.button3.TabIndex = 7;
            this.button3.Text = "Edit";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // perlBodyTxt
            // 
            this.perlBodyTxt.Location = new System.Drawing.Point(546, 501);
            this.perlBodyTxt.Multiline = true;
            this.perlBodyTxt.Name = "perlBodyTxt";
            this.perlBodyTxt.Size = new System.Drawing.Size(336, 131);
            this.perlBodyTxt.TabIndex = 8;
            // 
            // perlHeaderTxt
            // 
            this.perlHeaderTxt.Location = new System.Drawing.Point(547, 423);
            this.perlHeaderTxt.Multiline = true;
            this.perlHeaderTxt.Name = "perlHeaderTxt";
            this.perlHeaderTxt.Size = new System.Drawing.Size(336, 72);
            this.perlHeaderTxt.TabIndex = 9;
            // 
            // perlFooterTxt
            // 
            this.perlFooterTxt.Location = new System.Drawing.Point(546, 638);
            this.perlFooterTxt.Multiline = true;
            this.perlFooterTxt.Name = "perlFooterTxt";
            this.perlFooterTxt.Size = new System.Drawing.Size(336, 72);
            this.perlFooterTxt.TabIndex = 10;
            // 
            // listView2
            // 
            this.listView2.Location = new System.Drawing.Point(12, 423);
            this.listView2.Name = "listView2";
            this.listView2.Size = new System.Drawing.Size(512, 287);
            this.listView2.TabIndex = 11;
            this.listView2.UseCompatibleStateImageBehavior = false;
            // 
            // verifyBtn
            // 
            this.verifyBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.verifyBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.verifyBtn.Location = new System.Drawing.Point(564, 283);
            this.verifyBtn.Name = "verifyBtn";
            this.verifyBtn.Size = new System.Drawing.Size(78, 38);
            this.verifyBtn.TabIndex = 12;
            this.verifyBtn.Text = "Verify";
            this.verifyBtn.UseVisualStyleBackColor = false;
            this.verifyBtn.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(882, 334);
            this.Controls.Add(this.verifyBtn);
            this.Controls.Add(this.listView2);
            this.Controls.Add(this.perlFooterTxt);
            this.Controls.Add(this.perlHeaderTxt);
            this.Controls.Add(this.perlBodyTxt);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dbugTxt);
            this.Controls.Add(this.footerTxt);
            this.Controls.Add(this.headerTxt);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.bodyTxt);
            this.Name = "Form1";
            this.Text = "Bonus Item Manager";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox bodyTxt;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.TextBox headerTxt;
        private System.Windows.Forms.TextBox footerTxt;
        private System.Windows.Forms.TextBox dbugTxt;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox perlBodyTxt;
        private System.Windows.Forms.TextBox perlHeaderTxt;
        private System.Windows.Forms.TextBox perlFooterTxt;
        private System.Windows.Forms.ListView listView2;
        private System.Windows.Forms.Button verifyBtn;
    }
}

