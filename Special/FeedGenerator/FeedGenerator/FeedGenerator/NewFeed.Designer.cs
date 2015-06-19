namespace FeedGenerator
{
    partial class NewFeed
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
            this.label1 = new System.Windows.Forms.Label();
            this.feedNameTxt = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.itemNumberTxt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.loadItemDataBtn = new System.Windows.Forms.Button();
            this.availableLVI = new System.Windows.Forms.ListView();
            this.newFeedLVI = new System.Windows.Forms.ListView();
            this.feedSaveBtn = new System.Windows.Forms.Button();
            this.customColumnNameTxt = new System.Windows.Forms.TextBox();
            this.customColumnCodeTxt = new System.Windows.Forms.TextBox();
            this.customColumnAddBtn = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.feedUpdateBtn = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(12, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Feed Name:";
            // 
            // feedNameTxt
            // 
            this.feedNameTxt.Location = new System.Drawing.Point(83, 19);
            this.feedNameTxt.Name = "feedNameTxt";
            this.feedNameTxt.Size = new System.Drawing.Size(129, 20);
            this.feedNameTxt.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(218, 22);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(157, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "(Froogle, NexTag, PartsFroogle)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(12, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(67, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "ItemNumber:";
            // 
            // itemNumberTxt
            // 
            this.itemNumberTxt.Location = new System.Drawing.Point(82, 54);
            this.itemNumberTxt.Name = "itemNumberTxt";
            this.itemNumberTxt.Size = new System.Drawing.Size(130, 20);
            this.itemNumberTxt.TabIndex = 4;
            this.itemNumberTxt.Text = "ADJ1410-C";
            this.itemNumberTxt.Leave += new System.EventHandler(this.itemNumberTxt_Leave);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(218, 57);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(203, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "(Random item that would be on this feed.)";
            // 
            // loadItemDataBtn
            // 
            this.loadItemDataBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.loadItemDataBtn.ForeColor = System.Drawing.Color.Black;
            this.loadItemDataBtn.Location = new System.Drawing.Point(83, 80);
            this.loadItemDataBtn.Name = "loadItemDataBtn";
            this.loadItemDataBtn.Size = new System.Drawing.Size(122, 23);
            this.loadItemDataBtn.TabIndex = 7;
            this.loadItemDataBtn.Text = "Load Item Data";
            this.loadItemDataBtn.UseVisualStyleBackColor = false;
            this.loadItemDataBtn.Click += new System.EventHandler(this.loadItemDataBtn_Click);
            // 
            // availableLVI
            // 
            this.availableLVI.Location = new System.Drawing.Point(6, 125);
            this.availableLVI.Name = "availableLVI";
            this.availableLVI.Size = new System.Drawing.Size(369, 382);
            this.availableLVI.TabIndex = 8;
            this.availableLVI.UseCompatibleStateImageBehavior = false;
            this.availableLVI.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.availableLVI_MouseDoubleClick);
            // 
            // newFeedLVI
            // 
            this.newFeedLVI.Location = new System.Drawing.Point(408, 125);
            this.newFeedLVI.Name = "newFeedLVI";
            this.newFeedLVI.Size = new System.Drawing.Size(369, 382);
            this.newFeedLVI.TabIndex = 9;
            this.newFeedLVI.UseCompatibleStateImageBehavior = false;
            this.newFeedLVI.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.newFeedLVI_MouseDoubleClick);
            // 
            // feedSaveBtn
            // 
            this.feedSaveBtn.BackColor = System.Drawing.Color.Lime;
            this.feedSaveBtn.ForeColor = System.Drawing.Color.Black;
            this.feedSaveBtn.Location = new System.Drawing.Point(695, 547);
            this.feedSaveBtn.Name = "feedSaveBtn";
            this.feedSaveBtn.Size = new System.Drawing.Size(82, 39);
            this.feedSaveBtn.TabIndex = 10;
            this.feedSaveBtn.Text = "Save";
            this.feedSaveBtn.UseVisualStyleBackColor = false;
            this.feedSaveBtn.Click += new System.EventHandler(this.feedSaveBtn_Click);
            // 
            // customColumnNameTxt
            // 
            this.customColumnNameTxt.Location = new System.Drawing.Point(12, 38);
            this.customColumnNameTxt.Name = "customColumnNameTxt";
            this.customColumnNameTxt.Size = new System.Drawing.Size(128, 20);
            this.customColumnNameTxt.TabIndex = 12;
            // 
            // customColumnCodeTxt
            // 
            this.customColumnCodeTxt.Location = new System.Drawing.Point(146, 38);
            this.customColumnCodeTxt.Name = "customColumnCodeTxt";
            this.customColumnCodeTxt.Size = new System.Drawing.Size(229, 20);
            this.customColumnCodeTxt.TabIndex = 13;
            // 
            // customColumnAddBtn
            // 
            this.customColumnAddBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.customColumnAddBtn.ForeColor = System.Drawing.Color.Black;
            this.customColumnAddBtn.Location = new System.Drawing.Point(380, 37);
            this.customColumnAddBtn.Name = "customColumnAddBtn";
            this.customColumnAddBtn.Size = new System.Drawing.Size(36, 22);
            this.customColumnAddBtn.TabIndex = 14;
            this.customColumnAddBtn.Text = "Add";
            this.customColumnAddBtn.UseVisualStyleBackColor = false;
            this.customColumnAddBtn.Click += new System.EventHandler(this.customColumnAddBtn_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(39, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 13);
            this.label6.TabIndex = 15;
            this.label6.Text = "Column Name";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(227, 22);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "Code Behind";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.customColumnCodeTxt);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.customColumnNameTxt);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.customColumnAddBtn);
            this.groupBox1.Location = new System.Drawing.Point(6, 513);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(427, 82);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add Custom Column:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(101, 110);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(155, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Choose a column (double click)";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(497, 109);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(218, 13);
            this.label8.TabIndex = 19;
            this.label8.Text = "New feed\'s columns (double click to remove)";
            // 
            // feedUpdateBtn
            // 
            this.feedUpdateBtn.BackColor = System.Drawing.Color.Cyan;
            this.feedUpdateBtn.ForeColor = System.Drawing.Color.Black;
            this.feedUpdateBtn.Location = new System.Drawing.Point(695, 547);
            this.feedUpdateBtn.Name = "feedUpdateBtn";
            this.feedUpdateBtn.Size = new System.Drawing.Size(82, 39);
            this.feedUpdateBtn.TabIndex = 20;
            this.feedUpdateBtn.Text = "Update";
            this.feedUpdateBtn.UseVisualStyleBackColor = false;
            // 
            // NewFeed
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gray;
            this.ClientSize = new System.Drawing.Size(784, 598);
            this.Controls.Add(this.feedUpdateBtn);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.feedSaveBtn);
            this.Controls.Add(this.newFeedLVI);
            this.Controls.Add(this.availableLVI);
            this.Controls.Add(this.loadItemDataBtn);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.itemNumberTxt);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.feedNameTxt);
            this.Controls.Add(this.label1);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NewFeed";
            this.Text = "New Feed Creator";
            this.Load += new System.EventHandler(this.NewFeed_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox feedNameTxt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox itemNumberTxt;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button loadItemDataBtn;
        private System.Windows.Forms.ListView availableLVI;
        private System.Windows.Forms.ListView newFeedLVI;
        private System.Windows.Forms.Button feedSaveBtn;
        private System.Windows.Forms.TextBox customColumnNameTxt;
        private System.Windows.Forms.TextBox customColumnCodeTxt;
        private System.Windows.Forms.Button customColumnAddBtn;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button feedUpdateBtn;
    }
}