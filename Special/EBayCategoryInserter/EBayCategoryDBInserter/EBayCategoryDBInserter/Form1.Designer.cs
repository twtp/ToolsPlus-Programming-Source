namespace EBayCategoryDBInserter
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
            this.label1 = new System.Windows.Forms.Label();
            this.catLVI = new System.Windows.Forms.ListView();
            this.addToDBBtn = new System.Windows.Forms.Button();
            this.statusLbl = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(3, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(370, 63);
            this.label1.TabIndex = 0;
            this.label1.Text = "This program inserts a user filtered list of categories and their id\'s into the E" +
    "BayHomeGardenConversionID table for later linking Business && Industrial categor" +
    "ies to them.";
            // 
            // catLVI
            // 
            this.catLVI.FullRowSelect = true;
            this.catLVI.GridLines = true;
            this.catLVI.Location = new System.Drawing.Point(6, 55);
            this.catLVI.Name = "catLVI";
            this.catLVI.Size = new System.Drawing.Size(530, 223);
            this.catLVI.TabIndex = 1;
            this.catLVI.UseCompatibleStateImageBehavior = false;
            this.catLVI.View = System.Windows.Forms.View.Details;
            // 
            // addToDBBtn
            // 
            this.addToDBBtn.Location = new System.Drawing.Point(379, 13);
            this.addToDBBtn.Name = "addToDBBtn";
            this.addToDBBtn.Size = new System.Drawing.Size(157, 36);
            this.addToDBBtn.TabIndex = 2;
            this.addToDBBtn.Text = "Add All To DB";
            this.addToDBBtn.UseVisualStyleBackColor = true;
            this.addToDBBtn.Click += new System.EventHandler(this.addToDBBtn_Click);
            // 
            // statusLbl
            // 
            this.statusLbl.AutoSize = true;
            this.statusLbl.Location = new System.Drawing.Point(12, 281);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(40, 13);
            this.statusLbl.TabIndex = 3;
            this.statusLbl.Text = "Status:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(544, 297);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.addToDBBtn);
            this.Controls.Add(this.catLVI);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "EBay Category Inserter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListView catLVI;
        private System.Windows.Forms.Button addToDBBtn;
        private System.Windows.Forms.Label statusLbl;
    }
}

