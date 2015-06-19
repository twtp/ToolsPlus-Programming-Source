namespace MoveEbayCatsFromTableToAnother
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
            this.mainLVI = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // mainLVI
            // 
            this.mainLVI.FullRowSelect = true;
            this.mainLVI.GridLines = true;
            this.mainLVI.Location = new System.Drawing.Point(33, 52);
            this.mainLVI.Name = "mainLVI";
            this.mainLVI.Size = new System.Drawing.Size(311, 170);
            this.mainLVI.TabIndex = 0;
            this.mainLVI.UseCompatibleStateImageBehavior = false;
            this.mainLVI.View = System.Windows.Forms.View.Details;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(613, 356);
            this.Controls.Add(this.mainLVI);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView mainLVI;
    }
}

