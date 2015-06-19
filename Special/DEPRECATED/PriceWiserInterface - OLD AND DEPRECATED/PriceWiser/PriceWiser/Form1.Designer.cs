namespace WisePricer
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
            this.components = new System.ComponentModel.Container();
            this.secondTmr = new System.Windows.Forms.Timer(this.components);
            this.statusLVI = new System.Windows.Forms.ListView();
            this.priceWiserLVI = new System.Windows.Forms.ListView();
            this.timeLbl = new System.Windows.Forms.Label();
            this.topImg = new System.Windows.Forms.PictureBox();
            this.processingLbl = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.topImg)).BeginInit();
            this.SuspendLayout();
            // 
            // secondTmr
            // 
            this.secondTmr.Interval = 1000;
            this.secondTmr.Tick += new System.EventHandler(this.secondTmr_Tick);
            // 
            // statusLVI
            // 
            this.statusLVI.FullRowSelect = true;
            this.statusLVI.GridLines = true;
            this.statusLVI.Location = new System.Drawing.Point(12, 45);
            this.statusLVI.Name = "statusLVI";
            this.statusLVI.Size = new System.Drawing.Size(539, 135);
            this.statusLVI.TabIndex = 0;
            this.statusLVI.UseCompatibleStateImageBehavior = false;
            this.statusLVI.View = System.Windows.Forms.View.Details;
            // 
            // priceWiserLVI
            // 
            this.priceWiserLVI.FullRowSelect = true;
            this.priceWiserLVI.GridLines = true;
            this.priceWiserLVI.Location = new System.Drawing.Point(12, 184);
            this.priceWiserLVI.Name = "priceWiserLVI";
            this.priceWiserLVI.Size = new System.Drawing.Size(539, 162);
            this.priceWiserLVI.TabIndex = 2;
            this.priceWiserLVI.UseCompatibleStateImageBehavior = false;
            this.priceWiserLVI.View = System.Windows.Forms.View.Details;
            // 
            // timeLbl
            // 
            this.timeLbl.AutoSize = true;
            this.timeLbl.BackColor = System.Drawing.Color.Transparent;
            this.timeLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeLbl.ForeColor = System.Drawing.Color.White;
            this.timeLbl.Location = new System.Drawing.Point(358, 9);
            this.timeLbl.Name = "timeLbl";
            this.timeLbl.Size = new System.Drawing.Size(175, 13);
            this.timeLbl.TabIndex = 3;
            this.timeLbl.Text = "Time: 12-12-2000 1:00:00 AM";
            // 
            // topImg
            // 
            this.topImg.Image = global::WisePricer.Properties.Resources.wisepricer;
            this.topImg.Location = new System.Drawing.Point(1, 1);
            this.topImg.Name = "topImg";
            this.topImg.Size = new System.Drawing.Size(698, 38);
            this.topImg.TabIndex = 4;
            this.topImg.TabStop = false;
            // 
            // processingLbl
            // 
            this.processingLbl.AutoSize = true;
            this.processingLbl.BackColor = System.Drawing.Color.Transparent;
            this.processingLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processingLbl.ForeColor = System.Drawing.Color.White;
            this.processingLbl.Location = new System.Drawing.Point(158, 24);
            this.processingLbl.Name = "processingLbl";
            this.processingLbl.Size = new System.Drawing.Size(86, 13);
            this.processingLbl.TabIndex = 5;
            this.processingLbl.Text = "ProcessingLbl";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(580, 83);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 44);
            this.button1.TabIndex = 6;
            this.button1.Text = "List FTP";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(575, 248);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(80, 38);
            this.button2.TabIndex = 7;
            this.button2.Text = "test upload";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(575, 47);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(97, 25);
            this.button3.TabIndex = 8;
            this.button3.Text = "Pull from Wiser";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(679, 352);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.processingLbl);
            this.Controls.Add(this.timeLbl);
            this.Controls.Add(this.priceWiserLVI);
            this.Controls.Add(this.statusLVI);
            this.Controls.Add(this.topImg);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form1";
            this.Text = "WisePricer CSV SMC";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.topImg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer secondTmr;
        private System.Windows.Forms.ListView statusLVI;
        private System.Windows.Forms.ListView priceWiserLVI;
        private System.Windows.Forms.Label timeLbl;
        private System.Windows.Forms.PictureBox topImg;
        private System.Windows.Forms.Label processingLbl;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}

