namespace EBayLimitsInfo
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
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label16 = new System.Windows.Forms.Label();
            this.apiBuildLbl = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.apiVersionLbl = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.acknowledgeLbl = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.timeStampLbl = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.encodingLbl = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.xmlVersionLbl = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.authTokenLbl = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.callNameLbl = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.siteIDLbl = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.certNameLbl = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.appNameLbl = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.devNameLbl = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.compatibilityLevelLbl = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.userAgentLbl = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.contentTypeLbl = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.URLLbl = new System.Windows.Forms.Label();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.label11 = new System.Windows.Forms.Label();
            this.resultsLVI = new System.Windows.Forms.ListView();
            this.viewBtn = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.versionLbl = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(12, 270);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 48);
            this.button1.TabIndex = 0;
            this.button1.Text = "Query EBay";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.apiBuildLbl);
            this.groupBox1.Controls.Add(this.label17);
            this.groupBox1.Controls.Add(this.apiVersionLbl);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.acknowledgeLbl);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.timeStampLbl);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.encodingLbl);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.xmlVersionLbl);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.authTokenLbl);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.callNameLbl);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.siteIDLbl);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.certNameLbl);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.appNameLbl);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.devNameLbl);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.compatibilityLevelLbl);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.userAgentLbl);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.contentTypeLbl);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.URLLbl);
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(948, 252);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Connection Info:";
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label16.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label16.Location = new System.Drawing.Point(8, 233);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(63, 13);
            this.label16.TabIndex = 31;
            this.label16.Text = "API Build:";
            // 
            // apiBuildLbl
            // 
            this.apiBuildLbl.AutoSize = true;
            this.apiBuildLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.apiBuildLbl.Location = new System.Drawing.Point(108, 233);
            this.apiBuildLbl.Name = "apiBuildLbl";
            this.apiBuildLbl.Size = new System.Drawing.Size(27, 13);
            this.apiBuildLbl.TabIndex = 30;
            this.apiBuildLbl.Text = "N/A";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label17.Location = new System.Drawing.Point(8, 220);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(77, 13);
            this.label17.TabIndex = 29;
            this.label17.Text = "API Version:";
            // 
            // apiVersionLbl
            // 
            this.apiVersionLbl.AutoSize = true;
            this.apiVersionLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.apiVersionLbl.Location = new System.Drawing.Point(108, 220);
            this.apiVersionLbl.Name = "apiVersionLbl";
            this.apiVersionLbl.Size = new System.Drawing.Size(27, 13);
            this.apiVersionLbl.TabIndex = 28;
            this.apiVersionLbl.Text = "N/A";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label15.Location = new System.Drawing.Point(8, 207);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(94, 13);
            this.label15.TabIndex = 27;
            this.label15.Text = "Ackknowledge:";
            // 
            // acknowledgeLbl
            // 
            this.acknowledgeLbl.AutoSize = true;
            this.acknowledgeLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.acknowledgeLbl.Location = new System.Drawing.Point(108, 207);
            this.acknowledgeLbl.Name = "acknowledgeLbl";
            this.acknowledgeLbl.Size = new System.Drawing.Size(27, 13);
            this.acknowledgeLbl.TabIndex = 26;
            this.acknowledgeLbl.Text = "N/A";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label14.Location = new System.Drawing.Point(8, 194);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(73, 13);
            this.label14.TabIndex = 25;
            this.label14.Text = "TimeStamp:";
            // 
            // timeStampLbl
            // 
            this.timeStampLbl.AutoSize = true;
            this.timeStampLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.timeStampLbl.Location = new System.Drawing.Point(108, 194);
            this.timeStampLbl.Name = "timeStampLbl";
            this.timeStampLbl.Size = new System.Drawing.Size(27, 13);
            this.timeStampLbl.TabIndex = 24;
            this.timeStampLbl.Text = "N/A";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label13.Location = new System.Drawing.Point(8, 181);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(64, 13);
            this.label13.TabIndex = 23;
            this.label13.Text = "Encoding:";
            // 
            // encodingLbl
            // 
            this.encodingLbl.AutoSize = true;
            this.encodingLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.encodingLbl.Location = new System.Drawing.Point(108, 181);
            this.encodingLbl.Name = "encodingLbl";
            this.encodingLbl.Size = new System.Drawing.Size(27, 13);
            this.encodingLbl.TabIndex = 22;
            this.encodingLbl.Text = "N/A";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.label12.Location = new System.Drawing.Point(8, 168);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(82, 13);
            this.label12.TabIndex = 21;
            this.label12.Text = "XML Version:";
            // 
            // xmlVersionLbl
            // 
            this.xmlVersionLbl.AutoSize = true;
            this.xmlVersionLbl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.xmlVersionLbl.Location = new System.Drawing.Point(108, 168);
            this.xmlVersionLbl.Name = "xmlVersionLbl";
            this.xmlVersionLbl.Size = new System.Drawing.Size(27, 13);
            this.xmlVersionLbl.TabIndex = 20;
            this.xmlVersionLbl.Text = "N/A";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(415, 32);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(77, 13);
            this.label10.TabIndex = 19;
            this.label10.Text = "Auth Token:";
            // 
            // authTokenLbl
            // 
            this.authTokenLbl.Location = new System.Drawing.Point(498, 32);
            this.authTokenLbl.Name = "authTokenLbl";
            this.authTokenLbl.Size = new System.Drawing.Size(435, 205);
            this.authTokenLbl.TabIndex = 18;
            this.authTokenLbl.Text = "N/A";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(6, 145);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(68, 13);
            this.label9.TabIndex = 17;
            this.label9.Text = "Call Name:";
            // 
            // callNameLbl
            // 
            this.callNameLbl.AutoSize = true;
            this.callNameLbl.ForeColor = System.Drawing.Color.White;
            this.callNameLbl.Location = new System.Drawing.Point(108, 145);
            this.callNameLbl.Name = "callNameLbl";
            this.callNameLbl.Size = new System.Drawing.Size(27, 13);
            this.callNameLbl.TabIndex = 16;
            this.callNameLbl.Text = "N/A";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.White;
            this.label8.Location = new System.Drawing.Point(6, 132);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(50, 13);
            this.label8.TabIndex = 15;
            this.label8.Text = "Site ID:";
            // 
            // siteIDLbl
            // 
            this.siteIDLbl.AutoSize = true;
            this.siteIDLbl.ForeColor = System.Drawing.Color.White;
            this.siteIDLbl.Location = new System.Drawing.Point(108, 132);
            this.siteIDLbl.Name = "siteIDLbl";
            this.siteIDLbl.Size = new System.Drawing.Size(27, 13);
            this.siteIDLbl.TabIndex = 14;
            this.siteIDLbl.Text = "N/A";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.White;
            this.label7.Location = new System.Drawing.Point(6, 119);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(70, 13);
            this.label7.TabIndex = 13;
            this.label7.Text = "Cert Name:";
            // 
            // certNameLbl
            // 
            this.certNameLbl.AutoSize = true;
            this.certNameLbl.ForeColor = System.Drawing.Color.White;
            this.certNameLbl.Location = new System.Drawing.Point(108, 119);
            this.certNameLbl.Name = "certNameLbl";
            this.certNameLbl.Size = new System.Drawing.Size(27, 13);
            this.certNameLbl.TabIndex = 12;
            this.certNameLbl.Text = "N/A";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(6, 106);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(69, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "App Name:";
            // 
            // appNameLbl
            // 
            this.appNameLbl.AutoSize = true;
            this.appNameLbl.ForeColor = System.Drawing.Color.White;
            this.appNameLbl.Location = new System.Drawing.Point(108, 106);
            this.appNameLbl.Name = "appNameLbl";
            this.appNameLbl.Size = new System.Drawing.Size(27, 13);
            this.appNameLbl.TabIndex = 10;
            this.appNameLbl.Text = "N/A";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.White;
            this.label5.Location = new System.Drawing.Point(6, 93);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Dev Name:";
            // 
            // devNameLbl
            // 
            this.devNameLbl.AutoSize = true;
            this.devNameLbl.ForeColor = System.Drawing.Color.White;
            this.devNameLbl.Location = new System.Drawing.Point(108, 93);
            this.devNameLbl.Name = "devNameLbl";
            this.devNameLbl.Size = new System.Drawing.Size(27, 13);
            this.devNameLbl.TabIndex = 8;
            this.devNameLbl.Text = "N/A";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(6, 80);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Compatibility:";
            // 
            // compatibilityLevelLbl
            // 
            this.compatibilityLevelLbl.AutoSize = true;
            this.compatibilityLevelLbl.ForeColor = System.Drawing.Color.White;
            this.compatibilityLevelLbl.Location = new System.Drawing.Point(108, 80);
            this.compatibilityLevelLbl.Name = "compatibilityLevelLbl";
            this.compatibilityLevelLbl.Size = new System.Drawing.Size(27, 13);
            this.compatibilityLevelLbl.TabIndex = 6;
            this.compatibilityLevelLbl.Text = "N/A";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(6, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "User Agent:";
            // 
            // userAgentLbl
            // 
            this.userAgentLbl.AutoSize = true;
            this.userAgentLbl.ForeColor = System.Drawing.Color.White;
            this.userAgentLbl.Location = new System.Drawing.Point(108, 58);
            this.userAgentLbl.Name = "userAgentLbl";
            this.userAgentLbl.Size = new System.Drawing.Size(27, 13);
            this.userAgentLbl.TabIndex = 4;
            this.userAgentLbl.Text = "N/A";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(6, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Content-type:";
            // 
            // contentTypeLbl
            // 
            this.contentTypeLbl.AutoSize = true;
            this.contentTypeLbl.ForeColor = System.Drawing.Color.White;
            this.contentTypeLbl.Location = new System.Drawing.Point(108, 45);
            this.contentTypeLbl.Name = "contentTypeLbl";
            this.contentTypeLbl.Size = new System.Drawing.Size(27, 13);
            this.contentTypeLbl.TabIndex = 2;
            this.contentTypeLbl.Text = "N/A";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(6, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "URL";
            // 
            // URLLbl
            // 
            this.URLLbl.AutoSize = true;
            this.URLLbl.ForeColor = System.Drawing.Color.White;
            this.URLLbl.Location = new System.Drawing.Point(108, 32);
            this.URLLbl.Name = "URLLbl";
            this.URLLbl.Size = new System.Drawing.Size(27, 13);
            this.URLLbl.TabIndex = 0;
            this.URLLbl.Text = "N/A";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(977, 398);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(139, 160);
            this.webBrowser1.TabIndex = 2;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(462, 284);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(78, 24);
            this.label11.TabIndex = 3;
            this.label11.Text = "Results";
            // 
            // resultsLVI
            // 
            this.resultsLVI.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.resultsLVI.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.resultsLVI.FullRowSelect = true;
            this.resultsLVI.Location = new System.Drawing.Point(12, 322);
            this.resultsLVI.Name = "resultsLVI";
            this.resultsLVI.Size = new System.Drawing.Size(948, 365);
            this.resultsLVI.TabIndex = 4;
            this.resultsLVI.UseCompatibleStateImageBehavior = false;
            this.resultsLVI.View = System.Windows.Forms.View.Details;
            this.resultsLVI.SelectedIndexChanged += new System.EventHandler(this.resultsLVI_SelectedIndexChanged);
            // 
            // viewBtn
            // 
            this.viewBtn.BackColor = System.Drawing.Color.Teal;
            this.viewBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.viewBtn.ForeColor = System.Drawing.Color.White;
            this.viewBtn.Location = new System.Drawing.Point(831, 270);
            this.viewBtn.Name = "viewBtn";
            this.viewBtn.Size = new System.Drawing.Size(129, 46);
            this.viewBtn.TabIndex = 5;
            this.viewBtn.Text = "Raw Data";
            this.viewBtn.UseVisualStyleBackColor = false;
            this.viewBtn.Click += new System.EventHandler(this.viewBtn_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.Color.Black;
            this.button2.Location = new System.Drawing.Point(167, 270);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(138, 48);
            this.button2.TabIndex = 6;
            this.button2.Text = "Print Screen";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // versionLbl
            // 
            this.versionLbl.AutoSize = true;
            this.versionLbl.ForeColor = System.Drawing.Color.White;
            this.versionLbl.Location = new System.Drawing.Point(870, 694);
            this.versionLbl.Name = "versionLbl";
            this.versionLbl.Size = new System.Drawing.Size(45, 13);
            this.versionLbl.TabIndex = 7;
            this.versionLbl.Text = "Version:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(973, 706);
            this.Controls.Add(this.versionLbl);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.viewBtn);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.resultsLVI);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "EBay Limitations and Info";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label URLLbl;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label contentTypeLbl;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label userAgentLbl;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label compatibilityLevelLbl;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label devNameLbl;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label appNameLbl;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label certNameLbl;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label siteIDLbl;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label callNameLbl;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label authTokenLbl;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ListView resultsLVI;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label apiBuildLbl;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label apiVersionLbl;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label acknowledgeLbl;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label timeStampLbl;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label encodingLbl;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label xmlVersionLbl;
        private System.Windows.Forms.Button viewBtn;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label versionLbl;
    }
}

