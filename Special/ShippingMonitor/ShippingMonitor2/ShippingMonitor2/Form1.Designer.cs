namespace ShippingMonitor2
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
            System.Windows.Forms.Timer listViewScrollTmr;
            this.TopCaptionLbl = new System.Windows.Forms.Label();
            this.packersLVI = new System.Windows.Forms.ListView();
            this.tpClock1 = new TPClock.TPClock();
            this.totesInfoTicker = new TPTickerCtl.TPTicker();
            this.funFactsTicker = new TPTickerCtl.TPTicker();
            this.otherOrdersTicker = new TPTickerCtl.TPTicker();
            this.yahooTicker = new TPTickerCtl.TPTicker();
            this.ebayPMTicker = new TPTickerCtl.TPTicker();
            this.ebayAMTicker = new TPTickerCtl.TPTicker();
            this.testTicker = new TPTickerCtl.TPTicker();
            this.tickerFlipTmr = new System.Windows.Forms.Timer(this.components);
            this.tpTransferAlert1 = new TPTransferAlert.TPTransferAlert();
            this.tppoReceivedAlert1 = new TPPOReceivedAlert.TPPOReceivedAlert();
            this.tpRestockingAlert1 = new TPRestockingAlert.TPRestockingAlert();
            this.tpWillCall1 = new TPWillCall.TPWillCall();
            listViewScrollTmr = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // TopCaptionLbl
            // 
            this.TopCaptionLbl.BackColor = System.Drawing.Color.Transparent;
            this.TopCaptionLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TopCaptionLbl.Location = new System.Drawing.Point(12, 0);
            this.TopCaptionLbl.Name = "TopCaptionLbl";
            this.TopCaptionLbl.Size = new System.Drawing.Size(861, 76);
            this.TopCaptionLbl.TabIndex = 1;
            this.TopCaptionLbl.Text = "Tools Plus Warehouse Status";
            this.TopCaptionLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // packersLVI
            // 
            this.packersLVI.BackgroundImage = global::ShippingMonitor2.Properties.Resources.backgroundshort2;
            this.packersLVI.BackgroundImageTiled = true;
            this.packersLVI.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.packersLVI.Location = new System.Drawing.Point(92, 331);
            this.packersLVI.Name = "packersLVI";
            this.packersLVI.Size = new System.Drawing.Size(134, 94);
            this.packersLVI.TabIndex = 6;
            this.packersLVI.UseCompatibleStateImageBehavior = false;
            this.packersLVI.View = System.Windows.Forms.View.Details;
            // 
            // tpClock1
            // 
            this.tpClock1.Location = new System.Drawing.Point(250, 300);
            this.tpClock1.Name = "tpClock1";
            this.tpClock1.Size = new System.Drawing.Size(200, 73);
            this.tpClock1.TabIndex = 9;
            // 
            // totesInfoTicker
            // 
            this.totesInfoTicker.BackColor = System.Drawing.Color.Transparent;
            this.totesInfoTicker.Location = new System.Drawing.Point(0, 250);
            this.totesInfoTicker.Name = "totesInfoTicker";
            this.totesInfoTicker.Size = new System.Drawing.Size(210, 48);
            this.totesInfoTicker.TabIndex = 8;
            // 
            // funFactsTicker
            // 
            this.funFactsTicker.BackColor = System.Drawing.Color.Transparent;
            this.funFactsTicker.Location = new System.Drawing.Point(0, 200);
            this.funFactsTicker.Name = "funFactsTicker";
            this.funFactsTicker.Size = new System.Drawing.Size(210, 48);
            this.funFactsTicker.TabIndex = 7;
            // 
            // otherOrdersTicker
            // 
            this.otherOrdersTicker.BackColor = System.Drawing.Color.Transparent;
            this.otherOrdersTicker.Location = new System.Drawing.Point(0, 150);
            this.otherOrdersTicker.Name = "otherOrdersTicker";
            this.otherOrdersTicker.Size = new System.Drawing.Size(210, 48);
            this.otherOrdersTicker.TabIndex = 5;
            // 
            // yahooTicker
            // 
            this.yahooTicker.BackColor = System.Drawing.Color.Transparent;
            this.yahooTicker.Location = new System.Drawing.Point(0, 100);
            this.yahooTicker.Name = "yahooTicker";
            this.yahooTicker.Size = new System.Drawing.Size(210, 48);
            this.yahooTicker.TabIndex = 4;
            // 
            // ebayPMTicker
            // 
            this.ebayPMTicker.BackColor = System.Drawing.Color.Transparent;
            this.ebayPMTicker.Location = new System.Drawing.Point(0, 50);
            this.ebayPMTicker.Name = "ebayPMTicker";
            this.ebayPMTicker.Size = new System.Drawing.Size(210, 48);
            this.ebayPMTicker.TabIndex = 3;
            // 
            // ebayAMTicker
            // 
            this.ebayAMTicker.BackColor = System.Drawing.Color.Transparent;
            this.ebayAMTicker.Location = new System.Drawing.Point(0, 10);
            this.ebayAMTicker.Name = "ebayAMTicker";
            this.ebayAMTicker.Size = new System.Drawing.Size(210, 48);
            this.ebayAMTicker.TabIndex = 2;
            // 
            // testTicker
            // 
            this.testTicker.BackColor = System.Drawing.Color.Transparent;
            this.testTicker.Location = new System.Drawing.Point(271, 200);
            this.testTicker.Name = "testTicker";
            this.testTicker.Size = new System.Drawing.Size(210, 48);
            this.testTicker.TabIndex = 10;
            // 
            // tickerFlipTmr
            // 
            this.tickerFlipTmr.Interval = 5000;
            this.tickerFlipTmr.Tick += new System.EventHandler(this.tickerFlipTmr_Tick);
            // 
            // listViewScrollTmr
            // 
            listViewScrollTmr.Interval = 1000;
            listViewScrollTmr.Tick += new System.EventHandler(this.listViewScrollTmr_Tick);
            // 
            // tpTransferAlert1
            // 
            this.tpTransferAlert1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.tpTransferAlert1.Location = new System.Drawing.Point(357, 258);
            this.tpTransferAlert1.Name = "tpTransferAlert1";
            this.tpTransferAlert1.Size = new System.Drawing.Size(123, 30);
            this.tpTransferAlert1.TabIndex = 13;
            // 
            // tppoReceivedAlert1
            // 
            this.tppoReceivedAlert1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.tppoReceivedAlert1.Location = new System.Drawing.Point(414, 389);
            this.tppoReceivedAlert1.Name = "tppoReceivedAlert1";
            this.tppoReceivedAlert1.Size = new System.Drawing.Size(100, 45);
            this.tppoReceivedAlert1.TabIndex = 14;
            // 
            // tpRestockingAlert1
            // 
            this.tpRestockingAlert1.BackColor = System.Drawing.Color.Green;
            this.tpRestockingAlert1.Location = new System.Drawing.Point(265, 392);
            this.tpRestockingAlert1.Name = "tpRestockingAlert1";
            this.tpRestockingAlert1.Size = new System.Drawing.Size(130, 43);
            this.tpRestockingAlert1.TabIndex = 12;
            // 
            // tpWillCall1
            // 
            this.tpWillCall1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.tpWillCall1.Location = new System.Drawing.Point(321, 93);
            this.tpWillCall1.Name = "tpWillCall1";
            this.tpWillCall1.Size = new System.Drawing.Size(129, 78);
            this.tpWillCall1.TabIndex = 11;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::ShippingMonitor2.Properties.Resources.background;
            this.ClientSize = new System.Drawing.Size(526, 447);
            this.Controls.Add(this.tpRestockingAlert1);
            this.Controls.Add(this.tpWillCall1);
            this.Controls.Add(this.tppoReceivedAlert1);
            this.Controls.Add(this.tpTransferAlert1);
            this.Controls.Add(this.testTicker);
            this.Controls.Add(this.tpClock1);
            this.Controls.Add(this.totesInfoTicker);
            this.Controls.Add(this.funFactsTicker);
            this.Controls.Add(this.packersLVI);
            this.Controls.Add(this.otherOrdersTicker);
            this.Controls.Add(this.yahooTicker);
            this.Controls.Add(this.ebayPMTicker);
            this.Controls.Add(this.ebayAMTicker);
            this.Controls.Add(this.TopCaptionLbl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public TPTickerCtl.TPTicker ebayAMTicker;
        public TPTickerCtl.TPTicker ebayPMTicker;
        public TPTickerCtl.TPTicker yahooTicker;
        public TPTickerCtl.TPTicker otherOrdersTicker;
        public TPTickerCtl.TPTicker funFactsTicker;
        public System.Windows.Forms.ListView packersLVI;
        public TPTickerCtl.TPTicker totesInfoTicker;
        public TPTickerCtl.TPTicker testTicker;
        private System.Windows.Forms.Timer tickerFlipTmr;
        public TPPOReceivedAlert.TPPOReceivedAlert tppoReceivedAlert1;
        public TPClock.TPClock tpClock1;
        public TPRestockingAlert.TPRestockingAlert tpRestockingAlert1;
        public TPTransferAlert.TPTransferAlert tpTransferAlert1;
        public TPWillCall.TPWillCall tpWillCall1;
        public System.Windows.Forms.Label TopCaptionLbl;
    }
}

