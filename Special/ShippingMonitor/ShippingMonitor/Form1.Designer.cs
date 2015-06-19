namespace ShippingMonitor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.TopCaptionLbl = new System.Windows.Forms.Label();
            this.OrderCheckTmr = new System.Windows.Forms.Timer(this.components);
            this.OrdersListView = new System.Windows.Forms.ListView();
            this.PackersListView = new System.Windows.Forms.ListView();
            this.MorningOrdersLbl = new System.Windows.Forms.Label();
            this.MorningOrdersTmr = new System.Windows.Forms.Timer(this.components);
            this.AfternoonOrdersTmr = new System.Windows.Forms.Timer(this.components);
            this.AfternoonOrdersLbl = new System.Windows.Forms.Label();
            this.willCallLbl = new System.Windows.Forms.Label();
            this.WillCallOrdersTmr = new System.Windows.Forms.Timer(this.components);
            this.restockingStagingTmr = new System.Windows.Forms.Timer(this.components);
            this.restockingLbl = new System.Windows.Forms.Label();
            this.amEBayLbl = new System.Windows.Forms.Label();
            this.pmEbayLbl = new System.Windows.Forms.Label();
            this.totesLoadedLbl = new System.Windows.Forms.Label();
            this.totesTmr = new System.Windows.Forms.Timer(this.components);
            this.afternoonBackLbl = new System.Windows.Forms.Label();
            this.morningBackLbl = new System.Windows.Forms.Label();
            this.willCallFlasherTmr = new System.Windows.Forms.Timer(this.components);
            this.funFactHeaderLbl = new System.Windows.Forms.Label();
            this.funFactBackLbl = new System.Windows.Forms.Label();
            this.funFactLbl = new System.Windows.Forms.Label();
            this.funFactTmr = new System.Windows.Forms.Timer(this.components);
            this.transferLbl = new System.Windows.Forms.Label();
            this.yahooHeaderLbl = new System.Windows.Forms.Label();
            this.yahooOrdersBackLbl = new System.Windows.Forms.Label();
            this.yahooOrdersLbl = new System.Windows.Forms.Label();
            this.otherOrdersLbl = new System.Windows.Forms.Label();
            this.otherOrdersBackLbl = new System.Windows.Forms.Label();
            this.otherOrdersHeaderLbl = new System.Windows.Forms.Label();
            this.YahooOrdersTmr = new System.Windows.Forms.Timer(this.components);
            this.OtherOrdersTmr = new System.Windows.Forms.Timer(this.components);
            this.leftHeaderPic = new System.Windows.Forms.PictureBox();
            this.rightHeaderPic = new System.Windows.Forms.PictureBox();
            this.poReceivedLbl = new System.Windows.Forms.Label();
            this.yahooOlderOrdersLbl = new System.Windows.Forms.Label();
            this.yahooOlderHeaderLbl = new System.Windows.Forms.Label();
            this.yahooOlderBackLbl = new System.Windows.Forms.Label();
            this.timeLbl = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.leftHeaderPic)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rightHeaderPic)).BeginInit();
            this.SuspendLayout();
            // 
            // TopCaptionLbl
            // 
            this.TopCaptionLbl.BackColor = System.Drawing.Color.RoyalBlue;
            this.TopCaptionLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TopCaptionLbl.Location = new System.Drawing.Point(12, 9);
            this.TopCaptionLbl.Name = "TopCaptionLbl";
            this.TopCaptionLbl.Size = new System.Drawing.Size(861, 76);
            this.TopCaptionLbl.TabIndex = 0;
            this.TopCaptionLbl.Text = "Tools Plus Warehouse Status";
            this.TopCaptionLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // OrderCheckTmr
            // 
            this.OrderCheckTmr.Enabled = true;
            this.OrderCheckTmr.Interval = 15000;
            this.OrderCheckTmr.Tick += new System.EventHandler(this.OrderCheckTmr_Tick);
            // 
            // OrdersListView
            // 
            this.OrdersListView.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.OrdersListView.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OrdersListView.FullRowSelect = true;
            this.OrdersListView.GridLines = true;
            this.OrdersListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.OrdersListView.Location = new System.Drawing.Point(115, 194);
            this.OrdersListView.Name = "OrdersListView";
            this.OrdersListView.Size = new System.Drawing.Size(468, 97);
            this.OrdersListView.TabIndex = 1;
            this.OrdersListView.UseCompatibleStateImageBehavior = false;
            this.OrdersListView.View = System.Windows.Forms.View.Details;
            this.OrdersListView.Visible = false;
            // 
            // PackersListView
            // 
            this.PackersListView.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.PackersListView.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PackersListView.FullRowSelect = true;
            this.PackersListView.GridLines = true;
            this.PackersListView.Location = new System.Drawing.Point(115, 367);
            this.PackersListView.Name = "PackersListView";
            this.PackersListView.Size = new System.Drawing.Size(468, 97);
            this.PackersListView.TabIndex = 2;
            this.PackersListView.UseCompatibleStateImageBehavior = false;
            this.PackersListView.View = System.Windows.Forms.View.Details;
            // 
            // MorningOrdersLbl
            // 
            this.MorningOrdersLbl.AutoSize = true;
            this.MorningOrdersLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.MorningOrdersLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MorningOrdersLbl.Location = new System.Drawing.Point(12, 94);
            this.MorningOrdersLbl.Name = "MorningOrdersLbl";
            this.MorningOrdersLbl.Size = new System.Drawing.Size(806, 55);
            this.MorningOrdersLbl.TabIndex = 3;
            this.MorningOrdersLbl.Text = "Outstanding Orders This Morning...";
            this.MorningOrdersLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // MorningOrdersTmr
            // 
            this.MorningOrdersTmr.Enabled = true;
            this.MorningOrdersTmr.Interval = 15;
            this.MorningOrdersTmr.Tick += new System.EventHandler(this.MorningOrdersTmr_Tick);
            // 
            // AfternoonOrdersTmr
            // 
            this.AfternoonOrdersTmr.Enabled = true;
            this.AfternoonOrdersTmr.Interval = 15;
            this.AfternoonOrdersTmr.Tick += new System.EventHandler(this.AfternoonOrdersTmr_Tick);
            // 
            // AfternoonOrdersLbl
            // 
            this.AfternoonOrdersLbl.AutoSize = true;
            this.AfternoonOrdersLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.AfternoonOrdersLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AfternoonOrdersLbl.Location = new System.Drawing.Point(26, 215);
            this.AfternoonOrdersLbl.Name = "AfternoonOrdersLbl";
            this.AfternoonOrdersLbl.Size = new System.Drawing.Size(843, 55);
            this.AfternoonOrdersLbl.TabIndex = 4;
            this.AfternoonOrdersLbl.Text = "Outstanding Orders This Afternoon...";
            this.AfternoonOrdersLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // willCallLbl
            // 
            this.willCallLbl.BackColor = System.Drawing.Color.OrangeRed;
            this.willCallLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.willCallLbl.Location = new System.Drawing.Point(407, 55);
            this.willCallLbl.Name = "willCallLbl";
            this.willCallLbl.Size = new System.Drawing.Size(456, 185);
            this.willCallLbl.TabIndex = 5;
            this.willCallLbl.Text = "WILL CALL";
            this.willCallLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.willCallLbl.Visible = false;
            // 
            // WillCallOrdersTmr
            // 
            this.WillCallOrdersTmr.Enabled = true;
            this.WillCallOrdersTmr.Interval = 50000;
            this.WillCallOrdersTmr.Tick += new System.EventHandler(this.WillCallOrdersTmr_Tick);
            // 
            // restockingStagingTmr
            // 
            this.restockingStagingTmr.Enabled = true;
            this.restockingStagingTmr.Interval = 60000;
            this.restockingStagingTmr.Tick += new System.EventHandler(this.restockingStagingTmr_Tick);
            // 
            // restockingLbl
            // 
            this.restockingLbl.BackColor = System.Drawing.Color.LimeGreen;
            this.restockingLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.restockingLbl.Location = new System.Drawing.Point(209, 173);
            this.restockingLbl.Name = "restockingLbl";
            this.restockingLbl.Size = new System.Drawing.Size(456, 185);
            this.restockingLbl.TabIndex = 6;
            this.restockingLbl.Text = "RESTOCKING";
            this.restockingLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.restockingLbl.Visible = false;
            // 
            // amEBayLbl
            // 
            this.amEBayLbl.AutoSize = true;
            this.amEBayLbl.BackColor = System.Drawing.Color.Blue;
            this.amEBayLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.amEBayLbl.ForeColor = System.Drawing.Color.White;
            this.amEBayLbl.Location = new System.Drawing.Point(26, 294);
            this.amEBayLbl.Name = "amEBayLbl";
            this.amEBayLbl.Size = new System.Drawing.Size(331, 55);
            this.amEBayLbl.TabIndex = 7;
            this.amEBayLbl.Text = "AM EBay Tot:";
            this.amEBayLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pmEbayLbl
            // 
            this.pmEbayLbl.AutoSize = true;
            this.pmEbayLbl.BackColor = System.Drawing.Color.Blue;
            this.pmEbayLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pmEbayLbl.ForeColor = System.Drawing.Color.White;
            this.pmEbayLbl.Location = new System.Drawing.Point(26, 358);
            this.pmEbayLbl.Name = "pmEbayLbl";
            this.pmEbayLbl.Size = new System.Drawing.Size(331, 55);
            this.pmEbayLbl.TabIndex = 8;
            this.pmEbayLbl.Text = "PM EBay Tot:";
            this.pmEbayLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // totesLoadedLbl
            // 
            this.totesLoadedLbl.AutoSize = true;
            this.totesLoadedLbl.BackColor = System.Drawing.Color.SeaGreen;
            this.totesLoadedLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totesLoadedLbl.ForeColor = System.Drawing.Color.White;
            this.totesLoadedLbl.Location = new System.Drawing.Point(26, 413);
            this.totesLoadedLbl.Name = "totesLoadedLbl";
            this.totesLoadedLbl.Size = new System.Drawing.Size(345, 55);
            this.totesLoadedLbl.TabIndex = 9;
            this.totesLoadedLbl.Text = "Totes Loaded:";
            this.totesLoadedLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // totesTmr
            // 
            this.totesTmr.Enabled = true;
            this.totesTmr.Interval = 30000;
            this.totesTmr.Tick += new System.EventHandler(this.totesTmr_Tick);
            // 
            // afternoonBackLbl
            // 
            this.afternoonBackLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.afternoonBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.afternoonBackLbl.Location = new System.Drawing.Point(41, 160);
            this.afternoonBackLbl.Name = "afternoonBackLbl";
            this.afternoonBackLbl.Size = new System.Drawing.Size(770, 55);
            this.afternoonBackLbl.TabIndex = 10;
            this.afternoonBackLbl.Text = "                                     ";
            this.afternoonBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // morningBackLbl
            // 
            this.morningBackLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.morningBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.morningBackLbl.Location = new System.Drawing.Point(41, 94);
            this.morningBackLbl.Name = "morningBackLbl";
            this.morningBackLbl.Size = new System.Drawing.Size(770, 55);
            this.morningBackLbl.TabIndex = 11;
            this.morningBackLbl.Text = "                                     ";
            this.morningBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // willCallFlasherTmr
            // 
            this.willCallFlasherTmr.Interval = 3000;
            this.willCallFlasherTmr.Tick += new System.EventHandler(this.willCallFlasherTmr_Tick);
            // 
            // funFactHeaderLbl
            // 
            this.funFactHeaderLbl.AutoSize = true;
            this.funFactHeaderLbl.BackColor = System.Drawing.Color.RoyalBlue;
            this.funFactHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.funFactHeaderLbl.ForeColor = System.Drawing.Color.White;
            this.funFactHeaderLbl.Location = new System.Drawing.Point(26, 468);
            this.funFactHeaderLbl.Name = "funFactHeaderLbl";
            this.funFactHeaderLbl.Size = new System.Drawing.Size(122, 55);
            this.funFactHeaderLbl.TabIndex = 12;
            this.funFactHeaderLbl.Text = "Info:";
            this.funFactHeaderLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // funFactBackLbl
            // 
            this.funFactBackLbl.BackColor = System.Drawing.Color.MidnightBlue;
            this.funFactBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.funFactBackLbl.ForeColor = System.Drawing.Color.White;
            this.funFactBackLbl.Location = new System.Drawing.Point(154, 468);
            this.funFactBackLbl.Name = "funFactBackLbl";
            this.funFactBackLbl.Size = new System.Drawing.Size(682, 55);
            this.funFactBackLbl.TabIndex = 13;
            this.funFactBackLbl.Text = "          ";
            this.funFactBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // funFactLbl
            // 
            this.funFactLbl.AutoSize = true;
            this.funFactLbl.BackColor = System.Drawing.Color.MidnightBlue;
            this.funFactLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.funFactLbl.ForeColor = System.Drawing.Color.White;
            this.funFactLbl.Location = new System.Drawing.Point(616, 434);
            this.funFactLbl.Name = "funFactLbl";
            this.funFactLbl.Size = new System.Drawing.Size(245, 55);
            this.funFactLbl.TabIndex = 14;
            this.funFactLbl.Text = "        test  ";
            this.funFactLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // funFactTmr
            // 
            this.funFactTmr.Interval = 15;
            this.funFactTmr.Tick += new System.EventHandler(this.funFactTmr_Tick);
            // 
            // transferLbl
            // 
            this.transferLbl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.transferLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.transferLbl.Location = new System.Drawing.Point(129, 19);
            this.transferLbl.Name = "transferLbl";
            this.transferLbl.Size = new System.Drawing.Size(661, 52);
            this.transferLbl.TabIndex = 15;
            this.transferLbl.Text = "TRANSFER";
            this.transferLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.transferLbl.Visible = false;
            // 
            // yahooHeaderLbl
            // 
            this.yahooHeaderLbl.AutoSize = true;
            this.yahooHeaderLbl.BackColor = System.Drawing.Color.Blue;
            this.yahooHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooHeaderLbl.ForeColor = System.Drawing.Color.White;
            this.yahooHeaderLbl.Location = new System.Drawing.Point(26, 331);
            this.yahooHeaderLbl.Name = "yahooHeaderLbl";
            this.yahooHeaderLbl.Size = new System.Drawing.Size(269, 55);
            this.yahooHeaderLbl.TabIndex = 16;
            this.yahooHeaderLbl.Text = "Yahoo Tot:";
            this.yahooHeaderLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // yahooOrdersBackLbl
            // 
            this.yahooOrdersBackLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.yahooOrdersBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooOrdersBackLbl.Location = new System.Drawing.Point(154, 331);
            this.yahooOrdersBackLbl.Name = "yahooOrdersBackLbl";
            this.yahooOrdersBackLbl.Size = new System.Drawing.Size(770, 55);
            this.yahooOrdersBackLbl.TabIndex = 17;
            this.yahooOrdersBackLbl.Text = "                                     ";
            this.yahooOrdersBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // yahooOrdersLbl
            // 
            this.yahooOrdersLbl.AutoSize = true;
            this.yahooOrdersLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.yahooOrdersLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooOrdersLbl.Location = new System.Drawing.Point(154, 349);
            this.yahooOrdersLbl.Name = "yahooOrdersLbl";
            this.yahooOrdersLbl.Size = new System.Drawing.Size(518, 55);
            this.yahooOrdersLbl.TabIndex = 18;
            this.yahooOrdersLbl.Text = "Yahoo Orders             ";
            this.yahooOrdersLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // otherOrdersLbl
            // 
            this.otherOrdersLbl.AutoSize = true;
            this.otherOrdersLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.otherOrdersLbl.CausesValidation = false;
            this.otherOrdersLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.otherOrdersLbl.Location = new System.Drawing.Point(143, 215);
            this.otherOrdersLbl.Name = "otherOrdersLbl";
            this.otherOrdersLbl.Size = new System.Drawing.Size(498, 55);
            this.otherOrdersLbl.TabIndex = 21;
            this.otherOrdersLbl.Text = "Other Orders             ";
            this.otherOrdersLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // otherOrdersBackLbl
            // 
            this.otherOrdersBackLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.otherOrdersBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.otherOrdersBackLbl.Location = new System.Drawing.Point(116, 229);
            this.otherOrdersBackLbl.Name = "otherOrdersBackLbl";
            this.otherOrdersBackLbl.Size = new System.Drawing.Size(770, 55);
            this.otherOrdersBackLbl.TabIndex = 20;
            this.otherOrdersBackLbl.Text = "                                     ";
            this.otherOrdersBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // otherOrdersHeaderLbl
            // 
            this.otherOrdersHeaderLbl.AutoSize = true;
            this.otherOrdersHeaderLbl.BackColor = System.Drawing.Color.Blue;
            this.otherOrdersHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.otherOrdersHeaderLbl.ForeColor = System.Drawing.Color.White;
            this.otherOrdersHeaderLbl.Location = new System.Drawing.Point(-12, 229);
            this.otherOrdersHeaderLbl.Name = "otherOrdersHeaderLbl";
            this.otherOrdersHeaderLbl.Size = new System.Drawing.Size(249, 55);
            this.otherOrdersHeaderLbl.TabIndex = 19;
            this.otherOrdersHeaderLbl.Text = "Other Tot:";
            this.otherOrdersHeaderLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // YahooOrdersTmr
            // 
            this.YahooOrdersTmr.Interval = 15;
            this.YahooOrdersTmr.Tick += new System.EventHandler(this.YahooOrdersTmr_Tick);
            // 
            // OtherOrdersTmr
            // 
            this.OtherOrdersTmr.Interval = 15;
            this.OtherOrdersTmr.Tick += new System.EventHandler(this.OtherOrdersTmr_Tick);
            // 
            // leftHeaderPic
            // 
            this.leftHeaderPic.BackColor = System.Drawing.Color.RoyalBlue;
            this.leftHeaderPic.Image = ((System.Drawing.Image)(resources.GetObject("leftHeaderPic.Image")));
            this.leftHeaderPic.Location = new System.Drawing.Point(-2, 9);
            this.leftHeaderPic.Name = "leftHeaderPic";
            this.leftHeaderPic.Size = new System.Drawing.Size(110, 108);
            this.leftHeaderPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.leftHeaderPic.TabIndex = 22;
            this.leftHeaderPic.TabStop = false;
            // 
            // rightHeaderPic
            // 
            this.rightHeaderPic.BackColor = System.Drawing.Color.RoyalBlue;
            this.rightHeaderPic.Image = ((System.Drawing.Image)(resources.GetObject("rightHeaderPic.Image")));
            this.rightHeaderPic.Location = new System.Drawing.Point(763, 9);
            this.rightHeaderPic.Name = "rightHeaderPic";
            this.rightHeaderPic.Size = new System.Drawing.Size(110, 108);
            this.rightHeaderPic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.rightHeaderPic.TabIndex = 23;
            this.rightHeaderPic.TabStop = false;
            // 
            // poReceivedLbl
            // 
            this.poReceivedLbl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.poReceivedLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.poReceivedLbl.Location = new System.Drawing.Point(116, 85);
            this.poReceivedLbl.Name = "poReceivedLbl";
            this.poReceivedLbl.Size = new System.Drawing.Size(661, 52);
            this.poReceivedLbl.TabIndex = 24;
            this.poReceivedLbl.Text = "PO RECEIVED";
            this.poReceivedLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.poReceivedLbl.Visible = false;
            // 
            // yahooOlderOrdersLbl
            // 
            this.yahooOlderOrdersLbl.AutoSize = true;
            this.yahooOlderOrdersLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.yahooOlderOrdersLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooOlderOrdersLbl.Location = new System.Drawing.Point(209, 319);
            this.yahooOlderOrdersLbl.Name = "yahooOlderOrdersLbl";
            this.yahooOlderOrdersLbl.Size = new System.Drawing.Size(654, 55);
            this.yahooOlderOrdersLbl.TabIndex = 25;
            this.yahooOlderOrdersLbl.Text = "Yahoo Older Orders             ";
            this.yahooOlderOrdersLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // yahooOlderHeaderLbl
            // 
            this.yahooOlderHeaderLbl.AutoSize = true;
            this.yahooOlderHeaderLbl.BackColor = System.Drawing.Color.Blue;
            this.yahooOlderHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooOlderHeaderLbl.ForeColor = System.Drawing.Color.White;
            this.yahooOlderHeaderLbl.Location = new System.Drawing.Point(163, 331);
            this.yahooOlderHeaderLbl.Name = "yahooOlderHeaderLbl";
            this.yahooOlderHeaderLbl.Size = new System.Drawing.Size(405, 55);
            this.yahooOlderHeaderLbl.TabIndex = 26;
            this.yahooOlderHeaderLbl.Text = "Yahoo Older Tot:";
            this.yahooOlderHeaderLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // yahooOlderBackLbl
            // 
            this.yahooOlderBackLbl.BackColor = System.Drawing.Color.LightSkyBlue;
            this.yahooOlderBackLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.yahooOlderBackLbl.Location = new System.Drawing.Point(91, 331);
            this.yahooOlderBackLbl.Name = "yahooOlderBackLbl";
            this.yahooOlderBackLbl.Size = new System.Drawing.Size(770, 55);
            this.yahooOlderBackLbl.TabIndex = 27;
            this.yahooOlderBackLbl.Text = "                                     ";
            this.yahooOlderBackLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // timeLbl
            // 
            this.timeLbl.AutoSize = true;
            this.timeLbl.BackColor = System.Drawing.Color.MidnightBlue;
            this.timeLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 72F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeLbl.ForeColor = System.Drawing.Color.White;
            this.timeLbl.Location = new System.Drawing.Point(719, 30);
            this.timeLbl.Name = "timeLbl";
            this.timeLbl.Size = new System.Drawing.Size(462, 108);
            this.timeLbl.TabIndex = 28;
            this.timeLbl.Text = "12:00 PM";
            this.timeLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gray;
            this.BackgroundImage = global::ShippingMonitor.Properties.Resources.background;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(875, 530);
            this.Controls.Add(this.willCallLbl);
            this.Controls.Add(this.restockingLbl);
            this.Controls.Add(this.timeLbl);
            this.Controls.Add(this.yahooOlderBackLbl);
            this.Controls.Add(this.yahooOlderHeaderLbl);
            this.Controls.Add(this.yahooOlderOrdersLbl);
            this.Controls.Add(this.poReceivedLbl);
            this.Controls.Add(this.rightHeaderPic);
            this.Controls.Add(this.leftHeaderPic);
            this.Controls.Add(this.otherOrdersLbl);
            this.Controls.Add(this.otherOrdersBackLbl);
            this.Controls.Add(this.otherOrdersHeaderLbl);
            this.Controls.Add(this.yahooOrdersLbl);
            this.Controls.Add(this.yahooOrdersBackLbl);
            this.Controls.Add(this.yahooHeaderLbl);
            this.Controls.Add(this.transferLbl);
            this.Controls.Add(this.funFactLbl);
            this.Controls.Add(this.funFactBackLbl);
            this.Controls.Add(this.funFactHeaderLbl);
            this.Controls.Add(this.morningBackLbl);
            this.Controls.Add(this.afternoonBackLbl);
            this.Controls.Add(this.totesLoadedLbl);
            this.Controls.Add(this.pmEbayLbl);
            this.Controls.Add(this.amEBayLbl);
            this.Controls.Add(this.AfternoonOrdersLbl);
            this.Controls.Add(this.MorningOrdersLbl);
            this.Controls.Add(this.PackersListView);
            this.Controls.Add(this.OrdersListView);
            this.Controls.Add(this.TopCaptionLbl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.SizeChanged += new System.EventHandler(this.Form1_SizeChanged);
            ((System.ComponentModel.ISupportInitialize)(this.leftHeaderPic)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rightHeaderPic)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TopCaptionLbl;
        private System.Windows.Forms.Timer OrderCheckTmr;
        private System.Windows.Forms.ListView OrdersListView;
        private System.Windows.Forms.ListView PackersListView;
        private System.Windows.Forms.Label MorningOrdersLbl;
        private System.Windows.Forms.Timer MorningOrdersTmr;
        private System.Windows.Forms.Timer AfternoonOrdersTmr;
        private System.Windows.Forms.Label AfternoonOrdersLbl;
        private System.Windows.Forms.Label willCallLbl;
        private System.Windows.Forms.Timer WillCallOrdersTmr;
        private System.Windows.Forms.Timer restockingStagingTmr;
        private System.Windows.Forms.Label restockingLbl;
        private System.Windows.Forms.Label amEBayLbl;
        private System.Windows.Forms.Label pmEbayLbl;
        private System.Windows.Forms.Label totesLoadedLbl;
        private System.Windows.Forms.Timer totesTmr;
        private System.Windows.Forms.Label afternoonBackLbl;
        private System.Windows.Forms.Label morningBackLbl;
        private System.Windows.Forms.Timer willCallFlasherTmr;
        private System.Windows.Forms.Label funFactHeaderLbl;
        private System.Windows.Forms.Label funFactBackLbl;
        private System.Windows.Forms.Label funFactLbl;
        private System.Windows.Forms.Timer funFactTmr;
        private System.Windows.Forms.Label transferLbl;
        private System.Windows.Forms.Label yahooHeaderLbl;
        private System.Windows.Forms.Label yahooOrdersBackLbl;
        private System.Windows.Forms.Label yahooOrdersLbl;
        private System.Windows.Forms.Label otherOrdersLbl;
        private System.Windows.Forms.Label otherOrdersBackLbl;
        private System.Windows.Forms.Label otherOrdersHeaderLbl;
        private System.Windows.Forms.Timer YahooOrdersTmr;
        private System.Windows.Forms.Timer OtherOrdersTmr;
        private System.Windows.Forms.PictureBox leftHeaderPic;
        private System.Windows.Forms.PictureBox rightHeaderPic;
        private System.Windows.Forms.Label poReceivedLbl;
        private System.Windows.Forms.Label yahooOlderOrdersLbl;
        private System.Windows.Forms.Label yahooOlderHeaderLbl;
        private System.Windows.Forms.Label yahooOlderBackLbl;
        private System.Windows.Forms.Label timeLbl;
    }
}

