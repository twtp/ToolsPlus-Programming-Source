namespace LivePriceComparer
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.ebayBrowser = new System.Windows.Forms.WebBrowser();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.googleBrowser = new System.Windows.Forms.WebBrowser();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.yahooBrowser = new System.Windows.Forms.WebBrowser();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.showYahooStarRatingChk = new System.Windows.Forms.CheckBox();
            this.showYahooReviewsChk = new System.Windows.Forms.CheckBox();
            this.processYahooChk = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.showOnlyNewItemsChk = new System.Windows.Forms.CheckBox();
            this.filterBlacklistedChk = new System.Windows.Forms.CheckBox();
            this.showProductPricingChk = new System.Windows.Forms.CheckBox();
            this.showProductTitleChk = new System.Windows.Forms.CheckBox();
            this.showCompetitorNameChk = new System.Windows.Forms.CheckBox();
            this.showPictureChk = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio = new System.Windows.Forms.RadioButton();
            this.settingGoogleJustUseFirstResultSetRadio = new System.Windows.Forms.RadioButton();
            this.settingGoogleUseSavedGIDsOnlyRadio = new System.Windows.Forms.RadioButton();
            this.settingGoogleUseAllResultSetsRadio = new System.Windows.Forms.RadioButton();
            this.processGoogleChk = new System.Windows.Forms.CheckBox();
            this.settingGoogleSaveGIDsChk = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.ebayUseRawDataRadio = new System.Windows.Forms.RadioButton();
            this.ebayUseEBayAPIRadio = new System.Windows.Forms.RadioButton();
            this.processEBayChk = new System.Windows.Forms.CheckBox();
            this.showHandlingChk = new System.Windows.Forms.CheckBox();
            this.showFeedbackPercentChk = new System.Windows.Forms.CheckBox();
            this.showServicesChk = new System.Windows.Forms.CheckBox();
            this.showFeedbackScoreChk = new System.Windows.Forms.CheckBox();
            this.showOneDayChk = new System.Windows.Forms.CheckBox();
            this.showShippingTypeChk = new System.Windows.Forms.CheckBox();
            this.showExpeditedChk = new System.Windows.Forms.CheckBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.debugBrowser2 = new System.Windows.Forms.WebBrowser();
            this.debugBrowser = new System.Windows.Forms.WebBrowser();
            this.debugDisplayBtn = new System.Windows.Forms.Button();
            this.debugTxt = new System.Windows.Forms.TextBox();
            this.FindBtn = new System.Windows.Forms.Button();
            this.itemTxt = new System.Windows.Forms.TextBox();
            this.versionLbl = new System.Windows.Forms.Label();
            this.statusLbl = new System.Windows.Forms.Label();
            this.ebayStatus = new System.Windows.Forms.Label();
            this.googleStatus = new System.Windows.Forms.Label();
            this.yahooStatus = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(12, 53);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(398, 543);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.ebayBrowser);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(390, 517);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "EBay";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // ebayBrowser
            // 
            this.ebayBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ebayBrowser.Location = new System.Drawing.Point(3, 3);
            this.ebayBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.ebayBrowser.Name = "ebayBrowser";
            this.ebayBrowser.Size = new System.Drawing.Size(384, 511);
            this.ebayBrowser.TabIndex = 1;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.googleBrowser);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(390, 517);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Google";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // googleBrowser
            // 
            this.googleBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.googleBrowser.Location = new System.Drawing.Point(3, 3);
            this.googleBrowser.Name = "googleBrowser";
            this.googleBrowser.Size = new System.Drawing.Size(384, 511);
            this.googleBrowser.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.yahooBrowser);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(390, 517);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Yahoo";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // yahooBrowser
            // 
            this.yahooBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.yahooBrowser.Location = new System.Drawing.Point(0, 0);
            this.yahooBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.yahooBrowser.Name = "yahooBrowser";
            this.yahooBrowser.ScriptErrorsSuppressed = true;
            this.yahooBrowser.Size = new System.Drawing.Size(390, 517);
            this.yahooBrowser.TabIndex = 0;
            this.yahooBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.yahooBrowser_DocumentCompleted);
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.groupBox5);
            this.tabPage5.Controls.Add(this.groupBox2);
            this.tabPage5.Controls.Add(this.groupBox4);
            this.tabPage5.Controls.Add(this.groupBox3);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(390, 517);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Options";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.BackColor = System.Drawing.Color.Silver;
            this.groupBox5.Controls.Add(this.showYahooStarRatingChk);
            this.groupBox5.Controls.Add(this.showYahooReviewsChk);
            this.groupBox5.Controls.Add(this.processYahooChk);
            this.groupBox5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox5.Location = new System.Drawing.Point(267, 19);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(225, 97);
            this.groupBox5.TabIndex = 5;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Yahoo Settings";
            // 
            // showYahooStarRatingChk
            // 
            this.showYahooStarRatingChk.AutoSize = true;
            this.showYahooStarRatingChk.Checked = true;
            this.showYahooStarRatingChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showYahooStarRatingChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showYahooStarRatingChk.Location = new System.Drawing.Point(10, 67);
            this.showYahooStarRatingChk.Name = "showYahooStarRatingChk";
            this.showYahooStarRatingChk.Size = new System.Drawing.Size(109, 17);
            this.showYahooStarRatingChk.TabIndex = 15;
            this.showYahooStarRatingChk.Text = "Show Star Rating";
            this.showYahooStarRatingChk.UseVisualStyleBackColor = true;
            // 
            // showYahooReviewsChk
            // 
            this.showYahooReviewsChk.AutoSize = true;
            this.showYahooReviewsChk.Checked = true;
            this.showYahooReviewsChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showYahooReviewsChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showYahooReviewsChk.Location = new System.Drawing.Point(10, 44);
            this.showYahooReviewsChk.Name = "showYahooReviewsChk";
            this.showYahooReviewsChk.Size = new System.Drawing.Size(128, 17);
            this.showYahooReviewsChk.TabIndex = 14;
            this.showYahooReviewsChk.Text = "Show Reviews Count";
            this.showYahooReviewsChk.UseVisualStyleBackColor = true;
            // 
            // processYahooChk
            // 
            this.processYahooChk.AutoSize = true;
            this.processYahooChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processYahooChk.Location = new System.Drawing.Point(10, 21);
            this.processYahooChk.Name = "processYahooChk";
            this.processYahooChk.Size = new System.Drawing.Size(156, 17);
            this.processYahooChk.TabIndex = 13;
            this.processYahooChk.Text = "Process Yahoo Competitors";
            this.processYahooChk.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.Silver;
            this.groupBox2.Controls.Add(this.showOnlyNewItemsChk);
            this.groupBox2.Controls.Add(this.filterBlacklistedChk);
            this.groupBox2.Controls.Add(this.showProductPricingChk);
            this.groupBox2.Controls.Add(this.showProductTitleChk);
            this.groupBox2.Controls.Add(this.showCompetitorNameChk);
            this.groupBox2.Controls.Add(this.showPictureChk);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(267, 130);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(227, 197);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Global Settings:";
            // 
            // showOnlyNewItemsChk
            // 
            this.showOnlyNewItemsChk.AutoSize = true;
            this.showOnlyNewItemsChk.Checked = true;
            this.showOnlyNewItemsChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showOnlyNewItemsChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showOnlyNewItemsChk.Location = new System.Drawing.Point(15, 134);
            this.showOnlyNewItemsChk.Name = "showOnlyNewItemsChk";
            this.showOnlyNewItemsChk.Size = new System.Drawing.Size(130, 17);
            this.showOnlyNewItemsChk.TabIndex = 6;
            this.showOnlyNewItemsChk.Text = "Show Only New Items";
            this.showOnlyNewItemsChk.UseVisualStyleBackColor = true;
            this.showOnlyNewItemsChk.CheckedChanged += new System.EventHandler(this.showOnlyNewItemsChk_CheckedChanged);
            // 
            // filterBlacklistedChk
            // 
            this.filterBlacklistedChk.AutoSize = true;
            this.filterBlacklistedChk.Checked = true;
            this.filterBlacklistedChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.filterBlacklistedChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.filterBlacklistedChk.Location = new System.Drawing.Point(15, 21);
            this.filterBlacklistedChk.Name = "filterBlacklistedChk";
            this.filterBlacklistedChk.Size = new System.Drawing.Size(160, 17);
            this.filterBlacklistedChk.TabIndex = 0;
            this.filterBlacklistedChk.Text = "Hide Blacklisted Competitors";
            this.filterBlacklistedChk.UseVisualStyleBackColor = true;
            this.filterBlacklistedChk.CheckedChanged += new System.EventHandler(this.filterBlacklistedChk_CheckedChanged);
            // 
            // showProductPricingChk
            // 
            this.showProductPricingChk.AutoSize = true;
            this.showProductPricingChk.Checked = true;
            this.showProductPricingChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showProductPricingChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showProductPricingChk.Location = new System.Drawing.Point(15, 111);
            this.showProductPricingChk.Name = "showProductPricingChk";
            this.showProductPricingChk.Size = new System.Drawing.Size(128, 17);
            this.showProductPricingChk.TabIndex = 5;
            this.showProductPricingChk.Text = "Show Product Pricing";
            this.showProductPricingChk.UseVisualStyleBackColor = true;
            this.showProductPricingChk.CheckedChanged += new System.EventHandler(this.showProductPricingChk_CheckedChanged);
            // 
            // showProductTitleChk
            // 
            this.showProductTitleChk.AutoSize = true;
            this.showProductTitleChk.Checked = true;
            this.showProductTitleChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showProductTitleChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showProductTitleChk.Location = new System.Drawing.Point(15, 88);
            this.showProductTitleChk.Name = "showProductTitleChk";
            this.showProductTitleChk.Size = new System.Drawing.Size(116, 17);
            this.showProductTitleChk.TabIndex = 4;
            this.showProductTitleChk.Text = "Show Product Title";
            this.showProductTitleChk.UseVisualStyleBackColor = true;
            this.showProductTitleChk.CheckedChanged += new System.EventHandler(this.showProductTitleChk_CheckedChanged);
            // 
            // showCompetitorNameChk
            // 
            this.showCompetitorNameChk.AutoSize = true;
            this.showCompetitorNameChk.Checked = true;
            this.showCompetitorNameChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showCompetitorNameChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showCompetitorNameChk.Location = new System.Drawing.Point(15, 65);
            this.showCompetitorNameChk.Name = "showCompetitorNameChk";
            this.showCompetitorNameChk.Size = new System.Drawing.Size(137, 17);
            this.showCompetitorNameChk.TabIndex = 1;
            this.showCompetitorNameChk.Text = "Show Competitor Name";
            this.showCompetitorNameChk.UseVisualStyleBackColor = true;
            this.showCompetitorNameChk.CheckedChanged += new System.EventHandler(this.showCompetitorNameChk_CheckedChanged);
            // 
            // showPictureChk
            // 
            this.showPictureChk.AutoSize = true;
            this.showPictureChk.Checked = true;
            this.showPictureChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showPictureChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showPictureChk.Location = new System.Drawing.Point(15, 42);
            this.showPictureChk.Name = "showPictureChk";
            this.showPictureChk.Size = new System.Drawing.Size(89, 17);
            this.showPictureChk.TabIndex = 0;
            this.showPictureChk.Text = "Show Picture";
            this.showPictureChk.UseVisualStyleBackColor = true;
            this.showPictureChk.CheckedChanged += new System.EventHandler(this.showPictureChk_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.BackColor = System.Drawing.Color.Silver;
            this.groupBox4.Controls.Add(this.groupBox1);
            this.groupBox4.Controls.Add(this.processGoogleChk);
            this.groupBox4.Controls.Add(this.settingGoogleSaveGIDsChk);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(15, 309);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(225, 181);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Google Shopping Settings";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio);
            this.groupBox1.Controls.Add(this.settingGoogleJustUseFirstResultSetRadio);
            this.groupBox1.Controls.Add(this.settingGoogleUseSavedGIDsOnlyRadio);
            this.groupBox1.Controls.Add(this.settingGoogleUseAllResultSetsRadio);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 71);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 99);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Searching Preference";
            // 
            // settingGoogleUseSavedGIDsOrQueriedGIDsRadio
            // 
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.AutoSize = true;
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Enabled = false;
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Location = new System.Drawing.Point(13, 70);
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Name = "settingGoogleUseSavedGIDsOrQueriedGIDsRadio";
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Size = new System.Drawing.Size(180, 17);
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.TabIndex = 5;
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.Text = "Use saved GIDs or queried GIDs";
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.UseVisualStyleBackColor = true;
            this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio.CheckedChanged += new System.EventHandler(this.settingGoogleUseSavedGIDsOrQueriedGIDsRadio_CheckedChanged);
            // 
            // settingGoogleJustUseFirstResultSetRadio
            // 
            this.settingGoogleJustUseFirstResultSetRadio.AutoSize = true;
            this.settingGoogleJustUseFirstResultSetRadio.Checked = true;
            this.settingGoogleJustUseFirstResultSetRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingGoogleJustUseFirstResultSetRadio.Location = new System.Drawing.Point(13, 21);
            this.settingGoogleJustUseFirstResultSetRadio.Name = "settingGoogleJustUseFirstResultSetRadio";
            this.settingGoogleJustUseFirstResultSetRadio.Size = new System.Drawing.Size(128, 17);
            this.settingGoogleJustUseFirstResultSetRadio.TabIndex = 2;
            this.settingGoogleJustUseFirstResultSetRadio.TabStop = true;
            this.settingGoogleJustUseFirstResultSetRadio.Text = "Just use first result set";
            this.settingGoogleJustUseFirstResultSetRadio.UseVisualStyleBackColor = true;
            this.settingGoogleJustUseFirstResultSetRadio.CheckedChanged += new System.EventHandler(this.settingGoogleJustUseFirstResultSetRadio_CheckedChanged);
            // 
            // settingGoogleUseSavedGIDsOnlyRadio
            // 
            this.settingGoogleUseSavedGIDsOnlyRadio.AutoSize = true;
            this.settingGoogleUseSavedGIDsOnlyRadio.Enabled = false;
            this.settingGoogleUseSavedGIDsOnlyRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingGoogleUseSavedGIDsOnlyRadio.Location = new System.Drawing.Point(13, 53);
            this.settingGoogleUseSavedGIDsOnlyRadio.Name = "settingGoogleUseSavedGIDsOnlyRadio";
            this.settingGoogleUseSavedGIDsOnlyRadio.Size = new System.Drawing.Size(125, 17);
            this.settingGoogleUseSavedGIDsOnlyRadio.TabIndex = 3;
            this.settingGoogleUseSavedGIDsOnlyRadio.Text = "Use saved GIDs only";
            this.settingGoogleUseSavedGIDsOnlyRadio.UseVisualStyleBackColor = true;
            this.settingGoogleUseSavedGIDsOnlyRadio.CheckedChanged += new System.EventHandler(this.settingGoogleUseSavedGIDsOnlyRadio_CheckedChanged);
            // 
            // settingGoogleUseAllResultSetsRadio
            // 
            this.settingGoogleUseAllResultSetsRadio.AutoSize = true;
            this.settingGoogleUseAllResultSetsRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingGoogleUseAllResultSetsRadio.Location = new System.Drawing.Point(13, 37);
            this.settingGoogleUseAllResultSetsRadio.Name = "settingGoogleUseAllResultSetsRadio";
            this.settingGoogleUseAllResultSetsRadio.Size = new System.Drawing.Size(107, 17);
            this.settingGoogleUseAllResultSetsRadio.TabIndex = 4;
            this.settingGoogleUseAllResultSetsRadio.Text = "Use all result sets";
            this.settingGoogleUseAllResultSetsRadio.UseVisualStyleBackColor = true;
            this.settingGoogleUseAllResultSetsRadio.CheckedChanged += new System.EventHandler(this.settingGoogleUseAllResultSetsRadio_CheckedChanged);
            // 
            // processGoogleChk
            // 
            this.processGoogleChk.AutoSize = true;
            this.processGoogleChk.Checked = true;
            this.processGoogleChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.processGoogleChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processGoogleChk.Location = new System.Drawing.Point(10, 23);
            this.processGoogleChk.Name = "processGoogleChk";
            this.processGoogleChk.Size = new System.Drawing.Size(159, 17);
            this.processGoogleChk.TabIndex = 12;
            this.processGoogleChk.Text = "Process Google Competitors";
            this.processGoogleChk.UseVisualStyleBackColor = true;
            // 
            // settingGoogleSaveGIDsChk
            // 
            this.settingGoogleSaveGIDsChk.AutoSize = true;
            this.settingGoogleSaveGIDsChk.Checked = true;
            this.settingGoogleSaveGIDsChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.settingGoogleSaveGIDsChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.settingGoogleSaveGIDsChk.Location = new System.Drawing.Point(10, 46);
            this.settingGoogleSaveGIDsChk.Name = "settingGoogleSaveGIDsChk";
            this.settingGoogleSaveGIDsChk.Size = new System.Drawing.Size(118, 17);
            this.settingGoogleSaveGIDsChk.TabIndex = 6;
            this.settingGoogleSaveGIDsChk.Text = "Save Queried GIDs";
            this.settingGoogleSaveGIDsChk.UseVisualStyleBackColor = true;
            this.settingGoogleSaveGIDsChk.CheckedChanged += new System.EventHandler(this.settingGoogleSaveGIDsChk_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.Color.Silver;
            this.groupBox3.Controls.Add(this.groupBox6);
            this.groupBox3.Controls.Add(this.processEBayChk);
            this.groupBox3.Controls.Add(this.showHandlingChk);
            this.groupBox3.Controls.Add(this.showFeedbackPercentChk);
            this.groupBox3.Controls.Add(this.showServicesChk);
            this.groupBox3.Controls.Add(this.showFeedbackScoreChk);
            this.groupBox3.Controls.Add(this.showOneDayChk);
            this.groupBox3.Controls.Add(this.showShippingTypeChk);
            this.groupBox3.Controls.Add(this.showExpeditedChk);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(15, 19);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(227, 271);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "EBay Settings";
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.ebayUseRawDataRadio);
            this.groupBox6.Controls.Add(this.ebayUseEBayAPIRadio);
            this.groupBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox6.Location = new System.Drawing.Point(12, 203);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(200, 59);
            this.groupBox6.TabIndex = 6;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Searching Preference";
            // 
            // ebayUseRawDataRadio
            // 
            this.ebayUseRawDataRadio.AutoSize = true;
            this.ebayUseRawDataRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ebayUseRawDataRadio.Location = new System.Drawing.Point(13, 37);
            this.ebayUseRawDataRadio.Name = "ebayUseRawDataRadio";
            this.ebayUseRawDataRadio.Size = new System.Drawing.Size(95, 17);
            this.ebayUseRawDataRadio.TabIndex = 5;
            this.ebayUseRawDataRadio.Text = "Use Raw Data";
            this.ebayUseRawDataRadio.UseVisualStyleBackColor = true;
            this.ebayUseRawDataRadio.CheckedChanged += new System.EventHandler(this.ebayUseRawDataRadio_CheckedChanged);
            // 
            // ebayUseEBayAPIRadio
            // 
            this.ebayUseEBayAPIRadio.AutoSize = true;
            this.ebayUseEBayAPIRadio.Checked = true;
            this.ebayUseEBayAPIRadio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ebayUseEBayAPIRadio.Location = new System.Drawing.Point(13, 19);
            this.ebayUseEBayAPIRadio.Name = "ebayUseEBayAPIRadio";
            this.ebayUseEBayAPIRadio.Size = new System.Drawing.Size(91, 17);
            this.ebayUseEBayAPIRadio.TabIndex = 6;
            this.ebayUseEBayAPIRadio.TabStop = true;
            this.ebayUseEBayAPIRadio.Text = "Use eBay API";
            this.ebayUseEBayAPIRadio.UseVisualStyleBackColor = true;
            this.ebayUseEBayAPIRadio.CheckedChanged += new System.EventHandler(this.ebayUseEBayAPIRadio_CheckedChanged);
            // 
            // processEBayChk
            // 
            this.processEBayChk.AutoSize = true;
            this.processEBayChk.Checked = true;
            this.processEBayChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.processEBayChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processEBayChk.Location = new System.Drawing.Point(10, 21);
            this.processEBayChk.Name = "processEBayChk";
            this.processEBayChk.Size = new System.Drawing.Size(150, 17);
            this.processEBayChk.TabIndex = 11;
            this.processEBayChk.Text = "Process EBay Competitors";
            this.processEBayChk.UseVisualStyleBackColor = true;
            // 
            // showHandlingChk
            // 
            this.showHandlingChk.AutoSize = true;
            this.showHandlingChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showHandlingChk.Location = new System.Drawing.Point(10, 180);
            this.showHandlingChk.Name = "showHandlingChk";
            this.showHandlingChk.Size = new System.Drawing.Size(124, 17);
            this.showHandlingChk.TabIndex = 10;
            this.showHandlingChk.Text = "Show Handling Time";
            this.showHandlingChk.UseVisualStyleBackColor = true;
            this.showHandlingChk.CheckedChanged += new System.EventHandler(this.showHandlingChk_CheckedChanged);
            // 
            // showFeedbackPercentChk
            // 
            this.showFeedbackPercentChk.AutoSize = true;
            this.showFeedbackPercentChk.Checked = true;
            this.showFeedbackPercentChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.showFeedbackPercentChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showFeedbackPercentChk.Location = new System.Drawing.Point(10, 42);
            this.showFeedbackPercentChk.Name = "showFeedbackPercentChk";
            this.showFeedbackPercentChk.Size = new System.Drawing.Size(215, 17);
            this.showFeedbackPercentChk.TabIndex = 2;
            this.showFeedbackPercentChk.Text = "Show Competitor Feedback Percentage";
            this.showFeedbackPercentChk.UseVisualStyleBackColor = true;
            this.showFeedbackPercentChk.CheckedChanged += new System.EventHandler(this.showFeedbackPercentChk_CheckedChanged);
            // 
            // showServicesChk
            // 
            this.showServicesChk.AutoSize = true;
            this.showServicesChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showServicesChk.Location = new System.Drawing.Point(10, 157);
            this.showServicesChk.Name = "showServicesChk";
            this.showServicesChk.Size = new System.Drawing.Size(97, 17);
            this.showServicesChk.TabIndex = 9;
            this.showServicesChk.Text = "Show Services";
            this.showServicesChk.UseVisualStyleBackColor = true;
            this.showServicesChk.CheckedChanged += new System.EventHandler(this.showServicesChk_CheckedChanged);
            // 
            // showFeedbackScoreChk
            // 
            this.showFeedbackScoreChk.AutoSize = true;
            this.showFeedbackScoreChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showFeedbackScoreChk.Location = new System.Drawing.Point(10, 65);
            this.showFeedbackScoreChk.Name = "showFeedbackScoreChk";
            this.showFeedbackScoreChk.Size = new System.Drawing.Size(188, 17);
            this.showFeedbackScoreChk.TabIndex = 3;
            this.showFeedbackScoreChk.Text = "Show Competitor Feedback Score";
            this.showFeedbackScoreChk.UseVisualStyleBackColor = true;
            this.showFeedbackScoreChk.CheckedChanged += new System.EventHandler(this.showFeedbackScoreChk_CheckedChanged);
            // 
            // showOneDayChk
            // 
            this.showOneDayChk.AutoSize = true;
            this.showOneDayChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showOneDayChk.Location = new System.Drawing.Point(10, 134);
            this.showOneDayChk.Name = "showOneDayChk";
            this.showOneDayChk.Size = new System.Drawing.Size(98, 17);
            this.showOneDayChk.TabIndex = 8;
            this.showOneDayChk.Text = "Show One Day";
            this.showOneDayChk.UseVisualStyleBackColor = true;
            this.showOneDayChk.CheckedChanged += new System.EventHandler(this.showOneDayChk_CheckedChanged);
            // 
            // showShippingTypeChk
            // 
            this.showShippingTypeChk.AutoSize = true;
            this.showShippingTypeChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showShippingTypeChk.Location = new System.Drawing.Point(10, 88);
            this.showShippingTypeChk.Name = "showShippingTypeChk";
            this.showShippingTypeChk.Size = new System.Drawing.Size(124, 17);
            this.showShippingTypeChk.TabIndex = 6;
            this.showShippingTypeChk.Text = "Show Shipping Type";
            this.showShippingTypeChk.UseVisualStyleBackColor = true;
            this.showShippingTypeChk.CheckedChanged += new System.EventHandler(this.showShippingTypeChk_CheckedChanged);
            // 
            // showExpeditedChk
            // 
            this.showExpeditedChk.AutoSize = true;
            this.showExpeditedChk.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.showExpeditedChk.Location = new System.Drawing.Point(10, 111);
            this.showExpeditedChk.Name = "showExpeditedChk";
            this.showExpeditedChk.Size = new System.Drawing.Size(103, 17);
            this.showExpeditedChk.TabIndex = 7;
            this.showExpeditedChk.Text = "Show Expedited";
            this.showExpeditedChk.UseVisualStyleBackColor = true;
            this.showExpeditedChk.CheckedChanged += new System.EventHandler(this.showExpeditedChk_CheckedChanged);
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.button3);
            this.tabPage4.Controls.Add(this.button2);
            this.tabPage4.Controls.Add(this.button1);
            this.tabPage4.Controls.Add(this.debugBrowser2);
            this.tabPage4.Controls.Add(this.debugBrowser);
            this.tabPage4.Controls.Add(this.debugDisplayBtn);
            this.tabPage4.Controls.Add(this.debugTxt);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(390, 517);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Debug";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(464, 71);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(77, 22);
            this.button3.TabIndex = 8;
            this.button3.Text = "js injection 3";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(464, 43);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(77, 22);
            this.button2.TabIndex = 7;
            this.button2.Text = "js injection 2";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(464, 15);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(77, 22);
            this.button1.TabIndex = 6;
            this.button1.Text = "js injection 1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // debugBrowser2
            // 
            this.debugBrowser2.Location = new System.Drawing.Point(281, 165);
            this.debugBrowser2.MinimumSize = new System.Drawing.Size(20, 20);
            this.debugBrowser2.Name = "debugBrowser2";
            this.debugBrowser2.ScriptErrorsSuppressed = true;
            this.debugBrowser2.Size = new System.Drawing.Size(1109, 349);
            this.debugBrowser2.TabIndex = 5;
            // 
            // debugBrowser
            // 
            this.debugBrowser.Location = new System.Drawing.Point(16, 165);
            this.debugBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.debugBrowser.Name = "debugBrowser";
            this.debugBrowser.ScriptErrorsSuppressed = true;
            this.debugBrowser.Size = new System.Drawing.Size(249, 349);
            this.debugBrowser.TabIndex = 2;
            // 
            // debugDisplayBtn
            // 
            this.debugDisplayBtn.Location = new System.Drawing.Point(187, 119);
            this.debugDisplayBtn.Name = "debugDisplayBtn";
            this.debugDisplayBtn.Size = new System.Drawing.Size(91, 40);
            this.debugDisplayBtn.TabIndex = 1;
            this.debugDisplayBtn.Text = "Display In Browser";
            this.debugDisplayBtn.UseVisualStyleBackColor = true;
            // 
            // debugTxt
            // 
            this.debugTxt.Location = new System.Drawing.Point(16, 17);
            this.debugTxt.MaxLength = 200000;
            this.debugTxt.Multiline = true;
            this.debugTxt.Name = "debugTxt";
            this.debugTxt.Size = new System.Drawing.Size(434, 96);
            this.debugTxt.TabIndex = 0;
            // 
            // FindBtn
            // 
            this.FindBtn.Location = new System.Drawing.Point(157, 11);
            this.FindBtn.Name = "FindBtn";
            this.FindBtn.Size = new System.Drawing.Size(72, 28);
            this.FindBtn.TabIndex = 1;
            this.FindBtn.Text = "Find";
            this.FindBtn.UseVisualStyleBackColor = true;
            this.FindBtn.Click += new System.EventHandler(this.FindBtn_Click);
            // 
            // itemTxt
            // 
            this.itemTxt.Location = new System.Drawing.Point(12, 16);
            this.itemTxt.Name = "itemTxt";
            this.itemTxt.Size = new System.Drawing.Size(139, 20);
            this.itemTxt.TabIndex = 2;
            this.itemTxt.TextChanged += new System.EventHandler(this.itemTxt_TextChanged);
            // 
            // versionLbl
            // 
            this.versionLbl.AutoSize = true;
            this.versionLbl.Location = new System.Drawing.Point(270, 26);
            this.versionLbl.Name = "versionLbl";
            this.versionLbl.Size = new System.Drawing.Size(16, 13);
            this.versionLbl.TabIndex = 3;
            this.versionLbl.Text = "v.";
            // 
            // statusLbl
            // 
            this.statusLbl.AutoSize = true;
            this.statusLbl.Location = new System.Drawing.Point(9, 599);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(40, 13);
            this.statusLbl.TabIndex = 4;
            this.statusLbl.Text = "Status:";
            // 
            // ebayStatus
            // 
            this.ebayStatus.AutoSize = true;
            this.ebayStatus.Location = new System.Drawing.Point(303, 599);
            this.ebayStatus.Name = "ebayStatus";
            this.ebayStatus.Size = new System.Drawing.Size(27, 13);
            this.ebayStatus.TabIndex = 5;
            this.ebayStatus.Text = "eNo";
            // 
            // googleStatus
            // 
            this.googleStatus.AutoSize = true;
            this.googleStatus.Location = new System.Drawing.Point(330, 599);
            this.googleStatus.Name = "googleStatus";
            this.googleStatus.Size = new System.Drawing.Size(27, 13);
            this.googleStatus.TabIndex = 6;
            this.googleStatus.Text = "gNo";
            // 
            // yahooStatus
            // 
            this.yahooStatus.AutoSize = true;
            this.yahooStatus.Location = new System.Drawing.Point(357, 599);
            this.yahooStatus.Name = "yahooStatus";
            this.yahooStatus.Size = new System.Drawing.Size(26, 13);
            this.yahooStatus.TabIndex = 7;
            this.yahooStatus.Text = "yNo";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(418, 615);
            this.Controls.Add(this.yahooStatus);
            this.Controls.Add(this.googleStatus);
            this.Controls.Add(this.ebayStatus);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.versionLbl);
            this.Controls.Add(this.itemTxt);
            this.Controls.Add(this.FindBtn);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Follow Me Pricing";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.SizeChanged += new System.EventHandler(this.Form1_SizeChanged);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button FindBtn;
        private System.Windows.Forms.TextBox itemTxt;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.WebBrowser debugBrowser;
        private System.Windows.Forms.Button debugDisplayBtn;
        private System.Windows.Forms.TextBox debugTxt;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox showHandlingChk;
        private System.Windows.Forms.CheckBox showServicesChk;
        private System.Windows.Forms.CheckBox showOneDayChk;
        private System.Windows.Forms.CheckBox showExpeditedChk;
        private System.Windows.Forms.CheckBox showShippingTypeChk;
        private System.Windows.Forms.CheckBox showProductPricingChk;
        private System.Windows.Forms.CheckBox showProductTitleChk;
        private System.Windows.Forms.CheckBox showFeedbackScoreChk;
        private System.Windows.Forms.CheckBox showFeedbackPercentChk;
        private System.Windows.Forms.CheckBox showCompetitorNameChk;
        private System.Windows.Forms.CheckBox showPictureChk;
        private System.Windows.Forms.CheckBox filterBlacklistedChk;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label versionLbl;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.WebBrowser googleBrowser;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton settingGoogleUseSavedGIDsOrQueriedGIDsRadio;
        private System.Windows.Forms.RadioButton settingGoogleJustUseFirstResultSetRadio;
        private System.Windows.Forms.RadioButton settingGoogleUseSavedGIDsOnlyRadio;
        private System.Windows.Forms.RadioButton settingGoogleUseAllResultSetsRadio;
        private System.Windows.Forms.CheckBox settingGoogleSaveGIDsChk;
        private System.Windows.Forms.WebBrowser yahooBrowser;
        private System.Windows.Forms.CheckBox processYahooChk;
        private System.Windows.Forms.CheckBox processGoogleChk;
        private System.Windows.Forms.CheckBox processEBayChk;
        private System.Windows.Forms.WebBrowser ebayBrowser;
        private System.Windows.Forms.WebBrowser debugBrowser2;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.CheckBox showYahooStarRatingChk;
        private System.Windows.Forms.CheckBox showYahooReviewsChk;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.RadioButton ebayUseRawDataRadio;
        private System.Windows.Forms.RadioButton ebayUseEBayAPIRadio;
        private System.Windows.Forms.CheckBox showOnlyNewItemsChk;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label ebayStatus;
        private System.Windows.Forms.Label googleStatus;
        private System.Windows.Forms.Label yahooStatus;
        private System.Windows.Forms.Timer timer1;
    }
}

