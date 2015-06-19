namespace ReportGenerator
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
            this.poboStartDatePicker = new System.Windows.Forms.DateTimePicker();
            this.poboEndDatePicker = new System.Windows.Forms.DateTimePicker();
            this.POBackorderBox = new System.Windows.Forms.GroupBox();
            this.poboHiddenReportTitleChk = new System.Windows.Forms.CheckBox();
            this.poboCreateBtn = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.poboLineCodeEnd = new System.Windows.Forms.ComboBox();
            this.poboLineCodeStart = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.splashBox = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.reportLVI = new System.Windows.Forms.ListView();
            this.statusLbl = new System.Windows.Forms.Label();
            this.connectionTmr = new System.Windows.Forms.Timer(this.components);
            this.salesTxBox = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.salesTxReportEndDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label11 = new System.Windows.Forms.Label();
            this.salesTxReportStartDatePicker = new System.Windows.Forms.DateTimePicker();
            this.salesTxIdColumnChk = new System.Windows.Forms.CheckBox();
            this.salesTxCreateBtn = new System.Windows.Forms.Button();
            this.salesTxBlankLineBetweenOrdersChk = new System.Windows.Forms.CheckBox();
            this.salesTxSkipPartiallyInvoicedChk = new System.Windows.Forms.CheckBox();
            this.salesTxShowMiscShipLinesChk = new System.Windows.Forms.CheckBox();
            this.salesTxShowMiscLinesChk = new System.Windows.Forms.CheckBox();
            this.salesTxOrdersBelowMinChk = new System.Windows.Forms.CheckBox();
            this.salesTxIncludeCostChk = new System.Windows.Forms.CheckBox();
            this.salesTxCheckOdditiesChk = new System.Windows.Forms.CheckBox();
            this.salesTxCheckPOSChk = new System.Windows.Forms.CheckBox();
            this.salesTxCheckSOChk = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.salesTxItemsTxt = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.salesTxOrderMinTxt = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.salesTxCouponAmountTxt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.salesTxMinSaleItemCountTxt = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.salesTxEndDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.salesTxStartDatePicker = new System.Windows.Forms.DateTimePicker();
            this.previewBtn = new System.Windows.Forms.Button();
            this.statusHeaderLbl = new System.Windows.Forms.Label();
            this.orderDiscrepancyBox = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.orderDiscrepancyCreateBtn = new System.Windows.Forms.Button();
            this.label15 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label16 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.reportTreeBox = new System.Windows.Forms.GroupBox();
            this.reportTree = new System.Windows.Forms.TreeView();
            this.reportTitleLbl = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.exportCSVBtn = new System.Windows.Forms.Button();
            this.exportXLSBtn = new System.Windows.Forms.Button();
            this.exportTXTBtn = new System.Windows.Forms.Button();
            this.exportHTMLBtn = new System.Windows.Forms.Button();
            this.exportXMLBtn = new System.Windows.Forms.Button();
            this.exportJSONBtn = new System.Windows.Forms.Button();
            this.exportPnl = new System.Windows.Forms.Panel();
            this.exportWORDBtn = new System.Windows.Forms.Button();
            this.exportDOCXBtn = new System.Windows.Forms.Button();
            this.exportEMAILBtn = new System.Windows.Forms.Button();
            this.exportPDFBtn = new System.Windows.Forms.Button();
            this.exportEXCELBtn = new System.Windows.Forms.Button();
            this.fileSaveDlg = new System.Windows.Forms.SaveFileDialog();
            this.reportPnl = new System.Windows.Forms.Panel();
            this.mapPnl = new System.Windows.Forms.Panel();
            this.mapBrowser = new System.Windows.Forms.WebBrowser();
            this.itemSalesGeoBox = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.itemSalesGeoEndDatePicker = new System.Windows.Forms.DateTimePicker();
            this.itemSalesGeoStartDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label13 = new System.Windows.Forms.Label();
            this.itemSalesGeoItemNumberTxt = new System.Windows.Forms.TextBox();
            this.itemSalesGeoCreateBtn = new System.Windows.Forms.Button();
            this.packagesShippedBox = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.packagingShippedTrackingIDChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedPackagingUserChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedPackageTypeChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedDestinationAddressChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedPackageDimensionsChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedPostalCodeChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedBilledWeightChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedDeliveryChargeChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedServiceNameChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedProcessedDateChk = new System.Windows.Forms.CheckBox();
            this.packagesShippedXDaysTxt = new System.Windows.Forms.TextBox();
            this.packagesShippedLastDaysRadio = new System.Windows.Forms.RadioButton();
            this.packagesShippedDateRadio = new System.Windows.Forms.RadioButton();
            this.packagesShippedEndDateLbl = new System.Windows.Forms.Label();
            this.packagesShippedStartDateLbl = new System.Windows.Forms.Label();
            this.packagesShippedEndDatePicker = new System.Windows.Forms.DateTimePicker();
            this.packagesShippedStartDatePicker = new System.Windows.Forms.DateTimePicker();
            this.packagesShippedCreateBtn = new System.Windows.Forms.Button();
            this.p2LinecodeSalesInfoBox = new System.Windows.Forms.GroupBox();
            this.p2LinecodeSalesInfoCreateBtn = new System.Windows.Forms.Button();
            this.p2LinecodeSalesInfoHeaderLbl = new System.Windows.Forms.Label();
            this.p2LinecodesTxt = new System.Windows.Forms.TextBox();
            this.p2LinecodeSalesInfoItemChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoAddressChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoNameChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoPhoneChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoAccountCreationChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoLastPurchaseChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoSettingsBox = new System.Windows.Forms.GroupBox();
            this.p2LinecodeSalesInfoCustomerIDChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoEmailChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoWebsiteChk = new System.Windows.Forms.CheckBox();
            this.p2LinecodeSalesInfoSalespersonIDChk = new System.Windows.Forms.CheckBox();
            this.POBackorderBox.SuspendLayout();
            this.splashBox.SuspendLayout();
            this.salesTxBox.SuspendLayout();
            this.orderDiscrepancyBox.SuspendLayout();
            this.reportTreeBox.SuspendLayout();
            this.exportPnl.SuspendLayout();
            this.reportPnl.SuspendLayout();
            this.mapPnl.SuspendLayout();
            this.itemSalesGeoBox.SuspendLayout();
            this.packagesShippedBox.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.p2LinecodeSalesInfoBox.SuspendLayout();
            this.p2LinecodeSalesInfoSettingsBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // poboStartDatePicker
            // 
            this.poboStartDatePicker.Location = new System.Drawing.Point(15, 31);
            this.poboStartDatePicker.Name = "poboStartDatePicker";
            this.poboStartDatePicker.Size = new System.Drawing.Size(200, 20);
            this.poboStartDatePicker.TabIndex = 1;
            // 
            // poboEndDatePicker
            // 
            this.poboEndDatePicker.Location = new System.Drawing.Point(244, 31);
            this.poboEndDatePicker.Name = "poboEndDatePicker";
            this.poboEndDatePicker.Size = new System.Drawing.Size(200, 20);
            this.poboEndDatePicker.TabIndex = 2;
            // 
            // POBackorderBox
            // 
            this.POBackorderBox.BackColor = System.Drawing.Color.Silver;
            this.POBackorderBox.Controls.Add(this.poboHiddenReportTitleChk);
            this.POBackorderBox.Controls.Add(this.poboCreateBtn);
            this.POBackorderBox.Controls.Add(this.label2);
            this.POBackorderBox.Controls.Add(this.poboLineCodeEnd);
            this.POBackorderBox.Controls.Add(this.poboLineCodeStart);
            this.POBackorderBox.Controls.Add(this.label1);
            this.POBackorderBox.Controls.Add(this.poboStartDatePicker);
            this.POBackorderBox.Controls.Add(this.poboEndDatePicker);
            this.POBackorderBox.Location = new System.Drawing.Point(233, 35);
            this.POBackorderBox.Name = "POBackorderBox";
            this.POBackorderBox.Size = new System.Drawing.Size(547, 315);
            this.POBackorderBox.TabIndex = 3;
            this.POBackorderBox.TabStop = false;
            this.POBackorderBox.Text = "POs On Backorder Report:";
            // 
            // poboHiddenReportTitleChk
            // 
            this.poboHiddenReportTitleChk.AutoSize = true;
            this.poboHiddenReportTitleChk.Location = new System.Drawing.Point(15, 118);
            this.poboHiddenReportTitleChk.Name = "poboHiddenReportTitleChk";
            this.poboHiddenReportTitleChk.Size = new System.Drawing.Size(111, 17);
            this.poboHiddenReportTitleChk.TabIndex = 8;
            this.poboHiddenReportTitleChk.Text = "Show Report Title";
            this.poboHiddenReportTitleChk.UseVisualStyleBackColor = true;
            // 
            // poboCreateBtn
            // 
            this.poboCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.poboCreateBtn.Location = new System.Drawing.Point(451, 267);
            this.poboCreateBtn.Name = "poboCreateBtn";
            this.poboCreateBtn.Size = new System.Drawing.Size(89, 38);
            this.poboCreateBtn.TabIndex = 6;
            this.poboCreateBtn.Text = "Create Report";
            this.poboCreateBtn.UseVisualStyleBackColor = false;
            this.poboCreateBtn.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(116, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(16, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "to";
            // 
            // poboLineCodeEnd
            // 
            this.poboLineCodeEnd.FormattingEnabled = true;
            this.poboLineCodeEnd.Location = new System.Drawing.Point(159, 69);
            this.poboLineCodeEnd.Name = "poboLineCodeEnd";
            this.poboLineCodeEnd.Size = new System.Drawing.Size(78, 21);
            this.poboLineCodeEnd.TabIndex = 5;
            // 
            // poboLineCodeStart
            // 
            this.poboLineCodeStart.FormattingEnabled = true;
            this.poboLineCodeStart.Location = new System.Drawing.Point(15, 69);
            this.poboLineCodeStart.Name = "poboLineCodeStart";
            this.poboLineCodeStart.Size = new System.Drawing.Size(78, 21);
            this.poboLineCodeStart.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(221, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "to";
            // 
            // splashBox
            // 
            this.splashBox.BackColor = System.Drawing.Color.Silver;
            this.splashBox.Controls.Add(this.label3);
            this.splashBox.Location = new System.Drawing.Point(233, 34);
            this.splashBox.Name = "splashBox";
            this.splashBox.Size = new System.Drawing.Size(547, 315);
            this.splashBox.TabIndex = 5;
            this.splashBox.TabStop = false;
            this.splashBox.Text = "Report Generator";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(156, 140);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(250, 16);
            this.label3.TabIndex = 0;
            this.label3.Text = "Choose a report on the left to begin";
            // 
            // reportLVI
            // 
            this.reportLVI.AllowColumnReorder = true;
            this.reportLVI.Location = new System.Drawing.Point(8, 30);
            this.reportLVI.Name = "reportLVI";
            this.reportLVI.Size = new System.Drawing.Size(763, 289);
            this.reportLVI.TabIndex = 7;
            this.reportLVI.UseCompatibleStateImageBehavior = false;
            // 
            // statusLbl
            // 
            this.statusLbl.AutoSize = true;
            this.statusLbl.Location = new System.Drawing.Point(285, 13);
            this.statusLbl.Name = "statusLbl";
            this.statusLbl.Size = new System.Drawing.Size(40, 13);
            this.statusLbl.TabIndex = 10;
            this.statusLbl.Text = "Status:";
            // 
            // connectionTmr
            // 
            this.connectionTmr.Enabled = true;
            this.connectionTmr.Interval = 200;
            this.connectionTmr.Tick += new System.EventHandler(this.connectionTmr_Tick);
            // 
            // salesTxBox
            // 
            this.salesTxBox.BackColor = System.Drawing.Color.Silver;
            this.salesTxBox.Controls.Add(this.label4);
            this.salesTxBox.Controls.Add(this.salesTxReportEndDatePicker);
            this.salesTxBox.Controls.Add(this.label11);
            this.salesTxBox.Controls.Add(this.salesTxReportStartDatePicker);
            this.salesTxBox.Controls.Add(this.salesTxIdColumnChk);
            this.salesTxBox.Controls.Add(this.salesTxCreateBtn);
            this.salesTxBox.Controls.Add(this.salesTxBlankLineBetweenOrdersChk);
            this.salesTxBox.Controls.Add(this.salesTxSkipPartiallyInvoicedChk);
            this.salesTxBox.Controls.Add(this.salesTxShowMiscShipLinesChk);
            this.salesTxBox.Controls.Add(this.salesTxShowMiscLinesChk);
            this.salesTxBox.Controls.Add(this.salesTxOrdersBelowMinChk);
            this.salesTxBox.Controls.Add(this.salesTxIncludeCostChk);
            this.salesTxBox.Controls.Add(this.salesTxCheckOdditiesChk);
            this.salesTxBox.Controls.Add(this.salesTxCheckPOSChk);
            this.salesTxBox.Controls.Add(this.salesTxCheckSOChk);
            this.salesTxBox.Controls.Add(this.label10);
            this.salesTxBox.Controls.Add(this.salesTxItemsTxt);
            this.salesTxBox.Controls.Add(this.label5);
            this.salesTxBox.Controls.Add(this.salesTxOrderMinTxt);
            this.salesTxBox.Controls.Add(this.label6);
            this.salesTxBox.Controls.Add(this.salesTxCouponAmountTxt);
            this.salesTxBox.Controls.Add(this.label7);
            this.salesTxBox.Controls.Add(this.salesTxMinSaleItemCountTxt);
            this.salesTxBox.Controls.Add(this.label8);
            this.salesTxBox.Controls.Add(this.salesTxEndDatePicker);
            this.salesTxBox.Controls.Add(this.label9);
            this.salesTxBox.Controls.Add(this.salesTxStartDatePicker);
            this.salesTxBox.Location = new System.Drawing.Point(234, 30);
            this.salesTxBox.Name = "salesTxBox";
            this.salesTxBox.Size = new System.Drawing.Size(547, 315);
            this.salesTxBox.TabIndex = 11;
            this.salesTxBox.TabStop = false;
            this.salesTxBox.Text = "Sales Transaction Report:";
            this.salesTxBox.Enter += new System.EventHandler(this.salesTxBox_Enter);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(7, 95);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 13);
            this.label4.TabIndex = 47;
            this.label4.Text = "End Date";
            // 
            // salesTxReportEndDatePicker
            // 
            this.salesTxReportEndDatePicker.Location = new System.Drawing.Point(62, 91);
            this.salesTxReportEndDatePicker.Name = "salesTxReportEndDatePicker";
            this.salesTxReportEndDatePicker.Size = new System.Drawing.Size(200, 20);
            this.salesTxReportEndDatePicker.TabIndex = 46;
            this.salesTxReportEndDatePicker.Value = new System.DateTime(2015, 1, 30, 11, 17, 0, 0);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(4, 21);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(55, 13);
            this.label11.TabIndex = 45;
            this.label11.Text = "Start Date";
            // 
            // salesTxReportStartDatePicker
            // 
            this.salesTxReportStartDatePicker.Location = new System.Drawing.Point(62, 17);
            this.salesTxReportStartDatePicker.Name = "salesTxReportStartDatePicker";
            this.salesTxReportStartDatePicker.Size = new System.Drawing.Size(200, 20);
            this.salesTxReportStartDatePicker.TabIndex = 44;
            this.salesTxReportStartDatePicker.Value = new System.DateTime(2014, 12, 1, 11, 16, 0, 0);
            // 
            // salesTxIdColumnChk
            // 
            this.salesTxIdColumnChk.AutoSize = true;
            this.salesTxIdColumnChk.Checked = true;
            this.salesTxIdColumnChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxIdColumnChk.Location = new System.Drawing.Point(55, 297);
            this.salesTxIdColumnChk.Name = "salesTxIdColumnChk";
            this.salesTxIdColumnChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxIdColumnChk.Size = new System.Drawing.Size(105, 17);
            this.salesTxIdColumnChk.TabIndex = 43;
            this.salesTxIdColumnChk.Text = "Show ID Column";
            this.salesTxIdColumnChk.UseVisualStyleBackColor = true;
            // 
            // salesTxCreateBtn
            // 
            this.salesTxCreateBtn.AutoSize = true;
            this.salesTxCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.salesTxCreateBtn.Location = new System.Drawing.Point(450, 258);
            this.salesTxCreateBtn.Name = "salesTxCreateBtn";
            this.salesTxCreateBtn.Size = new System.Drawing.Size(83, 43);
            this.salesTxCreateBtn.TabIndex = 41;
            this.salesTxCreateBtn.Text = "Create Report";
            this.salesTxCreateBtn.UseVisualStyleBackColor = false;
            this.salesTxCreateBtn.Click += new System.EventHandler(this.salesTxCreateBtn_Click);
            // 
            // salesTxBlankLineBetweenOrdersChk
            // 
            this.salesTxBlankLineBetweenOrdersChk.AutoSize = true;
            this.salesTxBlankLineBetweenOrdersChk.Checked = true;
            this.salesTxBlankLineBetweenOrdersChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxBlankLineBetweenOrdersChk.Location = new System.Drawing.Point(5, 277);
            this.salesTxBlankLineBetweenOrdersChk.Name = "salesTxBlankLineBetweenOrdersChk";
            this.salesTxBlankLineBetweenOrdersChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxBlankLineBetweenOrdersChk.Size = new System.Drawing.Size(155, 17);
            this.salesTxBlankLineBetweenOrdersChk.TabIndex = 40;
            this.salesTxBlankLineBetweenOrdersChk.Text = "Blank Line Between Orders";
            this.salesTxBlankLineBetweenOrdersChk.UseVisualStyleBackColor = true;
            // 
            // salesTxSkipPartiallyInvoicedChk
            // 
            this.salesTxSkipPartiallyInvoicedChk.AutoSize = true;
            this.salesTxSkipPartiallyInvoicedChk.Checked = true;
            this.salesTxSkipPartiallyInvoicedChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxSkipPartiallyInvoicedChk.Location = new System.Drawing.Point(30, 258);
            this.salesTxSkipPartiallyInvoicedChk.Name = "salesTxSkipPartiallyInvoicedChk";
            this.salesTxSkipPartiallyInvoicedChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxSkipPartiallyInvoicedChk.Size = new System.Drawing.Size(130, 17);
            this.salesTxSkipPartiallyInvoicedChk.TabIndex = 39;
            this.salesTxSkipPartiallyInvoicedChk.Text = "Skip Partially Invoiced";
            this.salesTxSkipPartiallyInvoicedChk.UseVisualStyleBackColor = true;
            // 
            // salesTxShowMiscShipLinesChk
            // 
            this.salesTxShowMiscShipLinesChk.AutoSize = true;
            this.salesTxShowMiscShipLinesChk.Checked = true;
            this.salesTxShowMiscShipLinesChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxShowMiscShipLinesChk.Location = new System.Drawing.Point(57, 238);
            this.salesTxShowMiscShipLinesChk.Name = "salesTxShowMiscShipLinesChk";
            this.salesTxShowMiscShipLinesChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxShowMiscShipLinesChk.Size = new System.Drawing.Size(103, 17);
            this.salesTxShowMiscShipLinesChk.TabIndex = 38;
            this.salesTxShowMiscShipLinesChk.Text = "Misc. Ship Lines";
            this.salesTxShowMiscShipLinesChk.UseVisualStyleBackColor = true;
            // 
            // salesTxShowMiscLinesChk
            // 
            this.salesTxShowMiscLinesChk.AutoSize = true;
            this.salesTxShowMiscLinesChk.Location = new System.Drawing.Point(81, 217);
            this.salesTxShowMiscLinesChk.Name = "salesTxShowMiscLinesChk";
            this.salesTxShowMiscLinesChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxShowMiscLinesChk.Size = new System.Drawing.Size(79, 17);
            this.salesTxShowMiscLinesChk.TabIndex = 37;
            this.salesTxShowMiscLinesChk.Text = "Misc. Lines";
            this.salesTxShowMiscLinesChk.UseVisualStyleBackColor = true;
            // 
            // salesTxOrdersBelowMinChk
            // 
            this.salesTxOrdersBelowMinChk.AutoSize = true;
            this.salesTxOrdersBelowMinChk.Checked = true;
            this.salesTxOrdersBelowMinChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxOrdersBelowMinChk.Location = new System.Drawing.Point(51, 198);
            this.salesTxOrdersBelowMinChk.Name = "salesTxOrdersBelowMinChk";
            this.salesTxOrdersBelowMinChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxOrdersBelowMinChk.Size = new System.Drawing.Size(109, 17);
            this.salesTxOrdersBelowMinChk.TabIndex = 36;
            this.salesTxOrdersBelowMinChk.Text = "Orders Below Min";
            this.salesTxOrdersBelowMinChk.UseVisualStyleBackColor = true;
            // 
            // salesTxIncludeCostChk
            // 
            this.salesTxIncludeCostChk.AutoSize = true;
            this.salesTxIncludeCostChk.Location = new System.Drawing.Point(37, 178);
            this.salesTxIncludeCostChk.Name = "salesTxIncludeCostChk";
            this.salesTxIncludeCostChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxIncludeCostChk.Size = new System.Drawing.Size(123, 17);
            this.salesTxIncludeCostChk.TabIndex = 35;
            this.salesTxIncludeCostChk.Text = "Include Cost Column";
            this.salesTxIncludeCostChk.UseVisualStyleBackColor = true;
            // 
            // salesTxCheckOdditiesChk
            // 
            this.salesTxCheckOdditiesChk.AutoSize = true;
            this.salesTxCheckOdditiesChk.Location = new System.Drawing.Point(47, 158);
            this.salesTxCheckOdditiesChk.Name = "salesTxCheckOdditiesChk";
            this.salesTxCheckOdditiesChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxCheckOdditiesChk.Size = new System.Drawing.Size(113, 17);
            this.salesTxCheckOdditiesChk.TabIndex = 34;
            this.salesTxCheckOdditiesChk.Text = "Check for Oddities";
            this.salesTxCheckOdditiesChk.UseVisualStyleBackColor = true;
            // 
            // salesTxCheckPOSChk
            // 
            this.salesTxCheckPOSChk.AutoSize = true;
            this.salesTxCheckPOSChk.Location = new System.Drawing.Point(8, 138);
            this.salesTxCheckPOSChk.Name = "salesTxCheckPOSChk";
            this.salesTxCheckPOSChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxCheckPOSChk.Size = new System.Drawing.Size(152, 17);
            this.salesTxCheckPOSChk.TabIndex = 33;
            this.salesTxCheckPOSChk.Text = "Check Transaction in POS";
            this.salesTxCheckPOSChk.UseVisualStyleBackColor = true;
            // 
            // salesTxCheckSOChk
            // 
            this.salesTxCheckSOChk.AutoSize = true;
            this.salesTxCheckSOChk.Checked = true;
            this.salesTxCheckSOChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.salesTxCheckSOChk.Location = new System.Drawing.Point(46, 118);
            this.salesTxCheckSOChk.Name = "salesTxCheckSOChk";
            this.salesTxCheckSOChk.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.salesTxCheckSOChk.Size = new System.Drawing.Size(114, 17);
            this.salesTxCheckSOChk.TabIndex = 32;
            this.salesTxCheckSOChk.Text = "Check SO and AR";
            this.salesTxCheckSOChk.UseVisualStyleBackColor = true;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(200, 117);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(70, 13);
            this.label10.TabIndex = 31;
            this.label10.Text = "Inventory List";
            // 
            // salesTxItemsTxt
            // 
            this.salesTxItemsTxt.Location = new System.Drawing.Point(176, 129);
            this.salesTxItemsTxt.Multiline = true;
            this.salesTxItemsTxt.Name = "salesTxItemsTxt";
            this.salesTxItemsTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.salesTxItemsTxt.Size = new System.Drawing.Size(128, 177);
            this.salesTxItemsTxt.TabIndex = 30;
            this.salesTxItemsTxt.Text = resources.GetString("salesTxItemsTxt.Text");
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(303, 43);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 13);
            this.label5.TabIndex = 29;
            this.label5.Text = "Order Minimum";
            // 
            // salesTxOrderMinTxt
            // 
            this.salesTxOrderMinTxt.Location = new System.Drawing.Point(383, 40);
            this.salesTxOrderMinTxt.Name = "salesTxOrderMinTxt";
            this.salesTxOrderMinTxt.Size = new System.Drawing.Size(55, 20);
            this.salesTxOrderMinTxt.TabIndex = 28;
            this.salesTxOrderMinTxt.Text = "99";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(297, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(83, 13);
            this.label6.TabIndex = 27;
            this.label6.Text = "Coupon Amount";
            // 
            // salesTxCouponAmountTxt
            // 
            this.salesTxCouponAmountTxt.Location = new System.Drawing.Point(383, 17);
            this.salesTxCouponAmountTxt.Name = "salesTxCouponAmountTxt";
            this.salesTxCouponAmountTxt.Size = new System.Drawing.Size(55, 20);
            this.salesTxCouponAmountTxt.TabIndex = 26;
            this.salesTxCouponAmountTxt.Text = "25";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(294, 67);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(102, 13);
            this.label7.TabIndex = 25;
            this.label7.Text = "Min Sale Item Count";
            // 
            // salesTxMinSaleItemCountTxt
            // 
            this.salesTxMinSaleItemCountTxt.Location = new System.Drawing.Point(398, 64);
            this.salesTxMinSaleItemCountTxt.Name = "salesTxMinSaleItemCountTxt";
            this.salesTxMinSaleItemCountTxt.Size = new System.Drawing.Size(40, 20);
            this.salesTxMinSaleItemCountTxt.TabIndex = 24;
            this.salesTxMinSaleItemCountTxt.Text = "1";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(30, 69);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(50, 13);
            this.label8.TabIndex = 23;
            this.label8.Text = "Sale End";
            // 
            // salesTxEndDatePicker
            // 
            this.salesTxEndDatePicker.Location = new System.Drawing.Point(81, 65);
            this.salesTxEndDatePicker.Name = "salesTxEndDatePicker";
            this.salesTxEndDatePicker.Size = new System.Drawing.Size(200, 20);
            this.salesTxEndDatePicker.TabIndex = 22;
            this.salesTxEndDatePicker.Value = new System.DateTime(2015, 1, 30, 11, 17, 0, 0);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(27, 47);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 13);
            this.label9.TabIndex = 21;
            this.label9.Text = "Sale Start";
            // 
            // salesTxStartDatePicker
            // 
            this.salesTxStartDatePicker.Location = new System.Drawing.Point(81, 43);
            this.salesTxStartDatePicker.Name = "salesTxStartDatePicker";
            this.salesTxStartDatePicker.Size = new System.Drawing.Size(200, 20);
            this.salesTxStartDatePicker.TabIndex = 20;
            this.salesTxStartDatePicker.Value = new System.DateTime(2014, 12, 1, 11, 16, 0, 0);
            // 
            // previewBtn
            // 
            this.previewBtn.Font = new System.Drawing.Font("Wingdings 3", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.previewBtn.Location = new System.Drawing.Point(10, 359);
            this.previewBtn.Name = "previewBtn";
            this.previewBtn.Size = new System.Drawing.Size(769, 24);
            this.previewBtn.TabIndex = 12;
            this.previewBtn.Text = "q";
            this.previewBtn.UseVisualStyleBackColor = true;
            this.previewBtn.Click += new System.EventHandler(this.previewBtn_Click);
            // 
            // statusHeaderLbl
            // 
            this.statusHeaderLbl.AutoSize = true;
            this.statusHeaderLbl.Location = new System.Drawing.Point(246, 13);
            this.statusHeaderLbl.Name = "statusHeaderLbl";
            this.statusHeaderLbl.Size = new System.Drawing.Size(40, 13);
            this.statusHeaderLbl.TabIndex = 13;
            this.statusHeaderLbl.Text = "Status:";
            // 
            // orderDiscrepancyBox
            // 
            this.orderDiscrepancyBox.BackColor = System.Drawing.Color.Silver;
            this.orderDiscrepancyBox.Controls.Add(this.checkBox1);
            this.orderDiscrepancyBox.Controls.Add(this.orderDiscrepancyCreateBtn);
            this.orderDiscrepancyBox.Controls.Add(this.label15);
            this.orderDiscrepancyBox.Controls.Add(this.dateTimePicker1);
            this.orderDiscrepancyBox.Controls.Add(this.label16);
            this.orderDiscrepancyBox.Controls.Add(this.dateTimePicker2);
            this.orderDiscrepancyBox.Location = new System.Drawing.Point(231, 37);
            this.orderDiscrepancyBox.Name = "orderDiscrepancyBox";
            this.orderDiscrepancyBox.Size = new System.Drawing.Size(547, 315);
            this.orderDiscrepancyBox.TabIndex = 14;
            this.orderDiscrepancyBox.TabStop = false;
            this.orderDiscrepancyBox.Text = "Order Discrepancy Report:";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(30, 71);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(140, 17);
            this.checkBox1.TabIndex = 43;
            this.checkBox1.Text = "Sold less than sale price";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // orderDiscrepancyCreateBtn
            // 
            this.orderDiscrepancyCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.orderDiscrepancyCreateBtn.Location = new System.Drawing.Point(458, 263);
            this.orderDiscrepancyCreateBtn.Name = "orderDiscrepancyCreateBtn";
            this.orderDiscrepancyCreateBtn.Size = new System.Drawing.Size(83, 43);
            this.orderDiscrepancyCreateBtn.TabIndex = 41;
            this.orderDiscrepancyCreateBtn.Text = "Create Report";
            this.orderDiscrepancyCreateBtn.UseVisualStyleBackColor = false;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(241, 15);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(52, 13);
            this.label15.TabIndex = 23;
            this.label15.Text = "End Date";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(241, 31);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker1.TabIndex = 22;
            this.dateTimePicker1.Value = new System.DateTime(2015, 1, 30, 11, 17, 0, 0);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(12, 16);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(55, 13);
            this.label16.TabIndex = 21;
            this.label16.Text = "Start Date";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(12, 32);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker2.TabIndex = 20;
            this.dateTimePicker2.Value = new System.DateTime(2014, 12, 1, 11, 16, 0, 0);
            // 
            // reportTreeBox
            // 
            this.reportTreeBox.BackColor = System.Drawing.Color.Silver;
            this.reportTreeBox.Controls.Add(this.reportTree);
            this.reportTreeBox.Location = new System.Drawing.Point(10, 10);
            this.reportTreeBox.Name = "reportTreeBox";
            this.reportTreeBox.Size = new System.Drawing.Size(216, 343);
            this.reportTreeBox.TabIndex = 15;
            this.reportTreeBox.TabStop = false;
            this.reportTreeBox.Text = "Reports";
            // 
            // reportTree
            // 
            this.reportTree.Location = new System.Drawing.Point(7, 19);
            this.reportTree.Name = "reportTree";
            this.reportTree.Size = new System.Drawing.Size(203, 318);
            this.reportTree.TabIndex = 1;
            this.reportTree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.reportTree_AfterSelect);
            // 
            // reportTitleLbl
            // 
            this.reportTitleLbl.AutoSize = true;
            this.reportTitleLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.reportTitleLbl.Location = new System.Drawing.Point(26, 7);
            this.reportTitleLbl.Name = "reportTitleLbl";
            this.reportTitleLbl.Size = new System.Drawing.Size(139, 20);
            this.reportTitleLbl.TabIndex = 16;
            this.reportTitleLbl.Text = "Generated Report";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(5, 7);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(55, 13);
            this.label12.TabIndex = 17;
            this.label12.Text = "Export As:";
            // 
            // exportCSVBtn
            // 
            this.exportCSVBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportCSVBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportCSVBtn.Location = new System.Drawing.Point(117, 0);
            this.exportCSVBtn.Name = "exportCSVBtn";
            this.exportCSVBtn.Size = new System.Drawing.Size(43, 27);
            this.exportCSVBtn.TabIndex = 18;
            this.exportCSVBtn.Text = ".CSV";
            this.exportCSVBtn.UseVisualStyleBackColor = false;
            this.exportCSVBtn.Click += new System.EventHandler(this.exportCSVBtn_Click);
            // 
            // exportXLSBtn
            // 
            this.exportXLSBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportXLSBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportXLSBtn.Location = new System.Drawing.Point(163, 0);
            this.exportXLSBtn.Name = "exportXLSBtn";
            this.exportXLSBtn.Size = new System.Drawing.Size(43, 27);
            this.exportXLSBtn.TabIndex = 19;
            this.exportXLSBtn.Text = ".XLSX";
            this.exportXLSBtn.UseVisualStyleBackColor = false;
            this.exportXLSBtn.Click += new System.EventHandler(this.exportXLSBtn_Click);
            // 
            // exportTXTBtn
            // 
            this.exportTXTBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportTXTBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportTXTBtn.Location = new System.Drawing.Point(71, 0);
            this.exportTXTBtn.Name = "exportTXTBtn";
            this.exportTXTBtn.Size = new System.Drawing.Size(43, 27);
            this.exportTXTBtn.TabIndex = 21;
            this.exportTXTBtn.Text = ".TXT";
            this.exportTXTBtn.UseVisualStyleBackColor = false;
            this.exportTXTBtn.Click += new System.EventHandler(this.exportTXTBtn_Click);
            // 
            // exportHTMLBtn
            // 
            this.exportHTMLBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportHTMLBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportHTMLBtn.Location = new System.Drawing.Point(301, 0);
            this.exportHTMLBtn.Name = "exportHTMLBtn";
            this.exportHTMLBtn.Size = new System.Drawing.Size(43, 27);
            this.exportHTMLBtn.TabIndex = 20;
            this.exportHTMLBtn.Text = ".HTML";
            this.exportHTMLBtn.UseVisualStyleBackColor = false;
            this.exportHTMLBtn.Click += new System.EventHandler(this.exportHTMLBtn_Click);
            // 
            // exportXMLBtn
            // 
            this.exportXMLBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportXMLBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportXMLBtn.Location = new System.Drawing.Point(347, 0);
            this.exportXMLBtn.Name = "exportXMLBtn";
            this.exportXMLBtn.Size = new System.Drawing.Size(43, 27);
            this.exportXMLBtn.TabIndex = 22;
            this.exportXMLBtn.Text = ".XML";
            this.exportXMLBtn.UseVisualStyleBackColor = false;
            this.exportXMLBtn.Click += new System.EventHandler(this.exportXMLBtn_Click);
            // 
            // exportJSONBtn
            // 
            this.exportJSONBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportJSONBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportJSONBtn.Location = new System.Drawing.Point(393, 0);
            this.exportJSONBtn.Name = "exportJSONBtn";
            this.exportJSONBtn.Size = new System.Drawing.Size(43, 27);
            this.exportJSONBtn.TabIndex = 23;
            this.exportJSONBtn.Text = ".JSON";
            this.exportJSONBtn.UseVisualStyleBackColor = false;
            this.exportJSONBtn.Click += new System.EventHandler(this.exportJSONBtn_Click);
            // 
            // exportPnl
            // 
            this.exportPnl.Controls.Add(this.exportWORDBtn);
            this.exportPnl.Controls.Add(this.exportDOCXBtn);
            this.exportPnl.Controls.Add(this.exportEMAILBtn);
            this.exportPnl.Controls.Add(this.exportPDFBtn);
            this.exportPnl.Controls.Add(this.exportEXCELBtn);
            this.exportPnl.Controls.Add(this.exportJSONBtn);
            this.exportPnl.Controls.Add(this.label12);
            this.exportPnl.Controls.Add(this.exportXMLBtn);
            this.exportPnl.Controls.Add(this.exportCSVBtn);
            this.exportPnl.Controls.Add(this.exportTXTBtn);
            this.exportPnl.Controls.Add(this.exportXLSBtn);
            this.exportPnl.Controls.Add(this.exportHTMLBtn);
            this.exportPnl.Location = new System.Drawing.Point(180, 2);
            this.exportPnl.Name = "exportPnl";
            this.exportPnl.Size = new System.Drawing.Size(583, 27);
            this.exportPnl.TabIndex = 24;
            // 
            // exportWORDBtn
            // 
            this.exportWORDBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.exportWORDBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportWORDBtn.Location = new System.Drawing.Point(487, 0);
            this.exportWORDBtn.Name = "exportWORDBtn";
            this.exportWORDBtn.Size = new System.Drawing.Size(43, 27);
            this.exportWORDBtn.TabIndex = 28;
            this.exportWORDBtn.Text = "WORD";
            this.exportWORDBtn.UseVisualStyleBackColor = false;
            this.exportWORDBtn.Click += new System.EventHandler(this.exportWORDBtn_Click);
            // 
            // exportDOCXBtn
            // 
            this.exportDOCXBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportDOCXBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportDOCXBtn.Location = new System.Drawing.Point(209, 0);
            this.exportDOCXBtn.Name = "exportDOCXBtn";
            this.exportDOCXBtn.Size = new System.Drawing.Size(43, 27);
            this.exportDOCXBtn.TabIndex = 27;
            this.exportDOCXBtn.Text = ".DOCX";
            this.exportDOCXBtn.UseVisualStyleBackColor = false;
            this.exportDOCXBtn.Click += new System.EventHandler(this.exportDOCXBtn_Click);
            // 
            // exportEMAILBtn
            // 
            this.exportEMAILBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.exportEMAILBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportEMAILBtn.Location = new System.Drawing.Point(534, 0);
            this.exportEMAILBtn.Name = "exportEMAILBtn";
            this.exportEMAILBtn.Size = new System.Drawing.Size(43, 27);
            this.exportEMAILBtn.TabIndex = 26;
            this.exportEMAILBtn.Text = "EMAIL";
            this.exportEMAILBtn.UseVisualStyleBackColor = false;
            this.exportEMAILBtn.Click += new System.EventHandler(this.exportEMAILBtn_Click);
            // 
            // exportPDFBtn
            // 
            this.exportPDFBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.exportPDFBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportPDFBtn.Location = new System.Drawing.Point(255, 0);
            this.exportPDFBtn.Name = "exportPDFBtn";
            this.exportPDFBtn.Size = new System.Drawing.Size(43, 27);
            this.exportPDFBtn.TabIndex = 25;
            this.exportPDFBtn.Text = ".PDF";
            this.exportPDFBtn.UseVisualStyleBackColor = false;
            this.exportPDFBtn.Click += new System.EventHandler(this.exportPDFBtn_Click);
            // 
            // exportEXCELBtn
            // 
            this.exportEXCELBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.exportEXCELBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exportEXCELBtn.Location = new System.Drawing.Point(440, 0);
            this.exportEXCELBtn.Name = "exportEXCELBtn";
            this.exportEXCELBtn.Size = new System.Drawing.Size(43, 27);
            this.exportEXCELBtn.TabIndex = 24;
            this.exportEXCELBtn.Text = "EXCEL";
            this.exportEXCELBtn.UseVisualStyleBackColor = false;
            this.exportEXCELBtn.Click += new System.EventHandler(this.exportEXCELBtn_Click);
            // 
            // reportPnl
            // 
            this.reportPnl.Controls.Add(this.exportPnl);
            this.reportPnl.Controls.Add(this.reportTitleLbl);
            this.reportPnl.Controls.Add(this.reportLVI);
            this.reportPnl.Location = new System.Drawing.Point(3, 389);
            this.reportPnl.Name = "reportPnl";
            this.reportPnl.Size = new System.Drawing.Size(778, 319);
            this.reportPnl.TabIndex = 25;
            // 
            // mapPnl
            // 
            this.mapPnl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.mapPnl.Controls.Add(this.mapBrowser);
            this.mapPnl.ForeColor = System.Drawing.Color.White;
            this.mapPnl.Location = new System.Drawing.Point(3, 374);
            this.mapPnl.Name = "mapPnl";
            this.mapPnl.Size = new System.Drawing.Size(771, 331);
            this.mapPnl.TabIndex = 26;
            // 
            // mapBrowser
            // 
            this.mapBrowser.Location = new System.Drawing.Point(9, 3);
            this.mapBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.mapBrowser.Name = "mapBrowser";
            this.mapBrowser.Size = new System.Drawing.Size(755, 286);
            this.mapBrowser.TabIndex = 0;
            // 
            // itemSalesGeoBox
            // 
            this.itemSalesGeoBox.BackColor = System.Drawing.Color.Silver;
            this.itemSalesGeoBox.Controls.Add(this.label17);
            this.itemSalesGeoBox.Controls.Add(this.label14);
            this.itemSalesGeoBox.Controls.Add(this.itemSalesGeoEndDatePicker);
            this.itemSalesGeoBox.Controls.Add(this.itemSalesGeoStartDatePicker);
            this.itemSalesGeoBox.Controls.Add(this.label13);
            this.itemSalesGeoBox.Controls.Add(this.itemSalesGeoItemNumberTxt);
            this.itemSalesGeoBox.Controls.Add(this.itemSalesGeoCreateBtn);
            this.itemSalesGeoBox.Location = new System.Drawing.Point(234, 6);
            this.itemSalesGeoBox.Name = "itemSalesGeoBox";
            this.itemSalesGeoBox.Size = new System.Drawing.Size(547, 315);
            this.itemSalesGeoBox.TabIndex = 27;
            this.itemSalesGeoBox.TabStop = false;
            this.itemSalesGeoBox.Text = "Item Sales by Location:";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(17, 117);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(52, 13);
            this.label17.TabIndex = 47;
            this.label17.Text = "End Date";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(17, 74);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(55, 13);
            this.label14.TabIndex = 46;
            this.label14.Text = "Start Date";
            // 
            // itemSalesGeoEndDatePicker
            // 
            this.itemSalesGeoEndDatePicker.Location = new System.Drawing.Point(46, 130);
            this.itemSalesGeoEndDatePicker.Name = "itemSalesGeoEndDatePicker";
            this.itemSalesGeoEndDatePicker.Size = new System.Drawing.Size(200, 20);
            this.itemSalesGeoEndDatePicker.TabIndex = 45;
            // 
            // itemSalesGeoStartDatePicker
            // 
            this.itemSalesGeoStartDatePicker.Location = new System.Drawing.Point(46, 90);
            this.itemSalesGeoStartDatePicker.Name = "itemSalesGeoStartDatePicker";
            this.itemSalesGeoStartDatePicker.Size = new System.Drawing.Size(200, 20);
            this.itemSalesGeoStartDatePicker.TabIndex = 44;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(17, 44);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(67, 13);
            this.label13.TabIndex = 43;
            this.label13.Text = "Item Number";
            // 
            // itemSalesGeoItemNumberTxt
            // 
            this.itemSalesGeoItemNumberTxt.Location = new System.Drawing.Point(90, 40);
            this.itemSalesGeoItemNumberTxt.Name = "itemSalesGeoItemNumberTxt";
            this.itemSalesGeoItemNumberTxt.Size = new System.Drawing.Size(138, 20);
            this.itemSalesGeoItemNumberTxt.TabIndex = 42;
            // 
            // itemSalesGeoCreateBtn
            // 
            this.itemSalesGeoCreateBtn.AutoSize = true;
            this.itemSalesGeoCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.itemSalesGeoCreateBtn.Location = new System.Drawing.Point(450, 258);
            this.itemSalesGeoCreateBtn.Name = "itemSalesGeoCreateBtn";
            this.itemSalesGeoCreateBtn.Size = new System.Drawing.Size(83, 43);
            this.itemSalesGeoCreateBtn.TabIndex = 41;
            this.itemSalesGeoCreateBtn.Text = "Create Report";
            this.itemSalesGeoCreateBtn.UseVisualStyleBackColor = false;
            this.itemSalesGeoCreateBtn.Click += new System.EventHandler(this.itemSalesGeoCreateBtn_Click);
            // 
            // packagesShippedBox
            // 
            this.packagesShippedBox.BackColor = System.Drawing.Color.Silver;
            this.packagesShippedBox.Controls.Add(this.groupBox1);
            this.packagesShippedBox.Controls.Add(this.packagesShippedXDaysTxt);
            this.packagesShippedBox.Controls.Add(this.packagesShippedLastDaysRadio);
            this.packagesShippedBox.Controls.Add(this.packagesShippedDateRadio);
            this.packagesShippedBox.Controls.Add(this.packagesShippedEndDateLbl);
            this.packagesShippedBox.Controls.Add(this.packagesShippedStartDateLbl);
            this.packagesShippedBox.Controls.Add(this.packagesShippedEndDatePicker);
            this.packagesShippedBox.Controls.Add(this.packagesShippedStartDatePicker);
            this.packagesShippedBox.Controls.Add(this.packagesShippedCreateBtn);
            this.packagesShippedBox.Location = new System.Drawing.Point(227, 120);
            this.packagesShippedBox.Name = "packagesShippedBox";
            this.packagesShippedBox.Size = new System.Drawing.Size(547, 315);
            this.packagesShippedBox.TabIndex = 28;
            this.packagesShippedBox.TabStop = false;
            this.packagesShippedBox.Text = "Packages Shipped Via StarShip";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.packagingShippedTrackingIDChk);
            this.groupBox1.Controls.Add(this.packagesShippedPackagingUserChk);
            this.groupBox1.Controls.Add(this.packagesShippedPackageTypeChk);
            this.groupBox1.Controls.Add(this.packagesShippedDestinationAddressChk);
            this.groupBox1.Controls.Add(this.packagesShippedPackageDimensionsChk);
            this.groupBox1.Controls.Add(this.packagesShippedPostalCodeChk);
            this.groupBox1.Controls.Add(this.packagesShippedBilledWeightChk);
            this.groupBox1.Controls.Add(this.packagesShippedDeliveryChargeChk);
            this.groupBox1.Controls.Add(this.packagesShippedServiceNameChk);
            this.groupBox1.Controls.Add(this.packagesShippedProcessedDateChk);
            this.groupBox1.Location = new System.Drawing.Point(21, 180);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(391, 111);
            this.groupBox1.TabIndex = 51;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Choose Columns";
            // 
            // packagingShippedTrackingIDChk
            // 
            this.packagingShippedTrackingIDChk.AutoSize = true;
            this.packagingShippedTrackingIDChk.Location = new System.Drawing.Point(256, 44);
            this.packagingShippedTrackingIDChk.Name = "packagingShippedTrackingIDChk";
            this.packagingShippedTrackingIDChk.Size = new System.Drawing.Size(108, 17);
            this.packagingShippedTrackingIDChk.TabIndex = 9;
            this.packagingShippedTrackingIDChk.Text = "Tracking Number";
            this.packagingShippedTrackingIDChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedPackagingUserChk
            // 
            this.packagesShippedPackagingUserChk.AutoSize = true;
            this.packagesShippedPackagingUserChk.Location = new System.Drawing.Point(256, 27);
            this.packagesShippedPackagingUserChk.Name = "packagesShippedPackagingUserChk";
            this.packagesShippedPackagingUserChk.Size = new System.Drawing.Size(102, 17);
            this.packagesShippedPackagingUserChk.TabIndex = 8;
            this.packagesShippedPackagingUserChk.Text = "Packaging User";
            this.packagesShippedPackagingUserChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedPackageTypeChk
            // 
            this.packagesShippedPackageTypeChk.AutoSize = true;
            this.packagesShippedPackageTypeChk.Location = new System.Drawing.Point(124, 64);
            this.packagesShippedPackageTypeChk.Name = "packagesShippedPackageTypeChk";
            this.packagesShippedPackageTypeChk.Size = new System.Drawing.Size(96, 17);
            this.packagesShippedPackageTypeChk.TabIndex = 7;
            this.packagesShippedPackageTypeChk.Text = "Package Type";
            this.packagesShippedPackageTypeChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedDestinationAddressChk
            // 
            this.packagesShippedDestinationAddressChk.AutoSize = true;
            this.packagesShippedDestinationAddressChk.Location = new System.Drawing.Point(124, 82);
            this.packagesShippedDestinationAddressChk.Name = "packagesShippedDestinationAddressChk";
            this.packagesShippedDestinationAddressChk.Size = new System.Drawing.Size(120, 17);
            this.packagesShippedDestinationAddressChk.TabIndex = 6;
            this.packagesShippedDestinationAddressChk.Text = "Destination Address";
            this.packagesShippedDestinationAddressChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedPackageDimensionsChk
            // 
            this.packagesShippedPackageDimensionsChk.AutoSize = true;
            this.packagesShippedPackageDimensionsChk.Location = new System.Drawing.Point(124, 44);
            this.packagesShippedPackageDimensionsChk.Name = "packagesShippedPackageDimensionsChk";
            this.packagesShippedPackageDimensionsChk.Size = new System.Drawing.Size(126, 17);
            this.packagesShippedPackageDimensionsChk.TabIndex = 5;
            this.packagesShippedPackageDimensionsChk.Text = "Package Dimensions";
            this.packagesShippedPackageDimensionsChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedPostalCodeChk
            // 
            this.packagesShippedPostalCodeChk.AutoSize = true;
            this.packagesShippedPostalCodeChk.Checked = true;
            this.packagesShippedPostalCodeChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.packagesShippedPostalCodeChk.Location = new System.Drawing.Point(124, 26);
            this.packagesShippedPostalCodeChk.Name = "packagesShippedPostalCodeChk";
            this.packagesShippedPostalCodeChk.Size = new System.Drawing.Size(83, 17);
            this.packagesShippedPostalCodeChk.TabIndex = 4;
            this.packagesShippedPostalCodeChk.Text = "Postal Code";
            this.packagesShippedPostalCodeChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedBilledWeightChk
            // 
            this.packagesShippedBilledWeightChk.AutoSize = true;
            this.packagesShippedBilledWeightChk.Checked = true;
            this.packagesShippedBilledWeightChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.packagesShippedBilledWeightChk.Location = new System.Drawing.Point(16, 83);
            this.packagesShippedBilledWeightChk.Name = "packagesShippedBilledWeightChk";
            this.packagesShippedBilledWeightChk.Size = new System.Drawing.Size(88, 17);
            this.packagesShippedBilledWeightChk.TabIndex = 3;
            this.packagesShippedBilledWeightChk.Text = "Billed Weight";
            this.packagesShippedBilledWeightChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedDeliveryChargeChk
            // 
            this.packagesShippedDeliveryChargeChk.AutoSize = true;
            this.packagesShippedDeliveryChargeChk.Checked = true;
            this.packagesShippedDeliveryChargeChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.packagesShippedDeliveryChargeChk.Location = new System.Drawing.Point(16, 64);
            this.packagesShippedDeliveryChargeChk.Name = "packagesShippedDeliveryChargeChk";
            this.packagesShippedDeliveryChargeChk.Size = new System.Drawing.Size(101, 17);
            this.packagesShippedDeliveryChargeChk.TabIndex = 2;
            this.packagesShippedDeliveryChargeChk.Text = "Delivery Charge";
            this.packagesShippedDeliveryChargeChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedServiceNameChk
            // 
            this.packagesShippedServiceNameChk.AutoSize = true;
            this.packagesShippedServiceNameChk.Checked = true;
            this.packagesShippedServiceNameChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.packagesShippedServiceNameChk.Location = new System.Drawing.Point(16, 46);
            this.packagesShippedServiceNameChk.Name = "packagesShippedServiceNameChk";
            this.packagesShippedServiceNameChk.Size = new System.Drawing.Size(93, 17);
            this.packagesShippedServiceNameChk.TabIndex = 1;
            this.packagesShippedServiceNameChk.Text = "Service Name";
            this.packagesShippedServiceNameChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedProcessedDateChk
            // 
            this.packagesShippedProcessedDateChk.AutoSize = true;
            this.packagesShippedProcessedDateChk.Checked = true;
            this.packagesShippedProcessedDateChk.CheckState = System.Windows.Forms.CheckState.Checked;
            this.packagesShippedProcessedDateChk.Location = new System.Drawing.Point(16, 27);
            this.packagesShippedProcessedDateChk.Name = "packagesShippedProcessedDateChk";
            this.packagesShippedProcessedDateChk.Size = new System.Drawing.Size(102, 17);
            this.packagesShippedProcessedDateChk.TabIndex = 0;
            this.packagesShippedProcessedDateChk.Text = "Processed Date";
            this.packagesShippedProcessedDateChk.UseVisualStyleBackColor = true;
            // 
            // packagesShippedXDaysTxt
            // 
            this.packagesShippedXDaysTxt.Enabled = false;
            this.packagesShippedXDaysTxt.Location = new System.Drawing.Point(119, 129);
            this.packagesShippedXDaysTxt.MaxLength = 3;
            this.packagesShippedXDaysTxt.Name = "packagesShippedXDaysTxt";
            this.packagesShippedXDaysTxt.Size = new System.Drawing.Size(44, 20);
            this.packagesShippedXDaysTxt.TabIndex = 50;
            // 
            // packagesShippedLastDaysRadio
            // 
            this.packagesShippedLastDaysRadio.AutoSize = true;
            this.packagesShippedLastDaysRadio.Location = new System.Drawing.Point(22, 131);
            this.packagesShippedLastDaysRadio.Name = "packagesShippedLastDaysRadio";
            this.packagesShippedLastDaysRadio.Size = new System.Drawing.Size(85, 17);
            this.packagesShippedLastDaysRadio.TabIndex = 49;
            this.packagesShippedLastDaysRadio.Text = "Last xx Days";
            this.packagesShippedLastDaysRadio.UseVisualStyleBackColor = true;
            this.packagesShippedLastDaysRadio.CheckedChanged += new System.EventHandler(this.packagesShippedLastDaysRadio_CheckedChanged);
            // 
            // packagesShippedDateRadio
            // 
            this.packagesShippedDateRadio.AutoSize = true;
            this.packagesShippedDateRadio.Checked = true;
            this.packagesShippedDateRadio.Location = new System.Drawing.Point(21, 58);
            this.packagesShippedDateRadio.Name = "packagesShippedDateRadio";
            this.packagesShippedDateRadio.Size = new System.Drawing.Size(83, 17);
            this.packagesShippedDateRadio.TabIndex = 48;
            this.packagesShippedDateRadio.TabStop = true;
            this.packagesShippedDateRadio.Text = "Date Range";
            this.packagesShippedDateRadio.UseVisualStyleBackColor = true;
            this.packagesShippedDateRadio.CheckedChanged += new System.EventHandler(this.packagesShippedDateRadio_CheckedChanged);
            // 
            // packagesShippedEndDateLbl
            // 
            this.packagesShippedEndDateLbl.AutoSize = true;
            this.packagesShippedEndDateLbl.Location = new System.Drawing.Point(112, 76);
            this.packagesShippedEndDateLbl.Name = "packagesShippedEndDateLbl";
            this.packagesShippedEndDateLbl.Size = new System.Drawing.Size(52, 13);
            this.packagesShippedEndDateLbl.TabIndex = 47;
            this.packagesShippedEndDateLbl.Text = "End Date";
            // 
            // packagesShippedStartDateLbl
            // 
            this.packagesShippedStartDateLbl.AutoSize = true;
            this.packagesShippedStartDateLbl.Location = new System.Drawing.Point(112, 33);
            this.packagesShippedStartDateLbl.Name = "packagesShippedStartDateLbl";
            this.packagesShippedStartDateLbl.Size = new System.Drawing.Size(55, 13);
            this.packagesShippedStartDateLbl.TabIndex = 46;
            this.packagesShippedStartDateLbl.Text = "Start Date";
            // 
            // packagesShippedEndDatePicker
            // 
            this.packagesShippedEndDatePicker.Location = new System.Drawing.Point(141, 89);
            this.packagesShippedEndDatePicker.Name = "packagesShippedEndDatePicker";
            this.packagesShippedEndDatePicker.Size = new System.Drawing.Size(200, 20);
            this.packagesShippedEndDatePicker.TabIndex = 45;
            // 
            // packagesShippedStartDatePicker
            // 
            this.packagesShippedStartDatePicker.Location = new System.Drawing.Point(141, 49);
            this.packagesShippedStartDatePicker.Name = "packagesShippedStartDatePicker";
            this.packagesShippedStartDatePicker.Size = new System.Drawing.Size(200, 20);
            this.packagesShippedStartDatePicker.TabIndex = 44;
            // 
            // packagesShippedCreateBtn
            // 
            this.packagesShippedCreateBtn.AutoSize = true;
            this.packagesShippedCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.packagesShippedCreateBtn.Location = new System.Drawing.Point(450, 258);
            this.packagesShippedCreateBtn.Name = "packagesShippedCreateBtn";
            this.packagesShippedCreateBtn.Size = new System.Drawing.Size(83, 43);
            this.packagesShippedCreateBtn.TabIndex = 41;
            this.packagesShippedCreateBtn.Text = "Create Report";
            this.packagesShippedCreateBtn.UseVisualStyleBackColor = false;
            this.packagesShippedCreateBtn.Click += new System.EventHandler(this.packagesShippedCreateBtn_Click);
            // 
            // p2LinecodeSalesInfoBox
            // 
            this.p2LinecodeSalesInfoBox.BackColor = System.Drawing.Color.Silver;
            this.p2LinecodeSalesInfoBox.Controls.Add(this.p2LinecodeSalesInfoSettingsBox);
            this.p2LinecodeSalesInfoBox.Controls.Add(this.p2LinecodesTxt);
            this.p2LinecodeSalesInfoBox.Controls.Add(this.p2LinecodeSalesInfoHeaderLbl);
            this.p2LinecodeSalesInfoBox.Controls.Add(this.p2LinecodeSalesInfoCreateBtn);
            this.p2LinecodeSalesInfoBox.Location = new System.Drawing.Point(227, 5);
            this.p2LinecodeSalesInfoBox.Name = "p2LinecodeSalesInfoBox";
            this.p2LinecodeSalesInfoBox.Size = new System.Drawing.Size(547, 315);
            this.p2LinecodeSalesInfoBox.TabIndex = 52;
            this.p2LinecodeSalesInfoBox.TabStop = false;
            this.p2LinecodeSalesInfoBox.Text = "P2 Linecode Sales Info";
            // 
            // p2LinecodeSalesInfoCreateBtn
            // 
            this.p2LinecodeSalesInfoCreateBtn.AutoSize = true;
            this.p2LinecodeSalesInfoCreateBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.p2LinecodeSalesInfoCreateBtn.Location = new System.Drawing.Point(450, 258);
            this.p2LinecodeSalesInfoCreateBtn.Name = "p2LinecodeSalesInfoCreateBtn";
            this.p2LinecodeSalesInfoCreateBtn.Size = new System.Drawing.Size(83, 43);
            this.p2LinecodeSalesInfoCreateBtn.TabIndex = 41;
            this.p2LinecodeSalesInfoCreateBtn.Text = "Create Report";
            this.p2LinecodeSalesInfoCreateBtn.UseVisualStyleBackColor = false;
            this.p2LinecodeSalesInfoCreateBtn.Click += new System.EventHandler(this.p2LinecodeSalesInfoCreateBtn_Click);
            // 
            // p2LinecodeSalesInfoHeaderLbl
            // 
            this.p2LinecodeSalesInfoHeaderLbl.AutoSize = true;
            this.p2LinecodeSalesInfoHeaderLbl.Location = new System.Drawing.Point(17, 41);
            this.p2LinecodeSalesInfoHeaderLbl.Name = "p2LinecodeSalesInfoHeaderLbl";
            this.p2LinecodeSalesInfoHeaderLbl.Size = new System.Drawing.Size(81, 13);
            this.p2LinecodeSalesInfoHeaderLbl.TabIndex = 42;
            this.p2LinecodeSalesInfoHeaderLbl.Text = "Add Linecodes:";
            // 
            // p2LinecodesTxt
            // 
            this.p2LinecodesTxt.Location = new System.Drawing.Point(12, 57);
            this.p2LinecodesTxt.Multiline = true;
            this.p2LinecodesTxt.Name = "p2LinecodesTxt";
            this.p2LinecodesTxt.Size = new System.Drawing.Size(93, 98);
            this.p2LinecodesTxt.TabIndex = 43;
            // 
            // p2LinecodeSalesInfoItemChk
            // 
            this.p2LinecodeSalesInfoItemChk.AutoSize = true;
            this.p2LinecodeSalesInfoItemChk.Location = new System.Drawing.Point(283, 93);
            this.p2LinecodeSalesInfoItemChk.Name = "p2LinecodeSalesInfoItemChk";
            this.p2LinecodeSalesInfoItemChk.Size = new System.Drawing.Size(86, 17);
            this.p2LinecodeSalesInfoItemChk.TabIndex = 44;
            this.p2LinecodeSalesInfoItemChk.Text = "Item Number";
            this.p2LinecodeSalesInfoItemChk.UseVisualStyleBackColor = true;
            this.p2LinecodeSalesInfoItemChk.Visible = false;
            // 
            // p2LinecodeSalesInfoAddressChk
            // 
            this.p2LinecodeSalesInfoAddressChk.AutoSize = true;
            this.p2LinecodeSalesInfoAddressChk.Location = new System.Drawing.Point(16, 60);
            this.p2LinecodeSalesInfoAddressChk.Name = "p2LinecodeSalesInfoAddressChk";
            this.p2LinecodeSalesInfoAddressChk.Size = new System.Drawing.Size(64, 17);
            this.p2LinecodeSalesInfoAddressChk.TabIndex = 45;
            this.p2LinecodeSalesInfoAddressChk.Text = "Address";
            this.p2LinecodeSalesInfoAddressChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoNameChk
            // 
            this.p2LinecodeSalesInfoNameChk.AutoSize = true;
            this.p2LinecodeSalesInfoNameChk.Location = new System.Drawing.Point(16, 42);
            this.p2LinecodeSalesInfoNameChk.Name = "p2LinecodeSalesInfoNameChk";
            this.p2LinecodeSalesInfoNameChk.Size = new System.Drawing.Size(54, 17);
            this.p2LinecodeSalesInfoNameChk.TabIndex = 46;
            this.p2LinecodeSalesInfoNameChk.Text = "Name";
            this.p2LinecodeSalesInfoNameChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoPhoneChk
            // 
            this.p2LinecodeSalesInfoPhoneChk.AutoSize = true;
            this.p2LinecodeSalesInfoPhoneChk.Location = new System.Drawing.Point(16, 78);
            this.p2LinecodeSalesInfoPhoneChk.Name = "p2LinecodeSalesInfoPhoneChk";
            this.p2LinecodeSalesInfoPhoneChk.Size = new System.Drawing.Size(97, 17);
            this.p2LinecodeSalesInfoPhoneChk.TabIndex = 47;
            this.p2LinecodeSalesInfoPhoneChk.Text = "Phone Number";
            this.p2LinecodeSalesInfoPhoneChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoAccountCreationChk
            // 
            this.p2LinecodeSalesInfoAccountCreationChk.AutoSize = true;
            this.p2LinecodeSalesInfoAccountCreationChk.Location = new System.Drawing.Point(118, 60);
            this.p2LinecodeSalesInfoAccountCreationChk.Name = "p2LinecodeSalesInfoAccountCreationChk";
            this.p2LinecodeSalesInfoAccountCreationChk.Size = new System.Drawing.Size(134, 17);
            this.p2LinecodeSalesInfoAccountCreationChk.TabIndex = 48;
            this.p2LinecodeSalesInfoAccountCreationChk.Text = "Account Creation Date";
            this.p2LinecodeSalesInfoAccountCreationChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoLastPurchaseChk
            // 
            this.p2LinecodeSalesInfoLastPurchaseChk.AutoSize = true;
            this.p2LinecodeSalesInfoLastPurchaseChk.Location = new System.Drawing.Point(118, 78);
            this.p2LinecodeSalesInfoLastPurchaseChk.Name = "p2LinecodeSalesInfoLastPurchaseChk";
            this.p2LinecodeSalesInfoLastPurchaseChk.Size = new System.Drawing.Size(132, 17);
            this.p2LinecodeSalesInfoLastPurchaseChk.TabIndex = 49;
            this.p2LinecodeSalesInfoLastPurchaseChk.Text = "Date of Last Purchase";
            this.p2LinecodeSalesInfoLastPurchaseChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoSettingsBox
            // 
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoSalespersonIDChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoWebsiteChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoEmailChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoCustomerIDChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoLastPurchaseChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoItemChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoAccountCreationChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoAddressChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoPhoneChk);
            this.p2LinecodeSalesInfoSettingsBox.Controls.Add(this.p2LinecodeSalesInfoNameChk);
            this.p2LinecodeSalesInfoSettingsBox.Location = new System.Drawing.Point(10, 173);
            this.p2LinecodeSalesInfoSettingsBox.Name = "p2LinecodeSalesInfoSettingsBox";
            this.p2LinecodeSalesInfoSettingsBox.Size = new System.Drawing.Size(377, 116);
            this.p2LinecodeSalesInfoSettingsBox.TabIndex = 50;
            this.p2LinecodeSalesInfoSettingsBox.TabStop = false;
            this.p2LinecodeSalesInfoSettingsBox.Text = "Options";
            // 
            // p2LinecodeSalesInfoCustomerIDChk
            // 
            this.p2LinecodeSalesInfoCustomerIDChk.AutoSize = true;
            this.p2LinecodeSalesInfoCustomerIDChk.Location = new System.Drawing.Point(16, 23);
            this.p2LinecodeSalesInfoCustomerIDChk.Name = "p2LinecodeSalesInfoCustomerIDChk";
            this.p2LinecodeSalesInfoCustomerIDChk.Size = new System.Drawing.Size(84, 17);
            this.p2LinecodeSalesInfoCustomerIDChk.TabIndex = 50;
            this.p2LinecodeSalesInfoCustomerIDChk.Text = "Customer ID";
            this.p2LinecodeSalesInfoCustomerIDChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoEmailChk
            // 
            this.p2LinecodeSalesInfoEmailChk.AutoSize = true;
            this.p2LinecodeSalesInfoEmailChk.Location = new System.Drawing.Point(118, 23);
            this.p2LinecodeSalesInfoEmailChk.Name = "p2LinecodeSalesInfoEmailChk";
            this.p2LinecodeSalesInfoEmailChk.Size = new System.Drawing.Size(92, 17);
            this.p2LinecodeSalesInfoEmailChk.TabIndex = 51;
            this.p2LinecodeSalesInfoEmailChk.Text = "Email Address";
            this.p2LinecodeSalesInfoEmailChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoWebsiteChk
            // 
            this.p2LinecodeSalesInfoWebsiteChk.AutoSize = true;
            this.p2LinecodeSalesInfoWebsiteChk.Location = new System.Drawing.Point(118, 42);
            this.p2LinecodeSalesInfoWebsiteChk.Name = "p2LinecodeSalesInfoWebsiteChk";
            this.p2LinecodeSalesInfoWebsiteChk.Size = new System.Drawing.Size(112, 17);
            this.p2LinecodeSalesInfoWebsiteChk.TabIndex = 52;
            this.p2LinecodeSalesInfoWebsiteChk.Text = "Customer Website";
            this.p2LinecodeSalesInfoWebsiteChk.UseVisualStyleBackColor = true;
            // 
            // p2LinecodeSalesInfoSalespersonIDChk
            // 
            this.p2LinecodeSalesInfoSalespersonIDChk.AutoSize = true;
            this.p2LinecodeSalesInfoSalespersonIDChk.Location = new System.Drawing.Point(267, 23);
            this.p2LinecodeSalesInfoSalespersonIDChk.Name = "p2LinecodeSalesInfoSalespersonIDChk";
            this.p2LinecodeSalesInfoSalespersonIDChk.Size = new System.Drawing.Size(102, 17);
            this.p2LinecodeSalesInfoSalespersonIDChk.TabIndex = 53;
            this.p2LinecodeSalesInfoSalespersonIDChk.Text = "Sales Person ID";
            this.p2LinecodeSalesInfoSalespersonIDChk.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 718);
            this.Controls.Add(this.p2LinecodeSalesInfoBox);
            this.Controls.Add(this.packagesShippedBox);
            this.Controls.Add(this.itemSalesGeoBox);
            this.Controls.Add(this.mapPnl);
            this.Controls.Add(this.reportTreeBox);
            this.Controls.Add(this.statusHeaderLbl);
            this.Controls.Add(this.previewBtn);
            this.Controls.Add(this.statusLbl);
            this.Controls.Add(this.salesTxBox);
            this.Controls.Add(this.POBackorderBox);
            this.Controls.Add(this.splashBox);
            this.Controls.Add(this.orderDiscrepancyBox);
            this.Controls.Add(this.reportPnl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Report Generator";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.POBackorderBox.ResumeLayout(false);
            this.POBackorderBox.PerformLayout();
            this.splashBox.ResumeLayout(false);
            this.splashBox.PerformLayout();
            this.salesTxBox.ResumeLayout(false);
            this.salesTxBox.PerformLayout();
            this.orderDiscrepancyBox.ResumeLayout(false);
            this.orderDiscrepancyBox.PerformLayout();
            this.reportTreeBox.ResumeLayout(false);
            this.exportPnl.ResumeLayout(false);
            this.exportPnl.PerformLayout();
            this.reportPnl.ResumeLayout(false);
            this.reportPnl.PerformLayout();
            this.mapPnl.ResumeLayout(false);
            this.itemSalesGeoBox.ResumeLayout(false);
            this.itemSalesGeoBox.PerformLayout();
            this.packagesShippedBox.ResumeLayout(false);
            this.packagesShippedBox.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.p2LinecodeSalesInfoBox.ResumeLayout(false);
            this.p2LinecodeSalesInfoBox.PerformLayout();
            this.p2LinecodeSalesInfoSettingsBox.ResumeLayout(false);
            this.p2LinecodeSalesInfoSettingsBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker poboStartDatePicker;
        private System.Windows.Forms.DateTimePicker poboEndDatePicker;
        private System.Windows.Forms.GroupBox POBackorderBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox poboLineCodeEnd;
        private System.Windows.Forms.ComboBox poboLineCodeStart;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox splashBox;
        private System.Windows.Forms.Button poboCreateBtn;
        private System.Windows.Forms.ListView reportLVI;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label statusLbl;
        private System.Windows.Forms.CheckBox poboHiddenReportTitleChk;
        private System.Windows.Forms.Timer connectionTmr;
        private System.Windows.Forms.GroupBox salesTxBox;
        private System.Windows.Forms.Button previewBtn;
        private System.Windows.Forms.CheckBox salesTxBlankLineBetweenOrdersChk;
        private System.Windows.Forms.CheckBox salesTxSkipPartiallyInvoicedChk;
        private System.Windows.Forms.CheckBox salesTxShowMiscShipLinesChk;
        private System.Windows.Forms.CheckBox salesTxShowMiscLinesChk;
        private System.Windows.Forms.CheckBox salesTxOrdersBelowMinChk;
        private System.Windows.Forms.CheckBox salesTxIncludeCostChk;
        private System.Windows.Forms.CheckBox salesTxCheckOdditiesChk;
        private System.Windows.Forms.CheckBox salesTxCheckPOSChk;
        private System.Windows.Forms.CheckBox salesTxCheckSOChk;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox salesTxItemsTxt;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox salesTxOrderMinTxt;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox salesTxCouponAmountTxt;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox salesTxMinSaleItemCountTxt;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.DateTimePicker salesTxEndDatePicker;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.DateTimePicker salesTxStartDatePicker;
        private System.Windows.Forms.Button salesTxCreateBtn;
        private System.Windows.Forms.Label statusHeaderLbl;
        private System.Windows.Forms.CheckBox salesTxIdColumnChk;
        private System.Windows.Forms.GroupBox orderDiscrepancyBox;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button orderDiscrepancyCreateBtn;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.GroupBox reportTreeBox;
        private System.Windows.Forms.TreeView reportTree;
        private System.Windows.Forms.Label reportTitleLbl;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker salesTxReportEndDatePicker;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.DateTimePicker salesTxReportStartDatePicker;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button exportCSVBtn;
        private System.Windows.Forms.Button exportXLSBtn;
        private System.Windows.Forms.Button exportTXTBtn;
        private System.Windows.Forms.Button exportHTMLBtn;
        private System.Windows.Forms.Button exportXMLBtn;
        private System.Windows.Forms.Button exportJSONBtn;
        private System.Windows.Forms.Panel exportPnl;
        private System.Windows.Forms.SaveFileDialog fileSaveDlg;
        private System.Windows.Forms.Button exportEXCELBtn;
        private System.Windows.Forms.Button exportPDFBtn;
        private System.Windows.Forms.Button exportEMAILBtn;
        private System.Windows.Forms.Button exportWORDBtn;
        private System.Windows.Forms.Button exportDOCXBtn;
        private System.Windows.Forms.Panel reportPnl;
        private System.Windows.Forms.Panel mapPnl;
        private System.Windows.Forms.WebBrowser mapBrowser;
        private System.Windows.Forms.GroupBox itemSalesGeoBox;
        private System.Windows.Forms.Button itemSalesGeoCreateBtn;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.DateTimePicker itemSalesGeoEndDatePicker;
        private System.Windows.Forms.DateTimePicker itemSalesGeoStartDatePicker;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox itemSalesGeoItemNumberTxt;
        private System.Windows.Forms.GroupBox packagesShippedBox;
        private System.Windows.Forms.TextBox packagesShippedXDaysTxt;
        private System.Windows.Forms.RadioButton packagesShippedLastDaysRadio;
        private System.Windows.Forms.RadioButton packagesShippedDateRadio;
        private System.Windows.Forms.Label packagesShippedEndDateLbl;
        private System.Windows.Forms.Label packagesShippedStartDateLbl;
        private System.Windows.Forms.DateTimePicker packagesShippedEndDatePicker;
        private System.Windows.Forms.DateTimePicker packagesShippedStartDatePicker;
        private System.Windows.Forms.Button packagesShippedCreateBtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox packagesShippedDestinationAddressChk;
        private System.Windows.Forms.CheckBox packagesShippedPackageDimensionsChk;
        private System.Windows.Forms.CheckBox packagesShippedPostalCodeChk;
        private System.Windows.Forms.CheckBox packagesShippedBilledWeightChk;
        private System.Windows.Forms.CheckBox packagesShippedDeliveryChargeChk;
        private System.Windows.Forms.CheckBox packagesShippedServiceNameChk;
        private System.Windows.Forms.CheckBox packagesShippedProcessedDateChk;
        private System.Windows.Forms.CheckBox packagingShippedTrackingIDChk;
        private System.Windows.Forms.CheckBox packagesShippedPackagingUserChk;
        private System.Windows.Forms.CheckBox packagesShippedPackageTypeChk;
        private System.Windows.Forms.GroupBox p2LinecodeSalesInfoBox;
        private System.Windows.Forms.GroupBox p2LinecodeSalesInfoSettingsBox;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoLastPurchaseChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoItemChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoAccountCreationChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoAddressChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoPhoneChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoNameChk;
        private System.Windows.Forms.TextBox p2LinecodesTxt;
        private System.Windows.Forms.Label p2LinecodeSalesInfoHeaderLbl;
        private System.Windows.Forms.Button p2LinecodeSalesInfoCreateBtn;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoSalespersonIDChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoWebsiteChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoEmailChk;
        private System.Windows.Forms.CheckBox p2LinecodeSalesInfoCustomerIDChk;
    }
}

