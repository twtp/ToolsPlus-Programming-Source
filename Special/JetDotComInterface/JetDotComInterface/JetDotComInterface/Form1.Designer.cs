namespace JetDotComInterface
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
            this.apiUsernameTxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.apiPasswordTxt = new System.Windows.Forms.TextBox();
            this.getAllJetAcceptedOrdersBtn = new System.Windows.Forms.Button();
            this.titleLbl = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.apiMerchantTxt = new System.Windows.Forms.TextBox();
            this.apiInfoBox = new System.Windows.Forms.GroupBox();
            this.apiBearerTxt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.plusMinus1Lbl = new System.Windows.Forms.Label();
            this.currentTimeLbl = new System.Windows.Forms.Label();
            this.secondTmr = new System.Windows.Forms.Timer(this.components);
            this.statusLVI = new System.Windows.Forms.ListView();
            this.itemsLVI = new System.Windows.Forms.ListView();
            this.itemsOnJetLbl = new System.Windows.Forms.Label();
            this.getAllJetItemsBtn = new System.Windows.Forms.Button();
            this.readyOrdersLbl = new System.Windows.Forms.Label();
            this.readyOrderLVI = new System.Windows.Forms.ListView();
            this.pendingOrderLbl = new System.Windows.Forms.Label();
            this.pendingOrderLVI = new System.Windows.Forms.ListView();
            this.getAllJetPendingOrdersBtn = new System.Windows.Forms.Button();
            this.statusBox = new System.Windows.Forms.GroupBox();
            this.plusMinus2Lbl = new System.Windows.Forms.Label();
            this.itemInfoBox = new System.Windows.Forms.GroupBox();
            this.plusMinus3Lbl = new System.Windows.Forms.Label();
            this.orderInfoBox = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.plusMinus4Lbl = new System.Windows.Forms.Label();
            this.gotoJetDashboardBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.controlsBtn = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.processedOrdersBox = new System.Windows.Forms.GroupBox();
            this.processedOrdersLVI = new System.Windows.Forms.ListView();
            this.plusMinus5Lbl = new System.Windows.Forms.Label();
            this.processNextOrderBtn = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.updateIntervalTrk = new System.Windows.Forms.TrackBar();
            this.updateIntervalLbl = new System.Windows.Forms.Label();
            this.nextUpdateLbl = new System.Windows.Forms.Label();
            this.acknowledgedOrderLVI = new System.Windows.Forms.ListView();
            this.acknowledgedOrderLbl = new System.Windows.Forms.Label();
            this.inProgressOrderLVI = new System.Windows.Forms.ListView();
            this.inProgressOrderLbl = new System.Windows.Forms.Label();
            this.completedOrderLVI = new System.Windows.Forms.ListView();
            this.completedOrderLbl = new System.Windows.Forms.Label();
            this.getAckOrdersBtn = new System.Windows.Forms.Button();
            this.getProgressOrdersBtn = new System.Windows.Forms.Button();
            this.getCompletedOrdersBtn = new System.Windows.Forms.Button();
            this.getReturnOrdersBtn = new System.Windows.Forms.Button();
            this.returnOrdersBox = new System.Windows.Forms.GroupBox();
            this.plusMinus6Lbl = new System.Windows.Forms.Label();
            this.apiInfoBox.SuspendLayout();
            this.statusBox.SuspendLayout();
            this.itemInfoBox.SuspendLayout();
            this.orderInfoBox.SuspendLayout();
            this.processedOrdersBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.updateIntervalTrk)).BeginInit();
            this.returnOrdersBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // apiUsernameTxt
            // 
            this.apiUsernameTxt.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.apiUsernameTxt.ForeColor = System.Drawing.Color.White;
            this.apiUsernameTxt.Location = new System.Drawing.Point(141, 54);
            this.apiUsernameTxt.Name = "apiUsernameTxt";
            this.apiUsernameTxt.ReadOnly = true;
            this.apiUsernameTxt.Size = new System.Drawing.Size(324, 20);
            this.apiUsernameTxt.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(20, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Jet.com API Username";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(20, 83);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Jet.com API Password";
            // 
            // apiPasswordTxt
            // 
            this.apiPasswordTxt.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.apiPasswordTxt.ForeColor = System.Drawing.Color.White;
            this.apiPasswordTxt.Location = new System.Drawing.Point(141, 80);
            this.apiPasswordTxt.Name = "apiPasswordTxt";
            this.apiPasswordTxt.ReadOnly = true;
            this.apiPasswordTxt.Size = new System.Drawing.Size(324, 20);
            this.apiPasswordTxt.TabIndex = 2;
            // 
            // getAllJetAcceptedOrdersBtn
            // 
            this.getAllJetAcceptedOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getAllJetAcceptedOrdersBtn.Location = new System.Drawing.Point(778, 80);
            this.getAllJetAcceptedOrdersBtn.Name = "getAllJetAcceptedOrdersBtn";
            this.getAllJetAcceptedOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getAllJetAcceptedOrdersBtn.TabIndex = 5;
            this.getAllJetAcceptedOrdersBtn.Text = "Get All Jet.com Ready Orders";
            this.getAllJetAcceptedOrdersBtn.UseVisualStyleBackColor = false;
            this.getAllJetAcceptedOrdersBtn.Click += new System.EventHandler(this.getAllJetAcceptedOrdersBtn_Click);
            // 
            // titleLbl
            // 
            this.titleLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLbl.ForeColor = System.Drawing.Color.MediumPurple;
            this.titleLbl.Location = new System.Drawing.Point(12, 9);
            this.titleLbl.Name = "titleLbl";
            this.titleLbl.Size = new System.Drawing.Size(353, 24);
            this.titleLbl.TabIndex = 8;
            this.titleLbl.Text = "Jet.com Standalone Management Center";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(20, 31);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Jet.com API Merchant";
            // 
            // apiMerchantTxt
            // 
            this.apiMerchantTxt.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.apiMerchantTxt.ForeColor = System.Drawing.Color.White;
            this.apiMerchantTxt.Location = new System.Drawing.Point(141, 28);
            this.apiMerchantTxt.Name = "apiMerchantTxt";
            this.apiMerchantTxt.ReadOnly = true;
            this.apiMerchantTxt.Size = new System.Drawing.Size(324, 20);
            this.apiMerchantTxt.TabIndex = 9;
            // 
            // apiInfoBox
            // 
            this.apiInfoBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.apiInfoBox.Controls.Add(this.apiBearerTxt);
            this.apiInfoBox.Controls.Add(this.label4);
            this.apiInfoBox.Controls.Add(this.apiMerchantTxt);
            this.apiInfoBox.Controls.Add(this.apiUsernameTxt);
            this.apiInfoBox.Controls.Add(this.label3);
            this.apiInfoBox.Controls.Add(this.label1);
            this.apiInfoBox.Controls.Add(this.plusMinus1Lbl);
            this.apiInfoBox.Controls.Add(this.apiPasswordTxt);
            this.apiInfoBox.Controls.Add(this.label2);
            this.apiInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.apiInfoBox.ForeColor = System.Drawing.Color.White;
            this.apiInfoBox.Location = new System.Drawing.Point(16, 48);
            this.apiInfoBox.Name = "apiInfoBox";
            this.apiInfoBox.Size = new System.Drawing.Size(486, 20);
            this.apiInfoBox.TabIndex = 11;
            this.apiInfoBox.TabStop = false;
            this.apiInfoBox.Text = "       API Info";
            // 
            // apiBearerTxt
            // 
            this.apiBearerTxt.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.apiBearerTxt.ForeColor = System.Drawing.Color.White;
            this.apiBearerTxt.Location = new System.Drawing.Point(141, 106);
            this.apiBearerTxt.Multiline = true;
            this.apiBearerTxt.Name = "apiBearerTxt";
            this.apiBearerTxt.ReadOnly = true;
            this.apiBearerTxt.Size = new System.Drawing.Size(324, 88);
            this.apiBearerTxt.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(20, 109);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(103, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Jet.com Auth Bearer";
            // 
            // plusMinus1Lbl
            // 
            this.plusMinus1Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus1Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus1Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus1Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus1Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus1Lbl.Location = new System.Drawing.Point(1, 0);
            this.plusMinus1Lbl.Name = "plusMinus1Lbl";
            this.plusMinus1Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus1Lbl.TabIndex = 12;
            this.plusMinus1Lbl.Text = "+";
            this.plusMinus1Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus1Lbl.Click += new System.EventHandler(this.plusMinus1Lbl_Click);
            // 
            // currentTimeLbl
            // 
            this.currentTimeLbl.AutoSize = true;
            this.currentTimeLbl.ForeColor = System.Drawing.Color.MediumPurple;
            this.currentTimeLbl.Location = new System.Drawing.Point(420, 9);
            this.currentTimeLbl.Name = "currentTimeLbl";
            this.currentTimeLbl.Size = new System.Drawing.Size(73, 13);
            this.currentTimeLbl.TabIndex = 13;
            this.currentTimeLbl.Text = "Current Time: ";
            // 
            // secondTmr
            // 
            this.secondTmr.Enabled = true;
            this.secondTmr.Interval = 1000;
            this.secondTmr.Tick += new System.EventHandler(this.secondTmr_Tick);
            // 
            // statusLVI
            // 
            this.statusLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.statusLVI.ForeColor = System.Drawing.Color.White;
            this.statusLVI.FullRowSelect = true;
            this.statusLVI.GridLines = true;
            this.statusLVI.Location = new System.Drawing.Point(10, 27);
            this.statusLVI.MultiSelect = false;
            this.statusLVI.Name = "statusLVI";
            this.statusLVI.Size = new System.Drawing.Size(717, 168);
            this.statusLVI.TabIndex = 14;
            this.statusLVI.UseCompatibleStateImageBehavior = false;
            this.statusLVI.View = System.Windows.Forms.View.Details;
            // 
            // itemsLVI
            // 
            this.itemsLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.itemsLVI.ForeColor = System.Drawing.Color.White;
            this.itemsLVI.FullRowSelect = true;
            this.itemsLVI.GridLines = true;
            this.itemsLVI.Location = new System.Drawing.Point(23, 60);
            this.itemsLVI.MultiSelect = false;
            this.itemsLVI.Name = "itemsLVI";
            this.itemsLVI.Size = new System.Drawing.Size(197, 168);
            this.itemsLVI.TabIndex = 15;
            this.itemsLVI.UseCompatibleStateImageBehavior = false;
            this.itemsLVI.View = System.Windows.Forms.View.Details;
            // 
            // itemsOnJetLbl
            // 
            this.itemsOnJetLbl.AutoSize = true;
            this.itemsOnJetLbl.Location = new System.Drawing.Point(12, 27);
            this.itemsOnJetLbl.Name = "itemsOnJetLbl";
            this.itemsOnJetLbl.Size = new System.Drawing.Size(105, 13);
            this.itemsOnJetLbl.TabIndex = 16;
            this.itemsOnJetLbl.Text = "Items On Jet.com";
            // 
            // getAllJetItemsBtn
            // 
            this.getAllJetItemsBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getAllJetItemsBtn.Location = new System.Drawing.Point(778, 6);
            this.getAllJetItemsBtn.Name = "getAllJetItemsBtn";
            this.getAllJetItemsBtn.Size = new System.Drawing.Size(184, 38);
            this.getAllJetItemsBtn.TabIndex = 17;
            this.getAllJetItemsBtn.Text = "Get All Jet.com Items";
            this.getAllJetItemsBtn.UseVisualStyleBackColor = false;
            this.getAllJetItemsBtn.Click += new System.EventHandler(this.getAllJetItemsBtn_Click);
            // 
            // readyOrdersLbl
            // 
            this.readyOrdersLbl.AutoSize = true;
            this.readyOrdersLbl.Location = new System.Drawing.Point(166, 28);
            this.readyOrdersLbl.Name = "readyOrdersLbl";
            this.readyOrdersLbl.Size = new System.Drawing.Size(101, 13);
            this.readyOrdersLbl.TabIndex = 19;
            this.readyOrdersLbl.Text = "Ready To Import";
            // 
            // readyOrderLVI
            // 
            this.readyOrderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.readyOrderLVI.ForeColor = System.Drawing.Color.White;
            this.readyOrderLVI.FullRowSelect = true;
            this.readyOrderLVI.GridLines = true;
            this.readyOrderLVI.Location = new System.Drawing.Point(142, 58);
            this.readyOrderLVI.MultiSelect = false;
            this.readyOrderLVI.Name = "readyOrderLVI";
            this.readyOrderLVI.Size = new System.Drawing.Size(149, 168);
            this.readyOrderLVI.TabIndex = 18;
            this.readyOrderLVI.UseCompatibleStateImageBehavior = false;
            this.readyOrderLVI.View = System.Windows.Forms.View.Details;
            // 
            // pendingOrderLbl
            // 
            this.pendingOrderLbl.AutoSize = true;
            this.pendingOrderLbl.Location = new System.Drawing.Point(21, 28);
            this.pendingOrderLbl.Name = "pendingOrderLbl";
            this.pendingOrderLbl.Size = new System.Drawing.Size(94, 13);
            this.pendingOrderLbl.TabIndex = 21;
            this.pendingOrderLbl.Text = "Pending Orders";
            // 
            // pendingOrderLVI
            // 
            this.pendingOrderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.pendingOrderLVI.ForeColor = System.Drawing.Color.White;
            this.pendingOrderLVI.FullRowSelect = true;
            this.pendingOrderLVI.GridLines = true;
            this.pendingOrderLVI.Location = new System.Drawing.Point(4, 58);
            this.pendingOrderLVI.MultiSelect = false;
            this.pendingOrderLVI.Name = "pendingOrderLVI";
            this.pendingOrderLVI.Size = new System.Drawing.Size(136, 168);
            this.pendingOrderLVI.TabIndex = 20;
            this.pendingOrderLVI.UseCompatibleStateImageBehavior = false;
            this.pendingOrderLVI.View = System.Windows.Forms.View.Details;
            // 
            // getAllJetPendingOrdersBtn
            // 
            this.getAllJetPendingOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getAllJetPendingOrdersBtn.Location = new System.Drawing.Point(778, 43);
            this.getAllJetPendingOrdersBtn.Name = "getAllJetPendingOrdersBtn";
            this.getAllJetPendingOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getAllJetPendingOrdersBtn.TabIndex = 22;
            this.getAllJetPendingOrdersBtn.Text = "Get All Jet.com Pending Orders";
            this.getAllJetPendingOrdersBtn.UseVisualStyleBackColor = false;
            this.getAllJetPendingOrdersBtn.Click += new System.EventHandler(this.getAllJetPendingOrdersBtn_Click);
            // 
            // statusBox
            // 
            this.statusBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.statusBox.Controls.Add(this.plusMinus2Lbl);
            this.statusBox.Controls.Add(this.statusLVI);
            this.statusBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.statusBox.ForeColor = System.Drawing.Color.White;
            this.statusBox.Location = new System.Drawing.Point(16, 78);
            this.statusBox.Name = "statusBox";
            this.statusBox.Size = new System.Drawing.Size(737, 20);
            this.statusBox.TabIndex = 23;
            this.statusBox.TabStop = false;
            this.statusBox.Text = "       Status";
            this.statusBox.Enter += new System.EventHandler(this.statusBox_Enter);
            // 
            // plusMinus2Lbl
            // 
            this.plusMinus2Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus2Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus2Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus2Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus2Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus2Lbl.Location = new System.Drawing.Point(1, 0);
            this.plusMinus2Lbl.Name = "plusMinus2Lbl";
            this.plusMinus2Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus2Lbl.TabIndex = 13;
            this.plusMinus2Lbl.Text = "+";
            this.plusMinus2Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus2Lbl.Click += new System.EventHandler(this.plusMinus2Lbl_Click);
            // 
            // itemInfoBox
            // 
            this.itemInfoBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.itemInfoBox.Controls.Add(this.plusMinus3Lbl);
            this.itemInfoBox.Controls.Add(this.itemsLVI);
            this.itemInfoBox.Controls.Add(this.itemsOnJetLbl);
            this.itemInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.itemInfoBox.ForeColor = System.Drawing.Color.White;
            this.itemInfoBox.Location = new System.Drawing.Point(16, 108);
            this.itemInfoBox.Name = "itemInfoBox";
            this.itemInfoBox.Size = new System.Drawing.Size(238, 20);
            this.itemInfoBox.TabIndex = 24;
            this.itemInfoBox.TabStop = false;
            this.itemInfoBox.Text = "       Item Info";
            // 
            // plusMinus3Lbl
            // 
            this.plusMinus3Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus3Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus3Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus3Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus3Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus3Lbl.Location = new System.Drawing.Point(1, -1);
            this.plusMinus3Lbl.Name = "plusMinus3Lbl";
            this.plusMinus3Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus3Lbl.TabIndex = 25;
            this.plusMinus3Lbl.Text = "+";
            this.plusMinus3Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus3Lbl.Click += new System.EventHandler(this.plusMinus3Lbl_Click);
            // 
            // orderInfoBox
            // 
            this.orderInfoBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.orderInfoBox.Controls.Add(this.completedOrderLVI);
            this.orderInfoBox.Controls.Add(this.completedOrderLbl);
            this.orderInfoBox.Controls.Add(this.inProgressOrderLVI);
            this.orderInfoBox.Controls.Add(this.inProgressOrderLbl);
            this.orderInfoBox.Controls.Add(this.acknowledgedOrderLVI);
            this.orderInfoBox.Controls.Add(this.acknowledgedOrderLbl);
            this.orderInfoBox.Controls.Add(this.label5);
            this.orderInfoBox.Controls.Add(this.plusMinus4Lbl);
            this.orderInfoBox.Controls.Add(this.pendingOrderLbl);
            this.orderInfoBox.Controls.Add(this.readyOrderLVI);
            this.orderInfoBox.Controls.Add(this.readyOrdersLbl);
            this.orderInfoBox.Controls.Add(this.pendingOrderLVI);
            this.orderInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.orderInfoBox.ForeColor = System.Drawing.Color.White;
            this.orderInfoBox.Location = new System.Drawing.Point(16, 139);
            this.orderInfoBox.Name = "orderInfoBox";
            this.orderInfoBox.Size = new System.Drawing.Size(748, 20);
            this.orderInfoBox.TabIndex = 25;
            this.orderInfoBox.TabStop = false;
            this.orderInfoBox.Text = "       Order Info";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(104, 235);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(364, 13);
            this.label5.TabIndex = 27;
            this.label5.Text = "Pending orders are orders still passing fraud checks at Jet.com";
            // 
            // plusMinus4Lbl
            // 
            this.plusMinus4Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus4Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus4Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus4Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus4Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus4Lbl.Location = new System.Drawing.Point(1, 0);
            this.plusMinus4Lbl.Name = "plusMinus4Lbl";
            this.plusMinus4Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus4Lbl.TabIndex = 26;
            this.plusMinus4Lbl.Text = "+";
            this.plusMinus4Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus4Lbl.Click += new System.EventHandler(this.label5_Click);
            // 
            // gotoJetDashboardBtn
            // 
            this.gotoJetDashboardBtn.BackColor = System.Drawing.Color.SpringGreen;
            this.gotoJetDashboardBtn.Location = new System.Drawing.Point(962, 6);
            this.gotoJetDashboardBtn.Name = "gotoJetDashboardBtn";
            this.gotoJetDashboardBtn.Size = new System.Drawing.Size(184, 38);
            this.gotoJetDashboardBtn.TabIndex = 27;
            this.gotoJetDashboardBtn.Text = "Go to Jet.com Dashboard";
            this.gotoJetDashboardBtn.UseVisualStyleBackColor = false;
            this.gotoJetDashboardBtn.Click += new System.EventHandler(this.gotoJetDashboardBtn_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.SpringGreen;
            this.button1.Location = new System.Drawing.Point(962, 43);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(184, 38);
            this.button1.TabIndex = 28;
            this.button1.Text = "Go to Jet.com API Documentation";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // controlsBtn
            // 
            this.controlsBtn.BackColor = System.Drawing.Color.Indigo;
            this.controlsBtn.ForeColor = System.Drawing.Color.White;
            this.controlsBtn.Location = new System.Drawing.Point(661, 48);
            this.controlsBtn.Name = "controlsBtn";
            this.controlsBtn.Size = new System.Drawing.Size(92, 24);
            this.controlsBtn.TabIndex = 29;
            this.controlsBtn.Text = "Controls >";
            this.controlsBtn.UseVisualStyleBackColor = false;
            this.controlsBtn.Click += new System.EventHandler(this.controlsBtn_Click);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.SpringGreen;
            this.button2.Location = new System.Drawing.Point(962, 80);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(184, 38);
            this.button2.TabIndex = 30;
            this.button2.Text = "Email Jet.com Developer Team";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.SpringGreen;
            this.button3.Location = new System.Drawing.Point(962, 117);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(184, 38);
            this.button3.TabIndex = 31;
            this.button3.Text = "Jet.com Terms Of Service";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // processedOrdersBox
            // 
            this.processedOrdersBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.processedOrdersBox.Controls.Add(this.processedOrdersLVI);
            this.processedOrdersBox.Controls.Add(this.plusMinus5Lbl);
            this.processedOrdersBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.processedOrdersBox.ForeColor = System.Drawing.Color.White;
            this.processedOrdersBox.Location = new System.Drawing.Point(16, 174);
            this.processedOrdersBox.Name = "processedOrdersBox";
            this.processedOrdersBox.Size = new System.Drawing.Size(725, 20);
            this.processedOrdersBox.TabIndex = 32;
            this.processedOrdersBox.TabStop = false;
            this.processedOrdersBox.Text = "       Processed Orders";
            // 
            // processedOrdersLVI
            // 
            this.processedOrdersLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.processedOrdersLVI.ForeColor = System.Drawing.Color.White;
            this.processedOrdersLVI.FullRowSelect = true;
            this.processedOrdersLVI.GridLines = true;
            this.processedOrdersLVI.Location = new System.Drawing.Point(15, 31);
            this.processedOrdersLVI.Name = "processedOrdersLVI";
            this.processedOrdersLVI.Size = new System.Drawing.Size(689, 149);
            this.processedOrdersLVI.TabIndex = 34;
            this.processedOrdersLVI.UseCompatibleStateImageBehavior = false;
            this.processedOrdersLVI.View = System.Windows.Forms.View.Details;
            this.processedOrdersLVI.SelectedIndexChanged += new System.EventHandler(this.processedOrdersLVI_SelectedIndexChanged);
            // 
            // plusMinus5Lbl
            // 
            this.plusMinus5Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus5Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus5Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus5Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus5Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus5Lbl.Location = new System.Drawing.Point(1, -2);
            this.plusMinus5Lbl.Name = "plusMinus5Lbl";
            this.plusMinus5Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus5Lbl.TabIndex = 33;
            this.plusMinus5Lbl.Text = "+";
            this.plusMinus5Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus5Lbl.Click += new System.EventHandler(this.plusMinus5Lbl_Click);
            // 
            // processNextOrderBtn
            // 
            this.processNextOrderBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.processNextOrderBtn.Location = new System.Drawing.Point(778, 271);
            this.processNextOrderBtn.Name = "processNextOrderBtn";
            this.processNextOrderBtn.Size = new System.Drawing.Size(184, 38);
            this.processNextOrderBtn.TabIndex = 33;
            this.processNextOrderBtn.Text = "Import Orders";
            this.processNextOrderBtn.UseVisualStyleBackColor = false;
            this.processNextOrderBtn.Click += new System.EventHandler(this.processNextOrderBtn_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::JetDotComInterface.Properties.Resources.jetdotcom_logo;
            this.pictureBox1.Location = new System.Drawing.Point(396, 236);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(336, 142);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 26;
            this.pictureBox1.TabStop = false;
            // 
            // updateIntervalTrk
            // 
            this.updateIntervalTrk.Location = new System.Drawing.Point(878, 357);
            this.updateIntervalTrk.Maximum = 60;
            this.updateIntervalTrk.Minimum = 3;
            this.updateIntervalTrk.Name = "updateIntervalTrk";
            this.updateIntervalTrk.Size = new System.Drawing.Size(163, 45);
            this.updateIntervalTrk.TabIndex = 34;
            this.updateIntervalTrk.Value = 30;
            this.updateIntervalTrk.Scroll += new System.EventHandler(this.updateIntervalTrk_Scroll);
            // 
            // updateIntervalLbl
            // 
            this.updateIntervalLbl.AutoSize = true;
            this.updateIntervalLbl.ForeColor = System.Drawing.Color.White;
            this.updateIntervalLbl.Location = new System.Drawing.Point(896, 328);
            this.updateIntervalLbl.Name = "updateIntervalLbl";
            this.updateIntervalLbl.Size = new System.Drawing.Size(121, 13);
            this.updateIntervalLbl.TabIndex = 35;
            this.updateIntervalLbl.Text = "Update every x minutes.";
            // 
            // nextUpdateLbl
            // 
            this.nextUpdateLbl.AutoSize = true;
            this.nextUpdateLbl.ForeColor = System.Drawing.Color.White;
            this.nextUpdateLbl.Location = new System.Drawing.Point(893, 342);
            this.nextUpdateLbl.Name = "nextUpdateLbl";
            this.nextUpdateLbl.Size = new System.Drawing.Size(130, 13);
            this.nextUpdateLbl.TabIndex = 36;
            this.nextUpdateLbl.Text = "Next update in x seconds.";
            // 
            // acknowledgedOrderLVI
            // 
            this.acknowledgedOrderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.acknowledgedOrderLVI.ForeColor = System.Drawing.Color.White;
            this.acknowledgedOrderLVI.FullRowSelect = true;
            this.acknowledgedOrderLVI.GridLines = true;
            this.acknowledgedOrderLVI.Location = new System.Drawing.Point(293, 58);
            this.acknowledgedOrderLVI.MultiSelect = false;
            this.acknowledgedOrderLVI.Name = "acknowledgedOrderLVI";
            this.acknowledgedOrderLVI.Size = new System.Drawing.Size(149, 168);
            this.acknowledgedOrderLVI.TabIndex = 28;
            this.acknowledgedOrderLVI.UseCompatibleStateImageBehavior = false;
            this.acknowledgedOrderLVI.View = System.Windows.Forms.View.Details;
            // 
            // acknowledgedOrderLbl
            // 
            this.acknowledgedOrderLbl.AutoSize = true;
            this.acknowledgedOrderLbl.Location = new System.Drawing.Point(305, 28);
            this.acknowledgedOrderLbl.Name = "acknowledgedOrderLbl";
            this.acknowledgedOrderLbl.Size = new System.Drawing.Size(125, 13);
            this.acknowledgedOrderLbl.TabIndex = 29;
            this.acknowledgedOrderLbl.Text = "Acknowledged Order";
            // 
            // inProgressOrderLVI
            // 
            this.inProgressOrderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.inProgressOrderLVI.ForeColor = System.Drawing.Color.White;
            this.inProgressOrderLVI.FullRowSelect = true;
            this.inProgressOrderLVI.GridLines = true;
            this.inProgressOrderLVI.Location = new System.Drawing.Point(444, 58);
            this.inProgressOrderLVI.MultiSelect = false;
            this.inProgressOrderLVI.Name = "inProgressOrderLVI";
            this.inProgressOrderLVI.Size = new System.Drawing.Size(149, 168);
            this.inProgressOrderLVI.TabIndex = 30;
            this.inProgressOrderLVI.UseCompatibleStateImageBehavior = false;
            this.inProgressOrderLVI.View = System.Windows.Forms.View.Details;
            // 
            // inProgressOrderLbl
            // 
            this.inProgressOrderLbl.AutoSize = true;
            this.inProgressOrderLbl.Location = new System.Drawing.Point(463, 28);
            this.inProgressOrderLbl.Name = "inProgressOrderLbl";
            this.inProgressOrderLbl.Size = new System.Drawing.Size(112, 13);
            this.inProgressOrderLbl.TabIndex = 31;
            this.inProgressOrderLbl.Text = "In Progress Orders";
            // 
            // completedOrderLVI
            // 
            this.completedOrderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.completedOrderLVI.ForeColor = System.Drawing.Color.White;
            this.completedOrderLVI.FullRowSelect = true;
            this.completedOrderLVI.GridLines = true;
            this.completedOrderLVI.Location = new System.Drawing.Point(595, 58);
            this.completedOrderLVI.MultiSelect = false;
            this.completedOrderLVI.Name = "completedOrderLVI";
            this.completedOrderLVI.Size = new System.Drawing.Size(149, 168);
            this.completedOrderLVI.TabIndex = 32;
            this.completedOrderLVI.UseCompatibleStateImageBehavior = false;
            this.completedOrderLVI.View = System.Windows.Forms.View.Details;
            // 
            // completedOrderLbl
            // 
            this.completedOrderLbl.AutoSize = true;
            this.completedOrderLbl.Location = new System.Drawing.Point(626, 28);
            this.completedOrderLbl.Name = "completedOrderLbl";
            this.completedOrderLbl.Size = new System.Drawing.Size(107, 13);
            this.completedOrderLbl.TabIndex = 33;
            this.completedOrderLbl.Text = "Completed Orders";
            // 
            // getAckOrdersBtn
            // 
            this.getAckOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getAckOrdersBtn.Location = new System.Drawing.Point(778, 118);
            this.getAckOrdersBtn.Name = "getAckOrdersBtn";
            this.getAckOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getAckOrdersBtn.TabIndex = 38;
            this.getAckOrdersBtn.Text = "Get All Acknowleged Orders";
            this.getAckOrdersBtn.UseVisualStyleBackColor = false;
            this.getAckOrdersBtn.Click += new System.EventHandler(this.getAckOrdersBtn_Click);
            // 
            // getProgressOrdersBtn
            // 
            this.getProgressOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getProgressOrdersBtn.Location = new System.Drawing.Point(778, 155);
            this.getProgressOrdersBtn.Name = "getProgressOrdersBtn";
            this.getProgressOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getProgressOrdersBtn.TabIndex = 37;
            this.getProgressOrdersBtn.Text = "Get All In Progress Orders";
            this.getProgressOrdersBtn.UseVisualStyleBackColor = false;
            this.getProgressOrdersBtn.Click += new System.EventHandler(this.getProgressOrdersBtn_Click);
            // 
            // getCompletedOrdersBtn
            // 
            this.getCompletedOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getCompletedOrdersBtn.Location = new System.Drawing.Point(778, 193);
            this.getCompletedOrdersBtn.Name = "getCompletedOrdersBtn";
            this.getCompletedOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getCompletedOrdersBtn.TabIndex = 39;
            this.getCompletedOrdersBtn.Text = "Get All Completed Orders";
            this.getCompletedOrdersBtn.UseVisualStyleBackColor = false;
            this.getCompletedOrdersBtn.Click += new System.EventHandler(this.getCompletedOrdersBtn_Click);
            // 
            // getReturnOrdersBtn
            // 
            this.getReturnOrdersBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.getReturnOrdersBtn.Location = new System.Drawing.Point(778, 231);
            this.getReturnOrdersBtn.Name = "getReturnOrdersBtn";
            this.getReturnOrdersBtn.Size = new System.Drawing.Size(184, 38);
            this.getReturnOrdersBtn.TabIndex = 40;
            this.getReturnOrdersBtn.Text = "Get Return Requests";
            this.getReturnOrdersBtn.UseVisualStyleBackColor = false;
            // 
            // returnOrdersBox
            // 
            this.returnOrdersBox.BackColor = System.Drawing.Color.MediumSlateBlue;
            this.returnOrdersBox.Controls.Add(this.plusMinus6Lbl);
            this.returnOrdersBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.returnOrdersBox.ForeColor = System.Drawing.Color.White;
            this.returnOrdersBox.Location = new System.Drawing.Point(16, 205);
            this.returnOrdersBox.Name = "returnOrdersBox";
            this.returnOrdersBox.Size = new System.Drawing.Size(727, 20);
            this.returnOrdersBox.TabIndex = 41;
            this.returnOrdersBox.TabStop = false;
            this.returnOrdersBox.Text = "       Return Requests";
            // 
            // plusMinus6Lbl
            // 
            this.plusMinus6Lbl.BackColor = System.Drawing.Color.DarkSlateBlue;
            this.plusMinus6Lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.plusMinus6Lbl.Cursor = System.Windows.Forms.Cursors.Hand;
            this.plusMinus6Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.plusMinus6Lbl.ForeColor = System.Drawing.Color.White;
            this.plusMinus6Lbl.Location = new System.Drawing.Point(1, 0);
            this.plusMinus6Lbl.Name = "plusMinus6Lbl";
            this.plusMinus6Lbl.Size = new System.Drawing.Size(19, 18);
            this.plusMinus6Lbl.TabIndex = 12;
            this.plusMinus6Lbl.Text = "+";
            this.plusMinus6Lbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.plusMinus6Lbl.Click += new System.EventHandler(this.plusMinus6Lbl_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(903, 411);
            this.Controls.Add(this.returnOrdersBox);
            this.Controls.Add(this.getReturnOrdersBtn);
            this.Controls.Add(this.getCompletedOrdersBtn);
            this.Controls.Add(this.getAckOrdersBtn);
            this.Controls.Add(this.getProgressOrdersBtn);
            this.Controls.Add(this.orderInfoBox);
            this.Controls.Add(this.nextUpdateLbl);
            this.Controls.Add(this.updateIntervalLbl);
            this.Controls.Add(this.updateIntervalTrk);
            this.Controls.Add(this.processNextOrderBtn);
            this.Controls.Add(this.processedOrdersBox);
            this.Controls.Add(this.apiInfoBox);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.controlsBtn);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.gotoJetDashboardBtn);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.itemInfoBox);
            this.Controls.Add(this.statusBox);
            this.Controls.Add(this.getAllJetPendingOrdersBtn);
            this.Controls.Add(this.getAllJetItemsBtn);
            this.Controls.Add(this.currentTimeLbl);
            this.Controls.Add(this.titleLbl);
            this.Controls.Add(this.getAllJetAcceptedOrdersBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Jet.com SMC";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.apiInfoBox.ResumeLayout(false);
            this.apiInfoBox.PerformLayout();
            this.statusBox.ResumeLayout(false);
            this.itemInfoBox.ResumeLayout(false);
            this.itemInfoBox.PerformLayout();
            this.orderInfoBox.ResumeLayout(false);
            this.orderInfoBox.PerformLayout();
            this.processedOrdersBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.updateIntervalTrk)).EndInit();
            this.returnOrdersBox.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox apiUsernameTxt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox apiPasswordTxt;
        private System.Windows.Forms.Button getAllJetAcceptedOrdersBtn;
        private System.Windows.Forms.Label titleLbl;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox apiMerchantTxt;
        private System.Windows.Forms.GroupBox apiInfoBox;
        private System.Windows.Forms.Label plusMinus1Lbl;
        private System.Windows.Forms.Label currentTimeLbl;
        private System.Windows.Forms.Timer secondTmr;
        private System.Windows.Forms.TextBox apiBearerTxt;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListView statusLVI;
        private System.Windows.Forms.ListView itemsLVI;
        private System.Windows.Forms.Label itemsOnJetLbl;
        private System.Windows.Forms.Button getAllJetItemsBtn;
        private System.Windows.Forms.Label readyOrdersLbl;
        private System.Windows.Forms.ListView readyOrderLVI;
        private System.Windows.Forms.Label pendingOrderLbl;
        private System.Windows.Forms.ListView pendingOrderLVI;
        private System.Windows.Forms.Button getAllJetPendingOrdersBtn;
        private System.Windows.Forms.GroupBox statusBox;
        private System.Windows.Forms.Label plusMinus2Lbl;
        private System.Windows.Forms.GroupBox itemInfoBox;
        private System.Windows.Forms.Label plusMinus3Lbl;
        private System.Windows.Forms.GroupBox orderInfoBox;
        private System.Windows.Forms.Label plusMinus4Lbl;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button gotoJetDashboardBtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button controlsBtn;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox processedOrdersBox;
        private System.Windows.Forms.Label plusMinus5Lbl;
        private System.Windows.Forms.ListView processedOrdersLVI;
        private System.Windows.Forms.Button processNextOrderBtn;
        private System.Windows.Forms.TrackBar updateIntervalTrk;
        private System.Windows.Forms.Label updateIntervalLbl;
        private System.Windows.Forms.Label nextUpdateLbl;
        private System.Windows.Forms.ListView completedOrderLVI;
        private System.Windows.Forms.Label completedOrderLbl;
        private System.Windows.Forms.ListView inProgressOrderLVI;
        private System.Windows.Forms.Label inProgressOrderLbl;
        private System.Windows.Forms.ListView acknowledgedOrderLVI;
        private System.Windows.Forms.Label acknowledgedOrderLbl;
        private System.Windows.Forms.Button getAckOrdersBtn;
        private System.Windows.Forms.Button getProgressOrdersBtn;
        private System.Windows.Forms.Button getCompletedOrdersBtn;
        private System.Windows.Forms.Button getReturnOrdersBtn;
        private System.Windows.Forms.GroupBox returnOrdersBox;
        private System.Windows.Forms.Label plusMinus6Lbl;
    }
}

