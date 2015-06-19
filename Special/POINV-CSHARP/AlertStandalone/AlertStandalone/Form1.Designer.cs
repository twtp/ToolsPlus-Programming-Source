namespace AlertStandalone
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.totItemsWithAlertsLbl = new System.Windows.Forms.Label();
            this.totQtyAlertsLbl = new System.Windows.Forms.Label();
            this.totDateAlertsLbl = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.activeAlertsTotLbl = new System.Windows.Forms.Label();
            this.activeQtyAlertsLbl = new System.Windows.Forms.Label();
            this.activeDateAlertsLbl = new System.Windows.Forms.Label();
            this.qtyGroupBox = new System.Windows.Forms.GroupBox();
            this.dateGroupBox = new System.Windows.Forms.GroupBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.titleLbl = new System.Windows.Forms.Label();
            this.filterCombo = new System.Windows.Forms.ComboBox();
            this.itemNumberLbl = new System.Windows.Forms.Label();
            this.recordCountLbl = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.currentFilterLbl = new System.Windows.Forms.Label();
            this.aqohLbl = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.rightBell = new System.Windows.Forms.PictureBox();
            this.leftBell = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.dateAlertBell = new System.Windows.Forms.PictureBox();
            this.dateAlertSwitch = new System.Windows.Forms.PictureBox();
            this.dateAlertKnob = new System.Windows.Forms.PictureBox();
            this.qtyAlertBell = new System.Windows.Forms.PictureBox();
            this.aboveBelowSwitch = new System.Windows.Forms.PictureBox();
            this.aboveBelowKnob = new System.Windows.Forms.PictureBox();
            this.qtyAlertSwitch = new System.Windows.Forms.PictureBox();
            this.qtyAlertKnob = new System.Windows.Forms.PictureBox();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.qtyGroupBox.SuspendLayout();
            this.dateGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rightBell)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftBell)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertBell)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertSwitch)).BeginInit();
            this.dateAlertSwitch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertKnob)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertBell)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aboveBelowSwitch)).BeginInit();
            this.aboveBelowSwitch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.aboveBelowKnob)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertSwitch)).BeginInit();
            this.qtyAlertSwitch.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertKnob)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(242, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.settingsToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.settingsToolStripMenuItem.Text = "Settings";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // totItemsWithAlertsLbl
            // 
            this.totItemsWithAlertsLbl.AutoSize = true;
            this.totItemsWithAlertsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totItemsWithAlertsLbl.Location = new System.Drawing.Point(9, 77);
            this.totItemsWithAlertsLbl.Name = "totItemsWithAlertsLbl";
            this.totItemsWithAlertsLbl.Size = new System.Drawing.Size(170, 13);
            this.totItemsWithAlertsLbl.TabIndex = 1;
            this.totItemsWithAlertsLbl.Text = "Total Items With Alerts Configured:";
            // 
            // totQtyAlertsLbl
            // 
            this.totQtyAlertsLbl.AutoSize = true;
            this.totQtyAlertsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totQtyAlertsLbl.Location = new System.Drawing.Point(12, 90);
            this.totQtyAlertsLbl.Name = "totQtyAlertsLbl";
            this.totQtyAlertsLbl.Size = new System.Drawing.Size(159, 13);
            this.totQtyAlertsLbl.TabIndex = 2;
            this.totQtyAlertsLbl.Text = "Total Quantity Alerts Configured:";
            // 
            // totDateAlertsLbl
            // 
            this.totDateAlertsLbl.AutoSize = true;
            this.totDateAlertsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totDateAlertsLbl.Location = new System.Drawing.Point(12, 103);
            this.totDateAlertsLbl.Name = "totDateAlertsLbl";
            this.totDateAlertsLbl.Size = new System.Drawing.Size(143, 13);
            this.totDateAlertsLbl.TabIndex = 3;
            this.totDateAlertsLbl.Text = "Total Date Alerts Configured:";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.LightGreen;
            this.groupBox1.Controls.Add(this.activeAlertsTotLbl);
            this.groupBox1.Controls.Add(this.activeQtyAlertsLbl);
            this.groupBox1.Controls.Add(this.activeDateAlertsLbl);
            this.groupBox1.Controls.Add(this.totItemsWithAlertsLbl);
            this.groupBox1.Controls.Add(this.totDateAlertsLbl);
            this.groupBox1.Controls.Add(this.totQtyAlertsLbl);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 36);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(223, 70);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Alerts Info";
            // 
            // activeAlertsTotLbl
            // 
            this.activeAlertsTotLbl.AutoSize = true;
            this.activeAlertsTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.activeAlertsTotLbl.Location = new System.Drawing.Point(9, 24);
            this.activeAlertsTotLbl.Name = "activeAlertsTotLbl";
            this.activeAlertsTotLbl.Size = new System.Drawing.Size(96, 13);
            this.activeAlertsTotLbl.TabIndex = 6;
            this.activeAlertsTotLbl.Text = "Active Alerts Total:";
            // 
            // activeQtyAlertsLbl
            // 
            this.activeQtyAlertsLbl.AutoSize = true;
            this.activeQtyAlertsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.activeQtyAlertsLbl.Location = new System.Drawing.Point(16, 50);
            this.activeQtyAlertsLbl.Name = "activeQtyAlertsLbl";
            this.activeQtyAlertsLbl.Size = new System.Drawing.Size(114, 13);
            this.activeQtyAlertsLbl.TabIndex = 5;
            this.activeQtyAlertsLbl.Text = "Active Quantity Alerts: ";
            // 
            // activeDateAlertsLbl
            // 
            this.activeDateAlertsLbl.AutoSize = true;
            this.activeDateAlertsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.activeDateAlertsLbl.Location = new System.Drawing.Point(16, 37);
            this.activeDateAlertsLbl.Name = "activeDateAlertsLbl";
            this.activeDateAlertsLbl.Size = new System.Drawing.Size(98, 13);
            this.activeDateAlertsLbl.TabIndex = 4;
            this.activeDateAlertsLbl.Text = "Active Date Alerts: ";
            // 
            // qtyGroupBox
            // 
            this.qtyGroupBox.BackColor = System.Drawing.Color.PowderBlue;
            this.qtyGroupBox.Controls.Add(this.qtyAlertBell);
            this.qtyGroupBox.Controls.Add(this.aboveBelowSwitch);
            this.qtyGroupBox.Controls.Add(this.qtyAlertSwitch);
            this.qtyGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.qtyGroupBox.Location = new System.Drawing.Point(12, 246);
            this.qtyGroupBox.Name = "qtyGroupBox";
            this.qtyGroupBox.Size = new System.Drawing.Size(223, 223);
            this.qtyGroupBox.TabIndex = 5;
            this.qtyGroupBox.TabStop = false;
            this.qtyGroupBox.Text = "Qty Alert Settings";
            // 
            // dateGroupBox
            // 
            this.dateGroupBox.BackColor = System.Drawing.Color.PowderBlue;
            this.dateGroupBox.Controls.Add(this.dateAlertBell);
            this.dateGroupBox.Controls.Add(this.dateTimePicker1);
            this.dateGroupBox.Controls.Add(this.dateAlertSwitch);
            this.dateGroupBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateGroupBox.Location = new System.Drawing.Point(12, 475);
            this.dateGroupBox.Name = "dateGroupBox";
            this.dateGroupBox.Size = new System.Drawing.Size(223, 180);
            this.dateGroupBox.TabIndex = 6;
            this.dateGroupBox.TabStop = false;
            this.dateGroupBox.Text = "Date Alert Settings";
            this.dateGroupBox.Enter += new System.EventHandler(this.groupBox3_Enter);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dateTimePicker1.Location = new System.Drawing.Point(50, 85);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(131, 26);
            this.dateTimePicker1.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(57, 210);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(38, 26);
            this.button1.TabIndex = 8;
            this.button1.Text = "<";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(143, 210);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(38, 26);
            this.button2.TabIndex = 9;
            this.button2.Text = ">";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // titleLbl
            // 
            this.titleLbl.AutoSize = true;
            this.titleLbl.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.titleLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titleLbl.Location = new System.Drawing.Point(83, 0);
            this.titleLbl.Name = "titleLbl";
            this.titleLbl.Size = new System.Drawing.Size(71, 18);
            this.titleLbl.TabIndex = 14;
            this.titleLbl.Text = "ALERTS";
            // 
            // filterCombo
            // 
            this.filterCombo.FormattingEnabled = true;
            this.filterCombo.Items.AddRange(new object[] {
            "Show All Alerted Alerts",
            "Show Alerted Quantity Only",
            "Show Alerted Date Only",
            "Show All Quantity Alerts",
            "Show All Date Alerts"});
            this.filterCombo.Location = new System.Drawing.Point(481, 36);
            this.filterCombo.Name = "filterCombo";
            this.filterCombo.Size = new System.Drawing.Size(218, 21);
            this.filterCombo.TabIndex = 16;
            this.filterCombo.SelectedIndexChanged += new System.EventHandler(this.filterCombo_SelectedIndexChanged);
            // 
            // itemNumberLbl
            // 
            this.itemNumberLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.itemNumberLbl.Location = new System.Drawing.Point(49, 173);
            this.itemNumberLbl.Name = "itemNumberLbl";
            this.itemNumberLbl.Size = new System.Drawing.Size(142, 20);
            this.itemNumberLbl.TabIndex = 18;
            this.itemNumberLbl.Text = "Item";
            this.itemNumberLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // recordCountLbl
            // 
            this.recordCountLbl.Location = new System.Drawing.Point(96, 218);
            this.recordCountLbl.Name = "recordCountLbl";
            this.recordCountLbl.Size = new System.Drawing.Size(46, 13);
            this.recordCountLbl.TabIndex = 19;
            this.recordCountLbl.Text = "0/0";
            this.recordCountLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // button4
            // 
            this.button4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(143, 135);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(38, 26);
            this.button4.TabIndex = 21;
            this.button4.Text = ">";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(57, 135);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(38, 26);
            this.button5.TabIndex = 20;
            this.button5.Text = "<";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // currentFilterLbl
            // 
            this.currentFilterLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.currentFilterLbl.ForeColor = System.Drawing.Color.Red;
            this.currentFilterLbl.Location = new System.Drawing.Point(21, 116);
            this.currentFilterLbl.Name = "currentFilterLbl";
            this.currentFilterLbl.Size = new System.Drawing.Size(201, 20);
            this.currentFilterLbl.TabIndex = 23;
            this.currentFilterLbl.Text = "filter";
            this.currentFilterLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // aqohLbl
            // 
            this.aqohLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.aqohLbl.Location = new System.Drawing.Point(48, 193);
            this.aqohLbl.Name = "aqohLbl";
            this.aqohLbl.Size = new System.Drawing.Size(143, 15);
            this.aqohLbl.TabIndex = 24;
            this.aqohLbl.Text = "AQOH";
            this.aqohLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Image = global::AlertStandalone.Properties.Resources.checkmark;
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(20, 661);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(196, 42);
            this.button3.TabIndex = 15;
            this.button3.Text = "Apply Changes";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // rightBell
            // 
            this.rightBell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.rightBell.Image = global::AlertStandalone.Properties.Resources.alertbell;
            this.rightBell.Location = new System.Drawing.Point(151, 0);
            this.rightBell.Name = "rightBell";
            this.rightBell.Size = new System.Drawing.Size(22, 21);
            this.rightBell.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.rightBell.TabIndex = 12;
            this.rightBell.TabStop = false;
            // 
            // leftBell
            // 
            this.leftBell.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.leftBell.Image = global::AlertStandalone.Properties.Resources.alertbell2;
            this.leftBell.Location = new System.Drawing.Point(60, 0);
            this.leftBell.Name = "leftBell";
            this.leftBell.Size = new System.Drawing.Size(22, 21);
            this.leftBell.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.leftBell.TabIndex = 13;
            this.leftBell.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.pictureBox2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox2.Image = global::AlertStandalone.Properties.Resources.Exit_transparent;
            this.pictureBox2.Location = new System.Drawing.Point(212, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(30, 24);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 11;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // dateAlertBell
            // 
            this.dateAlertBell.BackColor = System.Drawing.Color.Transparent;
            this.dateAlertBell.Image = global::AlertStandalone.Properties.Resources.alertbell2;
            this.dateAlertBell.Location = new System.Drawing.Point(171, 13);
            this.dateAlertBell.Name = "dateAlertBell";
            this.dateAlertBell.Size = new System.Drawing.Size(45, 43);
            this.dateAlertBell.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.dateAlertBell.TabIndex = 15;
            this.dateAlertBell.TabStop = false;
            this.dateAlertBell.Visible = false;
            // 
            // dateAlertSwitch
            // 
            this.dateAlertSwitch.BackgroundImage = global::AlertStandalone.Properties.Resources.back1;
            this.dateAlertSwitch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.dateAlertSwitch.Controls.Add(this.dateAlertKnob);
            this.dateAlertSwitch.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dateAlertSwitch.Location = new System.Drawing.Point(65, 34);
            this.dateAlertSwitch.Name = "dateAlertSwitch";
            this.dateAlertSwitch.Size = new System.Drawing.Size(100, 33);
            this.dateAlertSwitch.TabIndex = 1;
            this.dateAlertSwitch.TabStop = false;
            this.dateAlertSwitch.Click += new System.EventHandler(this.dateAlertSwitch_Click);
            // 
            // dateAlertKnob
            // 
            this.dateAlertKnob.BackColor = System.Drawing.Color.Transparent;
            this.dateAlertKnob.BackgroundImage = global::AlertStandalone.Properties.Resources.knob;
            this.dateAlertKnob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.dateAlertKnob.Cursor = System.Windows.Forms.Cursors.Hand;
            this.dateAlertKnob.Location = new System.Drawing.Point(0, 0);
            this.dateAlertKnob.Name = "dateAlertKnob";
            this.dateAlertKnob.Size = new System.Drawing.Size(50, 33);
            this.dateAlertKnob.TabIndex = 2;
            this.dateAlertKnob.TabStop = false;
            this.dateAlertKnob.Click += new System.EventHandler(this.dateAlertKnob_Click);
            // 
            // qtyAlertBell
            // 
            this.qtyAlertBell.BackColor = System.Drawing.Color.Transparent;
            this.qtyAlertBell.Image = global::AlertStandalone.Properties.Resources.alertbell2;
            this.qtyAlertBell.Location = new System.Drawing.Point(172, 13);
            this.qtyAlertBell.Name = "qtyAlertBell";
            this.qtyAlertBell.Size = new System.Drawing.Size(45, 43);
            this.qtyAlertBell.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.qtyAlertBell.TabIndex = 14;
            this.qtyAlertBell.TabStop = false;
            this.qtyAlertBell.Visible = false;
            // 
            // aboveBelowSwitch
            // 
            this.aboveBelowSwitch.BackColor = System.Drawing.Color.Transparent;
            this.aboveBelowSwitch.BackgroundImage = global::AlertStandalone.Properties.Resources.aboveback2;
            this.aboveBelowSwitch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.aboveBelowSwitch.Controls.Add(this.aboveBelowKnob);
            this.aboveBelowSwitch.Cursor = System.Windows.Forms.Cursors.Hand;
            this.aboveBelowSwitch.Location = new System.Drawing.Point(55, 85);
            this.aboveBelowSwitch.Name = "aboveBelowSwitch";
            this.aboveBelowSwitch.Size = new System.Drawing.Size(100, 33);
            this.aboveBelowSwitch.TabIndex = 3;
            this.aboveBelowSwitch.TabStop = false;
            this.aboveBelowSwitch.Click += new System.EventHandler(this.aboveBelowSwitch_Click);
            // 
            // aboveBelowKnob
            // 
            this.aboveBelowKnob.BackColor = System.Drawing.Color.Transparent;
            this.aboveBelowKnob.BackgroundImage = global::AlertStandalone.Properties.Resources.knob;
            this.aboveBelowKnob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.aboveBelowKnob.Cursor = System.Windows.Forms.Cursors.Hand;
            this.aboveBelowKnob.Location = new System.Drawing.Point(0, 0);
            this.aboveBelowKnob.Name = "aboveBelowKnob";
            this.aboveBelowKnob.Size = new System.Drawing.Size(50, 33);
            this.aboveBelowKnob.TabIndex = 4;
            this.aboveBelowKnob.TabStop = false;
            this.aboveBelowKnob.Click += new System.EventHandler(this.aboveBelowKnob_Click);
            // 
            // qtyAlertSwitch
            // 
            this.qtyAlertSwitch.BackColor = System.Drawing.Color.Transparent;
            this.qtyAlertSwitch.BackgroundImage = global::AlertStandalone.Properties.Resources.aboveback2;
            this.qtyAlertSwitch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.qtyAlertSwitch.Controls.Add(this.qtyAlertKnob);
            this.qtyAlertSwitch.Cursor = System.Windows.Forms.Cursors.Hand;
            this.qtyAlertSwitch.Location = new System.Drawing.Point(55, 34);
            this.qtyAlertSwitch.Name = "qtyAlertSwitch";
            this.qtyAlertSwitch.Size = new System.Drawing.Size(100, 33);
            this.qtyAlertSwitch.TabIndex = 0;
            this.qtyAlertSwitch.TabStop = false;
            this.qtyAlertSwitch.Click += new System.EventHandler(this.qtyAlertSwitch_Click);
            // 
            // qtyAlertKnob
            // 
            this.qtyAlertKnob.BackColor = System.Drawing.Color.Transparent;
            this.qtyAlertKnob.BackgroundImage = global::AlertStandalone.Properties.Resources.knob_trans;
            this.qtyAlertKnob.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.qtyAlertKnob.Cursor = System.Windows.Forms.Cursors.Hand;
            this.qtyAlertKnob.Location = new System.Drawing.Point(0, 0);
            this.qtyAlertKnob.Name = "qtyAlertKnob";
            this.qtyAlertKnob.Size = new System.Drawing.Size(50, 33);
            this.qtyAlertKnob.TabIndex = 1;
            this.qtyAlertKnob.TabStop = false;
            this.qtyAlertKnob.Click += new System.EventHandler(this.qtyAlertKnob_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Silver;
            this.ClientSize = new System.Drawing.Size(242, 705);
            this.Controls.Add(this.aqohLbl);
            this.Controls.Add(this.currentFilterLbl);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.recordCountLbl);
            this.Controls.Add(this.itemNumberLbl);
            this.Controls.Add(this.filterCombo);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.rightBell);
            this.Controls.Add(this.titleLbl);
            this.Controls.Add(this.leftBell);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dateGroupBox);
            this.Controls.Add(this.qtyGroupBox);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Inventory Alerting System";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.qtyGroupBox.ResumeLayout(false);
            this.dateGroupBox.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rightBell)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftBell)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertBell)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertSwitch)).EndInit();
            this.dateAlertSwitch.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dateAlertKnob)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertBell)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aboveBelowSwitch)).EndInit();
            this.aboveBelowSwitch.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.aboveBelowKnob)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertSwitch)).EndInit();
            this.qtyAlertSwitch.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.qtyAlertKnob)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Label totItemsWithAlertsLbl;
        private System.Windows.Forms.Label totQtyAlertsLbl;
        private System.Windows.Forms.Label totDateAlertsLbl;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label activeAlertsTotLbl;
        private System.Windows.Forms.Label activeQtyAlertsLbl;
        private System.Windows.Forms.Label activeDateAlertsLbl;
        private System.Windows.Forms.GroupBox qtyGroupBox;
        private System.Windows.Forms.GroupBox dateGroupBox;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.PictureBox qtyAlertSwitch;
        private System.Windows.Forms.PictureBox qtyAlertKnob;
        private System.Windows.Forms.PictureBox dateAlertSwitch;
        private System.Windows.Forms.PictureBox dateAlertKnob;
        private System.Windows.Forms.PictureBox aboveBelowKnob;
        private System.Windows.Forms.PictureBox aboveBelowSwitch;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox rightBell;
        private System.Windows.Forms.PictureBox leftBell;
        private System.Windows.Forms.Label titleLbl;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ComboBox filterCombo;
        private System.Windows.Forms.Label itemNumberLbl;
        private System.Windows.Forms.Label recordCountLbl;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label currentFilterLbl;
        private System.Windows.Forms.PictureBox qtyAlertBell;
        private System.Windows.Forms.PictureBox dateAlertBell;
        private System.Windows.Forms.Label aqohLbl;
    }
}

