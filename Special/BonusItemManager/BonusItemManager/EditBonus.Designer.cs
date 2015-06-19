namespace BonusItemManager
{
    partial class EditBonus
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
            this.addBtn = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.priceTxt = new System.Windows.Forms.TextBox();
            this.priceRadio = new System.Windows.Forms.RadioButton();
            this.qtyRadio = new System.Windows.Forms.RadioButton();
            this.qtyTxt = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.itemsRadio = new System.Windows.Forms.RadioButton();
            this.itemsTxt = new System.Windows.Forms.TextBox();
            this.linesRadio = new System.Windows.Forms.RadioButton();
            this.linesTxt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cartTxt = new System.Windows.Forms.TextBox();
            this.bonusTxt = new System.Windows.Forms.TextBox();
            this.activeChk = new System.Windows.Forms.CheckBox();
            this.commentedChk = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.bonusItemNumberTxt = new System.Windows.Forms.TextBox();
            this.fixedQtyAmountTxt = new System.Windows.Forms.TextBox();
            this.buyOneGetOneRadio = new System.Windows.Forms.RadioButton();
            this.fixedQtyRadio = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // addBtn
            // 
            this.addBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.addBtn.Location = new System.Drawing.Point(305, 555);
            this.addBtn.Name = "addBtn";
            this.addBtn.Size = new System.Drawing.Size(86, 35);
            this.addBtn.TabIndex = 3;
            this.addBtn.Text = "Update";
            this.addBtn.UseVisualStyleBackColor = false;
            this.addBtn.Click += new System.EventHandler(this.addBtn_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox4);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Controls.Add(this.button5);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cartTxt);
            this.groupBox1.Controls.Add(this.bonusTxt);
            this.groupBox1.Controls.Add(this.activeChk);
            this.groupBox1.Controls.Add(this.commentedChk);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(385, 531);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Edit Bonus";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(253, 60);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(31, 22);
            this.button4.TabIndex = 20;
            this.button4.Text = "$";
            this.button4.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(216, 60);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(31, 22);
            this.button5.TabIndex = 19;
            this.button5.Text = "Qty";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(253, 11);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(31, 22);
            this.button3.TabIndex = 18;
            this.button3.Text = "$";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(216, 11);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(31, 22);
            this.button2.TabIndex = 17;
            this.button2.Text = "Qty";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.priceTxt);
            this.groupBox3.Controls.Add(this.priceRadio);
            this.groupBox3.Controls.Add(this.qtyRadio);
            this.groupBox3.Controls.Add(this.qtyTxt);
            this.groupBox3.Location = new System.Drawing.Point(10, 351);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(369, 75);
            this.groupBox3.TabIndex = 1;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Requirements";
            // 
            // priceTxt
            // 
            this.priceTxt.Location = new System.Drawing.Point(166, 19);
            this.priceTxt.Name = "priceTxt";
            this.priceTxt.Size = new System.Drawing.Size(84, 20);
            this.priceTxt.TabIndex = 10;
            // 
            // priceRadio
            // 
            this.priceRadio.AutoSize = true;
            this.priceRadio.Checked = true;
            this.priceRadio.Location = new System.Drawing.Point(18, 19);
            this.priceRadio.Name = "priceRadio";
            this.priceRadio.Size = new System.Drawing.Size(142, 17);
            this.priceRadio.TabIndex = 8;
            this.priceRadio.TabStop = true;
            this.priceRadio.Text = "Requirement: Total Price";
            this.priceRadio.UseVisualStyleBackColor = true;
            // 
            // qtyRadio
            // 
            this.qtyRadio.AutoSize = true;
            this.qtyRadio.Location = new System.Drawing.Point(18, 42);
            this.qtyRadio.Name = "qtyRadio";
            this.qtyRadio.Size = new System.Drawing.Size(134, 17);
            this.qtyRadio.TabIndex = 9;
            this.qtyRadio.Text = "Requirement: Total Qty";
            this.qtyRadio.UseVisualStyleBackColor = true;
            // 
            // qtyTxt
            // 
            this.qtyTxt.Enabled = false;
            this.qtyTxt.Location = new System.Drawing.Point(166, 42);
            this.qtyTxt.Name = "qtyTxt";
            this.qtyTxt.Size = new System.Drawing.Size(84, 20);
            this.qtyTxt.TabIndex = 11;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.itemsRadio);
            this.groupBox2.Controls.Add(this.itemsTxt);
            this.groupBox2.Controls.Add(this.linesRadio);
            this.groupBox2.Controls.Add(this.linesTxt);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(10, 168);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(369, 177);
            this.groupBox2.TabIndex = 16;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Applications";
            // 
            // itemsRadio
            // 
            this.itemsRadio.AutoSize = true;
            this.itemsRadio.Checked = true;
            this.itemsRadio.Location = new System.Drawing.Point(18, 20);
            this.itemsRadio.Name = "itemsRadio";
            this.itemsRadio.Size = new System.Drawing.Size(109, 17);
            this.itemsRadio.TabIndex = 3;
            this.itemsRadio.TabStop = true;
            this.itemsRadio.Text = "Applies To Item(s)";
            this.itemsRadio.UseVisualStyleBackColor = true;
            this.itemsRadio.CheckedChanged += new System.EventHandler(this.itemsRadio_CheckedChanged);
            // 
            // itemsTxt
            // 
            this.itemsTxt.Location = new System.Drawing.Point(48, 43);
            this.itemsTxt.Multiline = true;
            this.itemsTxt.Name = "itemsTxt";
            this.itemsTxt.Size = new System.Drawing.Size(301, 46);
            this.itemsTxt.TabIndex = 2;
            // 
            // linesRadio
            // 
            this.linesRadio.AutoSize = true;
            this.linesRadio.Location = new System.Drawing.Point(18, 95);
            this.linesRadio.Name = "linesRadio";
            this.linesRadio.Size = new System.Drawing.Size(109, 17);
            this.linesRadio.TabIndex = 4;
            this.linesRadio.Text = "Applies To Line(s)";
            this.linesRadio.UseVisualStyleBackColor = true;
            this.linesRadio.CheckedChanged += new System.EventHandler(this.linesRadio_CheckedChanged);
            // 
            // linesTxt
            // 
            this.linesTxt.Enabled = false;
            this.linesTxt.Location = new System.Drawing.Point(48, 118);
            this.linesTxt.Multiline = true;
            this.linesTxt.Name = "linesTxt";
            this.linesTxt.Size = new System.Drawing.Size(301, 46);
            this.linesTxt.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(252, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "(comma separated)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(252, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "(comma separated)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(24, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(163, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Cart Text (can contain javascript)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(173, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Bonus Title (can contain javascript)";
            // 
            // cartTxt
            // 
            this.cartTxt.Location = new System.Drawing.Point(27, 83);
            this.cartTxt.Name = "cartTxt";
            this.cartTxt.Size = new System.Drawing.Size(331, 20);
            this.cartTxt.TabIndex = 13;
            // 
            // bonusTxt
            // 
            this.bonusTxt.Location = new System.Drawing.Point(27, 34);
            this.bonusTxt.Name = "bonusTxt";
            this.bonusTxt.Size = new System.Drawing.Size(331, 20);
            this.bonusTxt.TabIndex = 12;
            // 
            // activeChk
            // 
            this.activeChk.AutoSize = true;
            this.activeChk.Location = new System.Drawing.Point(27, 141);
            this.activeChk.Name = "activeChk";
            this.activeChk.Size = new System.Drawing.Size(56, 17);
            this.activeChk.TabIndex = 1;
            this.activeChk.Text = "Active";
            this.activeChk.UseVisualStyleBackColor = true;
            // 
            // commentedChk
            // 
            this.commentedChk.AutoSize = true;
            this.commentedChk.Location = new System.Drawing.Point(27, 118);
            this.commentedChk.Name = "commentedChk";
            this.commentedChk.Size = new System.Drawing.Size(102, 17);
            this.commentedChk.TabIndex = 0;
            this.commentedChk.Text = "Commented Out";
            this.commentedChk.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.bonusItemNumberTxt);
            this.groupBox4.Controls.Add(this.fixedQtyAmountTxt);
            this.groupBox4.Controls.Add(this.buyOneGetOneRadio);
            this.groupBox4.Controls.Add(this.fixedQtyRadio);
            this.groupBox4.Location = new System.Drawing.Point(10, 432);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(369, 92);
            this.groupBox4.TabIndex = 22;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Bonus Details";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 65);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Bonus Item Number";
            // 
            // bonusItemNumberTxt
            // 
            this.bonusItemNumberTxt.Location = new System.Drawing.Point(115, 62);
            this.bonusItemNumberTxt.Name = "bonusItemNumberTxt";
            this.bonusItemNumberTxt.Size = new System.Drawing.Size(182, 20);
            this.bonusItemNumberTxt.TabIndex = 3;
            // 
            // fixedQtyAmountTxt
            // 
            this.fixedQtyAmountTxt.Location = new System.Drawing.Point(115, 18);
            this.fixedQtyAmountTxt.Name = "fixedQtyAmountTxt";
            this.fixedQtyAmountTxt.Size = new System.Drawing.Size(45, 20);
            this.fixedQtyAmountTxt.TabIndex = 2;
            // 
            // buyOneGetOneRadio
            // 
            this.buyOneGetOneRadio.AutoSize = true;
            this.buyOneGetOneRadio.Location = new System.Drawing.Point(17, 42);
            this.buyOneGetOneRadio.Name = "buyOneGetOneRadio";
            this.buyOneGetOneRadio.Size = new System.Drawing.Size(81, 17);
            this.buyOneGetOneRadio.TabIndex = 1;
            this.buyOneGetOneRadio.Text = "Buy 1 Get 1";
            this.buyOneGetOneRadio.UseVisualStyleBackColor = true;
            // 
            // fixedQtyRadio
            // 
            this.fixedQtyRadio.AutoSize = true;
            this.fixedQtyRadio.Checked = true;
            this.fixedQtyRadio.Location = new System.Drawing.Point(17, 19);
            this.fixedQtyRadio.Name = "fixedQtyRadio";
            this.fixedQtyRadio.Size = new System.Drawing.Size(92, 17);
            this.fixedQtyRadio.TabIndex = 0;
            this.fixedQtyRadio.TabStop = true;
            this.fixedQtyRadio.Text = "Fixed Quantity";
            this.fixedQtyRadio.UseVisualStyleBackColor = true;
            // 
            // EditBonus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(408, 602);
            this.Controls.Add(this.addBtn);
            this.Controls.Add(this.groupBox1);
            this.Name = "EditBonus";
            this.Text = "Edit Bonus";
            this.Load += new System.EventHandler(this.EditBonus_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button addBtn;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox priceTxt;
        private System.Windows.Forms.RadioButton priceRadio;
        private System.Windows.Forms.RadioButton qtyRadio;
        private System.Windows.Forms.TextBox qtyTxt;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton itemsRadio;
        private System.Windows.Forms.TextBox itemsTxt;
        private System.Windows.Forms.RadioButton linesRadio;
        private System.Windows.Forms.TextBox linesTxt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox cartTxt;
        private System.Windows.Forms.TextBox bonusTxt;
        private System.Windows.Forms.CheckBox activeChk;
        private System.Windows.Forms.CheckBox commentedChk;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox bonusItemNumberTxt;
        private System.Windows.Forms.TextBox fixedQtyAmountTxt;
        private System.Windows.Forms.RadioButton buyOneGetOneRadio;
        private System.Windows.Forms.RadioButton fixedQtyRadio;
    }
}