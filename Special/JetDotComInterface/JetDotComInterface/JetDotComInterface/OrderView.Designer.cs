namespace JetDotComInterface
{
    partial class OrderView
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
            this.orderNumberLbl = new System.Windows.Forms.Label();
            this.jetOrderNumberLbl = new System.Windows.Forms.Label();
            this.timeOfOrderLbl = new System.Windows.Forms.Label();
            this.orderTransmissionTimeLbl = new System.Windows.Forms.Label();
            this.orderInfoBox = new System.Windows.Forms.GroupBox();
            this.customerInfoBox = new System.Windows.Forms.GroupBox();
            this.customerNameLbl = new System.Windows.Forms.Label();
            this.customerPhoneLbl = new System.Windows.Forms.Label();
            this.recipientAddress1Lbl = new System.Windows.Forms.Label();
            this.recipientAddress2Lbl = new System.Windows.Forms.Label();
            this.recipientCityLbl = new System.Windows.Forms.Label();
            this.recipientStateLbl = new System.Windows.Forms.Label();
            this.recipientZipcodeLbl = new System.Windows.Forms.Label();
            this.shippingInfoBox = new System.Windows.Forms.GroupBox();
            this.recipientPhoneLbl = new System.Windows.Forms.Label();
            this.recipientNameLbl = new System.Windows.Forms.Label();
            this.orderLVI = new System.Windows.Forms.ListView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.differentItemsLbl = new System.Windows.Forms.Label();
            this.totQtyLbl = new System.Windows.Forms.Label();
            this.totCostLbl = new System.Windows.Forms.Label();
            this.totTaxLbl = new System.Windows.Forms.Label();
            this.totShippingLbl = new System.Windows.Forms.Label();
            this.totShippingTaxLbl = new System.Windows.Forms.Label();
            this.grandTotalLbl = new System.Windows.Forms.Label();
            this.orderFillableLbl = new System.Windows.Forms.Label();
            this.orderInfoBox.SuspendLayout();
            this.customerInfoBox.SuspendLayout();
            this.shippingInfoBox.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // orderNumberLbl
            // 
            this.orderNumberLbl.AutoSize = true;
            this.orderNumberLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.orderNumberLbl.Location = new System.Drawing.Point(17, 25);
            this.orderNumberLbl.Name = "orderNumberLbl";
            this.orderNumberLbl.Size = new System.Drawing.Size(76, 13);
            this.orderNumberLbl.TabIndex = 0;
            this.orderNumberLbl.Text = "Order Number:";
            // 
            // jetOrderNumberLbl
            // 
            this.jetOrderNumberLbl.AutoSize = true;
            this.jetOrderNumberLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.jetOrderNumberLbl.Location = new System.Drawing.Point(17, 42);
            this.jetOrderNumberLbl.Name = "jetOrderNumberLbl";
            this.jetOrderNumberLbl.Size = new System.Drawing.Size(93, 13);
            this.jetOrderNumberLbl.TabIndex = 1;
            this.jetOrderNumberLbl.Text = "Jet Order Number:";
            // 
            // timeOfOrderLbl
            // 
            this.timeOfOrderLbl.AutoSize = true;
            this.timeOfOrderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timeOfOrderLbl.Location = new System.Drawing.Point(17, 61);
            this.timeOfOrderLbl.Name = "timeOfOrderLbl";
            this.timeOfOrderLbl.Size = new System.Drawing.Size(76, 13);
            this.timeOfOrderLbl.TabIndex = 2;
            this.timeOfOrderLbl.Text = "Time Of Order:";
            // 
            // orderTransmissionTimeLbl
            // 
            this.orderTransmissionTimeLbl.AutoSize = true;
            this.orderTransmissionTimeLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.orderTransmissionTimeLbl.Location = new System.Drawing.Point(17, 77);
            this.orderTransmissionTimeLbl.Name = "orderTransmissionTimeLbl";
            this.orderTransmissionTimeLbl.Size = new System.Drawing.Size(126, 13);
            this.orderTransmissionTimeLbl.TabIndex = 3;
            this.orderTransmissionTimeLbl.Text = "Order Transmission Time:";
            // 
            // orderInfoBox
            // 
            this.orderInfoBox.BackColor = System.Drawing.Color.SlateBlue;
            this.orderInfoBox.Controls.Add(this.orderNumberLbl);
            this.orderInfoBox.Controls.Add(this.orderTransmissionTimeLbl);
            this.orderInfoBox.Controls.Add(this.jetOrderNumberLbl);
            this.orderInfoBox.Controls.Add(this.timeOfOrderLbl);
            this.orderInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.orderInfoBox.ForeColor = System.Drawing.Color.White;
            this.orderInfoBox.Location = new System.Drawing.Point(12, 12);
            this.orderInfoBox.Name = "orderInfoBox";
            this.orderInfoBox.Size = new System.Drawing.Size(322, 107);
            this.orderInfoBox.TabIndex = 4;
            this.orderInfoBox.TabStop = false;
            this.orderInfoBox.Text = "Order Info";
            // 
            // customerInfoBox
            // 
            this.customerInfoBox.BackColor = System.Drawing.Color.SlateBlue;
            this.customerInfoBox.Controls.Add(this.customerPhoneLbl);
            this.customerInfoBox.Controls.Add(this.customerNameLbl);
            this.customerInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.customerInfoBox.ForeColor = System.Drawing.Color.White;
            this.customerInfoBox.Location = new System.Drawing.Point(12, 125);
            this.customerInfoBox.Name = "customerInfoBox";
            this.customerInfoBox.Size = new System.Drawing.Size(322, 73);
            this.customerInfoBox.TabIndex = 4;
            this.customerInfoBox.TabStop = false;
            this.customerInfoBox.Text = "Customer Info";
            // 
            // customerNameLbl
            // 
            this.customerNameLbl.AutoSize = true;
            this.customerNameLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.customerNameLbl.Location = new System.Drawing.Point(17, 30);
            this.customerNameLbl.Name = "customerNameLbl";
            this.customerNameLbl.Size = new System.Drawing.Size(85, 13);
            this.customerNameLbl.TabIndex = 4;
            this.customerNameLbl.Text = "Customer Name:";
            // 
            // customerPhoneLbl
            // 
            this.customerPhoneLbl.AutoSize = true;
            this.customerPhoneLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.customerPhoneLbl.Location = new System.Drawing.Point(17, 44);
            this.customerPhoneLbl.Name = "customerPhoneLbl";
            this.customerPhoneLbl.Size = new System.Drawing.Size(88, 13);
            this.customerPhoneLbl.TabIndex = 5;
            this.customerPhoneLbl.Text = "Customer Phone:";
            this.customerPhoneLbl.Click += new System.EventHandler(this.customerPhoneLbl_Click);
            // 
            // recipientAddress1Lbl
            // 
            this.recipientAddress1Lbl.AutoSize = true;
            this.recipientAddress1Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientAddress1Lbl.Location = new System.Drawing.Point(16, 53);
            this.recipientAddress1Lbl.Name = "recipientAddress1Lbl";
            this.recipientAddress1Lbl.Size = new System.Drawing.Size(96, 13);
            this.recipientAddress1Lbl.TabIndex = 6;
            this.recipientAddress1Lbl.Text = "Recipient Address:";
            // 
            // recipientAddress2Lbl
            // 
            this.recipientAddress2Lbl.AutoSize = true;
            this.recipientAddress2Lbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientAddress2Lbl.Location = new System.Drawing.Point(111, 66);
            this.recipientAddress2Lbl.Name = "recipientAddress2Lbl";
            this.recipientAddress2Lbl.Size = new System.Drawing.Size(0, 13);
            this.recipientAddress2Lbl.TabIndex = 7;
            // 
            // recipientCityLbl
            // 
            this.recipientCityLbl.AutoSize = true;
            this.recipientCityLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientCityLbl.Location = new System.Drawing.Point(16, 79);
            this.recipientCityLbl.Name = "recipientCityLbl";
            this.recipientCityLbl.Size = new System.Drawing.Size(75, 13);
            this.recipientCityLbl.TabIndex = 8;
            this.recipientCityLbl.Text = "Recipient City:";
            // 
            // recipientStateLbl
            // 
            this.recipientStateLbl.AutoSize = true;
            this.recipientStateLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientStateLbl.Location = new System.Drawing.Point(16, 92);
            this.recipientStateLbl.Name = "recipientStateLbl";
            this.recipientStateLbl.Size = new System.Drawing.Size(83, 13);
            this.recipientStateLbl.TabIndex = 9;
            this.recipientStateLbl.Text = "Recipient State:";
            // 
            // recipientZipcodeLbl
            // 
            this.recipientZipcodeLbl.AutoSize = true;
            this.recipientZipcodeLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientZipcodeLbl.Location = new System.Drawing.Point(16, 106);
            this.recipientZipcodeLbl.Name = "recipientZipcodeLbl";
            this.recipientZipcodeLbl.Size = new System.Drawing.Size(97, 13);
            this.recipientZipcodeLbl.TabIndex = 10;
            this.recipientZipcodeLbl.Text = "Recipient Zipcode:";
            // 
            // shippingInfoBox
            // 
            this.shippingInfoBox.BackColor = System.Drawing.Color.SlateBlue;
            this.shippingInfoBox.Controls.Add(this.recipientPhoneLbl);
            this.shippingInfoBox.Controls.Add(this.recipientNameLbl);
            this.shippingInfoBox.Controls.Add(this.recipientZipcodeLbl);
            this.shippingInfoBox.Controls.Add(this.recipientAddress1Lbl);
            this.shippingInfoBox.Controls.Add(this.recipientStateLbl);
            this.shippingInfoBox.Controls.Add(this.recipientAddress2Lbl);
            this.shippingInfoBox.Controls.Add(this.recipientCityLbl);
            this.shippingInfoBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.shippingInfoBox.ForeColor = System.Drawing.Color.White;
            this.shippingInfoBox.Location = new System.Drawing.Point(12, 204);
            this.shippingInfoBox.Name = "shippingInfoBox";
            this.shippingInfoBox.Size = new System.Drawing.Size(322, 148);
            this.shippingInfoBox.TabIndex = 5;
            this.shippingInfoBox.TabStop = false;
            this.shippingInfoBox.Text = "Shipping Info";
            // 
            // recipientPhoneLbl
            // 
            this.recipientPhoneLbl.AutoSize = true;
            this.recipientPhoneLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientPhoneLbl.Location = new System.Drawing.Point(17, 40);
            this.recipientPhoneLbl.Name = "recipientPhoneLbl";
            this.recipientPhoneLbl.Size = new System.Drawing.Size(89, 13);
            this.recipientPhoneLbl.TabIndex = 12;
            this.recipientPhoneLbl.Text = "Recipient Phone:";
            // 
            // recipientNameLbl
            // 
            this.recipientNameLbl.AutoSize = true;
            this.recipientNameLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recipientNameLbl.Location = new System.Drawing.Point(17, 26);
            this.recipientNameLbl.Name = "recipientNameLbl";
            this.recipientNameLbl.Size = new System.Drawing.Size(86, 13);
            this.recipientNameLbl.TabIndex = 11;
            this.recipientNameLbl.Text = "Recipient Name:";
            // 
            // orderLVI
            // 
            this.orderLVI.BackColor = System.Drawing.Color.SlateBlue;
            this.orderLVI.ForeColor = System.Drawing.Color.White;
            this.orderLVI.FullRowSelect = true;
            this.orderLVI.GridLines = true;
            this.orderLVI.Location = new System.Drawing.Point(359, 37);
            this.orderLVI.Name = "orderLVI";
            this.orderLVI.Size = new System.Drawing.Size(484, 161);
            this.orderLVI.TabIndex = 6;
            this.orderLVI.UseCompatibleStateImageBehavior = false;
            this.orderLVI.View = System.Windows.Forms.View.Details;
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.Color.SlateBlue;
            this.groupBox1.Controls.Add(this.orderFillableLbl);
            this.groupBox1.Controls.Add(this.grandTotalLbl);
            this.groupBox1.Controls.Add(this.totShippingTaxLbl);
            this.groupBox1.Controls.Add(this.totShippingLbl);
            this.groupBox1.Controls.Add(this.totTaxLbl);
            this.groupBox1.Controls.Add(this.totCostLbl);
            this.groupBox1.Controls.Add(this.totQtyLbl);
            this.groupBox1.Controls.Add(this.differentItemsLbl);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.ForeColor = System.Drawing.Color.White;
            this.groupBox1.Location = new System.Drawing.Point(360, 206);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(483, 145);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Order Summary";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.SlateBlue;
            this.label1.Location = new System.Drawing.Point(572, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Order";
            // 
            // differentItemsLbl
            // 
            this.differentItemsLbl.AutoSize = true;
            this.differentItemsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.differentItemsLbl.Location = new System.Drawing.Point(23, 24);
            this.differentItemsLbl.Name = "differentItemsLbl";
            this.differentItemsLbl.Size = new System.Drawing.Size(78, 13);
            this.differentItemsLbl.TabIndex = 1;
            this.differentItemsLbl.Text = "Different Items:";
            // 
            // totQtyLbl
            // 
            this.totQtyLbl.AutoSize = true;
            this.totQtyLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totQtyLbl.Location = new System.Drawing.Point(23, 38);
            this.totQtyLbl.Name = "totQtyLbl";
            this.totQtyLbl.Size = new System.Drawing.Size(53, 13);
            this.totQtyLbl.TabIndex = 2;
            this.totQtyLbl.Text = "Total Qty:";
            // 
            // totCostLbl
            // 
            this.totCostLbl.AutoSize = true;
            this.totCostLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totCostLbl.Location = new System.Drawing.Point(23, 51);
            this.totCostLbl.Name = "totCostLbl";
            this.totCostLbl.Size = new System.Drawing.Size(58, 13);
            this.totCostLbl.TabIndex = 3;
            this.totCostLbl.Text = "Total Cost:";
            // 
            // totTaxLbl
            // 
            this.totTaxLbl.AutoSize = true;
            this.totTaxLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totTaxLbl.Location = new System.Drawing.Point(23, 64);
            this.totTaxLbl.Name = "totTaxLbl";
            this.totTaxLbl.Size = new System.Drawing.Size(55, 13);
            this.totTaxLbl.TabIndex = 4;
            this.totTaxLbl.Text = "Total Tax:";
            // 
            // totShippingLbl
            // 
            this.totShippingLbl.AutoSize = true;
            this.totShippingLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totShippingLbl.Location = new System.Drawing.Point(23, 77);
            this.totShippingLbl.Name = "totShippingLbl";
            this.totShippingLbl.Size = new System.Drawing.Size(78, 13);
            this.totShippingLbl.TabIndex = 5;
            this.totShippingLbl.Text = "Total Shipping:";
            // 
            // totShippingTaxLbl
            // 
            this.totShippingTaxLbl.AutoSize = true;
            this.totShippingTaxLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.totShippingTaxLbl.Location = new System.Drawing.Point(23, 90);
            this.totShippingTaxLbl.Name = "totShippingTaxLbl";
            this.totShippingTaxLbl.Size = new System.Drawing.Size(99, 13);
            this.totShippingTaxLbl.TabIndex = 6;
            this.totShippingTaxLbl.Text = "Total Shipping Tax:";
            // 
            // grandTotalLbl
            // 
            this.grandTotalLbl.AutoSize = true;
            this.grandTotalLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grandTotalLbl.Location = new System.Drawing.Point(23, 114);
            this.grandTotalLbl.Name = "grandTotalLbl";
            this.grandTotalLbl.Size = new System.Drawing.Size(97, 20);
            this.grandTotalLbl.TabIndex = 7;
            this.grandTotalLbl.Text = "Grand Total:";
            // 
            // orderFillableLbl
            // 
            this.orderFillableLbl.AutoSize = true;
            this.orderFillableLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.orderFillableLbl.Location = new System.Drawing.Point(322, 16);
            this.orderFillableLbl.Name = "orderFillableLbl";
            this.orderFillableLbl.Size = new System.Drawing.Size(146, 13);
            this.orderFillableLbl.TabIndex = 8;
            this.orderFillableLbl.Text = "This order is unknown fillable.";
            // 
            // OrderView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(861, 364);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.orderLVI);
            this.Controls.Add(this.shippingInfoBox);
            this.Controls.Add(this.customerInfoBox);
            this.Controls.Add(this.orderInfoBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "OrderView";
            this.Text = "Jet.com Order Viewer";
            this.Load += new System.EventHandler(this.OrderView_Load);
            this.orderInfoBox.ResumeLayout(false);
            this.orderInfoBox.PerformLayout();
            this.customerInfoBox.ResumeLayout(false);
            this.customerInfoBox.PerformLayout();
            this.shippingInfoBox.ResumeLayout(false);
            this.shippingInfoBox.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label orderNumberLbl;
        private System.Windows.Forms.Label jetOrderNumberLbl;
        private System.Windows.Forms.Label timeOfOrderLbl;
        private System.Windows.Forms.Label orderTransmissionTimeLbl;
        private System.Windows.Forms.GroupBox orderInfoBox;
        private System.Windows.Forms.GroupBox customerInfoBox;
        private System.Windows.Forms.Label customerPhoneLbl;
        private System.Windows.Forms.Label customerNameLbl;
        private System.Windows.Forms.Label recipientZipcodeLbl;
        private System.Windows.Forms.Label recipientStateLbl;
        private System.Windows.Forms.Label recipientCityLbl;
        private System.Windows.Forms.Label recipientAddress2Lbl;
        private System.Windows.Forms.Label recipientAddress1Lbl;
        private System.Windows.Forms.GroupBox shippingInfoBox;
        private System.Windows.Forms.Label recipientPhoneLbl;
        private System.Windows.Forms.Label recipientNameLbl;
        private System.Windows.Forms.ListView orderLVI;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label totCostLbl;
        private System.Windows.Forms.Label totQtyLbl;
        private System.Windows.Forms.Label differentItemsLbl;
        private System.Windows.Forms.Label orderFillableLbl;
        private System.Windows.Forms.Label grandTotalLbl;
        private System.Windows.Forms.Label totShippingTaxLbl;
        private System.Windows.Forms.Label totShippingLbl;
        private System.Windows.Forms.Label totTaxLbl;
    }
}