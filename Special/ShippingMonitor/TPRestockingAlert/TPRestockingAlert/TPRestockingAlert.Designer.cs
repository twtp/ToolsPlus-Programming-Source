namespace TPRestockingAlert
{
    partial class TPRestockingAlert
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.RestockingHeaderLbl = new System.Windows.Forms.Label();
            this.TotDifferentComponentsLbl = new System.Windows.Forms.Label();
            this.TotalQtyLbl = new System.Windows.Forms.Label();
            this.MondayTotLbl = new System.Windows.Forms.Label();
            this.TuesdayTotLbl = new System.Windows.Forms.Label();
            this.WednesdayTotLbl = new System.Windows.Forms.Label();
            this.ThursdayTotLbl = new System.Windows.Forms.Label();
            this.FridayTotLbl = new System.Windows.Forms.Label();
            this.SaturdayTotLbl = new System.Windows.Forms.Label();
            this.notPutAwayLbl = new System.Windows.Forms.Label();
            this.updateTmr = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // RestockingHeaderLbl
            // 
            this.RestockingHeaderLbl.BackColor = System.Drawing.Color.Transparent;
            this.RestockingHeaderLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.RestockingHeaderLbl.ForeColor = System.Drawing.Color.White;
            this.RestockingHeaderLbl.Location = new System.Drawing.Point(3, 0);
            this.RestockingHeaderLbl.Name = "RestockingHeaderLbl";
            this.RestockingHeaderLbl.Size = new System.Drawing.Size(607, 62);
            this.RestockingHeaderLbl.TabIndex = 0;
            this.RestockingHeaderLbl.Text = "Restocking";
            this.RestockingHeaderLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // TotDifferentComponentsLbl
            // 
            this.TotDifferentComponentsLbl.BackColor = System.Drawing.Color.Transparent;
            this.TotDifferentComponentsLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TotDifferentComponentsLbl.ForeColor = System.Drawing.Color.White;
            this.TotDifferentComponentsLbl.Location = new System.Drawing.Point(2, 65);
            this.TotDifferentComponentsLbl.Name = "TotDifferentComponentsLbl";
            this.TotDifferentComponentsLbl.Size = new System.Drawing.Size(304, 39);
            this.TotDifferentComponentsLbl.TabIndex = 4;
            this.TotDifferentComponentsLbl.Text = "Components:";
            // 
            // TotalQtyLbl
            // 
            this.TotalQtyLbl.AutoSize = true;
            this.TotalQtyLbl.BackColor = System.Drawing.Color.Transparent;
            this.TotalQtyLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TotalQtyLbl.ForeColor = System.Drawing.Color.White;
            this.TotalQtyLbl.Location = new System.Drawing.Point(298, 65);
            this.TotalQtyLbl.Name = "TotalQtyLbl";
            this.TotalQtyLbl.Size = new System.Drawing.Size(142, 33);
            this.TotalQtyLbl.TabIndex = 3;
            this.TotalQtyLbl.Text = "Total Qty:";
            // 
            // MondayTotLbl
            // 
            this.MondayTotLbl.AutoSize = true;
            this.MondayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.MondayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MondayTotLbl.ForeColor = System.Drawing.Color.White;
            this.MondayTotLbl.Location = new System.Drawing.Point(16, 122);
            this.MondayTotLbl.Name = "MondayTotLbl";
            this.MondayTotLbl.Size = new System.Drawing.Size(74, 29);
            this.MondayTotLbl.TabIndex = 5;
            this.MondayTotLbl.Text = "Older";
            this.MondayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // TuesdayTotLbl
            // 
            this.TuesdayTotLbl.AutoSize = true;
            this.TuesdayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.TuesdayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TuesdayTotLbl.ForeColor = System.Drawing.Color.White;
            this.TuesdayTotLbl.Location = new System.Drawing.Point(110, 122);
            this.TuesdayTotLbl.Name = "TuesdayTotLbl";
            this.TuesdayTotLbl.Size = new System.Drawing.Size(85, 29);
            this.TuesdayTotLbl.TabIndex = 6;
            this.TuesdayTotLbl.Text = "4 Days";
            this.TuesdayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // WednesdayTotLbl
            // 
            this.WednesdayTotLbl.AutoSize = true;
            this.WednesdayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.WednesdayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.WednesdayTotLbl.ForeColor = System.Drawing.Color.White;
            this.WednesdayTotLbl.Location = new System.Drawing.Point(221, 122);
            this.WednesdayTotLbl.Name = "WednesdayTotLbl";
            this.WednesdayTotLbl.Size = new System.Drawing.Size(85, 29);
            this.WednesdayTotLbl.TabIndex = 7;
            this.WednesdayTotLbl.Text = "3 Days";
            this.WednesdayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // ThursdayTotLbl
            // 
            this.ThursdayTotLbl.AutoSize = true;
            this.ThursdayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.ThursdayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ThursdayTotLbl.ForeColor = System.Drawing.Color.White;
            this.ThursdayTotLbl.Location = new System.Drawing.Point(328, 122);
            this.ThursdayTotLbl.Name = "ThursdayTotLbl";
            this.ThursdayTotLbl.Size = new System.Drawing.Size(85, 29);
            this.ThursdayTotLbl.TabIndex = 8;
            this.ThursdayTotLbl.Text = "2 Days";
            this.ThursdayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // FridayTotLbl
            // 
            this.FridayTotLbl.AutoSize = true;
            this.FridayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.FridayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FridayTotLbl.ForeColor = System.Drawing.Color.White;
            this.FridayTotLbl.Location = new System.Drawing.Point(440, 122);
            this.FridayTotLbl.Name = "FridayTotLbl";
            this.FridayTotLbl.Size = new System.Drawing.Size(73, 29);
            this.FridayTotLbl.TabIndex = 9;
            this.FridayTotLbl.Text = "1 Day";
            this.FridayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // SaturdayTotLbl
            // 
            this.SaturdayTotLbl.AutoSize = true;
            this.SaturdayTotLbl.BackColor = System.Drawing.Color.Transparent;
            this.SaturdayTotLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SaturdayTotLbl.ForeColor = System.Drawing.Color.White;
            this.SaturdayTotLbl.Location = new System.Drawing.Point(532, 122);
            this.SaturdayTotLbl.Name = "SaturdayTotLbl";
            this.SaturdayTotLbl.Size = new System.Drawing.Size(81, 29);
            this.SaturdayTotLbl.TabIndex = 10;
            this.SaturdayTotLbl.Text = "Today";
            this.SaturdayTotLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // notPutAwayLbl
            // 
            this.notPutAwayLbl.BackColor = System.Drawing.Color.Transparent;
            this.notPutAwayLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.notPutAwayLbl.ForeColor = System.Drawing.Color.White;
            this.notPutAwayLbl.Location = new System.Drawing.Point(0, 72);
            this.notPutAwayLbl.Name = "notPutAwayLbl";
            this.notPutAwayLbl.Size = new System.Drawing.Size(610, 29);
            this.notPutAwayLbl.TabIndex = 11;
            this.notPutAwayLbl.Text = "Not Put Away";
            this.notPutAwayLbl.Visible = false;
            // 
            // updateTmr
            // 
            this.updateTmr.Enabled = true;
            this.updateTmr.Interval = 57454;
            this.updateTmr.Tick += new System.EventHandler(this.updateTmr_Tick);
            // 
            // TPRestockingAlert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Green;
            this.Controls.Add(this.SaturdayTotLbl);
            this.Controls.Add(this.FridayTotLbl);
            this.Controls.Add(this.ThursdayTotLbl);
            this.Controls.Add(this.WednesdayTotLbl);
            this.Controls.Add(this.TuesdayTotLbl);
            this.Controls.Add(this.MondayTotLbl);
            this.Controls.Add(this.TotDifferentComponentsLbl);
            this.Controls.Add(this.TotalQtyLbl);
            this.Controls.Add(this.RestockingHeaderLbl);
            this.Controls.Add(this.notPutAwayLbl);
            this.Name = "TPRestockingAlert";
            this.Size = new System.Drawing.Size(613, 179);
            this.Load += new System.EventHandler(this.TPRestockingAlert_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label RestockingHeaderLbl;
        private System.Windows.Forms.Label TotDifferentComponentsLbl;
        private System.Windows.Forms.Label TotalQtyLbl;
        private System.Windows.Forms.Label MondayTotLbl;
        private System.Windows.Forms.Label TuesdayTotLbl;
        private System.Windows.Forms.Label WednesdayTotLbl;
        private System.Windows.Forms.Label ThursdayTotLbl;
        private System.Windows.Forms.Label FridayTotLbl;
        private System.Windows.Forms.Label SaturdayTotLbl;
        private System.Windows.Forms.Label notPutAwayLbl;
        private System.Windows.Forms.Timer updateTmr;
    }
}
