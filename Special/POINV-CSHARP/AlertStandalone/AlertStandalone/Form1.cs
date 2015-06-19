using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace AlertStandalone
{
    public partial class Form1 : Form
    {

        public struct ItemAlerts
        {
            public string ItemNumber;
            public List<AlertType> ActiveAlerts;
                                    
        }
        public struct AlertType
        {
            public string TriggerType;
            public decimal QtyThreshold;
            public decimal ActualQtyOnHand;
            public bool isBelow;
            public string DateThreshold;
            public string Notes;
        }
        public List<string> CurrentFilterList = new List<string>();
        public ItemAlerts CurrentItem = new ItemAlerts();
        public int CurrentIndex = 0;
       
        public List<string> ItemNumbersWithAlertsConfigured = new List<string>();
        public List<string> ItemNumbersWithActiveQtyAlerts = new List<string>();
        public List<string> ItemNumbersWithActiveDateAlerts = new List<string>();
        public int TotalQtyAlerts = 0;
        public int TotalDateAlerts = 0;


        public bool qtyKnobOn = false;
        public bool dateKnobOn = false;
        public bool aboveBelowKnobOn = false;

        public SQLCalls dbQuery = new SQLCalls();

        public settingsFrm newSettings = new settingsFrm();

        Watermarking.WaterMarkTextBox qtyBox = new Watermarking.WaterMarkTextBox();
        Watermarking.WaterMarkTextBox qtyNotesBox = new Watermarking.WaterMarkTextBox();
        Watermarking.WaterMarkTextBox dateNotesBox = new Watermarking.WaterMarkTextBox();


        public List<string> FilterList = new List<string>();
        public int FilterIndex = 0;


        public Form1()
        {


            InitializeComponent();
           
        }

        private void EnableApplyButton()
        {
            button3.Enabled = true;
        }
        private void DisableApplyButton()
        {
            button3.Enabled = false;
        }

        private void CheckToSave(object sender,EventArgs e)
        {
            if (button3.Enabled == true)
            {
                DialogResult res = MessageBox.Show("Do you wish to save your changes to the alert/s for Item " + CurrentItem.ItemNumber + "?", "Save Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
                {
                    //save alert changes...
                    button3_Click(sender,e);
                }
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            FilterList.Add("Show All Alerted Alerts");
            FilterList.Add("Show Alerted Quantity Only");
            FilterList.Add("Show Alerted Date Only");
            FilterList.Add("Show All Quantity Alerts");
            FilterList.Add("Show All Date Alerts");
            currentFilterLbl.Text = FilterList[FilterIndex];


//Show All Alerted Alerts
//Show Alerted Quantity Only
//Show Alerted Date Only
//Show All Quantity Alerts
//Show All Date Alerts


            //filterCombo.SelectedIndex= 0;
            leftBell.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            leftBell.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            rightBell.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            rightBell.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            titleLbl.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            titleLbl.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            menuStrip1.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            menuStrip1.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            
            
            qtyBox.Parent = this.qtyGroupBox;
            this.qtyGroupBox.Controls.Add(qtyBox);
            qtyBox.Width = 60;
            qtyBox.Left = aboveBelowSwitch.Left + (aboveBelowSwitch.Width / 2) - (qtyBox.Width / 2) + 5;
            qtyBox.Top = aboveBelowSwitch.Bottom + 10;
            qtyBox.TextAlign = HorizontalAlignment.Center;
            qtyBox.WaterMarkText = "  Qty";
            qtyBox.Visible = true;
            qtyBox.Enabled = true;
            
            qtyBox.BringToFront();
            //qtyBox.TextChanged += new EventHandler(qtyBox_TextChanged);
            qtyBox.KeyDown += new KeyEventHandler(qtyBox_TextChanged);
            

            
            qtyNotesBox.Multiline = true;
            qtyNotesBox.Parent = this.qtyGroupBox;
            qtyNotesBox.Left = 10;
            qtyNotesBox.Width = this.qtyGroupBox.Width - 20;
            qtyNotesBox.Top = qtyBox.Bottom + 10;
            qtyNotesBox.Height = 50;
            qtyNotesBox.WaterMarkText = "Enter Notes Here...";
            qtyNotesBox.Visible = true;
            qtyNotesBox.Enabled = true;
            this.qtyGroupBox.Controls.Add(qtyNotesBox);
            qtyNotesBox.BringToFront();
            qtyNotesBox.KeyDown +=new KeyEventHandler(qtyBox_TextChanged);


            dateNotesBox.Multiline = true;
            dateNotesBox.Parent = this.dateGroupBox;
            dateNotesBox.Left = 10;
            dateNotesBox.Width = this.dateGroupBox.Width - 20;
            dateNotesBox.Top = this.dateTimePicker1.Bottom + 10;
            dateNotesBox.Height = 50;
            dateNotesBox.WaterMarkText = "Enter Notes Here...";
            dateNotesBox.Visible = true;
            dateNotesBox.Enabled = true;
            this.dateGroupBox.Controls.Add(dateNotesBox);
            dateNotesBox.BringToFront();


            //newTextBox.BackColor = Color.Transparent;
            //dateTimePicker1.ValueChanged += new EventHandler(dateTimePicker1_ValueChanged);
            dateTimePicker1.KeyDown +=new KeyEventHandler(dateTimePicker1_ValueChanged);

            //qtyQtyTxt.BackColor = Color.Transparent;
            ItemNumbersWithAlertsConfigured = dbQuery.sqlQuery("SELECT DISTINCT ItemNumber FROM InventoryQuantityTriggers");

            ItemNumbersWithAlertsConfigured.ForEach(d => { d = d.Replace("|", ""); });


            totItemsWithAlertsLbl.Text = "Total Items With Alerts Configured: " + ItemNumbersWithAlertsConfigured.Count.ToString();
            totQtyAlertsLbl.Text = "Total Quantity Alerts Configured: " + dbQuery.sqlQuery("SELECT ItemNumber FROM InventoryQuantityTriggers WHERE TriggerType='qty'").Count.ToString();
            totDateAlertsLbl.Text = "Total Date Alerts Configured: " + dbQuery.sqlQuery("SELECT ItemNumber FROM InventoryQuantityTriggers WHERE TriggerType='date'").Count.ToString();


           // LoadNewQtyAlertData();
           // LoadNewDateAlertData();

            activeAlertsTotLbl.Text = "Active Alerts Total: " + (TotalDateAlerts + TotalQtyAlerts).ToString();

            //qtyAlertSwitch.Controls.Add(qtyAlertKnob);


            EnumerateCurrentItem(sender,e);



        }

        void qtyBox_TextChanged(object sender, EventArgs e)
        {
            if (sender.GetHashCode() != this.GetHashCode() & sender.GetHashCode() != this.button1.GetHashCode() & sender.GetHashCode() != this.button2.GetHashCode() & sender.GetHashCode() != this.qtyBox.GetHashCode())
            {
                EnableApplyButton();
            }
        }

        void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
           // if (sender.GetHashCode() != this.dateTimePicker1.GetHashCode())
           // {
                EnableApplyButton();
           // }
        }

        public void EnumerateCurrentItem(object sender, EventArgs e)
        {
            dateNotesBox.Text = "";
            dateTimePicker1.Value = DateTime.Now;
            //default: Show All Alerted Alerts
            
            if (FilterIndex == 0)
            {
               CurrentItem = GetAlertData(CurrentIndex, true, true, false);
            }

            //Show Alerted Quantity Only
            if (FilterIndex == 1)
            {
                CurrentItem = GetAlertData(CurrentIndex, true, false, false);
            }

            //Show Alerted Date Only
            if (FilterIndex == 2)
            {
                CurrentItem = GetAlertData(CurrentIndex, false, true, false);
            }
            
            //Show All Qty Alerts
            if (FilterIndex == 3)
            {
               
                CurrentItem = GetAlertData(CurrentIndex, true, false, true);
            }

            //Show All Date Alerts
            if (FilterIndex == 4)
            {
                
                CurrentItem = GetAlertData(CurrentIndex, false, true, true);
            }

            itemNumberLbl.Text = CurrentItem.ItemNumber;
            try
            {
                aqohLbl.Text = "AQOH: " + CurrentItem.ActiveAlerts[0].ActualQtyOnHand.ToString();
            }
            catch
            {
                aqohLbl.Text = "AQOH is not available.";
            }

            //turn all knobs off...
            if (qtyKnobOn == true)
            {
                qtyAlertKnob_Click(sender, e);
            }
            if (dateKnobOn == true)
            {
                dateAlertKnob_Click(sender, e);
            }
            if (aboveBelowKnobOn == true)
            {
                aboveBelowKnob_Click(sender, e);
            }
            qtyAlertBell.Visible = false;
            qtyGroupBox.BackColor = Color.PowderBlue;
            dateAlertBell.Visible = false;
            dateGroupBox.BackColor = Color.PowderBlue;

            foreach (AlertType alert in CurrentItem.ActiveAlerts)
            {
                

                if (alert.TriggerType == "qty")
                {
                    if (qtyKnobOn == false)
                    {
                        qtyAlertKnob_Click(sender, e);
                    }
                    if (alert.isBelow == true)
                    {
                        if (aboveBelowKnobOn == false)
                        {
                            aboveBelowKnob_Click(sender, e);
                        }
                      
                    }
                    else
                    {
                        if (aboveBelowKnobOn == true)
                        {
                            aboveBelowKnob_Click(sender, e);
                        }
                        
                    }
                    //qtyAlertKnob.Left += 50;

                    qtyBox.Text = alert.QtyThreshold.ToString();
                    qtyNotesBox.Text = alert.Notes;
                    if (alert.ActualQtyOnHand <= alert.QtyThreshold)
                    {
                        qtyAlertBell.Visible = true;
                        qtyGroupBox.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        qtyAlertBell.Visible = false;
                        if (qtyKnobOn == true)
                        {
                            qtyGroupBox.BackColor = Color.Khaki;
                        }
                        else
                        {
                            qtyGroupBox.BackColor = Color.PowderBlue;
                        }
                    }

                }
                if (alert.TriggerType == "date")
                {
                    //dateAlertKnob.Left += 50;
                    if (dateKnobOn == false)
                    {
                        dateAlertKnob_Click(sender, e);
                    }
                    dateNotesBox.Text = alert.Notes;
                    dateTimePicker1.Value = DateTime.Parse(alert.DateThreshold);
                    if (DateTime.Compare(DateTime.Parse(DateTime.Now.ToShortDateString()), DateTime.Parse(dateTimePicker1.Value.ToString())) >= 0)
                    {
                        dateAlertBell.Visible = true;
                        dateGroupBox.BackColor = Color.Firebrick;
                    }
                    else
                    {
                        dateAlertBell.Visible = false;
                        if (dateKnobOn == true)
                        {
                            dateGroupBox.BackColor = Color.Khaki;
                        }
                        else
                        {
                            dateGroupBox.BackColor = Color.PowderBlue;
                        }
                    }
                }
            }


            recordCountLbl.Text = (CurrentIndex + 1).ToString() + "/" + CurrentFilterList.Count.ToString();

        }


        public ItemAlerts GetAlertData(int index, bool showQtyAlerts, bool showDateAlerts, bool showInactiveAlerts)
        {
            //first, update the alert items...
            LoadNewDateAlertData();
            LoadNewQtyAlertData();

            
            CurrentFilterList = new List<string>();


            if (showDateAlerts == true)
            {
                foreach (string itm in ItemNumbersWithActiveDateAlerts)
                {
                //    bool isItemExisting = false;
                //    foreach (string itms in CurrentFilterList)
                //    {
                //        if (itms == itm)
                //        {
                //            isItemExisting = true;
                //            break;
                //        }                        
                //     
                //    }
                //    if (isItemExisting == false)
                //    {
                        CurrentFilterList.Add(itm);
                //    }

                }
            }
            if (showQtyAlerts == true)
            {
                foreach (string itm in ItemNumbersWithActiveQtyAlerts)
                {
                   
                            //bool isItemExisting = false;
                            //foreach (string itms in CurrentFilterList)
                            //{
                            //    if (itms == itm)
                            //    {
                            //        isItemExisting = true;
                            //        break;
                            //    }
                            //}
                            //if (isItemExisting == false)
                            //{
                                CurrentFilterList.Add(itm);
                            //}
     
                }
            }
            if (showInactiveAlerts == true)
            {
                foreach (string itm in ItemNumbersWithAlertsConfigured)
                {
                    bool isItemExisting = false;
                    foreach (string itms in CurrentFilterList)
                    {
                        if (itm == itms)
                        {
                            isItemExisting = true;
                            break;
                        }
                    }
                    if (isItemExisting == false)
                    {
                        CurrentFilterList.Add(itm);
                    }
                }
            }

           

            ItemAlerts currentItem = new ItemAlerts();
            AlertType qtyAlert = new AlertType();
            AlertType dateAlert = new AlertType();


            string CurrentItem_ID = CurrentFilterList[index].Split('|')[0];
 
            bool isQtyAlerted = false;
            bool isDateAlerted = false;
            string CurrentItem_Below = "";
            string CurrentItem_QtyThreshold = "";
            string CurrentItem_DateThreshold = "";
            string CurrentItem_QtyNotes = "";
            string CurrentItem_DateNotes = "";


            //question 1: is item date or qty alerted?
            //if (CurrentFilterList[index].Split('|')[1].Contains("/") == true)
            //{
            //    isDateAlerted = true;
           // }
           // else
           // {
          //      isQtyAlerted = true;
          //  }
            //question 2: does this item have more than one alert? and only search one way (to avoid dupes)
            //for (int x = index+1; x < CurrentFilterList.Count; x++)
            for (int x = 0; x <CurrentFilterList.Count;x++)
            {
                if (x != index)
                {
                    if (CurrentFilterList[x].Split('|')[0] == CurrentItem_ID)
                    {
                        //this item has multiple alerts...
                        if (CurrentFilterList[index].Split('|')[1].Contains("/") == true)
                        {
                            if (isDateAlerted == true)
                            {
                                MessageBox.Show("This item seems to have more than one date alert setup for it. Tell Tom as this shouldn't happen (at least yet)");
                            }
                            else
                            {
                                isDateAlerted = true;
                            }
                        }
                        else
                        {
                            if (isQtyAlerted == true)
                            {
                                MessageBox.Show("This item seems to have more than one qty alert setup for it. Tell Tom as this shouldn't happen (at least yet)");

                            }
                            else
                            {
                                isQtyAlerted = true;
                            }
                        }
                    }
                    else
                    {
                        //this item does not have multiple alerts. do nothing.
                    }
                }
            }



            currentItem.ItemNumber = CurrentItem_ID;
            currentItem.ActiveAlerts = new List<AlertType>();

            //if (isQtyAlerted == true)
           // {
                List<string> qtyResults = dbQuery.sqlQuery("SELECT QtyThreshold,Below,Notes FROM InventoryQuantityTriggers WHERE ItemNumber='" + CurrentItem_ID + "' AND TriggerType='qty'");
                if (qtyResults.Count > 0)
                {
                    
                       
                        CurrentItem_QtyThreshold = qtyResults[0].Split('|')[0];
                        CurrentItem_Below = qtyResults[0].Split('|')[1];
                        CurrentItem_QtyNotes = qtyResults[0].Split('|')[2];                        
                        qtyAlert.QtyThreshold = decimal.Parse(CurrentItem_QtyThreshold);
                        qtyAlert.isBelow = bool.Parse(CurrentItem_Below);
                        qtyAlert.Notes = CurrentItem_QtyNotes;
                        qtyAlert.TriggerType = "qty";

                        //now lets get the AQOH...
                        List<string> AQOH = dbQuery.sqlQuery("SELECT AQOH=SUM(QuantityOnHand)-SUM(QuantityOnSalesOrder) FROM InventoryQuantities WHERE ItemNumber='" + CurrentItem_ID + "'");
                        qtyAlert.ActualQtyOnHand = decimal.Parse(AQOH[0].Split('|')[0]);    

                        currentItem.ActiveAlerts.Add(qtyAlert);
                    
                }

            //}

           // if (isDateAlerted == true)
           // {
                List<string> dateResults = dbQuery.sqlQuery("SELECT DateThreshold,Notes FROM InventoryQuantityTriggers WHERE ItemNumber='" + CurrentItem_ID + "' AND TriggerType='date'");
                if (dateResults.Count > 0)
                {
                    CurrentItem_DateThreshold = dateResults[0].Split('|')[0];
                    CurrentItem_DateNotes = dateResults[0].Split('|')[1];
                    dateAlert.DateThreshold = CurrentItem_DateThreshold;
                    dateAlert.Notes = CurrentItem_DateNotes;
                    dateAlert.TriggerType = "date";
                    currentItem.ActiveAlerts.Add(dateAlert);
                }
           // }






            return currentItem;

        }


        void menuStrip1_MouseMove(object sender, MouseEventArgs e)
        {
            if (!Focused & newSettings.Visible == false)
            {
                Focus();
            }
        }

        public void LoadNewDateAlertData()
        {
            TotalDateAlerts = 0;
            ItemNumbersWithActiveDateAlerts = new List<string>();

            ItemNumbersWithActiveDateAlerts = dbQuery.sqlQuery("SELECT ItemNumber, DateThreshold, Notes FROM InventoryQuantityTriggers WHERE TriggerType='date' AND DateThreshold < '" + DateTime.Now.ToString() + "'");

            activeDateAlertsLbl.Text = "Active Date Alerts: " + ItemNumbersWithActiveDateAlerts.Count.ToString();
            TotalDateAlerts = ItemNumbersWithActiveDateAlerts.Count;
           


        }

        public void LoadNewQtyAlertData()
        {
            TotalQtyAlerts = 0;
            ItemNumbersWithActiveQtyAlerts = new List<string>();

            List<string> QtyAlerts = dbQuery.sqlQuery("SELECT ItemNumber, QtyThreshold,Below,Notes From InventoryQuantityTriggers WHERE TriggerType='qty'");
            foreach (string qtyConfiguredAlert in QtyAlerts)
            {
                if (qtyConfiguredAlert.Split('|')[0] == "SENFPRO42XP")
                {
                }
                int QtyThreshold = int.Parse(qtyConfiguredAlert.Split('|')[1]);
                List<string> getAQOH = dbQuery.sqlQuery("SELECT QuantityOnHand, QuantityOnSalesOrder FROM InventoryQuantities WHERE ItemNumber='" + qtyConfiguredAlert.Split('|')[0] + "'");
                int AQOH = int.Parse(getAQOH[0].Split('|')[0]) - int.Parse(getAQOH[0].Split('|')[1]);
                bool isBelow = bool.Parse(qtyConfiguredAlert.Split('|')[2]);

                if (isBelow == true)
                {
                    if (AQOH <= QtyThreshold)
                    {
                        ItemNumbersWithActiveQtyAlerts.Add(qtyConfiguredAlert);
                    }
                }
                else
                {
                    if (AQOH >= QtyThreshold)
                    {
                        ItemNumbersWithActiveQtyAlerts.Add(qtyConfiguredAlert);
                    }
                }


            }
            activeQtyAlertsLbl.Text = "Active Quantity Alerts: " + ItemNumbersWithActiveQtyAlerts.Count.ToString();
            TotalQtyAlerts = ItemNumbersWithActiveQtyAlerts.Count;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            CheckToSave(sender, e);
            CurrentIndex++;
            if (CurrentIndex >= CurrentFilterList.Count)
            {
                CurrentIndex = 0;
            }
            DisableApplyButton();
            EnumerateCurrentItem(sender,e);
            
        }

        private void dateAlertKnob_Click(object sender, EventArgs e)
        {
            if (dateKnobOn == false)
            {
                dateKnobOn = true;
                dateAlertKnob.Left += 50;
                dateAlertSwitch.BackgroundImage = AlertStandalone.Properties.Resources.back2;
            }
            else
            {
                dateKnobOn = false;
                dateAlertKnob.Left -= 50;
                dateAlertSwitch.BackgroundImage = AlertStandalone.Properties.Resources.back1;
            }
            if (sender.GetHashCode() != this.GetHashCode() & sender.GetHashCode() != this.button1.GetHashCode() & sender.GetHashCode() != this.button2.GetHashCode())
            {
                EnableApplyButton();
            }
        }

        private void qtyAlertKnob_Click(object sender, EventArgs e)
        {
            
            if (qtyKnobOn == false)
            {
                qtyKnobOn = true;
                qtyAlertKnob.Left += 50;
                qtyAlertSwitch.BackgroundImage = AlertStandalone.Properties.Resources.back2;
            }
            else
            {
                qtyKnobOn = false;
                qtyAlertKnob.Left -= 50;
                qtyAlertSwitch.BackgroundImage = AlertStandalone.Properties.Resources.back1;
            }
            if (sender.GetHashCode() != this.GetHashCode() & sender.GetHashCode() != this.button1.GetHashCode() & sender.GetHashCode() != this.button2.GetHashCode())
            {
                EnableApplyButton();
            }
        }

        private void dateAlertSwitch_Click(object sender, EventArgs e)
        {
            dateAlertKnob_Click(sender, e);

        }

        private void qtyAlertSwitch_Click(object sender, EventArgs e)
        {
            qtyAlertKnob_Click(sender, e);
        }

        private void aboveBelowKnob_Click(object sender, EventArgs e)
        {
            if (aboveBelowKnobOn == false)
            {
                aboveBelowKnobOn = true;
                aboveBelowKnob.Left += 50;
                aboveBelowSwitch.BackgroundImage = AlertStandalone.Properties.Resources.belowback2;
            }
            else
            {
                aboveBelowKnobOn = false;
                aboveBelowKnob.Left -= 50;
                aboveBelowSwitch.BackgroundImage = AlertStandalone.Properties.Resources.aboveback2;
            }
            if (sender.GetHashCode() != this.GetHashCode() & sender.GetHashCode() != this.button1.GetHashCode() & sender.GetHashCode() != this.button3.GetHashCode() & sender.GetHashCode() != this.button4.GetHashCode() & sender.GetHashCode() != this.button2.GetHashCode() & sender.GetHashCode() != this.button5.GetHashCode())
            {
                EnableApplyButton();
            }
        }

        private void aboveBelowSwitch_Click(object sender, EventArgs e)
        {
            aboveBelowKnob_Click(sender, e);
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            exitToolStripMenuItem_Click(sender, e);
        }


        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd,
                         int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        private static extern bool ReleaseCapture();

    

        private void menuStrip1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            newSettings = new settingsFrm();
            newSettings.Show();
            newSettings.Left = this.Left + 15;
            newSettings.Top = this.Top + 20;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CheckToSave(sender, e);
            CurrentIndex--;
            if (CurrentIndex < 0)
            {
                CurrentIndex = CurrentFilterList.Count - 1;
            }
            DisableApplyButton();
            EnumerateCurrentItem(sender,e);
        }

        private void filterCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (filterCombo.SelectedIndex == 3)
            {
                MessageBox.Show("This filter (Show All Qty Alerts) is still under construction...");
            }
            if (filterCombo.SelectedIndex == 4)
            {
                MessageBox.Show("This filter (Show All Date Alerts) is still under construction...");
            }
            
            CurrentIndex = 0;
            EnumerateCurrentItem(sender, e);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CheckToSave(sender,e);
            
            FilterIndex++;
            if (FilterIndex >= FilterList.Count)
            {
                FilterIndex = 0;
            }
            currentFilterLbl.Text = FilterList[FilterIndex];

            if (FilterIndex == 3)
            {
                MessageBox.Show("This filter (Show All Qty Alerts) is still under construction...");
            }
            if (FilterIndex == 4)
            {
                MessageBox.Show("This filter (Show All Date Alerts) is still under construction...");
            }
            CurrentIndex = 0;
            
            EnumerateCurrentItem(sender, e);
            DisableApplyButton();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            CheckToSave(sender,e);
            
            FilterIndex--;
            if (FilterIndex < 0)
            {
                FilterIndex = FilterList.Count - 1;
            }
            currentFilterLbl.Text = FilterList[FilterIndex];

            if (FilterIndex == 3)
            {
                MessageBox.Show("This filter (Show All Qty Alerts) is still under construction...");
            }
            if (FilterIndex == 4)
            {
                MessageBox.Show("This filter (Show All Date Alerts) is still under construction...");
            }
            CurrentIndex = 0;
            
            EnumerateCurrentItem(sender, e);
            DisableApplyButton();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (qtyKnobOn == true)
            {
                //We must update qty alert...
                int belowInt = 0;
                if (aboveBelowKnobOn == true)
                {
                    belowInt = 1;
                }
                else
                {
                    belowInt = 0;
                }
                dbQuery.sqlQuery("UPDATE InventoryQuantityTriggers SET QtyThreshold=" + qtyBox.Text + ",Below=" + belowInt.ToString() + ",Notes='" + qtyNotesBox.Text + "' WHERE ItemNumber='" + CurrentItem.ItemNumber + "' AND TriggerType='qty'");

            }
            else
            {
                //attempt to remove this alert if it exists...
                dbQuery.sqlQuery("DELETE FROM InventoryQuantityTriggers WHERE TriggerType='qty' AND ItemNumber='" + CurrentItem.ItemNumber + "'");
               
            }
            if (dateKnobOn == true)
            {
                //We must update date alert...
                dbQuery.sqlQuery("UPDATE InventoryQuantityTriggers SET DateThreshold='" + this.dateTimePicker1.Value.ToString() + "' WHERE ItemNumber='" + CurrentItem.ItemNumber + "' AND TriggerType='date'");
            }
            else
            {
                //attempt to remove this alert if it exists...
                dbQuery.sqlQuery("DELETE FROM InventoryQuantityTriggers WHERE ItemNumber='" + CurrentItem.ItemNumber + "' AND TriggerType='date'");
            }
        }



    }



}
