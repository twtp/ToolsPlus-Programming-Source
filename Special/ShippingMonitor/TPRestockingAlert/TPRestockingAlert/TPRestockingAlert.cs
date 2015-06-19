using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TPRestockingAlert
{
    public partial class TPRestockingAlert : UserControl
    {
        public TPRestockingAlert()
        {
            InitializeComponent();
            RefreshRestocking();
            this.SizeChanged += new EventHandler(TPRestockingAlert_SizeChanged);

        }

        void TPRestockingAlert_SizeChanged(object sender, EventArgs e)
        {
            RefreshRestocking();
        }

        private void TPRestockingAlert_Load(object sender, EventArgs e)
        {
            //GetTodaysPOsReceived();
            GetRestockingStatus();
        }
        private void RefreshRestocking()
        {
            RestockingHeaderLbl.Top = 0;
            RestockingHeaderLbl.Left = 0;
            RestockingHeaderLbl.Width = this.Width;
            //RestockingHeaderLbl.Height = this.Height / 3;

           

            //POCountLbl.Left = RestockingHeaderLbl.Right;
            //POCountLbl.Top = 0;
           // POCountLbl.Width = this.Width - POCountLbl.Left;
            //POCountLbl.Height = RestockingHeaderLbl.Height / 2;
           // POCountLbl.Refresh();

            //POSumLbl.Left = RestockingHeaderLbl.Right;
           // POSumLbl.Top = POCountLbl.Bottom;
           // POSumLbl.Width = this.Width - POSumLbl.Left;
          //  POSumLbl.Height = RestockingHeaderLbl.Height / 2;

            TotDifferentComponentsLbl.Top = this.RestockingHeaderLbl.Bottom-10;
            TotDifferentComponentsLbl.Left = 0;
            TotDifferentComponentsLbl.Width = this.Width / 2;
            TotDifferentComponentsLbl.Height = this.Height / 5;

            TotalQtyLbl.Left = TotDifferentComponentsLbl.Right;
            TotalQtyLbl.Top = TotDifferentComponentsLbl.Top;
            TotalQtyLbl.Width = this.Width - TotalQtyLbl.Left;
            TotalQtyLbl.Height = TotDifferentComponentsLbl.Height;

            notPutAwayLbl.Top = TotDifferentComponentsLbl.Bottom;
            notPutAwayLbl.Left = 0;
            notPutAwayLbl.Width = this.Width;
            //notPutAwayLbl.Height = this.Height / 10;

            int avgLabelWidth = (MondayTotLbl.Width + TuesdayTotLbl.Width + WednesdayTotLbl.Width + ThursdayTotLbl.Width + FridayTotLbl.Width + SaturdayTotLbl.Width);
            int freespace = this.Width - avgLabelWidth;
            

            MondayTotLbl.Top = TotDifferentComponentsLbl.Bottom;
            MondayTotLbl.Left = 0;
            TuesdayTotLbl.Top = MondayTotLbl.Top;
            TuesdayTotLbl.Left = MondayTotLbl.Right + (freespace / 6);
            WednesdayTotLbl.Top = TuesdayTotLbl.Top;
            WednesdayTotLbl.Left = TuesdayTotLbl.Right + (freespace/6);
            ThursdayTotLbl.Top = WednesdayTotLbl.Top;
            ThursdayTotLbl.Left = WednesdayTotLbl.Right + (freespace/6);
            FridayTotLbl.Top = ThursdayTotLbl.Top;
            FridayTotLbl.Left = ThursdayTotLbl.Right + (freespace/6);
            SaturdayTotLbl.Top = FridayTotLbl.Top;
            SaturdayTotLbl.Left = FridayTotLbl.Right + (freespace/6);



        }

        public void GetRestockingStatus()
        {
            List<string> OlderRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate<'" + DateTime.Now.AddDays(-5).ToShortDateString() + "' ORDER BY LastInventoriedDate");
            List<string> FourthDayRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate>'" + DateTime.Now.AddDays(-4).ToShortDateString() + "' AND LastInventoriedDate < '" + DateTime.Now.AddDays(-3).ToShortDateString() + "' ORDER BY LastInventoriedDate");
            List<string> ThirdDayRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate>'" + DateTime.Now.AddDays(-3).ToShortDateString() + "' AND LastInventoriedDate < '" + DateTime.Now.AddDays(-2).ToShortDateString() + "' ORDER BY LastInventoriedDate");
            List<string> SecondDayRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate>'" + DateTime.Now.AddDays(-2).ToShortDateString() + "' AND LastInventoriedDate < '" + DateTime.Now.AddDays(-1).ToShortDateString() + "' ORDER BY LastInventoriedDate");
            List<string> FirstDayRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate>'" + DateTime.Now.AddDays(-1).ToShortDateString() + "' AND LastInventoriedDate < '" + DateTime.Now.ToShortDateString() + "' ORDER BY LastInventoriedDate");
            List<string> ZeroDayRestockings = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ComponentID, Quantity,LastInventoriedDate FROM LocationContents WHERE LocationID=10830 AND LastInventoriedDate>'" + DateTime.Now.ToShortDateString() + "' AND LastInventoriedDate < '" + DateTime.Now.ToString() + "' ORDER BY LastInventoriedDate");
            if (OlderRestockings.Count > 0 | FourthDayRestockings.Count > 0 | ThirdDayRestockings.Count > 0 | SecondDayRestockings.Count > 0 | FirstDayRestockings.Count > 0 | ZeroDayRestockings.Count >0)
            {
                //do the processin...
                int totComponents = 0;
                int totQuantity = 0;
                int olderCount = 0;
                foreach (string older in OlderRestockings)
                {
                    olderCount += int.Parse(older.Split('|')[1]);                    
                }
                totQuantity += olderCount;
                totComponents += OlderRestockings.Count;
                MondayTotLbl.Text = "Older\r\n" + olderCount.ToString();
                int fourthCount = 0;
                foreach (string fourth in FourthDayRestockings)
                {
                    fourthCount += int.Parse(fourth.Split('|')[1]);
                }
                totQuantity += fourthCount;
                totComponents += FourthDayRestockings.Count;
                TuesdayTotLbl.Text = "4 Days\r\n" + fourthCount.ToString();
                int thirdCount = 0;
                foreach (string third in ThirdDayRestockings)
                {
                    thirdCount += int.Parse(third.Split('|')[1]);
                }
                totQuantity += thirdCount;
                totComponents += ThirdDayRestockings.Count;
                WednesdayTotLbl.Text = "3 Days\r\n" + thirdCount.ToString();
                int secondCount = 0;
                foreach (string second in SecondDayRestockings)
                {
                    secondCount += int.Parse(second.Split('|')[1]);
                }
                totQuantity += secondCount;
                totComponents += SecondDayRestockings.Count;
                ThursdayTotLbl.Text = "2 Days\r\n" + secondCount.ToString();
                int firstCount = 0;
                foreach (string first in FirstDayRestockings)
                {
                    firstCount += int.Parse(first.Split('|')[1]);
                }
                totQuantity += firstCount;
                totComponents += FirstDayRestockings.Count;
                FridayTotLbl.Text = "1 Day\r\n" + firstCount.ToString();
                int zeroCount = 0;
                foreach (string zero in ZeroDayRestockings)
                {
                    zeroCount += int.Parse(zero.Split('|')[1]);
                }
                totQuantity += zeroCount;
                totComponents += ZeroDayRestockings.Count;
                SaturdayTotLbl.Text = "Today\r\n" + zeroCount.ToString();

                TotalQtyLbl.Text = "Total Qty: " + totQuantity.ToString();
                TotDifferentComponentsLbl.Text = "Components: " + totComponents.ToString();
                    


                this.Visible = true;
                this.BringToFront();
            }
            else
            {
                this.Visible = false;
                this.SendToBack();
            }


        }

        public string GetTodaysPOsReceived()
        {
            
            //List<string> TodaysPOs = RestockingConnectivity.SQLCalls.sqlQuery("SELECT PONumber,POAmount,DateExported,VendorNumber FROM PurchaseOrders WHERE DateExported>'" + DateTime.Now.ToShortDateString() + "' AND ShipToAddress1='60 Scott Rd' AND Finalized=1");
            List<string> TodaysPOs = RestockingConnectivity.SQLCalls.sqlQuery("SELECT TOP 1 PONumber,POAmount,DateExported,VendorNumber,ID FROM PurchaseOrders WHERE DateExported>'11-03-2014' AND ShipToAddress1='60 Scott Rd' AND Finalized=1");
            if (TodaysPOs.Count > 1)
            {
                return showMultiPO(TodaysPOs);
            }
            else if (TodaysPOs.Count == 1)
            {
                return showSinglePO(TodaysPOs[0]);
            }
            else
            {
                return "";
            }
        }
        private void getPOReceivingDetails(List<string> PONumber)
        {
            int totComponents = 0;
            int totQty = 0;
            foreach (string po in PONumber)
            {
                List<string> poList = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ID FROM ReceivingWorksheetHeader WHERE PurchaseOrderNo='" + po + "'");
                if (poList.Count > 0)
                {
                    List<string> poDetails = RestockingConnectivity.SQLCalls.sqlQuery("SELECT ItemNumber,ActualReceivedQty FROM PurchaseOrderLines WHERE HeaderID=" + poList[0].Split('|')[0]);
                    if (poDetails.Count > 0)
                    {
                        totComponents = poDetails.Count;
                        foreach (string item in poDetails)
                        {
                            totQty += int.Parse(item.Split('|')[1]);
                        }

                    }
                }
            }
            TotDifferentComponentsLbl.Text = "Components: " + totComponents.ToString();
            TotalQtyLbl.Text = "Total Qty: " + totQty.ToString();
        }
        private string showSinglePO(string singlePO)
        {
            List<string> tmpPOList = new List<string>();
            tmpPOList.Add(singlePO);
            getPOReceivingDetails(tmpPOList);
            decimal singlePOCost = decimal.Parse(singlePO.Split('|')[1]);
            string SinglePOCost = singlePOCost.ToString("#.##");
            if (SinglePOCost == "")
            {
                SinglePOCost = "N/A";
            }
            //this.POCountLbl.Text = "PO: " + singlePO.Split('|')[0];
            //this.POSumLbl.Text = "Net Worth: $" + SinglePOCost;
            string SinglePOStr = "RECEIVED IN - #" + singlePO.Split('|')[0] + " $" + SinglePOCost + " " + singlePO.Split('|')[3];

            return SinglePOStr;
        }
        private string showMultiPO(List<string> multiplePO)
        {
            decimal RunningPrice = 0;
            int poCount =0;
            foreach (string po in multiplePO)
            {
                try
                {
                    if (po.Split('|')[1].Length > 0)
                    {
                        poCount++;
                        RunningPrice += decimal.Parse(po.Split('|')[1]);
                    }
                }
                catch
                {
                }
            }
            //this.POCountLbl.Text = "POs: " + poCount.ToString();
            //this.POSumLbl.Text = "Net Worth: $" + RunningPrice.ToString("#.##");
            
            string MultiPOStr = "RECEIVED IN " + multiplePO.Count.ToString() + " POs ($" + RunningPrice.ToString("#.##") + ")";
            return MultiPOStr;

        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
            GetRestockingStatus();
        }



    }
}
