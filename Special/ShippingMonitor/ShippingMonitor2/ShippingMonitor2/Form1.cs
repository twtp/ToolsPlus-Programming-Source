using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ShippingMonitor2
{
    public partial class Form1 : Form
    {
        public Screen selectedScreen;
        private const int TickerVerticalSizeDivisor = 12;
        private const int TickerHorizontalSizeSubtractor = 75;
        private const int HeaderVerticalSizeDivisor = 8;
        private const int ListViewVerticalOffset = 10;

        public bool amEbayRestocking = false;
        public bool pmEbayRestocking = false;
        public bool yahooRestocking = false;
        public bool otherRestocking = false;

      


        private static int ListViewAutoScrollLine = 0;
        public Form1()
        {
            
            InitializeComponent();
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            
        }
        public ListView GetListView()
        {
            return packersLVI;
        }
        private void SizeForm()
        {
            
            this.Left = selectedScreen.Bounds.Left;
            this.Top = selectedScreen.Bounds.Top;
            this.Width = selectedScreen.Bounds.Width;
            this.Height = selectedScreen.Bounds.Height;

            TopCaptionLbl.Top = 0;
            TopCaptionLbl.Left = 0;
            TopCaptionLbl.Width = this.Width;
            TopCaptionLbl.Height = this.Height / HeaderVerticalSizeDivisor;

            ebayAMTicker.Left = 25;
            ebayAMTicker.Top = TopCaptionLbl.Bottom;
            ebayAMTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            ebayAMTicker.Height = this.Height / TickerVerticalSizeDivisor;

            ebayPMTicker.Left = 25;
            ebayPMTicker.Top = ebayAMTicker.Bottom;
            ebayPMTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            ebayPMTicker.Height = this.Height / TickerVerticalSizeDivisor;

            yahooTicker.Left = 25;
            yahooTicker.Top = ebayPMTicker.Bottom;
            yahooTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            yahooTicker.Height = this.Height / TickerVerticalSizeDivisor;

            otherOrdersTicker.Left = 25;
            otherOrdersTicker.Top = yahooTicker.Bottom;
            otherOrdersTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            otherOrdersTicker.Height = this.Height / TickerVerticalSizeDivisor;

            funFactsTicker.Left = 25;
            funFactsTicker.Top = otherOrdersTicker.Bottom;
            funFactsTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            funFactsTicker.Height = this.Height / TickerVerticalSizeDivisor;

            testTicker.Left = 25;
            testTicker.Top = otherOrdersTicker.Bottom;
            testTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            testTicker.Height = this.Height / TickerVerticalSizeDivisor;
            testTicker.Visible = false;

            funFactsTicker.BringToFront();

            totesInfoTicker.Parent = this;
            totesInfoTicker.Left = 25;
            totesInfoTicker.Top = funFactsTicker.Bottom;
            totesInfoTicker.Width = this.Width - TickerHorizontalSizeSubtractor;
            totesInfoTicker.Height = this.Height / TickerVerticalSizeDivisor;
            

            packersLVI.Left = 0;
            packersLVI.Top = totesInfoTicker.Bottom + ListViewVerticalOffset;
            packersLVI.Height = this.Height - funFactsTicker.Bottom;
            packersLVI.Width = this.Width;
            packersLVI.Scrollable = false;

            tpClock1.Width = this.Width / 4;
            tpClock1.Height = TopCaptionLbl.Height;
            tpClock1.Left = this.Width - tpClock1.Width;
            tpClock1.Top = 0;
            tpClock1.SetBackgroundColor(Color.Transparent);

            tpWillCall1.Width = (int)(this.Width / 3.25f);
            tpWillCall1.Height = TopCaptionLbl.Height;
            tpWillCall1.Left = this.Width - tpWillCall1.Width;
            tpWillCall1.Top = 0;

            tpRestockingAlert1.Width = (int)(this.Width / 3.25f);
            tpRestockingAlert1.Height = TopCaptionLbl.Height;
            tpRestockingAlert1.Left = 0;
            tpRestockingAlert1.Top = 0;

            tpTransferAlert1.Top = TopCaptionLbl.Height /2;
            tpTransferAlert1.Left = tpRestockingAlert1.Right;
            tpTransferAlert1.Width = this.Width - tpWillCall1.Width - tpRestockingAlert1.Width;
            tpTransferAlert1.Height = TopCaptionLbl.Height / 2;

            tppoReceivedAlert1.Top = 0;
            tppoReceivedAlert1.Left = tpRestockingAlert1.Right;
            tppoReceivedAlert1.Width = tpTransferAlert1.Width;
            tppoReceivedAlert1.Height = TopCaptionLbl.Height / 2;

        }

        private void Form1_Load(object sender, EventArgs e)
        {            
            SizeForm();
            //ebayPMTicker.SetHeaderSubText("Tot:");
            ebayAMTicker.SetTickerMethod(ebayAMThread);
            ebayPMTicker.SetTickerMethod(ebayPMThread);
            yahooTicker.SetTickerMethod(yahooThread);
            otherOrdersTicker.SetTickerMethod(otherOrdersThread);
            funFactsTicker.SetTickerMethod(funFactsThread);
          
            totesInfoTicker.SetTickerMethod(totesInfoThread);

            PackingTime.PackingTimes.ResetListView(ref packersLVI, this.Width);
            PackingTime.PackingTimes.UpdatePackers(ref packersLVI, this.Width);

            Rectangle picCroppedRegion = new Rectangle(0, Properties.Resources.background.Height / 2, Properties.Resources.background.Width-1, Properties.Resources.background.Height-1);
            packersLVI.BackgroundImage = Properties.Resources.backgroundshort;

            tpWillCall1.SetBackColor(Color.FromArgb(255, 192, 0, 0));
            tpWillCall1.SetForeColor(Color.White);


            //this.packersLVI.BackgroundImage = this.BackgroundImage;
            //this.packersLVI.BackgroundImageTiled = true;
            //ebayAMTicker.isEnabled = true;
            //ebayPMTicker.isEnabled = true;
            //ebayPMTicker.DisableSubHeader();
            
            //ebayPMTicker.showSubHeader = true;
        }



        public TPTickerCtl.TPTicker.TickerData ebayAMThread()
        {
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();   
            List<string> morningOrderList = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "' AND TimeInserted < '" + DateTime.Now.ToShortDateString() + " 12:00:00.000'");
            string tickerText = "";
            int horizontalPicks = 0;
            int verticalPicks = 0;
            int flowRackPicks = 0;
            int longPicks = 0;
            int tunnelPicks = 0;
            int upstairsPicks = 0;
            int containerPicks = 0;
            int palletizedPicks = 0;
            int restockingPicks = 0;

            if (morningOrderList.Count > 0)
            {


                foreach (string amEbayOrder in morningOrderList)
                {
                    tickerText += amEbayOrder.Split('|')[0] + ", ";
                    
                    List<string> getOrderInfo = Connectivity.SQLCalls.sqlQuery("SELECT LocationContents.ComponentID,LocationTypeID,Quantity,QuantityOrdered FROM LocationContents,LocationMaster,WhseTaskSalesOrderLines WHERE WhseTaskSalesOrderLines.HeaderID=" + amEbayOrder.Split('|')[1] + " AND LocationMaster.ID = LocationContents.LocationID AND QuantityOrdered < Quantity AND LocationContents.ComponentID IN (SELECT ComponentID FROM WhseTaskSalesOrderLines WHERE HeaderID=" + amEbayOrder.Split('|')[1] + ") ORDER BY CASE WHEN LocationTypeID=2 THEN 0 WHEN LocationTypeID=3 THEN 1 WHEN LocationTypeID=1 THEN 2 WHEN LocationTypeID=4 THEN 3 WHEN LocationTypeID=16 THEN 4  WHEN LocationTypeID=17 THEN 5 WHEN LocationTypeID=18 THEN 6 WHEN LocationTypeID = 0 THEN 7 WHEN LocationTypeID=14 THEN 8 ELSE 9 END");
                    if (getOrderInfo.Count > 0)
                    {
                        string lastComponent = "";
                        

                        //lol, so that long ass sql query basically got us everything we need to know...
                        //if something was returned, it means that location has enough quantity to fulfill the order...
                        //however, the locations are also order by FPZ so to speak. So pull the top one an ur good!
                        foreach(string result in getOrderInfo)
                        {
                            if (result.Split('|')[0] != lastComponent)
                            {
                                lastComponent = result.Split('|')[0];
                                string locTypeID = result.Split('|')[1];
                                if (locTypeID == "2")
                                    horizontalPicks++;
                                if (locTypeID == "3")
                                    verticalPicks++;
                                if (locTypeID == "1")
                                    flowRackPicks++;
                                if (locTypeID == "4")
                                    longPicks++;
                                if (locTypeID == "16")
                                    tunnelPicks++;
                                if (locTypeID == "17")
                                    upstairsPicks++;
                                if (locTypeID == "18")
                                    containerPicks++;
                                if (locTypeID == "0")
                                    palletizedPicks++;
                                if (locTypeID == "14")
                                    restockingPicks++;

                            }
                        }


                    }

                }
                tickerText = tickerText.Substring(0, tickerText.Length - 2);

            }
            TD.Enabled = true;            
            TD.HeaderText = "AM EBay: " + morningOrderList.Count.ToString() + "\r\n";
            TD.HeaderSubText = "";
            if (horizontalPicks > 0)
            {
                TD.HeaderSubText += " HOR:" + horizontalPicks.ToString() + "  ";
            }
            if (verticalPicks > 0)
            {
                TD.HeaderSubText += " VER:" + verticalPicks.ToString() + "  ";
            }
            if (flowRackPicks > 0)
            {
                TD.HeaderSubText += " FLO:" + flowRackPicks.ToString() + "  ";
            }
            if (longPicks > 0)
            {
                TD.HeaderSubText += " LNG:" + longPicks.ToString() + "  ";
            }
            if (tunnelPicks > 0)
            {
                TD.HeaderSubText += " TNL:" + tunnelPicks.ToString() + "  ";
            }
            if (upstairsPicks > 0)
            {
                TD.HeaderSubText += " UST:" + upstairsPicks.ToString() + "  ";
            }
            if (containerPicks > 0)
            {
                TD.HeaderSubText += " CON:" + containerPicks.ToString() + "  ";
            }
            if (restockingPicks > 0)
            {
                amEbayRestocking = true;
                TD.HeaderSubText += " R/S:" + restockingPicks.ToString() + "  ";
            }
            else
            {
                amEbayRestocking = false;
            }
            
            
           
            

            //this is where we must get locations and components there...


            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.OrangeRed;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 60;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            TD.ShowSubHeader = true;
            TD.TickerText = tickerText;
            return TD;
            
                    
        }

        public TPTickerCtl.TPTicker.TickerData ebayPMThread()
        {
            int horizontalPicks = 0;
            int verticalPicks = 0;
            int flowRackPicks = 0;
            int longPicks = 0;
            int tunnelPicks = 0;
            int upstairsPicks = 0;
            int containerPicks = 0;
            int palletizedPicks = 0;
            int restockingPicks = 0;

            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();
            List<string> afternoonOrderList = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + " 12:00:00.000'");
            string tickerText = "";
            if (afternoonOrderList.Count > 0)
            {
                foreach (string pmEbayOrder in afternoonOrderList)
                {
                    tickerText += pmEbayOrder.Split('|')[0] + ", ";

                    List<string> getOrderInfo = Connectivity.SQLCalls.sqlQuery("SELECT LocationContents.ComponentID,LocationTypeID,Quantity,QuantityOrdered FROM LocationContents,LocationMaster,WhseTaskSalesOrderLines WHERE WhseTaskSalesOrderLines.HeaderID=" + pmEbayOrder.Split('|')[1] + " AND LocationMaster.ID = LocationContents.LocationID AND QuantityOrdered < Quantity AND LocationContents.ComponentID IN (SELECT ComponentID FROM WhseTaskSalesOrderLines WHERE HeaderID=" + pmEbayOrder.Split('|')[1] + ") ORDER BY CASE WHEN LocationTypeID=2 THEN 0 WHEN LocationTypeID=3 THEN 1 WHEN LocationTypeID=1 THEN 2 WHEN LocationTypeID=4 THEN 3 WHEN LocationTypeID=16 THEN 4  WHEN LocationTypeID=17 THEN 5 WHEN LocationTypeID=18 THEN 6 WHEN LocationTypeID = 0 THEN 7 WHEN LocationTypeID=14 THEN 8 ELSE 9 END");
                    if (getOrderInfo.Count > 0)
                    {
                        string lastComponent = "";


                        //lol, so that long ass sql query basically got us everything we need to know...
                        //if something was returned, it means that location has enough quantity to fulfill the order...
                        //however, the locations are also order by FPZ so to speak. So pull the top one an ur good!
                        foreach (string result in getOrderInfo)
                        {
                            if (result.Split('|')[0] != lastComponent)
                            {
                                lastComponent = result.Split('|')[0];
                                string locTypeID = result.Split('|')[1];
                                if (locTypeID == "2")
                                    horizontalPicks++;
                                if (locTypeID == "3")
                                    verticalPicks++;
                                if (locTypeID == "1")
                                    flowRackPicks++;
                                if (locTypeID == "4")
                                    longPicks++;
                                if (locTypeID == "16")
                                    tunnelPicks++;
                                if (locTypeID == "17")
                                    upstairsPicks++;
                                if (locTypeID == "18")
                                    containerPicks++;
                                if (locTypeID == "0")
                                    palletizedPicks++;
                                if (locTypeID == "14")
                                    restockingPicks++;

                            }
                        }


                    }
                }
                tickerText = tickerText.Substring(0, tickerText.Length - 2);
            }
            
            TD.Enabled = true;
            TD.HeaderText = "PM EBay: " + afternoonOrderList.Count.ToString() + "\r\n";
            TD.HeaderSubText = "";

            if (horizontalPicks > 0)
            {
                TD.HeaderSubText += " HOR:" + horizontalPicks.ToString() + "  ";
            }
            if (verticalPicks > 0)
            {
                TD.HeaderSubText += " VER:" + verticalPicks.ToString() + "  ";
            }
            if (flowRackPicks > 0)
            {
                TD.HeaderSubText += " FLO:" + flowRackPicks.ToString() + "  ";
            }
            if (longPicks > 0)
            {
                TD.HeaderSubText += " LNG:" + longPicks.ToString() + "  ";
            }
            if (tunnelPicks > 0)
            {
                TD.HeaderSubText += " TNL:" + tunnelPicks.ToString() + "  ";
            }
            if (upstairsPicks > 0)
            {
                TD.HeaderSubText += " UST:" + upstairsPicks.ToString() + "  ";
            }
            if (containerPicks > 0)
            {
                TD.HeaderSubText += " CON:" + containerPicks.ToString() + "  ";
            }
            if (restockingPicks > 0)
            {
                pmEbayRestocking = true;
                TD.HeaderSubText += " R/S:" + restockingPicks.ToString() + "  ";
            }
            else
            {
                pmEbayRestocking = false;
            }


            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.Green;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 56;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            TD.ShowSubHeader = true;
            TD.TickerText = tickerText;
            return TD;


        }
        public TPTickerCtl.TPTicker.TickerData yahooThread()
        {
            int horizontalPicks = 0;
            int verticalPicks = 0;
            int flowRackPicks = 0;
            int longPicks = 0;
            int tunnelPicks = 0;
            int upstairsPicks = 0;
            int containerPicks = 0;
            int palletizedPicks = 0;
            int restockingPicks = 0;
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();
            List<string> yahooOrderList = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'Y%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            string tickerText = "";
            if (yahooOrderList.Count > 0)
            {
                foreach (string yahooOrder in yahooOrderList)
                {
                    tickerText += yahooOrder.Split('|')[0] + ", ";
                    List<string> getOrderInfo = Connectivity.SQLCalls.sqlQuery("SELECT LocationContents.ComponentID,LocationTypeID,Quantity,QuantityOrdered FROM LocationContents,LocationMaster,WhseTaskSalesOrderLines WHERE WhseTaskSalesOrderLines.HeaderID=" + yahooOrder.Split('|')[1] + " AND LocationMaster.ID = LocationContents.LocationID AND QuantityOrdered < Quantity AND LocationContents.ComponentID IN (SELECT ComponentID FROM WhseTaskSalesOrderLines WHERE HeaderID=" + yahooOrder.Split('|')[1] + ") ORDER BY CASE WHEN LocationTypeID=2 THEN 0 WHEN LocationTypeID=3 THEN 1 WHEN LocationTypeID=1 THEN 2 WHEN LocationTypeID=4 THEN 3 WHEN LocationTypeID=16 THEN 4  WHEN LocationTypeID=17 THEN 5 WHEN LocationTypeID=18 THEN 6 WHEN LocationTypeID = 0 THEN 7 WHEN LocationTypeID=14 THEN 8 ELSE 9 END");
                    if (getOrderInfo.Count > 0)
                    {
                        string lastComponent = "";


                        //lol, so that long ass sql query basically got us everything we need to know...
                        //if something was returned, it means that location has enough quantity to fulfill the order...
                        //however, the locations are also order by FPZ so to speak. So pull the top one an ur good!
                        foreach (string result in getOrderInfo)
                        {
                            if (result.Split('|')[0] != lastComponent)
                            {
                                lastComponent = result.Split('|')[0];
                                string locTypeID = result.Split('|')[1];
                                if (locTypeID == "2")
                                    horizontalPicks++;
                                if (locTypeID == "3")
                                    verticalPicks++;
                                if (locTypeID == "1")
                                    flowRackPicks++;
                                if (locTypeID == "4")
                                    longPicks++;
                                if (locTypeID == "16")
                                    tunnelPicks++;
                                if (locTypeID == "17")
                                    upstairsPicks++;
                                if (locTypeID == "18")
                                    containerPicks++;
                                if (locTypeID == "0")
                                    palletizedPicks++;
                                if (locTypeID == "14")
                                    restockingPicks++;

                            }
                        }


                    }
                }
                tickerText = tickerText.Substring(0, tickerText.Length - 2);
            }
            TD.Enabled = true;
            TD.HeaderText = "Yahoo: " + yahooOrderList.Count.ToString() + "\r\n";
            TD.HeaderSubText = "";
            if (horizontalPicks > 0)
            {
                TD.HeaderSubText += " HOR:" + horizontalPicks.ToString() + "  ";
            }
            if (verticalPicks > 0)
            {
                TD.HeaderSubText += " VER:" + verticalPicks.ToString() + "  ";
            }
            if (flowRackPicks > 0)
            {
                TD.HeaderSubText += " FLO:" + flowRackPicks.ToString() + "  ";
            }
            if (longPicks > 0)
            {
                TD.HeaderSubText += " LNG:" + longPicks.ToString() + "  ";
            }
            if (tunnelPicks > 0)
            {
                TD.HeaderSubText += " TNL:" + tunnelPicks.ToString() + "  ";
            }
            if (upstairsPicks > 0)
            {
                TD.HeaderSubText += " UST:" + upstairsPicks.ToString() + "  ";
            }
            if (containerPicks > 0)
            {
                TD.HeaderSubText += " CON:" + containerPicks.ToString() + "  ";
            }
            if (restockingPicks > 0)
            {
                yahooRestocking = true;
                TD.HeaderSubText += " R/S:" + restockingPicks.ToString() + "  ";
            }
            else
            {
                yahooRestocking = false;
            }
            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.Orange;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 51;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            TD.ShowSubHeader = true;
            TD.TickerText = tickerText;
            return TD;


        }

        public TPTickerCtl.TPTicker.TickerData otherOrdersThread()
        {
            int horizontalPicks = 0;
            int verticalPicks = 0;
            int flowRackPicks = 0;
            int longPicks = 0;
            int tunnelPicks = 0;
            int upstairsPicks = 0;
            int containerPicks = 0;
            int palletizedPicks = 0;
            int restockingPicks = 0;
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();
            List<string> otherOrderList = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo NOT LIKE 'Y%' AND SalesOrderNo NOT LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            string tickerText = "";
            if (otherOrderList.Count > 0)
            {
                foreach (string otherOrder in otherOrderList)
                {
                    tickerText += otherOrder.Split('|')[0] + ", ";
                    List<string> getOrderInfo = Connectivity.SQLCalls.sqlQuery("SELECT LocationContents.ComponentID,LocationTypeID,Quantity,QuantityOrdered FROM LocationContents,LocationMaster,WhseTaskSalesOrderLines WHERE WhseTaskSalesOrderLines.HeaderID=" + otherOrder.Split('|')[1] + " AND LocationMaster.ID = LocationContents.LocationID AND QuantityOrdered < Quantity AND LocationContents.ComponentID IN (SELECT ComponentID FROM WhseTaskSalesOrderLines WHERE HeaderID=" + otherOrder.Split('|')[1] + ") ORDER BY CASE WHEN LocationTypeID=2 THEN 0 WHEN LocationTypeID=3 THEN 1 WHEN LocationTypeID=1 THEN 2 WHEN LocationTypeID=4 THEN 3 WHEN LocationTypeID=16 THEN 4  WHEN LocationTypeID=17 THEN 5 WHEN LocationTypeID=18 THEN 6 WHEN LocationTypeID = 0 THEN 7 WHEN LocationTypeID=14 THEN 8 ELSE 9 END");
                    if (getOrderInfo.Count > 0)
                    {
                        string lastComponent = "";


                        //lol, so that long ass sql query basically got us everything we need to know...
                        //if something was returned, it means that location has enough quantity to fulfill the order...
                        //however, the locations are also order by FPZ so to speak. So pull the top one an ur good!
                        foreach (string result in getOrderInfo)
                        {
                            if (result.Split('|')[0] != lastComponent)
                            {
                                lastComponent = result.Split('|')[0];
                                string locTypeID = result.Split('|')[1];
                                if (locTypeID == "2")
                                    horizontalPicks++;
                                if (locTypeID == "3")
                                    verticalPicks++;
                                if (locTypeID == "1")
                                    flowRackPicks++;
                                if (locTypeID == "4")
                                    longPicks++;
                                if (locTypeID == "16")
                                    tunnelPicks++;
                                if (locTypeID == "17")
                                    upstairsPicks++;
                                if (locTypeID == "18")
                                    containerPicks++;
                                if (locTypeID == "0")
                                    palletizedPicks++;
                                if (locTypeID == "14")
                                    restockingPicks++;

                            }
                        }


                    }
                }
                tickerText = tickerText.Substring(0, tickerText.Length - 2);
            }
            TD.Enabled = true;
            TD.HeaderText = "Others: " + otherOrderList.Count.ToString() + "\r\n";
            TD.HeaderSubText = "";
            if (horizontalPicks > 0)
            {
                TD.HeaderSubText += " HOR:" + horizontalPicks.ToString() + "  ";
            }
            if (verticalPicks > 0)
            {
                TD.HeaderSubText += " VER:" + verticalPicks.ToString() + "  ";
            }
            if (flowRackPicks > 0)
            {
                TD.HeaderSubText += " FLO:" + flowRackPicks.ToString() + "  ";
            }
            if (longPicks > 0)
            {
                TD.HeaderSubText += " LNG:" + longPicks.ToString() + "  ";
            }
            if (tunnelPicks > 0)
            {
                TD.HeaderSubText += " TNL:" + tunnelPicks.ToString() + "  ";
            }
            if (upstairsPicks > 0)
            {
                TD.HeaderSubText += " UST:" + upstairsPicks.ToString() + "  ";
            }
            if (containerPicks > 0)
            {
                TD.HeaderSubText += " CON:" + containerPicks.ToString() + "  ";
            }
            if (restockingPicks > 0)
            {
                otherRestocking = true;
                TD.HeaderSubText += " R/S:" + restockingPicks.ToString() + "  ";
            }
            else
            {
                otherRestocking = false;
            }
            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.Orange;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 45;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            TD.ShowSubHeader = true;
            TD.TickerText = tickerText;
            return TD;


        }

        public TPTickerCtl.TPTicker.TickerData funFactsThread()
        {
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();
            //List<string> otherOrderList = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo NOT LIKE 'Y%' AND SalesOrderNo NOT LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
             string tickerText = "";
            if (amEbayRestocking == true)
            {
                tickerText += "Restocking has components to be put to fill one or more AM EBay orders.                    ";
            }
            if (pmEbayRestocking == true)
            {
                tickerText += "Restocking has components to be put fill to one or more PM Ebay orders.                    ";
            }
            if (yahooRestocking == true)
            {
                tickerText += "Restocking has components to be put to fill one or more Yahoo orders.                    ";
            }
            if (otherRestocking == true)
            {
                tickerText += "Restocking has components to be put to fill one or more Other orders.                    ";
            }
/*
            TimeSpan fastestPack = new TimeSpan();
            string fastestPacker = "";


            ListView myView = new ListView();
            string test = "";
            System.Windows.Forms.ListView.ListViewItemCollection itemList = new ListView.ListViewItemCollection(myView);


            packersLVI.Invoke((MethodInvoker)(()=> itemList = packersLVI.Items));
            
           //the above is acting like a reference, not a copy...





            //myView.Invoke((MethodInvoker)(()=> myView = GetListView()));            

            MessageBox.Show(test);
            
            //Invoke((MethodInvoker)(() => GetListView(out myView)));
            //WHYYYYY
 
            foreach (ListViewItem lvi in itemList)
            {
                TimeSpan ts = TimeSpan.Parse(lvi.SubItems[2].Text);
                if (ts.TotalSeconds > 0)
                {
                    if (ts.TotalSeconds < fastestPack.TotalSeconds)
                    {
                        fastestPack = ts;
                        fastestPacker = lvi.Text;
                    }
                }

            }

            tickerText += "Fastest Pack Job - " + fastestPacker + " - " + fastestPack.TotalSeconds + " seconds.";
    */       
            
            TD.Enabled = true;
            TD.HeaderText = "Fun Facts:";
            //TD.HeaderSubText = "Tot: " + otherOrderList.Count.ToString() + " Last Update: " + DateTime.Now.ToString();
            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            //TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetHeaderBackgroundColor = Color.DarkGreen;
            TD.SetHeaderTextForeColor = Color.White;
            //TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.Navy;
            TD.SetTickerForegroundColor = Color.White;
            
            //TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 45;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            TD.ShowSubHeader = false;
            TD.TickerText = tickerText;
            return TD;


        }

        public TPTickerCtl.TPTicker.TickerData totesInfoThread()
        {
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT LocationID FROM LocationContents WHERE LocationID IN (SELECT ID FROM LocationMaster WHERE WarehouseID=5 AND LocationTypeID=11)");
            List<string> partials = Connectivity.SQLCalls.sqlQuery("SELECT QuantityOrdered,QuantityBackordered,QuantityPicked,QuantityStaged,QuantityShipped,SalesOrderNo,CurrentStatus,TimeInserted FROM WhseTaskSalesOrderLines INNER JOIN WhseTaskSalesOrderHeader ON WhseTaskSalesOrderHeader.ID=HeaderID WHERE QuantityPicked > 0 AND QuantityPicked < QuantityOrdered AND CurrentStatus<6 AND CurrentHoldReason=1 AND WarehouseID=5");
            string tickerText = "";            
            TD.Enabled = true;
            TD.HeaderText = "Totes Info:";

            TD.HeaderSubText = "Tot: ";
            TD.SetEndTickerImage = Properties.Resources.tickerBackgroundEnd;
            TD.SetHeaderBackgroundImage = Properties.Resources.tickerHeaderBackground2;
            TD.SetTickerBackgroundImage = Properties.Resources.tickerBackgroundMiddle;
            TD.SetTickerBackgroundColor = Color.Orange;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 40;
            //TD.SetEndTickerColor;
            //TD.SetHeaderBackgroundColor;
            int completeableOrders = 0;
            foreach (string row in partials)
            {
                int QtyOrdered = int.Parse(row.Split('|')[0]);
                int QtyBackordered = int.Parse(row.Split('|')[1]);
                int QtyPicked = int.Parse(row.Split('|')[2]);
                int QtyStaged = int.Parse(row.Split('|')[3]);
                int QtyShipped = int.Parse(row.Split('|')[4]);
                string ItemNumber = row.Split('|')[5];
                if (QtyPicked <= QtyOrdered)
                {
                    //see if we have them in stock.. 
                    
                    List<string> itemComponents = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID,Quantity FROM InventoryComponentMap WHERE ItemNumber='" + ItemNumber + "'");
                    bool notCompletableFlag = false;
                    if (itemComponents.Count > 0)
                    {
                        foreach (string component in itemComponents)
                        {
                            int componentQty = int.Parse(component.Split('|')[1]);
                            string componentID = component.Split('|')[0];

                            //now see if we have enough of this component's qty
                            List<string> itemQty = Connectivity.SQLCalls.sqlQuery("SELECT Quantity FROM LocationContents WHERE ComponentID=" + componentID);
                            if (itemQty.Count > 0)
                            {
                                int retQty = 0;
                                foreach (string ret in itemQty)
                                {
                                    retQty += int.Parse(ret.Split('|')[0]);
                                }
                                if (retQty <= (componentQty * (QtyOrdered - QtyPicked)))
                                {
                                    notCompletableFlag = true;
                                    break;
                                }
                            }
                            if (notCompletableFlag == true)
                            {
                                break;
                            }
                            else
                            {
                                completeableOrders += 1;
                            }
                        }


                    }

                }

            }
            

            tickerText = "Totes Loaded: " + results.Count.ToString() + "    Partially Picked: " + partials.Count.ToString() + " (" + completeableOrders.ToString() + " Completable)";

            TD.ShowSubHeader = false;
            TD.TickerText = tickerText;
            return TD;


        }

        private void tickerFlipTmr_Tick(object sender, EventArgs e)
        {
            if (testTicker.Visible != false | funFactsTicker.Visible != false)
            {
                if (testTicker.Visible == false)
                {
                    funFactsTicker.SendToBack();
                    funFactsTicker.Visible = false;
                    testTicker.Visible = true;
                    testTicker.BringToFront();
                }
                else
                {
                    testTicker.SendToBack();
                    testTicker.Visible = false;
                    funFactsTicker.Visible = true;
                    funFactsTicker.BringToFront();
                }
            }
        }

        private void listViewScrollTmr_Tick(object sender, EventArgs e)
        {
            if (ListViewAutoScrollLine < packersLVI.Items.Count)
            {

                SendMessage(packersLVI.Handle, (uint)WM_VSCROLL, (System.UIntPtr)ScrollEventType.SmallIncrement, (System.IntPtr)0);
                ListViewAutoScrollLine++;
            }
            else
            {
                SendMessage(packersLVI.Handle, (uint)WM_VSCROLL, (System.UIntPtr)ScrollEventType.First, (System.IntPtr)0);
                ListViewAutoScrollLine = 0;
            }
            //packersLVI.AutoScrollOffset = new Point(packersLVI.AutoScrollOffset.X,packersLVI.AutoScrollOffset.Y + 10);
 
          
               
            
            
        }

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);
        private const int WM_VSCROLL = 0x115;


    }
}
