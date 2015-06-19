using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ShippingMonitor
{
    public partial class Form1 : Form
    {


        public struct FormLayout
        {
            public bool ShowTransfers;
            public bool ShowPOs;
            public bool ShowWillCalls;
            public bool ShowAMEBayOrders;
            public bool ShowPMEBayOrders;
            public bool ShowYahooOrders;
            public bool ShowOtherOrders;
            public bool ShowToteStatus;
            public bool ShowFunFacts;
        }
        public FormLayout FormLayoutInfo = new FormLayout();

        //public string MorningOrders;
        //public string AfternoonOrders;
        public Form splashScreen;

        public static void ListItemSorter(object sender, int ColumnNumber)
        {
            
            ListView list = (ListView)sender;
            int total = list.Items.Count;
            list.BeginUpdate();
            ListViewItem[] items = new ListViewItem[total];
            for (int i = 0; i < total; i++)
            {
                int count = list.Items.Count;
                int minIdx = 0;
                for (int j = 1; j < count; j++)
                    if (list.Items[j].SubItems[ColumnNumber].Text.CompareTo(list.Items[minIdx].SubItems[ColumnNumber].Text) > 0)
                        minIdx = j;
                items[i] = list.Items[minIdx];
                list.Items.RemoveAt(minIdx);
            }
            list.Items.AddRange(items);
            list.EndUpdate();
        }


        public Screen CurrentMonitor;
        public int TimerCount = 0;
        public const int UpdateInSeconds = 60;
        public const int OrdersToShow = 5;
        
        public Form1()
        {
            this.SendToBack();
            InitializeComponent();
           

        }

        public void Form1_Load(object sender, EventArgs e)
        {
            
           
        }

        public void Form1_LoadThisStuff(ref ProgressBar progress)
        {
                
            progress.Value = 5;
            progress.Refresh();
            FunFacts.newFunFacts = new FunFacts.FunFactStruct();
            progress.Value = 10;
            progress.Refresh();
            
            MorningOrders.LoadMorningOrders(ref afternoonBackLbl, ref AfternoonOrdersLbl, ref MorningOrdersLbl, ref morningBackLbl, ref amEBayLbl, ref pmEbayLbl, this.Width);
            
            progress.Value = 20;
            progress.Refresh();
            SetFormToMonitor(CurrentMonitor);
            RefreshScreen();
            this.SendToBack();
            progress.Value = 30;
            progress.Refresh();
            PackingTimes.ResetListView(ref PackersListView,this.Width);
            progress.Value = 35;
            progress.Refresh();
            PackingTimes.UpdatePackers(ref PackersListView, this.Width);
            progress.Value = 55;
            progress.Refresh();
            ListItemSorter(PackersListView, 1);
            progress.Value = 70;
            progress.Refresh();
            if (FormLayoutInfo.ShowYahooOrders == true)
            {
                YahooOrders.GetYahooOrders(ref yahooHeaderLbl, ref yahooOrdersLbl, ref YahooOrdersTmr);
                yahooOlderOrdersLbl.Text = YahooOrders.GetYahooOrdersDaysDetails();
                yahooOlderOrdersLbl.Refresh();
                
            }
            progress.Value = 80;
            progress.Refresh();
            if (FormLayoutInfo.ShowOtherOrders == true)
            {
                OtherOrders.GetOtherOrders(ref otherOrdersHeaderLbl, ref otherOrdersLbl, ref OtherOrdersTmr);
            }
            progress.Value = 100;
            progress.Refresh();
            System.Threading.Thread.Sleep(1000);
            //loadingPic.Visible = false;
            //loadingPic.SendToBack();
            timeLbl.Text = DateTime.Now.ToString("hh:mm tt");
            
        }




        private void RefreshScreen()
        {
            this.Refresh();
            TopCaptionLbl.Left = 0;
            TopCaptionLbl.Top = 0;
            TopCaptionLbl.Width = this.Width;
            TopCaptionLbl.Height = this.Height / 6;
            OrdersListView.Left = 0;
            OrdersListView.Top = TopCaptionLbl.Bottom;
            OrdersListView.Width = this.Width;
            OrdersListView.Height = ((this.Height - TopCaptionLbl.Height) / 2) - 10;
            PackersListView.Top = OrdersListView.Bottom + 25;
            PackersListView.Width = this.Width;
            PackersListView.Left = 0;
            PackersListView.Height = this.Bottom - PackersListView.Top + 25;
            MorningOrdersLbl.Top = TopCaptionLbl.Bottom;
            AfternoonOrdersLbl.Top = MorningOrdersLbl.Bottom;
           // willCallLbl.Visible = false;
            willCallLbl.Top = 0;
            willCallLbl.Left = this.Width - willCallLbl.Width;
            willCallLbl.Height = TopCaptionLbl.Height;
            restockingLbl.Top = 0;
            restockingLbl.Left = 0;
            restockingLbl.Height = TopCaptionLbl.Height;
            amEBayLbl.Left = 0;
            amEBayLbl.Top = MorningOrdersLbl.Top;
            amEBayLbl.Height = MorningOrdersLbl.Height;
            pmEbayLbl.Left = 0;
            pmEbayLbl.Top = AfternoonOrdersLbl.Top;
            pmEbayLbl.Height = AfternoonOrdersLbl.Height;
           
            morningBackLbl.Top = MorningOrdersLbl.Top;
            morningBackLbl.Left = 0;
            morningBackLbl.Width = this.Width;
            morningBackLbl.Height = MorningOrdersLbl.Height;
            morningBackLbl.SendToBack();
            morningBackLbl.BackColor = MorningOrdersLbl.BackColor;
            afternoonBackLbl.Top = AfternoonOrdersLbl.Top;
            afternoonBackLbl.Left = 0;
            afternoonBackLbl.Height = AfternoonOrdersLbl.Height;
            afternoonBackLbl.Width = this.Width;
            afternoonBackLbl.BackColor = AfternoonOrdersLbl.BackColor;
            afternoonBackLbl.SendToBack();
           
            transferLbl.Top = (int)(TopCaptionLbl.Height * 0.25f) + 10;
            transferLbl.Width = (this.Width / 2);
            transferLbl.Left = (this.Width / 2) - (transferLbl.Width / 2);
            transferLbl.Height = (int)(TopCaptionLbl.Height * 0.75f) -10;
            yahooHeaderLbl.Left = 0;
            yahooHeaderLbl.Top = pmEbayLbl.Bottom;
            yahooHeaderLbl.Height = pmEbayLbl.Height;
            yahooHeaderLbl.SendToBack();
            yahooOrdersBackLbl.Top = yahooHeaderLbl.Top;
            yahooOrdersBackLbl.Left = 0;
            yahooOrdersBackLbl.Width = this.Width;
            yahooOrdersBackLbl.Height = yahooHeaderLbl.Height;
            yahooOrdersBackLbl.SendToBack();
            yahooOrdersLbl.Top = yahooHeaderLbl.Top;
            yahooOrdersLbl.Left = 0;
            yahooOrdersLbl.Width = this.Width;
            yahooOrdersLbl.Height = yahooHeaderLbl.Height;
            yahooOrdersLbl.BringToFront();
            yahooOlderOrdersLbl.Top = yahooOrdersLbl.Bottom;
            yahooOlderOrdersLbl.Left = 0;
            yahooOlderOrdersLbl.Width = this.Width;
            yahooOlderOrdersLbl.Height = yahooOrdersLbl.Height;
            otherOrdersHeaderLbl.Left = 0;
            otherOrdersHeaderLbl.Top = yahooOrdersLbl.Bottom + 50;
            otherOrdersHeaderLbl.Height = yahooHeaderLbl.Height;
            otherOrdersHeaderLbl.SendToBack();
            otherOrdersBackLbl.Left = 0;
            otherOrdersBackLbl.Top = otherOrdersHeaderLbl.Top;
            otherOrdersBackLbl.Width = this.Width;
            otherOrdersBackLbl.Height = otherOrdersHeaderLbl.Height;
            otherOrdersBackLbl.SendToBack();
            otherOrdersLbl.Top = otherOrdersHeaderLbl.Top;
            otherOrdersLbl.Left = 0;
            otherOrdersLbl.Height = otherOrdersHeaderLbl.Height;
            otherOrdersLbl.Width = this.Width;
            totesLoadedLbl.Left = 0;
            totesLoadedLbl.Top = otherOrdersLbl.Bottom + 50;
            totesLoadedLbl.Height = otherOrdersLbl.Height;
            funFactHeaderLbl.Left = 0;
            funFactHeaderLbl.Top = totesLoadedLbl.Bottom + 25;
            funFactHeaderLbl.Height = totesLoadedLbl.Height;
            funFactHeaderLbl.SendToBack();
            funFactBackLbl.Top = funFactHeaderLbl.Top;
            funFactBackLbl.Left = 0;
            funFactBackLbl.Width = this.Width;
            funFactBackLbl.SendToBack();
            funFactLbl.Top = funFactBackLbl.Top;
            funFactLbl.Left = this.Width + 200;
            funFactLbl.Height = funFactBackLbl.Height;
            funFactLbl.BringToFront();
            funFactBackLbl.BackColor = funFactLbl.BackColor;
            leftHeaderPic.Left = 0;
            leftHeaderPic.Top = 0;
            leftHeaderPic.Height = TopCaptionLbl.Height;
            leftHeaderPic.Width = leftHeaderPic.Height;
            rightHeaderPic.Left = this.Width-(rightHeaderPic.Width*2)+28;
            rightHeaderPic.Top = 0;
            rightHeaderPic.Height = TopCaptionLbl.Height;
            rightHeaderPic.Width = rightHeaderPic.Height;
            poReceivedLbl.Top = transferLbl.Top - poReceivedLbl.Height;
            poReceivedLbl.Left = transferLbl.Left;
            poReceivedLbl.Width = transferLbl.Width;
            yahooOlderOrdersLbl.Top = yahooOlderHeaderLbl.Top;
            yahooOlderOrdersLbl.Left = 0;
            yahooOlderOrdersLbl.Height = yahooHeaderLbl.Height;
            yahooOlderOrdersLbl.BringToFront();
            yahooOlderHeaderLbl.Left = 0;
            yahooOlderHeaderLbl.Top = yahooHeaderLbl.Bottom;
            yahooOlderHeaderLbl.Height = yahooHeaderLbl.Height;
            yahooOlderHeaderLbl.BringToFront();
            yahooOlderBackLbl.Left = 0;
            yahooOlderBackLbl.Width = this.Width;
            yahooOlderBackLbl.Top = yahooOlderHeaderLbl.Top;
            timeLbl.Top = 0;
            timeLbl.Width = this.Width / 4;
            timeLbl.Left = this.Width - timeLbl.Width;
            timeLbl.Height = TopCaptionLbl.Height;
            rightHeaderPic.Left += (this.Width - rightHeaderPic.Right);
            rightHeaderPic.Refresh();
            this.Refresh();
        }

        public void SetFormToMonitor(Screen TheMonitor)
        {          
                Point ScreenCoords = new Point();
                ScreenCoords.X = TheMonitor.WorkingArea.Left;
                ScreenCoords.Y = TheMonitor.WorkingArea.Top;
                this.Left = ScreenCoords.X;
                this.Top = ScreenCoords.Y;
                this.ResizeRedraw = true;
                this.WindowState = FormWindowState.Maximized;
                this.Refresh();
                EnumerateOrdersToShip();
        }

        public List<Screen> GetMonitors()
        {
            List<Screen> ScreenList = new List<Screen>();
            try
            {                
                foreach (Screen screen in Screen.AllScreens)
                {
                    ScreenList.Add(screen);
                }
            }
            catch
            {
            }

            return ScreenList;
        }

        private void OrderCheckTmr_Tick(object sender, EventArgs e)
        {           
            TimerCount++;
            if (TimerCount >= UpdateInSeconds)
            {
                TimerCount = 0;
                //EnumerateOrdersToShip();
                PackingTimes.UpdatePackers(ref PackersListView, this.Width);
                ListItemSorter(PackersListView, 1);
            }

        }
        private void EnumeratePackers(ref ListView listView1, object form1)
        {
            PackingTimes.ResetListView(ref listView1,this.Width);

        }
        private void EnumerateOrdersToShip()
        {
            ResetListView();
            
           // SQLCalls totedOrdersToShip = new SQLCalls();
            //List<string> TotedOrders = totedOrdersToShip.sqlQuery("SELECT WhseTaskSalesOrderHeader.SalesOrderNo, WhseTaskSalesOrderHeader.ID, WhseTaskSalesOrderHeader.CurrentStatus, WhseTaskSalesOrderHeader.CurrentHoldReason, WhseTaskSalesOrderHeader.TimeInserted FROM WhseTaskSalesOrderHeader WHERE CurrentStatus=3 AND CurrentHoldReason=1 ORDER BY WhseTaskSalesOrderHeader.TimeInserted");
            //ProcessOrdersToShip(TotedOrders);

        }

        private void ProcessOrdersToShip(List<string> OrderList)
        {
            int OrderCount = OrderList.Count;

            for (int xCount = 0; xCount < OrderCount; xCount++)
            {
                if (GetSalesOrderDetails(OrderList[xCount].Split('|')[1]) == true)
                {
                    ListViewItem orderedShipment = new ListViewItem();
                    orderedShipment.Text = OrderList[xCount].Split('|')[0];

                    ListViewItem.ListViewSubItem orderedShipmentDate = new ListViewItem.ListViewSubItem();
                    orderedShipmentDate.Text = OrderList[xCount].Split('|')[4].Split(' ')[0];

                    TimeSpan ColorOffset = DateTime.Now - DateTime.Parse(orderedShipmentDate.Text);

                    //MessageBox.Show("Color Offset Is: " + ColorOffset.Hours.ToString());
                    int theColor = ColorOffset.Days * 50;
                    if (theColor > 205)
                    {
                        theColor = 205;
                    }
                    if (theColor < 0)
                    {
                        theColor =0;
                    }

                    orderedShipment.BackColor = Color.FromArgb(50+theColor, 50, 50);
                    orderedShipment.ForeColor = Color.White;
                    ListViewItem.ListViewSubItem orderedQueueTime = new ListViewItem.ListViewSubItem();
                    orderedQueueTime.Text = "N/A";

                    orderedShipment.SubItems.Add(orderedShipmentDate);
                    orderedShipment.SubItems.Add(orderedQueueTime);
                    OrdersListView.Items.Add(orderedShipment);
                }

            }
           // MessageBox.Show("Total Orders: " + OrderCount.ToString());
        }

        private void ResetListView()
        {
            OrdersListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.None);
            OrdersListView.Items.Clear();
            OrdersListView.Columns.Clear();
            
            ColumnHeader orderNumberHeader = new ColumnHeader();
            orderNumberHeader.Text = "Order #";
            orderNumberHeader.TextAlign = HorizontalAlignment.Center;
            orderNumberHeader.Width = this.Width / 3;
            ColumnHeader dateOrderedHeader = new ColumnHeader();
            dateOrderedHeader.Text = "Date Ordered";
            dateOrderedHeader.TextAlign = HorizontalAlignment.Center;
            dateOrderedHeader.Width = this.Width / 3;
            ColumnHeader queueTimeHeader = new ColumnHeader();
            queueTimeHeader.Text = "Queued Time";
            queueTimeHeader.TextAlign = HorizontalAlignment.Center;
            queueTimeHeader.Width = this.Width / 3;

            OrdersListView.Columns.Add(orderNumberHeader);
            OrdersListView.Columns.Add(dateOrderedHeader);
            OrdersListView.Columns.Add(queueTimeHeader);


        }
        

        private bool GetSalesOrderDetails(string SO_HeaderID)
        {
            SQLCalls getDetails = new SQLCalls();
            List<string> details = getDetails.sqlQuery("SELECT ComponentID,QuantityOrdered,QuantityBackordered,QuantityPicked,QuantityStaged,QuantityShipped FROM WhseTaskSalesOrderLines WHERE HeaderID=" + SO_HeaderID);
            if (details.Count == 0)
            {
                return false;
            }
            return IsSalesOrderQueued(details);

        }
        private bool IsSalesOrderQueued(List<string> SalesOrderInfo)
        {
            int QuantityOrdered = 0;
            int QuantityBackordered = 0;
            int QuantityPicked = 0;
            int QuantityStaged = 0;
            int QuantityShipped = 0;

            foreach (string OrderLine in SalesOrderInfo)
            {
                QuantityOrdered = int.Parse(OrderLine.Split('|')[1]);
                QuantityBackordered = int.Parse(OrderLine.Split('|')[2]);
                QuantityPicked = int.Parse(OrderLine.Split('|')[3]);
                QuantityStaged = int.Parse(OrderLine.Split('|')[4]);
                QuantityShipped = int.Parse(OrderLine.Split('|')[5]);
                if (QuantityOrdered > QuantityShipped)
                {
                    return true;
                }
            }

            return false;

            
        }

        private void MorningOrdersTmr_Tick(object sender, EventArgs e)
        {
            MorningOrders.MorningOrders_Tick(ref MorningOrdersLbl,ref morningBackLbl,ref afternoonBackLbl,ref AfternoonOrdersLbl,ref amEBayLbl,ref pmEbayLbl,this.Width);


        }
        
       
        
        private void AfternoonOrdersTmr_Tick(object sender, EventArgs e)
        {
            AfternoonOrders.AfternoonOrders_Tick(ref MorningOrdersLbl, ref morningBackLbl, ref AfternoonOrdersLbl, ref pmEbayLbl, this.Width);
            

        }

        private void WillCallOrdersTmr_Tick(object sender, EventArgs e)
        {
            WillCall.LoadWillCalls(ref willCallLbl, ref willCallFlasherTmr);
            OtherOrders.GetOtherOrders(ref otherOrdersHeaderLbl, ref otherOrdersLbl, ref OtherOrdersTmr);
        }

        private void restockingStagingTmr_Tick(object sender, EventArgs e)
        {
            yahooOlderOrdersLbl.Text = YahooOrders.GetYahooOrdersDaysDetails();
            yahooOlderOrdersLbl.Refresh();
            LoadRestockingStaging();
            if (FormLayoutInfo.ShowTransfers == true)
            {
                Transfer.CheckForTransfer(ref this.transferLbl);
            }
            if (FormLayoutInfo.ShowPOs == true)
            {
                poReceivedLbl.Text = Receiving.GetTodaysPOsReceived();
            }
            if (poReceivedLbl.Text.Length > 0)
            {
                poReceivedLbl.Visible = true;
            }
            else
            {
                poReceivedLbl.Visible = false;
            }

        }
        private void LoadRestockingStaging()
        {
            SQLCalls newCall = new SQLCalls();
            List<string> results = newCall.sqlQuery("SELECT ComponentID,Quantity FROM LocationContents WHERE LocationID=10830");
            if (results.Count > 0)
            {
                int totComponents = results.Count;
                int totQty = 0;
                foreach (string str in results)
                {
                    totQty += int.Parse(str.Split('|')[1]);
                }

                restockingLbl.Text = "RESTOCKING\r\nItems: " + totComponents.ToString();
                restockingLbl.Text += "\r\nQty: " + totQty.ToString();
                restockingLbl.Visible = true;
                restockingLbl.BringToFront();
            }
            else
            {
                restockingLbl.Visible = false;
            }
        }

        private void totesTmr_Tick(object sender, EventArgs e)
        {
            timeLbl.Text = DateTime.Now.ToString("hh:mm tt");

            if (FormLayoutInfo.ShowFunFacts == true)
            {
                FunFacts.LoadFunFacts(ref PackersListView, ref funFactLbl, ref funFactTmr);
            }
            if (FormLayoutInfo.ShowToteStatus == true)
            {
                ToteInfo.LoadTotesLoaded(ref totesLoadedLbl);
            }
        }
        


        



        private void funFactTmr_Tick(object sender, EventArgs e)
        {
            
            FunFacts.FunFactTimer_Tick(ref funFactHeaderLbl, ref funFactLbl, this.Width);
        }

        private void willCallFlasherTmr_Tick(object sender, EventArgs e)
        {
            WillCall.willCallFlasherTmr_Tick(ref willCallFlasherTmr, ref willCallLbl);
        }

        private void YahooOrdersTmr_Tick(object sender, EventArgs e)
        {
            YahooOrders.YahooOrders_Tick(ref yahooOrdersLbl, ref yahooOrdersBackLbl, ref yahooHeaderLbl, ref YahooOrdersTmr, this.Width);
            YahooOrders.YahooOlderOrders_Tick(ref yahooOlderOrdersLbl, ref yahooOlderHeaderLbl, ref yahooOlderBackLbl, this.Width);
            
        }

        private void OtherOrdersTmr_Tick(object sender, EventArgs e)
        {
            OtherOrders.OtherOrders_Tick(ref otherOrdersLbl, ref otherOrdersBackLbl, ref otherOrdersHeaderLbl, ref OtherOrdersTmr, this.Width);
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {

        }

        private void StartupTmr_Tick(object sender, EventArgs e)
        {
      
        }

        private void Form1_Shown(object sender, EventArgs e)
        {


            //Form1_LoadThisStuff();
            

        }

    }
}
