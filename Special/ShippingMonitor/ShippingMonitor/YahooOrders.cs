using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ShippingMonitor
{
    public static class YahooOrders
    {
        public static List<string> yahooOrderList = new List<string>();


        public static string yahooOrders = "";

        public static string yahooOrderDayDetails = "";

        public static void YahooThread()
        {

            SQLCalls newCall = new SQLCalls();
            yahooOrderList = newCall.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'Y%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            if (yahooOrderList.Count > 0)
            {
                yahooOrders = "";
                foreach (string order in yahooOrderList)
                {
                    yahooOrders += order.Split('|')[0] + ",";
                }
                yahooOrders = yahooOrders.Substring(0, yahooOrders.Length - 1);

            }
        }
        
        public static void GetYahooOrders(ref Label yahooOrderHeader,ref Label yahooOrdersLbl,ref System.Windows.Forms.Timer yahooOrderTmr)
        {
            
            
            
            //THREAD HERE
            Thread yahooThread = new Thread(new ThreadStart(YahooThread));
            yahooThread.Start();
            
            while (yahooThread.ThreadState != ThreadState.Running)
            {
                Application.DoEvents();
            }
           
            yahooOrderHeader.BringToFront();
            yahooOrdersLbl.Text = yahooOrders;
            yahooOrderHeader.Text = "Yahoo Tot: " + yahooOrderList.Count.ToString() + OtherOrders.GetOrderPickLocations(yahooOrderList);
            yahooOrderTmr.Enabled = true;
        }

        public static string GetYahooOrdersDaysDetails()
        {

            
            Thread orderXDaysAgo = new Thread(new ThreadStart(GetYahooOrdersThreadFromXDaysThread));
            orderXDaysAgo.Start();
            while (orderXDaysAgo.ThreadState != ThreadState.Running)
            {
                Application.DoEvents();
            }
            return yahooOrderDayDetails;
        }
        public static void GetYahooOrdersThreadFromXDaysThread()
        {
           List<string> SalesOrders =  Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo FROM WhseTaskSalesOrderHeader WHERE TimeInserted<'" + DateTime.Now.AddDays(-1).ToShortDateString() + "' AND SalesOrderNo LIKE 'Y%' AND CurrentHoldReason=1 AND CurrentStatus < 4 ORDER BY TimeInserted ASC");
           //yahooOrderDayDetails = "1 day old: " + SalesOrders.Count.ToString();
           //SalesOrders = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo FROM WhseTaskSalesOrderHeader WHERE TimeInserted>'" + DateTime.Now.AddDays(-2).ToShortDateString() + "' AND TimeInserted <'" + DateTime.Now.AddDays(-2).ToShortDateString() + " 23:59:00.000' AND SalesOrderNo LIKE 'Y%'");
           //yahooOrderDayDetails += "2 day old: " + SalesOrders.Count.ToString();
           //SalesOrders = Connectivity.SQLCalls.sqlQuery("SELECT SalesOrderNo FROM WhseTaskSalesOrderHeader WHERE TimeInserted<'" + DateTime.Now.AddDays(-3).ToShortDateString() + "' AND SalesOrderNo LIKE 'Y%'");
           //yahooOrderDayDetails += "older: " + SalesOrders.Count.ToString();
           yahooOrderDayDetails = "";
            foreach(string so in SalesOrders)
            {
                yahooOrderDayDetails += so.Split('|')[0] + ","; 
            }
        }
        public static void YahooOlderOrders_Tick(ref Label yahooOrderDayLbl, ref Label yahooOrderDayHeaderLbl, ref Label yahooOrderBackLbl, int formWidth)
        {
            yahooOrderDayLbl.Left -= 5;
            if (yahooOrderDayLbl.Right <= 0)
            {
                yahooOrderDayLbl.Left = formWidth + 200;

            }
            yahooOrderDayLbl.Refresh();
        }

        public static void YahooOrders_Tick(ref Label yahooOrdersLbl, ref Label yahooOrdersBackLbl, ref Label yahooOrderHeaderLbl,ref System.Windows.Forms.Timer yahooOrderTmr, int formWidth)
        {
            
            yahooOrdersLbl.Left -= 5;
           
            if (yahooOrdersLbl.Right <= 0)
            {
                yahooOrdersLbl.Left = formWidth + 200;
                GetYahooOrders(ref yahooOrderHeaderLbl, ref yahooOrdersLbl,ref yahooOrderTmr);
            }
            yahooOrdersLbl.Refresh();
        }

    }

}
