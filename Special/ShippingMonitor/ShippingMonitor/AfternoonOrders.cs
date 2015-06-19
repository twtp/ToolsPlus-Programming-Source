using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;

namespace ShippingMonitor
{
    public static class AfternoonOrders
    {
        public static string afternoonOrders = "";
        public static List<string> afternoonOrderList = new List<string>();

        public static void AfternoonOrdersThread()
        {
            SQLCalls newCall = new SQLCalls();
            afternoonOrderList = newCall.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + " 12:00:00.000'");
        }


        public static void LoadAfternoonOrders(ref Label AfternoonOrdersLbl, ref Label pmEbayLbl)
        {
            afternoonOrders = "";

            Thread afternoonThread = new Thread(new ThreadStart(AfternoonOrdersThread));
            afternoonThread.Start();
            while (afternoonThread.ThreadState != ThreadState.Running)
            {
                Application.DoEvents();
            }
            //THREAD HERE...
            foreach (string order in afternoonOrderList)
            {
                afternoonOrders += order.Split('|')[0] + ", ";
            }
            try
            {
                afternoonOrders = afternoonOrders.Substring(0, afternoonOrders.Length - 2);
            }
            catch
            {
                afternoonOrders = "";
            }
            AfternoonOrdersLbl.Text = afternoonOrders;
            pmEbayLbl.Text = "PM EBay Tot: " + afternoonOrderList.Count.ToString();
            pmEbayLbl.Text += OtherOrders.GetOrderPickLocations(afternoonOrderList);
            pmEbayLbl.Refresh();



            AfternoonOrdersLbl.Refresh();
        }


        



        public static void AfternoonOrders_Tick(ref Label MorningOrdersLbl,ref Label morningBackLbl, ref Label AfternoonOrdersLbl,ref Label pmEbayLbl, int formWidth)
        {
            string test = DateTime.Now.ToString("tt");

            if (test == "PM")
            {
                MorningOrdersLbl.BackColor = Color.Orange;
                morningBackLbl.BackColor = MorningOrdersLbl.BackColor;
                AfternoonOrdersLbl.Visible = true;
                pmEbayLbl.Visible = true;
                AfternoonOrdersLbl.Left -= 5;
                if (AfternoonOrdersLbl.Right <= 0)
                {
                    AfternoonOrdersLbl.Left = formWidth + 200;
                    LoadAfternoonOrders(ref AfternoonOrdersLbl,ref pmEbayLbl);

                }
                AfternoonOrdersLbl.Refresh();
            }
            else
            {
                AfternoonOrdersLbl.Left = formWidth + 200;
            }
        }

    }
}
