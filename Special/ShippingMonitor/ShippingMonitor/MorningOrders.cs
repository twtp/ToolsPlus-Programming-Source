using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;

namespace ShippingMonitor
{
    public static class MorningOrders
    {

        public static List<string> morningOrderList = new List<string>();
        public static string morningOrders;


        public static void MorningOrders_Tick(ref Label MorningOrdersLbl, ref Label morningBackLbl,ref Label afternoonBackLbl, ref Label AfternoonOrdersLbl, ref Label amEBayLbl, ref Label pmEbayLbl, int formWidth )
        {
            string test = DateTime.Now.ToString("tt");
            if (test == "AM")
                MorningOrdersLbl.BackColor = Color.LightBlue;
            morningBackLbl.BackColor = MorningOrdersLbl.BackColor;

            MorningOrdersLbl.Left -= 5;

            if (MorningOrdersLbl.Right <= 0)
            {
                MorningOrdersLbl.Left = formWidth + 200;
                LoadMorningOrders(ref afternoonBackLbl, ref AfternoonOrdersLbl, ref MorningOrdersLbl,ref morningBackLbl, ref amEBayLbl, ref pmEbayLbl, formWidth);

            }
            MorningOrdersLbl.Refresh();
        }

        public static void MorningThread()
        {
            SQLCalls newCall = new SQLCalls();
            morningOrderList = newCall.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "' AND TimeInserted < '" + DateTime.Now.ToShortDateString() + " 12:00:00.000'");           

        }
        
        public static void LoadMorningOrders(ref Label afternoonBackLbl,ref Label AfternoonOrdersLbl, ref Label MorningOrdersLbl,ref Label MorningBackLbl, ref Label amEBayLbl, ref Label pmEbayLbl, int formWidth)
        {
            string test = DateTime.Now.ToString("tt");
            if (test == "AM")
            {
                afternoonBackLbl.Visible = false;
                AfternoonOrdersLbl.Visible = false;
                pmEbayLbl.Visible = false;
            }
            else
            {

                afternoonBackLbl.Visible = true;
                AfternoonOrdersLbl.Visible = true;
                pmEbayLbl.Visible = true;

            }
            morningOrders = "";

            Thread morningThread = new Thread(new ThreadStart(MorningThread));
            morningThread.Start();
            
 //call thread here...
            //
            //
            //
            //
            while (morningThread.IsAlive ==true)
            {
                Application.DoEvents();
            }
            if (morningOrderList.Count > 0)
            {
                amEBayLbl.Visible = true;
                MorningOrdersLbl.Visible = true;
                MorningBackLbl.Visible = true;
                morningOrders = "";
                foreach (string order in morningOrderList)
                {
                    morningOrders += order.Split('|')[0] + ", ";
                }
                try
                {
                    morningOrders = morningOrders.Substring(0, morningOrders.Length - 2);
                }
                catch
                {
                    morningOrders = "";
                }
                MorningOrdersLbl.Text = morningOrders;
                amEBayLbl.Text = "AM EBay Tot: " + morningOrderList.Count.ToString() + " " + OtherOrders.GetOrderPickLocations(morningOrderList);
                amEBayLbl.Refresh();
                MorningOrdersLbl.Refresh();
            }
            else
            {
                MorningOrdersLbl.Left = formWidth + 200;
                amEBayLbl.Text = "AM EBay Tot: 0";
                amEBayLbl.Visible = false;
                MorningOrdersLbl.Visible = false;
                MorningBackLbl.Visible = false;
            }
            amEBayLbl.Text += OtherOrders.GetOrderPickLocations(morningOrderList);


        }
    }
}
