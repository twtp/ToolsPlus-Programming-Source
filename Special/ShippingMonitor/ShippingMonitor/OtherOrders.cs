using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ShippingMonitor
{
    public static class OtherOrders
    {
        public static List<string> otherOrderList = new List<string>();
        public static List<string> orderInfo = new List<string>();

        public static string otherOrders = "";


        public static void otherThread()
        {
            SQLCalls newCall = new SQLCalls();
            otherOrderList = newCall.sqlQuery("SELECT SalesOrderNo,ID FROM WhseTaskSalesOrderHeader WHERE ID IN (SELECT HeaderID FROM WhseTaskSalesOrderLines WHERE QuantityPicked < QuantityOrdered) AND SalesOrderNo NOT LIKE 'Y%' AND SalesOrderNo NOT LIKE 'E%' AND TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            otherOrders = "";
            if (otherOrderList.Count > 0)
            {
                foreach (string order in otherOrderList)
                {
                    otherOrders += order.Split('|')[0] + ",";
                }
                otherOrders = otherOrders.Substring(0, otherOrders.Length - 1);

            }
        }
        public static void GetOtherOrders(ref Label otherOrderHeader, ref Label otherOrdersLbl, ref System.Windows.Forms.Timer otherOrderTmr)
        {
            //otherOrders = "";

            Thread startOtherThread = new Thread(new ThreadStart(otherThread));
            startOtherThread.Start();
            while (startOtherThread.ThreadState != ThreadState.Running)
            {
                Application.DoEvents();
            }


            otherOrderHeader.BringToFront();
            otherOrdersLbl.Text = otherOrders;
            otherOrderHeader.Text = "Other Tot: " + otherOrderList.Count.ToString() + GetOrderPickLocations(otherOrderList);
            otherOrderTmr.Enabled = true;
        }

        public static string GetOrderPickLocations(List<string> OrdersList)
        {
           
            
            int Horizontal = 0;
            int HorizontalQty = 0;
            int Vertical = 0;
            int VerticalQty = 0;
            int FlowRack = 0;
            int FlowRackQty = 0;
            int Longs = 0;
            int LongsQty = 0;
            int Tunnel = 0;
            int TunnelQty = 0;
            int Palletized = 0;
            int PalletizedQty = 0;

            SQLCalls newCall = new SQLCalls();
            foreach (string order in OrdersList)
            {
                string OrderID = order.Split('|')[1];
                //this statement takes care of finding if location has enough qty, so iterate locations by priority high to low.
                orderInfo = newCall.sqlQuery("SELECT WhseTaskSalesOrderLines.ComponentID,LocationContents.LocationID,LocationContents.Quantity,LocationTypeID FROM WhseTaskSalesOrderLines INNER JOIN LocationContents ON LocationContents.ComponentID=WhseTaskSalesOrderLines.ComponentID INNER JOIN LocationMaster ON LocationMaster.ID=LocationContents.LocationID WHERE HeaderID=" + OrderID + " AND QuantityOrdered > QuantityPicked AND QuantityOrdered > LocationContents.Quantity");
                if (orderInfo.Count > 0)
                {
                    foreach (string component in orderInfo)
                    {
                        if (component.Split('|')[3] == "0")
                        {
                            Palletized++;
                            PalletizedQty += int.Parse(component.Split('|')[2]);
                            break;
                        }
                        if (component.Split('|')[3] == "1")
                        {
                            FlowRack++;
                            FlowRackQty += int.Parse(component.Split('|')[2]);
                            break;
                        }
                        if (component.Split('|')[3] == "2")
                        {
                            Horizontal++;
                            HorizontalQty += int.Parse(component.Split('|')[2]);
                            break;
                        }
                        if (component.Split('|')[3] == "3")
                        {
                            Vertical++;
                            VerticalQty += int.Parse(component.Split('|')[2]);
                            break;
                        }
                        if (component.Split('|')[3] == "4")
                        {
                            Longs++;
                            LongsQty += int.Parse(component.Split('|')[2]);
                            break;
                        }
                        if (component.Split('|')[3] == "16")
                        {
                            Tunnel++;
                            TunnelQty += int.Parse(component.Split('|')[2]);
                            break;
                        }


                    }
                }
            }

            string returnString = "";

            if (Horizontal > 0)
            {
                returnString += " H:" + Horizontal.ToString() + "(" + HorizontalQty.ToString() + ")";
            }
            if (Vertical > 0)
            {
                returnString += " V:" + Vertical.ToString() + "(" + VerticalQty.ToString() + ")";
            }
            if (FlowRack > 0)
            {
                returnString += " F:" + FlowRack.ToString() + "(" + FlowRackQty.ToString() + ")";
            }
            if (Longs > 0)
            {
                returnString += " L:" + Longs.ToString() + "(" + LongsQty.ToString() + ")";
            }
            if (Tunnel > 0)
            {
                returnString += " T:" + Tunnel.ToString() + "(" + TunnelQty.ToString() + ")";
            }
            if (Palletized > 0)
            {
                returnString += " P:" + Palletized.ToString() + "(" + PalletizedQty.ToString() + ")";
            } 

            return returnString;

        }
        public static void OtherOrders_Tick(ref Label otherOrdersLbl, ref Label otherOrdersBackLbl, ref Label otherOrderHeaderLbl, ref System.Windows.Forms.Timer otherOrderTmr, int formWidth)
        {
            if (otherOrdersLbl.Width > formWidth - otherOrderHeaderLbl.Width)
            {
                otherOrdersLbl.Left -= 5;

                if (otherOrdersLbl.Right <= 0)
                {
                    otherOrdersLbl.Left = formWidth + 200;

                }
            }
            else
            {
                otherOrdersLbl.Left = otherOrderHeaderLbl.Right;
            }

                otherOrdersLbl.Refresh();
            
        }
    }
}
