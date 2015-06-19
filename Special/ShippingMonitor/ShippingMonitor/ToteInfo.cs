using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ShippingMonitor
{
    public static class ToteInfo
    {
        public static List<string> results = new List<string>();
        public static List<string> partials = new List<string>();

        public static void totesThread()
        {
            SQLCalls newCall = new SQLCalls();
            results = newCall.sqlQuery("SELECT DISTINCT LocationID FROM LocationContents WHERE LocationID IN (SELECT ID FROM LocationMaster WHERE WarehouseID=5 AND LocationTypeID=11)");
            partials = newCall.sqlQuery("SELECT QuantityOrdered,QuantityBackordered,QuantityPicked,QuantityStaged,QuantityShipped,SalesOrderNo,CurrentStatus,TimeInserted FROM WhseTaskSalesOrderLines INNER JOIN WhseTaskSalesOrderHeader ON WhseTaskSalesOrderHeader.ID=HeaderID WHERE QuantityPicked > 0 AND QuantityPicked < QuantityOrdered AND CurrentStatus<6 AND CurrentHoldReason=1 AND WarehouseID=5");

        }
        
        public static void LoadTotesLoaded(ref Label totesLoadedLbl)
        {
            //THREAD HERE
            Thread totesThreading = new Thread(new ThreadStart(totesThread));
            totesThreading.Start();
            while (totesThreading.ThreadState != ThreadState.Running)
            {
                Application.DoEvents();
            }
           totesLoadedLbl.Text = "Totes Loaded: " + results.Count.ToString() + "    Partially Picked: " + partials.Count.ToString() + " (" + CompletableOrders(partials).ToString() + " Completable)";
        }
        private static int CompletableOrders(List<string> dbResults)
        {
            int completeableOrders = 0;
            foreach (string row in dbResults)
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
                    SQLCalls newCall = new SQLCalls();
                    List<string> itemComponents = newCall.sqlQuery("SELECT ComponentID,Quantity FROM InventoryComponentMap WHERE ItemNumber='" + ItemNumber + "'");
                    bool notCompletableFlag = false;
                    if (itemComponents.Count > 0)
                    {
                        foreach (string component in itemComponents)
                        {
                            int componentQty = int.Parse(component.Split('|')[1]);
                            string componentID = component.Split('|')[0];

                            //now see if we have enough of this component's qty
                            List<string> itemQty = newCall.sqlQuery("SELECT Quantity FROM LocationContents WHERE ComponentID=" + componentID);
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

            return completeableOrders;
        }

    }
}
