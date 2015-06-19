using System;
using System.Collections.Generic;
using System.Text;

namespace ShippingMonitor
{
    public static class Receiving
    {
        public static string GetTodaysPOsReceived()
        {
            SQLCalls newCall = new SQLCalls();
            List<string> TodaysPOs = newCall.sqlQuery("SELECT PONumber,POAmount,DateExported,VendorNumber FROM PurchaseOrders WHERE DateExported>'" + DateTime.Now.ToShortDateString() + "' AND ShipToAddress1='60 Scott Rd' AND Finalized=1");
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

        private static string showSinglePO(string singlePO)
        {
            decimal singlePOCost = decimal.Parse(singlePO.Split('|')[1]);
            string SinglePOCost = singlePOCost.ToString("#.##");
            if (SinglePOCost == "")
            {
                SinglePOCost = "N/A";
            }
            string SinglePOStr = "RECEIVED IN - #" + singlePO.Split('|')[0] + " $" + SinglePOCost + " " + singlePO.Split('|')[3];

            return SinglePOStr;
        }
        private static string showMultiPO(List<string> multiplePO)
        {
            decimal RunningPrice = 0;
            foreach (string po in multiplePO)
            {
                try{
                    if (po.Split('|')[1].Length > 0)
                    {
                        RunningPrice += decimal.Parse(po.Split('|')[1]);
                    }
                }
                catch{
                }
            }
            string MultiPOStr = "RECEIVED IN " + multiplePO.Count.ToString() + " POs ($" + RunningPrice.ToString("#.##") +")";
            return MultiPOStr;

        }



    }
}
