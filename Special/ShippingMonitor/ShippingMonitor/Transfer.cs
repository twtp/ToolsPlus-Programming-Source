using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ShippingMonitor
{
    public static class Transfer
    {
        public static void CheckForTransfer(ref Label transferLbl)
        {
            List<string> transfers = new List<string>();
            SQLCalls newCall = new SQLCalls();
            transfers = newCall.sqlQuery("SELECT TransferNumber FROM WhseTaskTransferHeader WHERE WarehouseID=5 AND IsComplete=0");
            if (transfers.Count == 1)
            {
                transferLbl.Visible = true;
                transferLbl.BringToFront();
                transferLbl.Text = "TRANSFER - " + transfers[0].Split('|')[0];                
            }
            else if (transfers.Count > 1)
            {               
                transferLbl.Visible = true;
                transferLbl.BringToFront();
                transferLbl.Text = "TRANSFERS (" + transfers.Count.ToString() + ") - \r\n";
                foreach (string transfer in transfers)
                {
                    transferLbl.Text += transfer.Split('|')[0] + ",";
                }
                transferLbl.Text = transferLbl.Text.Substring(0, transferLbl.Text.Length - 1);
            }
            else
            {
                transferLbl.Visible = false;
            }
            
        }
    }
}
