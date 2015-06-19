using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TPPOReceivedAlert
{
    public partial class TPPOReceivedAlert : UserControl
    {
        public TPPOReceivedAlert()
        {
            InitializeComponent();
            this.SizeChanged += new EventHandler(TPPOReceivedAlert_SizeChanged);
        }

        void TPPOReceivedAlert_SizeChanged(object sender, EventArgs e)
        {
            RefreshControl();
        }

        private void TPPOReceivedAlert_Load(object sender, EventArgs e)
        {
            RefreshControl();
            poReceivedThread();
        }
        private void RefreshControl()
        {
            poHeaderLbl.Top = 0;
            poHeaderLbl.Height = this.Height;
           
            poReceivedHeaderLbl.Top = 0;
            poReceivedHeaderLbl.Left = poHeaderLbl.Right;
            poReceivedHeaderLbl.Width = this.Width;
            poReceivedHeaderLbl.Height = this.Height;
        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
            poReceivedThread();
        }
        private void poReceivedThread()
        {
            //poReceivedHeaderLbl.Text = "PO Received: "
            List<string> receivedPOs = POConnectivity.SQLCalls.sqlQuery("SELECT ID,PurchaseOrderNo FROM HandheldReceipt WHERE ReceivedTime > '" + DateTime.Now.AddMinutes(-10).ToString() + "'");
            string PONumber = "";
            int TotComponentsReceived = 0;
            int TotQtyReceived = 0;

            if (receivedPOs.Count > 0)
            {
                if (receivedPOs.Count == 1)
                {
                    PONumber = receivedPOs[0].Split('|')[1];
                    List<string> componentDetails = POConnectivity.SQLCalls.sqlQuery("SELECT ReceivedQty FROM HandheldReceiptLine WHERE HeaderID=" + receivedPOs[0].Split('|')[0]);
                    if (componentDetails.Count > 0)
                    {
                        TotComponentsReceived = componentDetails.Count;
                        foreach (string c in componentDetails)
                        {
                            TotQtyReceived += int.Parse(c.Split('|')[0]);
                        }
                    }
                    else
                    {
                        PONumber = "Cannot find detailed info..";
                    }
                    
                    poReceivedHeaderLbl.Text = "PO Received: #" + PONumber + " -\r\nComponents: " + TotComponentsReceived.ToString() + " - Total Qty: " + TotQtyReceived.ToString();
                }
                else
                {
                    //multiple POs..
                    List<string> multiplePOs = POConnectivity.SQLCalls.sqlQuery("SELECT PurchaseOrderNo From HandheldReceipt WHERE ReceivedTime > '" + DateTime.Now.AddMinutes(-10).ToString() + "'");
                    string poList = "";
                    foreach (string po in multiplePOs)
                    {
                        poList += "#" + po.Split('|')[0] + ", ";
                    }
                    poList = poList.Substring(0, poList.Length - 2);
                    poReceivedHeaderLbl.Text = "POs Received (" + multiplePOs.Count.ToString() + "): " + poList;
                }

                this.Visible = true;
            }
            else
            {
                this.Visible = false;
            }
            


        }

        

    }
}
