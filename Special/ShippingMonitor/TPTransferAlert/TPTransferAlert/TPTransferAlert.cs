using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TPTransferAlert
{
    public partial class TPTransferAlert : UserControl
    {
        public TPTransferAlert()
        {
            InitializeComponent();
            this.SizeChanged += new EventHandler(TPTransferAlert_SizeChanged);
        }

        void TPTransferAlert_SizeChanged(object sender, EventArgs e)
        {
            refreshControl();
        }

        private void TPTransferAlert_Load(object sender, EventArgs e)
        {
            refreshControl();
            checkForTransfersThread();
        }

        private void refreshControl()
        {
            transferHeaderLbl.Left = 0;
            transferHeaderLbl.Top = 0;
            //transferHeaderLbl.Width = this.Width/5;
            //transferHeaderLbl.Height = this.Height;

            transferDetailsLbl.Left = transferHeaderLbl.Right;
            transferDetailsLbl.Top = transferHeaderLbl.Top;
            transferHeaderLbl.Height = transferHeaderLbl.Height;
            transferHeaderLbl.Width = this.Width - transferHeaderLbl.Width;
        }

        private void checkForTransfersThread()
        {
            int componentCount = 0;
            int totQty = 0;
            int totPicked = 0;
            List<string> transfers = TransferConnectivity.SQLCalls.sqlQuery("SELECT ID,Description,TimeInserted,TransferNumber FROM WhseTaskTransferHeader WHERE IsComplete=0 AND TimeInserted>'" + DateTime.Now.ToShortDateString() + "'");
            if (transfers.Count > 0)
            {
                if (transfers.Count == 1)
                {
                    List<string> trans = TransferConnectivity.SQLCalls.sqlQuery("SELECT ComponentID,QuantityRequested,QuantityPicked FROM WhseTaskTransferLines WHERE HeaderID=" + transfers[0].Split('|')[0]);
                    if (trans.Count > 0)
                    {

                        componentCount = trans.Count;
                        foreach (string t in trans)
                        {
                            totQty += int.Parse(t.Split('|')[1]);
                            totPicked += int.Parse(t.Split('|')[2]);
                        }
                        transferDetailsLbl.Text = "#" + transfers[0].Split('|')[3] + "- Tot Comp: " + componentCount.ToString() + " - Tot Qty: " + totQty.ToString() + "\r\n" + transfers[0].Split('|')[1];
                        transferDetailsLbl.Text += "Total Picked: " + totPicked.ToString() + "/" + totQty.ToString();

                    }
                    else
                    {
                        //an empty transfer?? wtf??
                        transferDetailsLbl.Text = "Error obtaining transfer data.";

                    }

                    
                }
                else
                {
                    List<string> tx = TransferConnectivity.SQLCalls.sqlQuery("SELECT ID, TimeInserted, TransferNumber FROM WhseTaskTransferHeader WHERE IsComplete=0 AND TimeInserted>'" + DateTime.Now.ToShortDateString() + "'");
                    if (tx.Count > 0)
                    {
                        int totComponents = 0;
                        int totalQty = 0;
                        string transfersList = "";
                        foreach (string t in tx)
                        {
                            transfersList += "#" + t.Split('|')[2] + ", ";
                            List<string> transferData = TransferConnectivity.SQLCalls.sqlQuery("SELECT QuantityRequested, QuantityPicked FROM WhseTaskTransferLines WHERE HeaderID=" + t.Split('|')[0]);
                            totComponents += transferData.Count;
                            if (transferData.Count > 0)
                            {
                                foreach (string dataLine in transferData)
                                {
                                    totalQty += int.Parse(dataLine.Split('|')[0]);
                                }
                            }
                        }
                        transfersList = transfersList.Substring(0, transfersList.Length - 2);
                        transferHeaderLbl.Text = "Transfers (" + tx.Count.ToString() + ")";
                        transferDetailsLbl.Text = transfersList  + "\r\nComponents: " + totComponents.ToString() + "  -  Qty: " + totalQty.ToString();
                    }
                    else
                    {
                        transferDetailsLbl.Text = "Transfer: Failed to pull multiple transfer data.";
                    }
                }
                this.Visible = true;
            }
            else
            {
                this.Visible = false;
            }
        }




        private void transferHeaderLbl_Click(object sender, EventArgs e)
        {

        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
            checkForTransfersThread();
        }

        private void transferHeaderLbl_Click_1(object sender, EventArgs e)
        {

        }
    }
}
