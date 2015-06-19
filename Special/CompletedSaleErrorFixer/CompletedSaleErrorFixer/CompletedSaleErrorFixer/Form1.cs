using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CompletedSaleErrorFixer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            pictureBox1.Refresh();
            setupListView();

            if (textBox1.Text.Contains("-") == true)
            {
                textBox1.Text = textBox1.Text.Split('-')[1];
            }
            string SONumber = Step0_GetSONumber(textBox1.Text);
            if (SONumber.Length > 0)
            {
                string HeaderID = Step1_DoesOrderExist(SONumber);
                if (HeaderID.Length > 0)
                {
                    string LineID = Step2_VerifyNotificationFlagIsOff(HeaderID);
                    if (LineID.Length > 0)
                    {
                        bool finalres = Step3_SetCustomerNotified(HeaderID, LineID);
                        if (finalres == true)
                        {
                            MessageBox.Show("Order #" + SONumber + " has been fixed!");
                        }
                        else
                        {
                            MessageBox.Show("Did not run successfully");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Did not run successfully");
                    }
                }
                else
                {
                    MessageBox.Show("Did not run successfully.");
                }
            }
            else
            {
                MessageBox.Show("Could not interpret Transaction ID.");
            }
            pictureBox1.Visible = false;

        }

        private void setupListView()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();

            ColumnHeader step = new ColumnHeader();
            ColumnHeader status = new ColumnHeader();
            ColumnHeader details = new ColumnHeader();
            step.Text = "Step";
            status.Text = "Status";
            details.Text = "Details";
            step.Width = listView1.Width / 3 - 5;
            status.Width = listView1.Width / 3 - 5;
            details.Width = listView1.Width / 3 - 5;
            

            listView1.Columns.Add(step);
            listView1.Columns.Add(status);
            listView1.Columns.Add(details);

            listView1.Refresh();



        }

        private string Step0_GetSONumber(string transactionID)
        {
            ListViewItem soNumber = new ListViewItem();
            soNumber.Text = "1. Obtain SO";
            List<string> soNum = CompletedSaleErrorFixer.SQLCalls.sqlQuery("SELECT SalesOrderNo FROM EBayTransactionLog WHERE TransactionID='" + transactionID + "'");
            if (soNum.Count > 0)
            {
                soNumber.SubItems.Add("YES");
                soNumber.SubItems.Add("SO: " + soNum[0].Split('|')[0]);
                listView1.Items.Add(soNumber);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return soNum[0].Split('|')[0];
                
            }
            else
            {
                soNumber.SubItems.Add("NO");
                soNumber.SubItems.Add("Couldn't find SO #");
                listView1.Items.Add(soNumber);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return "";
            }
        }

        private string Step1_DoesOrderExist(string orderNo)
        {
            ListViewItem existsLVI = new ListViewItem();
            existsLVI.Text = "2. Order Exists";
            List<string> exists = CompletedSaleErrorFixer.SQLCalls.sqlQuery("SELECT ID,VendorNumber FROM PurchaseOrders WHERE SalesOrderNo='" + orderNo + "'");
            if (exists.Count > 0)
            {
                              
                existsLVI.SubItems.Add("YES");
                existsLVI.SubItems.Add("Vendor: " + exists[0].Split('|')[1]);
                listView1.Items.Add(existsLVI);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return exists[0].Split('|')[0];
            }
            else
            {
                existsLVI.SubItems.Add("NO");
                listView1.Items.Add(existsLVI);
                return "";
            }
            

        }

        private string Step2_VerifyNotificationFlagIsOff(string headerID)
        {
            ListViewItem verifyLVI = new ListViewItem();
            verifyLVI.Text = "3. Verify Issue";
            List<string> verify = CompletedSaleErrorFixer.SQLCalls.sqlQuery("SELECT ID,CustomerNotified FROM PurchaseOrderTracking WHERE HeaderID=" + headerID + " AND CustomerNotified=0");
            if (verify.Count > 0)
            {
                if (verify[0].Split('|')[1] == "False")
                {
                    verifyLVI.SubItems.Add("YES");
                    verifyLVI.SubItems.Add("Line ID: " + verify[0].Split('|')[0]);
                    listView1.Items.Add(verifyLVI);
                    listView1.Refresh();
                    System.Threading.Thread.Sleep(2000);
                    return verify[0].Split('|')[0];
                }
                else
                {
                    verifyLVI.SubItems.Add("NO");
                    listView1.Items.Add(verifyLVI);
                    listView1.Refresh();
                    System.Threading.Thread.Sleep(2000);
                    return "";
                }
            }
            else
            {
                MessageBox.Show("Could not find tracking info for headerID #" + headerID);
                verifyLVI.SubItems.Add("NO");
                listView1.Items.Add(verifyLVI);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return "";
            }
        }

        private bool Step3_SetCustomerNotified(string headerID, string lineID)
        {

            ListViewItem notified = new ListViewItem();
            
            notified.Text = "4. Fixing Issue";
            try
            {
                List<string> notify = CompletedSaleErrorFixer.SQLCalls.sqlQuery("UPDATE PurchaseOrderTracking SET CustomerNotified=1 WHERE HeaderID=" + headerID + " AND ID=" + lineID);
                notified.SubItems.Add("YES");
                notified.SubItems.Add("Result: Problem Fixed");
                listView1.Items.Add(notified);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return true;
            }
            catch
            {
                notified.SubItems.Add("NO");
                notified.SubItems.Add("Result: Not Fixed");
                listView1.Items.Add(notified);
                listView1.Refresh();
                System.Threading.Thread.Sleep(2000);
                return false;
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            // pictureBox1.Image = Image.FromFile("C:\\users\\tomwestbrook\\documents\\toolsplus programming source\\picture repo\\blackloading.gif");
        
            setupListView();
            SetInfoLabel();


        }
        private void SetInfoLabel()
        {
            string info = "";
            info = "  This program should only be used when you get an 'ebay completesale api error' and it's description being 'Tracking number(s) in the request already exists either for a different buyer or for a different seller'.";
            info += "If you do have this error, copy the OrderID number from the XML response in email to the below input field. This should be about 25 chars long with a dash in the middle however only the right side of the dash is actually used.";

            infoLbl.Text = info;
        }

        private void infoLbl_Click(object sender, EventArgs e)
        {

        }


    }
}
