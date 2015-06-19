using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        public const string VERSION = "Version: 1.0.1";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SQLCalls sqlCall = new SQLCalls();
            List<string> results = sqlCall.sqlQuery("SELECT SalesOrderNo,BuyerID,ItemID,Quantity,TimeInserted FROM EBayTransactionLog WHERE TransactionID='" + TXIDTxt.Text + "'");
            resultsTxt.Text = "";
            if (results.Count > 0)
            {
                //get the ItemNumber from EBayItemID...
                List<string> results2 = sqlCall.sqlQuery("SELECT ItemNumber FROM PartNumbers WHERE EBayItemID=" + results[0].Split('|')[2]);
                
                //get the UnitPrice of current sale...
                List<string> unitPrice = sqlCall.sqlQuery("SELECT ID,ProductCode,UnitPrice FROM SORealTimeImportOrders WHERE OrderNumber='" + results[0].Split('|')[0] + "' AND LineType='ITEM'");

              

                resultsTxt.Text = "Sales Order #" + results[0].Split('|')[0] + "\r\n";
                resultsTxt.Text += "Buyer ID: " + results[0].Split('|')[1] + "\r\n";
                resultsTxt.Text += "-----------------------------" + "\r\n";


                if (unitPrice.Count > 0)
                {

                    foreach (string units in unitPrice)
                    {
                        resultsTxt.Text += "Item Number: " + results2[0].Split('|')[0] + "\r\n";
                        resultsTxt.Text += "Quantity: " + results[0].Split('|')[3] + "\r\n";
                        resultsTxt.Text += "Unit Price: " + units.Split('|')[2] + "\r\n";
                        resultsTxt.Text += "-----------------------------" + "\r\n";

                    }
                }
                else
                {
                    resultsTxt.Text += "No items found on current order.\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                }

                List<string> getClientData = sqlCall.sqlQuery("SELECT * FROM SORealTimeImportCustomers WHERE OrderNumber='" + results[0].Split('|')[0] + "'");
                if (getClientData.Count > 0)
                {
                    resultsTxt.Text += "Customer ID: " + getClientData[0].Split('|')[2] + "\r\n";
                    resultsTxt.Text += "Order Date: " + getClientData[0].Split('|')[3] + "\r\n";
                    resultsTxt.Text += "Order Status: " + getClientData[0].Split('|')[5] + "\r\n";
                    resultsTxt.Text += "Hold Reason: " + getClientData[0].Split('|')[6] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "Billing Name: " + getClientData[0].Split('|')[8] + "\r\n";
                    resultsTxt.Text += "Bill First Name: " + getClientData[0].Split('|')[9] + "\r\n";
                    resultsTxt.Text += "Bill Last Name: " + getClientData[0].Split('|')[10] + "\r\n";
                    resultsTxt.Text += "Bill Company: " + getClientData[0].Split('|')[11] + "\r\n";
                    resultsTxt.Text += "Bill Address 1: " + getClientData[0].Split('|')[12] + "\r\n";
                    resultsTxt.Text += "Bill Address 2: " + getClientData[0].Split('|')[13] + "\r\n";
                    resultsTxt.Text += "Billing City: " + getClientData[0].Split('|')[14] + "\r\n";
                    resultsTxt.Text += "Billing State: " + getClientData[0].Split('|')[15] + "\r\n";
                    resultsTxt.Text += "Billing Country: " + getClientData[0].Split('|')[16] + "\r\n";
                    resultsTxt.Text += "Billing Zip: " + getClientData[0].Split('|')[17] + "\r\n";
                    resultsTxt.Text += "Billing Phone: " + getClientData[0].Split('|')[18] + "\r\n";
                    resultsTxt.Text += "Billing Address Verification: " + getClientData[0].Split('|')[19] + "\r\n";
                    resultsTxt.Text += "Billing Is Residential: " + getClientData[0].Split('|')[20] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "ShipTo Name: " + getClientData[0].Split('|')[21] + "\r\n";
                    resultsTxt.Text += "ShipTo First Name: " + getClientData[0].Split('|')[22] + "\r\n";
                    resultsTxt.Text += "ShipTo Last Name: " + getClientData[0].Split('|')[23] + "\r\n";
                    resultsTxt.Text += "ShipTo Company: " + getClientData[0].Split('|')[24] + "\r\n";
                    resultsTxt.Text += "ShipTo Address 1: " + getClientData[0].Split('|')[25] + "\r\n";
                    resultsTxt.Text += "ShipTo Address 2: " + getClientData[0].Split('|')[26] + "\r\n";
                    resultsTxt.Text += "ShipTo City: " + getClientData[0].Split('|')[27] + "\r\n";
                    resultsTxt.Text += "ShipTo State: " + getClientData[0].Split('|')[28] + "\r\n";
                    resultsTxt.Text += "ShipTo Country: " + getClientData[0].Split('|')[29] + "\r\n";
                    resultsTxt.Text += "ShipTo Zip: " + getClientData[0].Split('|')[30] + "\r\n";
                    resultsTxt.Text += "ShipTo Phone: " + getClientData[0].Split('|')[31] + "\r\n";
                    resultsTxt.Text += "ShipTo Address Verification: " + getClientData[0].Split('|')[32] + "\r\n";
                    resultsTxt.Text += "ShipTo Is Residential: " + getClientData[0].Split('|')[33] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "Email: " + getClientData[0].Split('|')[34] + "\r\n";
                    resultsTxt.Text += "Customer PO Number: " + getClientData[0].Split('|')[35] + "\r\n";
                    resultsTxt.Text += "Payment Type: " + getClientData[0].Split('|')[36] + "\r\n";
                    resultsTxt.Text += "Payment Reference #: " + getClientData[0].Split('|')[37] + "\r\n";
                    resultsTxt.Text += "Ship Via: " + getClientData[0].Split('|')[38] + "\r\n";
                    resultsTxt.Text += "Warehouse Code: " + getClientData[0].Split('|')[41] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "PayPal Auth Amount: " + getClientData[0].Split('|')[48] + "\r\n";
                    resultsTxt.Text += "PayPal Auth Max: " + getClientData[0].Split('|')[49] + "\r\n";
                    resultsTxt.Text += "PayPal Payer Verified: " + getClientData[0].Split('|')[50] + "\r\n";
                    resultsTxt.Text += "PayPal Address Verified: " + getClientData[0].Split('|')[51] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "UDF Order Status: " + getClientData[0].Split('|')[52] + "\r\n";
                    resultsTxt.Text += "UDF External Order ID: " + getClientData[0].Split('|')[53] + "\r\n";
                    resultsTxt.Text += "UDF External User ID: " + getClientData[0].Split('|')[54] + "\r\n";
                    resultsTxt.Text += "-----------------------------" + "\r\n";
                    resultsTxt.Text += "Exported: " + getClientData[0].Split('|')[55] + "\r\n";
                    resultsTxt.Text += "Time Inserted: " + getClientData[0].Split('|')[56] + "\r\n";




                    
                }
                else
                {
                    resultsTxt.Text += "No customer info found.\r\n";
                }
                resultsTxt.Text += "Time Inserted: " + results[0].Split('|')[4] + "\r\n";
            }
            else
            {
                resultsTxt.Text = "No results in database.";
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            versionLbl.Text = VERSION;
        }
    }
}
