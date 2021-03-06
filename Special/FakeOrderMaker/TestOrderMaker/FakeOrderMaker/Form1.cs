﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace FakeOrderMaker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sqlInsert = GenerateSQLInsertionHeader();
            textBox58.Text = sqlInsert;
        }
        private string GenerateSQLInsertionHeader()
        {
            //OrderID (autogenerated;dont assign a value)
            //OrderNumber 
            //ExternalOrderID
            //ExternalUserID
            //OrderDate
            //OrderType
            //OrderStatus
            //HoldReason
            //ShipName
            //ShipFirstName
            //ShipLastName
            //ShipCompany
            //ShipAddress1
            //ShipAddress2
            //ShipCity
            //ShipState
            //ShipCountry
            //ShipZip
            //ShipPhone
            //BillName
            //BillFirstName
            //BillLastName
            //BillCompany
            //BillAddress1
            //BillAddress2
            //BillCity
            //BillState
            //BillCountry
            //BillZip
            //BillPhone
            //EmailAddress
            //CustomerPONumber
            //PaymentInfo
            //PaymentMethodString
            //WarehouseCode
            //ShipMethod
            //ShipMethodString
            //Comments
            //BankName
            //BankPhone
            //IPAddress
            //TimeInserted
            //Processed
            //Emailed
            //Exported
            //ExportedTime

            string OrderNumber = textBox1.Text == "NULL" ? "NULL" : "'" + textBox1.Text + "'";
            //string ARDivisionNumber = "'" + textBox2.Text + "'";
            string ExternalOrderID = "'" + textBox2.Text + "'";
            //string CustomerNumber = "'" + textBox3.Text + "'";
            string ExternalUserID = "'" + textBox3.Text + "'";

            string OrderDate = "'" + textBox4.Text + "'";
            string OrderType = "'" + textBox5.Text + "'";
            string OrderStatus = "'" + textBox6.Text + "'";
            //string HoldReasonCode = textBox7.Text == "NULL" ? "NULL" : "'" + textBox7.Text + "'";            
            string HoldReason = textBox7.Text == "NULL" ? "NULL" : "'" + textBox7.Text + "'";   
            //string TermsCode = "'" + textBox8.Text + "'";
            
            
            string ShipName = "'" + textBox22.Text + "'";
            string ShipFirstName = "'" + textBox23.Text + "'";
            string ShipLastName = "'" + textBox24.Text + "'";
            string ShipCompany = textBox25.Text == "NULL" ? "NULL" : "'" + textBox25.Text + "'";
            string ShipAddr1 = "'" + textBox26.Text + "'";
            string ShipAddr2 = textBox27.Text == "NULL" ? "NULL" : "'" + textBox27.Text + "'";
            string ShipCity = "'" + textBox28.Text + "'";
            string ShipState = "'" + textBox29.Text + "'";
            string ShipCountry = "'" + textBox30.Text + "'";
            string ShipZip = "'" + textBox31.Text + "'";
            string ShipPhone = "'" + textBox32.Text + "'";
            //string ShipAddressValidationStatus = textBox33.Text == "NULL" ? "NULL" : "'" + textBox33.Text + "'";
            //string ShipAddressResidential = textBox34.Text == "NULL" ? "NULL" : textBox34.Text;
            string BillName = "'" + textBox9.Text + "'";
            string BillFirstName = "'" + textBox10.Text + "'";
            string BillLastName = "'" + textBox11.Text + "'";
            string BillCompany = textBox12.Text == "NULL" ? "NULL" : "'" + textBox12.Text + "'";
            string BillAddr1 = "'" + textBox13.Text + "'";
            string BillAddr2 = textBox14.Text == "NULL" ? "NULL" : "'" + textBox14.Text + "'";
            string BillCity = "'" + textBox15.Text + "'";
            string BillState = "'" + textBox16.Text + "'";
            string BillCountry = "'" + textBox17.Text + "'";
            string BillZip = "'" + textBox18.Text + "'";
            string BillPhone = "'" + textBox19.Text + "'";
            //string BillAddressValidationStatus = textBox20.Text == "NULL" ? "NULL" : "'" + textBox20.Text + "'";
            //string BillAddressResidential = textBox21.Text == "NULL" ? "NULL" : textBox21.Text;
            string Email = "'" + textBox35.Text + "'";
            string CustomerPONumber = "'" + textBox36.Text + "'";
            string PaymentInfo = "'" + textBox37.Text + "'";
            string PaymentMethodString = "'" + textBox38.Text + "'";
            //string PaymentType = "'" + textBox37.Text + "'";
            //string PaymentReferenceNo = textBox38.Text == "NULL" ? "NULL" : "'" + textBox38.Text + "'";
            string WarehouseCode = "'" + textBox42.Text + "'";
            string ShipMethod = "'" + textBox39.Text + "'";
            //string ShipVia = "'" + textBox39.Text + "'";
            string ShipMethodString = "'" + textBox40.Text + "'";
            //string TaxSchedule = "'" + textBox40.Text + "'";
            //string TaxScheduleForOrder = textBox41.Text == "NULL" ? "NULL" : "'" + textBox41.Text + "'";
            string Comments = "'" + textBox66.Text + "'";
            
            //string CreditCardNumber = "'" + textBox43.Text + "'";
            //string ExpireDateMonth = "'" + textBox44.Text + "'";
            //string ExpireDateYear = "'" + textBox45.Text + "'";
            //string CVV = "'" + textBox46.Text + "'";
            string BankName = textBox47.Text == "NULL" ? "NULL" : "'" + textBox47.Text + "'";
            string BankPhone = textBox48.Text == "NULL" ? "NULL" : "'" + textBox48.Text + "'";
            string IPAddress = textBox67.Text == "NULL" ? "NULL" : "'" + textBox67.Text + "'";
                
            //string PayPalAuthAmt = textBox49.Text == "NULL" ? "NULL" : textBox49.Text;
            //string PayPalAuthMax = textBox50.Text == "NULL" ? "NULL" : textBox50.Text;
            //string PayPalPayerVerified = textBox51.Text == "NULL" ? "NULL" : textBox51.Text;
            //string PayPalAddressVerified = textBox52.Text == "NULL" ? "NULL" : textBox52.Text;
            //string UDFOrderStatus = "'" + textBox53.Text + "'";
            //string UDFExternalOrderID = textBox54.Text == "NULL" ? "NULL" : "'" + textBox54.Text + "'";
            //string UDFExternalUserID = textBox55.Text == "NULL" ? "NULL" : "'" + textBox55.Text + "'";
            
            string TimeInserted = "'" + textBox57.Text + "'";
            string Processed = textBox68.Text;
            string Emailed = textBox69.Text;
            string Exported = textBox56.Text;
            string ExportedTime = textBox70.Text == "NULL" ? "NULL" : "'" + textBox70.Text + "'";



            string sql = "INSERT INTO SalesOrders (OrderNumber,ExternalOrderID,ExternalUserID,OrderDate,OrderType,OrderStatus,HoldReason,ShipName,ShipFirstName,ShipLastName,ShipCompany,ShipAddress1,ShipAddress2,ShipCity,ShipState,ShipCountry,ShipZip,ShipPhone,BillName,BillFirstName,BillLastName,BillCompany,BillAddress1,BillAddress2,BillCity,BillState,BillCountry,BillZip,BillPhone,EmailAddress,CustomerPONumber,PaymentInfo,PaymentMethodString,WarehouseCode,ShipMethod,ShipMethodString,Comments,BankName,BankPhone,IPAddress,TimeInserted,Processed,Emailed,Exported,ExportedTime) VALUES(";
            sql += OrderNumber + "," + ExternalOrderID + "," + ExternalUserID + "," + OrderDate + "," + OrderType + "," + OrderStatus + "," +
                HoldReason + "," + ShipName + "," + ShipFirstName + "," + ShipLastName + "," + ShipCompany + "," + ShipAddr1 + "," +
                ShipAddr2 + "," + ShipCity + "," + ShipState + "," + ShipCountry + "," + ShipZip + "," + ShipPhone + "," + BillName + "," +
                BillFirstName + "," + BillLastName + "," + BillCompany + "," + BillAddr1 + "," + BillAddr2 + "," + BillCity + "," +
                BillState + "," + BillCountry + "," + BillZip + "," + BillPhone + "," + Email + "," + CustomerPONumber + "," + PaymentInfo + "," +
                PaymentMethodString + "," + WarehouseCode + "," + ShipMethod + "," + ShipMethodString + "," + Comments + "," + BankName + "," +
                BankPhone + "," + IPAddress + "," + TimeInserted + "," + Processed + "," + Emailed + "," + Exported + "," + ExportedTime + ")";
            return sql;
                




            /*string sql = "INSERT INTO SORealTimeImportCustomers (OrderNumber,ARDivisionNumber,CustomerNumber,OrderDate,OrderType,OrderStatus,HoldReasonCode,TermsCode,BillName,BillFirstName,BillLastName,BillCompany,BillAddr1,BillAddr2,BillCity,BillState,BillCountry,BillZip,BillPhone,BillAddressValidationStatus,BillAddressResidential,ShipName,ShipFirstName,ShipLastName,ShipCompany,ShipAddr1,ShipAddr2,ShipCity,ShipState,ShipCountry,ShipZip,ShipPhone,ShipAddressValidationStatus,ShipAddressResidential,Email,CustomerPONumber,PaymentType,PaymentReferenceNo,ShipVia,TaxSchedule,TaxScheduleForOrder,WarehouseCode,CreditCardNumber,ExpireDateMonth,ExpireDateYear,CVV,BankName,BankPhone,PayPalAuthAmt,PayPalAuthMax,PayPalPayerVerified,PayPalAddressVerified,UDFOrderStatus,UDFExternalOrderID,UDFExternalUserID,Exported,TimeInserted) VALUES(";
            sql += OrderNumber + "," + ARDivisionNumber + "," + CustomerNumber + "," + OrderDate + "," + OrderType + "," + OrderStatus + "," +
                HoldReasonCode + "," + TermsCode + "," + BillName + "," + BillFirstName + "," + BillLastName + "," + BillCompany + "," +
                BillAddr1 + "," + BillAddr2 + "," + BillCity + "," + BillState + "," + BillCountry + "," + BillZip + "," +
                BillPhone + "," + BillAddressValidationStatus + "," + BillAddressResidential + "," + ShipName + "," + ShipFirstName + "," +
                ShipLastName + "," + ShipCompany + "," + ShipAddr1 + "," + ShipAddr2 + "," + ShipCity + "," + ShipState + "," +
                ShipCountry + "," + ShipZip + "," + ShipPhone + "," + ShipAddressValidationStatus + "," + ShipAddressResidential + "," +
                Email + "," + CustomerPONumber + "," + PaymentType + "," + PaymentReferenceNo + "," + ShipVia + "," + TaxSchedule + "," +
                TaxScheduleForOrder + "," + WarehouseCode + "," + CreditCardNumber + "," + ExpireDateMonth + "," + ExpireDateYear + "," +
                CVV + "," + BankName + "," + BankPhone + "," + PayPalAuthAmt + "," + PayPalAuthMax + "," + PayPalPayerVerified + "," +
                PayPalAddressVerified + "," + UDFOrderStatus + "," + UDFExternalOrderID + "," + UDFExternalUserID + "," +
                Exported + "," + TimeInserted + ")";
            */
            return sql;



        }
        private string DeprecatedGenerateSQLInsertionHeader()
        {
            
            string OrderNumber = textBox1.Text == "NULL" ? "NULL" : "'" + textBox1.Text + "'";
            string ARDivisionNumber = "'" + textBox2.Text + "'";
            string CustomerNumber = "'" + textBox3.Text + "'";
            string OrderDate = "'" + textBox4.Text + "'";
            string OrderType = "'" + textBox5.Text + "'";
            string OrderStatus = "'" + textBox6.Text + "'";
            string HoldReasonCode = textBox7.Text == "NULL" ? "NULL" : "'" + textBox7.Text + "'";
            string TermsCode = "'" + textBox8.Text + "'";
            string BillName = "'" + textBox9.Text + "'";
            string BillFirstName = "'" + textBox10.Text + "'";
            string BillLastName = "'" + textBox11.Text + "'";
            string BillCompany = textBox12.Text == "NULL" ? "NULL" : "'" + textBox12.Text + "'";
            string BillAddr1 = "'" + textBox13.Text + "'";
            string BillAddr2 = textBox14.Text == "NULL" ? "NULL" : "'" + textBox14.Text + "'";
            string BillCity = "'" + textBox15.Text + "'";
            string BillState = "'" + textBox16.Text + "'";
            string BillCountry = "'" + textBox17.Text + "'";
            string BillZip = "'" + textBox18.Text + "'";
            string BillPhone = "'" + textBox19.Text + "'";
            string BillAddressValidationStatus = textBox20.Text == "NULL" ? "NULL" : "'" + textBox20.Text + "'";
            string BillAddressResidential = textBox21.Text == "NULL" ? "NULL" : textBox21.Text;
            string ShipName = "'" + textBox22.Text + "'";
            string ShipFirstName = "'" + textBox23.Text + "'";
            string ShipLastName = "'" + textBox24.Text + "'";
            string ShipCompany = textBox25.Text == "NULL" ? "NULL" : "'" + textBox25.Text + "'";
            string ShipAddr1 = "'" + textBox26.Text + "'";
            string ShipAddr2 = textBox27.Text == "NULL" ? "NULL" : "'" + textBox27.Text + "'";
            string ShipCity = "'" + textBox28.Text + "'";
            string ShipState = "'" + textBox29.Text + "'";
            string ShipCountry = "'" + textBox30.Text + "'";
            string ShipZip = "'" + textBox31.Text + "'";
            string ShipPhone = "'" + textBox32.Text + "'";
            string ShipAddressValidationStatus = textBox33.Text == "NULL" ? "NULL" : "'" + textBox33.Text + "'";
            string ShipAddressResidential = textBox34.Text == "NULL" ? "NULL" : textBox34.Text;
            string Email = "'" + textBox35.Text + "'";
            string CustomerPONumber = "'" + textBox36.Text + "'";
            string PaymentType = "'" + textBox37.Text + "'";
            string PaymentReferenceNo = textBox38.Text == "NULL" ? "NULL" : "'" + textBox38.Text + "'";
            string ShipVia = "'" + textBox39.Text + "'";
            string TaxSchedule = "'" + textBox40.Text + "'";
            string TaxScheduleForOrder = textBox41.Text == "NULL" ? "NULL" : "'" + textBox41.Text + "'";
            string WarehouseCode = "'" + textBox42.Text + "'";
            string CreditCardNumber = "'" + textBox43.Text + "'";
            string ExpireDateMonth = "'" + textBox44.Text + "'";
            string ExpireDateYear = "'" + textBox45.Text + "'";
            string CVV = "'" + textBox46.Text + "'";
            string BankName = textBox47.Text == "NULL" ? "NULL" : "'" + textBox47.Text + "'";
            string BankPhone = textBox48.Text == "NULL" ? "NULL" : "'" + textBox48.Text + "'";
            string PayPalAuthAmt = textBox49.Text == "NULL" ? "NULL" : textBox49.Text;
            string PayPalAuthMax = textBox50.Text == "NULL" ? "NULL" : textBox50.Text;
            string PayPalPayerVerified = textBox51.Text == "NULL" ? "NULL" : textBox51.Text;
            string PayPalAddressVerified = textBox52.Text == "NULL" ? "NULL" : textBox52.Text;
            string UDFOrderStatus = "'" + textBox53.Text + "'";
            string UDFExternalOrderID = textBox54.Text == "NULL" ? "NULL" : "'" + textBox54.Text + "'";
            string UDFExternalUserID = textBox55.Text == "NULL" ? "NULL" : "'" + textBox55.Text + "'";
            string Exported = textBox56.Text;
            string TimeInserted = "'" + textBox57.Text + "'";

            string sql = "INSERT INTO SORealTimeImportCustomers (OrderNumber,ARDivisionNumber,CustomerNumber,OrderDate,OrderType,OrderStatus,HoldReasonCode,TermsCode,BillName,BillFirstName,BillLastName,BillCompany,BillAddr1,BillAddr2,BillCity,BillState,BillCountry,BillZip,BillPhone,BillAddressValidationStatus,BillAddressResidential,ShipName,ShipFirstName,ShipLastName,ShipCompany,ShipAddr1,ShipAddr2,ShipCity,ShipState,ShipCountry,ShipZip,ShipPhone,ShipAddressValidationStatus,ShipAddressResidential,Email,CustomerPONumber,PaymentType,PaymentReferenceNo,ShipVia,TaxSchedule,TaxScheduleForOrder,WarehouseCode,CreditCardNumber,ExpireDateMonth,ExpireDateYear,CVV,BankName,BankPhone,PayPalAuthAmt,PayPalAuthMax,PayPalPayerVerified,PayPalAddressVerified,UDFOrderStatus,UDFExternalOrderID,UDFExternalUserID,Exported,TimeInserted) VALUES(";
            sql += OrderNumber + "," + ARDivisionNumber + "," + CustomerNumber + "," + OrderDate + "," + OrderType + "," + OrderStatus + "," +
                HoldReasonCode + "," + TermsCode + "," + BillName + "," + BillFirstName + "," + BillLastName + "," + BillCompany + "," +
                BillAddr1 + "," + BillAddr2 + "," + BillCity + "," + BillState + "," + BillCountry + "," + BillZip + "," +
                BillPhone + "," + BillAddressValidationStatus + "," + BillAddressResidential + "," + ShipName + "," + ShipFirstName + "," +
                ShipLastName + "," + ShipCompany + "," + ShipAddr1 + "," + ShipAddr2 + "," + ShipCity + "," + ShipState + "," +
                ShipCountry + "," + ShipZip + "," + ShipPhone + "," + ShipAddressValidationStatus + "," + ShipAddressResidential + "," +
                Email + "," + CustomerPONumber + "," + PaymentType + "," + PaymentReferenceNo + "," + ShipVia + "," + TaxSchedule + "," +
                TaxScheduleForOrder + "," + WarehouseCode + "," + CreditCardNumber + "," + ExpireDateMonth + "," + ExpireDateYear + "," +
                CVV + "," + BankName + "," + BankPhone + "," + PayPalAuthAmt + "," + PayPalAuthMax + "," + PayPalPayerVerified + "," +
                PayPalAddressVerified + "," + UDFOrderStatus + "," + UDFExternalOrderID + "," + UDFExternalUserID + "," +
                Exported + "," + TimeInserted + ")";

            return sql;



        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT IDiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" + textBox61.Text + "'");
            if (results.Count > 0)
            {
                decimal price = decimal.Parse(results[0].Split('|')[0]);
                textBox63.Text = price.ToString("0.00");
            }
            else
            {
                MessageBox.Show("Could not pull an internet price for item #" + textBox60.Text);
            }
        }
        private bool VerifyLineFields()
        {
            if (textBox59.Text.Length > 32000)
            {
                return false;
            }
            if (textBox60.Text.Length > 32000)
            {
                return false;
            }
            if (textBox61.Text.Length > 30)
            {
                return false;
            }
            try
            {
                int.Parse(textBox62.Text);
            }
            catch
            {
                return false;
            }
            try { decimal.Parse(textBox63.Text); }
            catch { return false; }
            if (textBox64.Text.Length > 32000)
            {
                return false;
            }

            return true;
        }
        private bool VerifyHeaderFields()
        {
            if (textBox1.Text.Length > 7 | textBox1.Text.Length < 7)
            {
                return false;
            }
            if (textBox2.Text.Length > 50)
            {
                return false;
            }
            if (textBox3.Text.Length > 50)
            {
                return false;
            }
            if (textBox4.Text[2] != '/' | textBox4.Text[5] != '/')
            {
                MessageBox.Show("Order Date Must Match Format- 01/01/2000");
                return false;
            }
            else
            {
                try
                {
                    DateTime.Parse(textBox4.Text);
                    
                }
                catch
                {
                    MessageBox.Show("Order Date Must Match Format- 01/01/2000");
                    return false;
                }
            }
            if (textBox5.Text.Length > 1)
            {
                return false;
            }
            if (textBox6.Text.Length > 1)
            {
                return false;
            }
            if (textBox7.Text.Length > 5)
            {
                return false;
            }
            if (textBox22.Text.Length > 30)
            {
                return false;
            }
            if (textBox23.Text.Length > 30)
            {
                return false;
            }
            if (textBox24.Text.Length > 30)
            {
                return false;
            }

            return true;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            bool test = VerifyLineFields();
            if (test == false)
            {
                MessageBox.Show("One of your line data input fields contains the wrong data. Please check over and try again.");
                return; 
            }

            //this ends up a bit tricky. first we query the ordernumber in the salesorder table to get the orderid.
            //next we use that here. the tricky part about this is we must generate it before we know the orderid to begin with!!!
            //hence the placeholder tag {ORDERID}

            //LineID    
            string OrderID = "{ORDERID}";
            int LineNumber = listBox1.Items.Count;
            string LineType = "'" + textBox59.Text + "'";
            string ProductCode = "'" + textBox60.Text + "'";
            string ItemNumber = textBox61.Text == "NULL" ? "NULL" : "'" + textBox61.Text + "'";
            string Quantity = textBox62.Text;
            string UnitPrice = textBox63.Text;
            string Comment = textBox64.Text == "NULL" ? "NULL" : "'" + textBox64.Text + "'";
            string TimeInserted = "CURRENT_TIMESTAMP";

            string sqlStatement = "INSERT INTO SalesOrderLines (OrderID,LineNumber,LineType,ProductCode,ItemNumber,Quantity," + 
                "UnitPrice,Comment,TimeInserted) VALUES(" + OrderID + "," + LineNumber.ToString() + "," + LineType + "," +
                ProductCode + "," + ItemNumber + "," + Quantity.ToString() + "," + UnitPrice.ToString() + "," + Comment +
                "," + TimeInserted + ")";
            listBox1.Items.Add(sqlStatement);

        }
        private void oldButton2_Click(object sender, EventArgs e)
        {
            string OrderNumber = "'" + textBox1.Text + "'";
            int LineNumber = listBox1.Items.Count;
            string LineType = "'" + textBox59.Text + "'";
            string ProductCode = "'" + textBox60.Text + "'";
            string QtyOrdered = textBox61.Text == "NULL" ? "NULL" : textBox61.Text;
            string UnitPrice = textBox62.Text == "NULL" ? "NULL" : textBox62.Text;
            string Comment1 = textBox63.Text == "NULL" ? "NULL" : "'" + textBox63.Text + "'";
            string Comment2 = textBox64.Text == "NULL" ? "NULL" : "'" + textBox64.Text + "'";

            string sql = "INSERT INTO SORealTimeImportOrders (OrderNumber,LineNumber,LineType,ProductCode,QtyOrdered,UnitPrice,Comment1,Comment2) VALUES(" +
                OrderNumber + "," + LineNumber + "," + LineType + "," + ProductCode + "," + QtyOrdered + "," + UnitPrice + "," +
                Comment1 + "," + Comment2 + ")";
            listBox1.Items.Add(sql);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Are you sure you want to insert this Fake Order into the database?", "Are You Sure?", MessageBoxButtons.YesNo);
            if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
            {

                if (textBox58.Text.Length <= 0)
                {
                    MessageBox.Show("You must have generated a header before inserting to database. You need to press the \"Obtain SQL Insert Query\" in the \"Order Header\" tab section.");
                    return;
                }
                if (listBox1.Items.Count <= 0)
                {
                    MessageBox.Show("You must generate order lines before inserting to database. Do this in the \"Order Lines\" tab.");
                    return;
                }

                //Connectivity.SQLCalls.sqlQuery(textBox58.Text);
                Connectivity.VPSCalls.QueryDB(textBox58.Text);


                //now that we have inserted the header, we must retrieve it's id before continuing. We need to inject the id into
                //the listbox items below, replace placeholder {ORDERID}

                Connectivity.VPSCalls.VPSResult result = Connectivity.VPSCalls.QueryDB("SELECT OrderID FROM SalesOrders WHERE OrderNumber='" + textBox1.Text + "'");

                string results = result.Results;
                string errors = result.Errors;

                if (results != "")
                {
                    string orderID = results.Split('\n')[1];
                    foreach (string itm in listBox1.Items)
                    {
                        Connectivity.SQLCalls.sqlQuery(itm.Replace("{ORDERID}",orderID));
                    }
                }
                if (errors != "")
                {
                    MessageBox.Show("Error (Also, you need to delete the header that was just inserted before this error): " + errors);
                }                
                MessageBox.Show("Failed to insert order #" + textBox1.Text + ". Tom has disabled this control intentionally until a better testing time.");
                listBox1.Items.Clear();
                listBox1.Refresh();
                textBox58.Text = "";
            }
        }
        private void SetToolTips()
        {
            ToolTip tt = new ToolTip();
            tt.SetToolTip(textBox1, "OrderNumber: VARCHAR(7), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox2, "ExternalOrderID: VARCHAR(50), NULL, DEFAULT=NULL");
            tt.SetToolTip(textBox3, "ExternalUserID: VARCHAR(50), NULL, DEFAULT=NULL");
            tt.SetToolTip(textBox4, "OrderDate: VARCHAR(10), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox5, "OrderType: CHAR(1), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox6, "OrderStatus: CHAR(1), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox7, "HoldReason: VARCHAR(5), NOT NULL, NO DEFAULT");
            
            //22-32
            tt.SetToolTip(textBox22, "ShipName: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox23, "ShipFirstName: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox24, "ShipLastName: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox25, "ShipCompany: VARCHAR(30), NULL, DEFAULT=NULL");
            tt.SetToolTip(textBox26, "ShipAddress1: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox27, "ShipAddress2: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox28, "ShipCity: VARCHAR(20), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox29, "ShipState: VARCHAR(2), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox30, "ShipCountry: VARCHAR(2), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox31, "ShipZip: VARCHAR(10), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox32, "ShipPhone: VARCHAR(17), NOT NULL, NO DEFAULT");
            
            //9-19
            tt.SetToolTip(textBox9, "BillName: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox10, "BillFirstName: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox11, "BillLastName: VARCHAR(30), NOT NULL,NO DEFAULT");
            tt.SetToolTip(textBox12, "BillCompany: VARCHAR(30), NULL, DEFAULT=NULL");
            tt.SetToolTip(textBox13, "BillAddress1: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox14, "BillAddress2: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox15, "BillCity: VARCHAR(20), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox16, "BillState: VARCHAR(2), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox17, "BillCountry: VARCHAR(2), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox18, "BillZip: VARCHAR(10), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox19, "BillPhone: VARCHAR(17), NOT NULL, NO DEFAULT");
            
            //35-38
            //42,39,40,66
            tt.SetToolTip(textBox35, "EmailAddress: VARCHAR(50), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox36, "CustomerPONumber: VARCHAR(15), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox37, "PaymentInfo: VARCHAR(150), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox38, "PaymentMethodString: VARCHAR(50), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox42, "WarehouseCode: VARCHAR(3), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox39, "ShipMethod: VARCHAR(15), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox40, "ShipMethodString: VARCHAR(100), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox66, "Comments: TEXT, NOT NULL, DEFAULT=NULL");

            //47,48,67
            tt.SetToolTip(textBox47, "BankName: VARCHAR(50), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox48, "BankPhone: VARCHAR(30), NOT NULL, NO DEFAULT");
            tt.SetToolTip(textBox67, "IPAddress: VARCHAR(15), NOT NULL, NO DEFAULT");

            //57,68,69,56,70
            tt.SetToolTip(textBox57, "TimeInserted: TIMESTAMP, NOT NULL, DEFAULT=CURRENT_TIMESTAMP");
            tt.SetToolTip(textBox68, "Processed: TINYINT(4), NOT NULL, DEFAULT=0");
            tt.SetToolTip(textBox69, "Emailed: TINYINT(4), NOT NULL, DEFAULT=0");
            tt.SetToolTip(textBox56, "Exported: TINYINT(4), NOT NULL, DEFAULT=0");
            tt.SetToolTip(textBox70, "ExportedTime: DATETIME, NULL, DEFAULT=NULL");


        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormClosed += new FormClosedEventHandler(Form1_FormClosed);
            SetToolTips();
        }

        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            treeView1.Refresh();            
            treeView1 = Connectivity.VPSCalls.GetDatabaseTree(treeView1);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Connectivity.VPSCalls.VPSResult vpsResult = Connectivity.VPSCalls.QueryDB(textBox65.Text);
            if (vpsResult.Results.Length > 0)
            {
                ThrowDataInLVI(vpsResult.Results);
            }
            if (vpsResult.Errors.Length > 0)
            {
                string[] Errors = vpsResult.Errors.Split('\n');
                foreach (string err in Errors)
                {
                    listBox2.Items.Add("VPS Error: " + err);
                }
            }
        }
        private void ClearLVI()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();

            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.View = View.Details;

            listView1.Refresh();
        }
        private void ThrowDataInLVI(string returnedData)
        {
            ClearLVI();
            //each field is denoted by \t, each line is denoted by \n...
            string[] Lines = returnedData.Split('\n');

            int counter = 0;
            foreach (string line in Lines)
            {
                if (counter == 0)
                {
                    //header... 
                    string[] headerColumns = line.Split('\t');

                    listBox2.Refresh();
                    listBox2.Items.Add("Found " + headerColumns.GetLength(0).ToString() + " Columns.");
                    foreach (string col in headerColumns)
                    {
                        ColumnHeader tmpCol = new ColumnHeader();
                        tmpCol.Text = col;
                        tmpCol.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                        listView1.Columns.Add(tmpCol);
                    }
                }
                else
                {
                    ListViewItem tmpItm = new ListViewItem();
                    string[] rowFields = line.Split('\t');
                    int counter2 = 0;
                    foreach (string field in rowFields)
                    {
                        if (counter2 == 0)
                        {
                            tmpItm.Text = field;
                        }
                        else
                        {
                            tmpItm.SubItems.Add(field);
                        }
                        counter2++;
                    }
                    listView1.Items.Add(tmpItm);
                }
                counter++;
            }

            listBox2.Items.Add("Pulled " + (Lines.GetLength(0) - 1).ToString() + " Lines.");
            listBox2.Refresh();
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Connectivity.VPSCalls.VPSResult res = Connectivity.VPSCalls.QueryDB("SELECT EBayItemID FROM Items WHERE ItemNumber='" + textBox61.Text + "'");

            string results = res.Results;
            string errors = res.Errors;
            if (results != "")
            {
                textBox60.Text = results.Split('\n')[1];
            }
            if (errors != "")
            {
                textBox60.Text = "";
            }


        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}
