﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Collections;

namespace SalesTransactionsReport
{
    public partial class Form1 : Form
    {

        public struct SalesTransactions
        {
            public bool FoundCoupon;
            public bool AboveOrderMin;
            public string SalesType;
            public string InvoiceNumber;
            public string InvoiceDate;
            public string LineType;
            public string ItemNumber;
            public string Description;
            public string Warehouse;
            public string Quantity;
            public string UnitPrice;
            public string UnitCost;
            public string ChargeCode;
            public string Extension;
            public string SalesOrderNo;
            public string SalesOrderDate;
            public string SalesOrderStatus;
            public string CustomerName;
            public string HasInvoices;
            public string HasOpen;
            public string AmountDue;
            public List<SalesTransactions> Lines;

            
        }
        public List<SalesTransactions> SalesReportData = new List<SalesTransactions>();

        public struct SaleTransactionReportData
        {
            public Dictionary<string, dynamic> OrdersFromHistory;
            public Dictionary<string, dynamic> SalesFromHistory;
            public Dictionary<string, dynamic> UnknownFromHistory;
            //public Dictionary<string, dynamic> InvoicedLinesHeader;
            //public Dictionary<string, dynamic> InvoicedLinesLines;
            public void InitializeReport()
            {
                OrdersFromHistory = new Dictionary<string,dynamic>();
                SalesFromHistory = new Dictionary<string,dynamic>();
                UnknownFromHistory = new Dictionary<string, dynamic>();
                //InvoicedLinesHeader = new Dictionary<string, dynamic>();
                //InvoicedLinesLines = new Dictionary<string, dynamic>();                
            }

        }

        public SaleTransactionReportData ReportData = new SaleTransactionReportData();
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private List<SalesTransactions> FilterResults(List<SalesTransactions> Results)
        {
            List<SalesTransactions> tmpTxLst = new List<SalesTransactions>();
            foreach (SalesTransactions ST in Results)
            {
                
                foreach (string itm in textBox1.Text.Split(new string[]{"\r\n"},StringSplitOptions.RemoveEmptyEntries))
                {
                    if (itm.ToUpper() == ST.ItemNumber.ToUpper())
                    {
                        tmpTxLst.Add(ST);
                        break;
                    }
                }
            }
            return tmpTxLst;
        }

        public class ListViewItemComparer : IComparer
        {
            private int col = 0;

            public ListViewItemComparer(int column)
            {
                col = column;
            }
            public int Compare(object x, object y)
            {
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ReportData.InitializeReport();
        }



        private void LVItoCSV(string FilePath, ListView ListViewControl)
        {
            string tmpFile = "";
            int Cols = ListViewControl.Columns.Count;

            foreach (ColumnHeader ch in ListViewControl.Columns)
            {
                if (ch.Text.Contains("\"") == true | ch.Text.Contains(",") == true)
                {
                    tmpFile += "\"" + ch.Text + "\",";
                }
                else
                {
                    tmpFile += ch.Text + ",";
                }
            }
            tmpFile += "\r\n";


            foreach (ListViewItem lvi in ListViewControl.Items)
            {
                for (int x = 0; x < Cols; x++)
                {
                    if (lvi.SubItems[x].Text.Contains("\"") == true | lvi.SubItems[x].Text.Contains(",") ==true)
                    {
                        tmpFile += "\"" + lvi.SubItems[x].Text + "\",";
                    }
                    else
                    {
                        tmpFile += lvi.SubItems[x].Text + ",";
                    }
                }
                tmpFile += "\r\n";
            }

            System.IO.File.WriteAllText(FilePath, tmpFile);
            statusLbl.Text = "Status: File saved to \"" + FilePath + "\"!";
        }



        private void SetupLVI()
        {
            reportLVI.Items.Clear();
            reportLVI.Columns.Clear();
            reportLVI.Clear();
            reportLVI.Refresh();
            reportLVI.FullRowSelect = true;
            reportLVI.GridLines = true;
            reportLVI.View = View.Details;
            reportLVI.ListViewItemSorter = new ListViewItemComparer(6);

            ColumnHeader IDCol = new ColumnHeader();
            ColumnHeader aboveOrderMinCol = new ColumnHeader();
            ColumnHeader foundCouponCol = new ColumnHeader();
            ColumnHeader typeCol = new ColumnHeader();
            ColumnHeader orderNumberCol = new ColumnHeader();
            ColumnHeader orderDateCol = new ColumnHeader();
            ColumnHeader customerNameCol = new ColumnHeader();
            ColumnHeader invoiceNumberCol = new ColumnHeader();
            ColumnHeader invoiceDateCol = new ColumnHeader();
            ColumnHeader itemNumberCol = new ColumnHeader();
            ColumnHeader quantityCol = new ColumnHeader();
            ColumnHeader unitPriceCol = new ColumnHeader();
            ColumnHeader unitCostCol = new ColumnHeader();
            ColumnHeader mappPriceCol = new ColumnHeader();
            ColumnHeader yahooPriceCol = new ColumnHeader();
            ColumnHeader storePriceCol = new ColumnHeader();

            IDCol.Text = "ID";
            aboveOrderMinCol.Text = "Above Order Min?";
            foundCouponCol.Text = "Found Coupon?";
            typeCol.Text = "Type";
            orderNumberCol.Text = "Order Number";
            orderDateCol.Text = "Order Date";
            customerNameCol.Text = "Customer Name";
            invoiceNumberCol.Text = "Invoice Number";
            invoiceDateCol.Text = "Invoice Date";
            itemNumberCol.Text = "Item Number";
            quantityCol.Text = "Quantity";
            unitPriceCol.Text = "Unit Price";
            unitCostCol.Text = "Unit Cost";
            mappPriceCol.Text = "MAPP";
            yahooPriceCol.Text = "Yahoo Price";
            storePriceCol.Text = "Store Price";

            reportLVI.Columns.Add(IDCol);

                reportLVI.Columns.Add(aboveOrderMinCol);          

                reportLVI.Columns.Add(foundCouponCol);

            reportLVI.Columns.Add(typeCol);
            reportLVI.Columns.Add(orderNumberCol);
            reportLVI.Columns.Add(orderDateCol);
            reportLVI.Columns.Add(customerNameCol);
            reportLVI.Columns.Add(invoiceNumberCol);
            reportLVI.Columns.Add(invoiceDateCol);
            reportLVI.Columns.Add(itemNumberCol);
            reportLVI.Columns.Add(quantityCol);
            reportLVI.Columns.Add(unitPriceCol);
            if (checkBox4.Checked == true)
            {
                reportLVI.Columns.Add(unitCostCol);
            }
            reportLVI.Columns.Add(mappPriceCol);
            reportLVI.Columns.Add(yahooPriceCol);
            reportLVI.Columns.Add(storePriceCol);



        }

        private void generateBtn_Click(object sender, EventArgs e)
        {
            generateBtn.Enabled = false;
            newReportBtn.Visible = true;
            generateBtn.Refresh();
            newReportBtn.Refresh();
            SetupLVI();

            if (checkBox1.Checked == true)
            {
                RetrieveOrdersFromHistory();
            }
            if (checkBox2.Checked == true)
            {
                RetrieveSalesFromHistory();
            }
            if (checkBox3.Checked == true)
            {
                RetrieveUnknownsFromHistory();
            }
            //SalesReportData = FilterResults(SalesReportData);
            //shoot out data to excel...

            //inform end user process completed!


            FillLVIFromReportData();
            
            
            
            //List<string> mappAndPrices = Connectivity.SQLCalls.sqlQuery("SELECT MAPP,IDiscountMarkupPriceRate1,DiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber=''");

            
            //DataTable dt = ParseResults(result);
            
           // WriteToLVI(dt);
        }
        private void FillLVIFromReportData()
        {
            SetupLVI();
            

            
            int counter = 1;
            foreach (SalesTransactions ST in SalesReportData)
            {

                string ItemNumber = ST.ItemNumber;
                decimal MAPP = getMAPP(ItemNumber);
                decimal Yahoo = getINETPrice(ItemNumber);
                decimal Store = getStorePrice(ItemNumber);


                ListViewItem newItem = new ListViewItem();
                newItem.Text = counter.ToString();
                //newItem.SubItems.Add("Not Coded 100%");
                newItem.SubItems.Add(ST.AboveOrderMin == true ? "Y" : "");
                newItem.SubItems.Add("Not Coded 100%");
                //newItem.SubItems.Add("Not Coded 100%");
                newItem.SubItems.Add(ST.SalesType);
              /*  if (ST.SalesOrderNo != null)
                {
                    if (ST.SalesOrderNo.Length == 0)
                    {
                        newItem.SubItems.Add("n/a");
                    }
                    else
                    {
                        string character = ST.SalesOrderNo.Substring(0, 2);
                        try
                        {
                            //S/Os can be 007xxxx while POSs can be 02xxxxx
                            int test = int.Parse(character);
                            if (test == 2)
                            {
                                newItem.SubItems.Add("pos");
                            }
                            else
                            {
                                newItem.SubItems.Add("s/o");
                            }
                        }
                        catch
                        {
                            newItem.SubItems.Add("s/o");
                        }
                    }
                }
                else
                {
                    newItem.SubItems.Add("n/a");
                }
                */
                newItem.SubItems.Add(ST.SalesOrderNo);
                newItem.SubItems.Add(ST.SalesOrderDate);
                newItem.SubItems.Add(ST.CustomerName);
                newItem.SubItems.Add(ST.InvoiceNumber);
                newItem.SubItems.Add(ST.InvoiceDate);
                newItem.SubItems.Add(ST.ItemNumber);
                try
                {
                    newItem.SubItems.Add(decimal.Parse(ST.Quantity).ToString("0"));
                }
                catch { newItem.SubItems.Add("0"); }
                try
                {
                    newItem.SubItems.Add(decimal.Parse(ST.UnitPrice).ToString("$0.00"));
                }
                catch { newItem.SubItems.Add("0"); }
                try
                {
                    if (decimal.Parse(MAPP.ToString()) > 0)
                    {
                        newItem.SubItems.Add(decimal.Parse(MAPP.ToString()).ToString("$0.00"));
                    }
                    else { newItem.SubItems.Add("-"); }
                }
                catch { newItem.SubItems.Add("-"); }
                try
                {
                    if (decimal.Parse(Yahoo.ToString()) > 0)
                    {
                    newItem.SubItems.Add(decimal.Parse(Yahoo.ToString()).ToString("$0.00"));
                    }
                    else{ newItem.SubItems.Add("-");}

                }
                catch { newItem.SubItems.Add("-"); }
                try
                {
                    if (decimal.Parse(Store.ToString()) > 0)
                    {
                        newItem.SubItems.Add(decimal.Parse(Store.ToString()).ToString("$0.00"));
                    }
                    else{newItem.SubItems.Add("-");
                    }
                }
                catch { newItem.SubItems.Add("-"); }



                reportLVI.Items.Add(newItem);
                counter++;
            }
        }
        private void RetrieveOrdersFromHistory()
        {
            //call #1... get orders history from between the two selected dates on the form...
            //statusLbl.Text = "Status: Retrieving Orders From History From MAS200 (and ten years later...)";
           // statusLbl.Refresh();


            string startDateYear = (startDate.Value.Year.ToString().Length == 4 ? startDate.Value.Year.ToString() : "20" + startDate.Value.Year.ToString());
            string startDateMonth = (startDate.Value.Month.ToString().Length == 2 ? startDate.Value.Month.ToString() : "0" + startDate.Value.Month.ToString());
            string startDateDay = (startDate.Value.Day.ToString().Length == 2 ? startDate.Value.Day.ToString() : "0" + startDate.Value.Day.ToString()); 
            string endDateYear = (endDate.Value.Year.ToString().Length == 4? endDate.Value.Year.ToString() : "20" + endDate.Value.Year.ToString());
            string endDateMonth = (endDate.Value.Month.ToString().Length == 2? endDate.Value.Month.ToString():"0" + endDate.Value.Month.ToString());
            string endDateDay = (endDate.Value.Day.ToString().Length == 2 ? endDate.Value.Day.ToString() : "0" + endDate.Value.Day.ToString());




            string query = "SELECT SalesOrderNo, OrderDate, OrderStatus, BillToName, LastInvoiceNo FROM SO_SalesOrderHistoryHeader WHERE OrderDate>={ d '" + startDateYear + "-" + startDateMonth + "-" + startDateDay + "'} AND OrderDate<={ d '" + endDateYear + "-" + endDateMonth + "-" + endDateDay + "'}";

            string ordersFromHistory = Connectivity.MASCalls.Retrieve(query);
            if (ordersFromHistory == null)
            {
                return;
            }
            ReportData.OrdersFromHistory = JSONtoDictionary(ordersFromHistory);
            int totRecords = ((System.Collections.ArrayList)ReportData.OrdersFromHistory["rows"]).Count;
            int counter = 0;
            foreach (var item in ReportData.OrdersFromHistory["rows"])
            {
                counter++;
                
                statusLbl.Text = "Status: Processing Orders From History (" + counter.ToString() + "/" + totRecords.ToString() + ")";
                statusLbl.Refresh();
                //need to handle null values
                string SO = ((System.Collections.ArrayList)item)[0].ToString();
                string OrderStatus = ((System.Collections.ArrayList)item)[3] == null ? "" : ((System.Collections.ArrayList)item)[3].ToString();

                if (OrderStatus != "X" & OrderStatus != "Z")
                {
                    if (SO == "0070822")
                    {
                    }

                    //first does this order have open lines or closed lines or both?
                    if (OrderStatus != "C")
                    {
                        //has unfinished business... open lines.
                        GetOpenLines(SO);
                    }

                        if (((System.Collections.ArrayList)item)[4] != null)
                        {
                            if (((System.Collections.ArrayList)item)[4].ToString() != "")
                            {
                                //has completed business.. closed lines.
                                GetInvoicedLines(((System.Collections.ArrayList)item)[0].ToString(), ((System.Collections.ArrayList)item)[1].ToString(), (((System.Collections.ArrayList)item)[3] == null ? "" : ((System.Collections.ArrayList)item)[3].ToString().ToString()));

                            }
                        }
                    



/*
                    //NOT ONE OR THE OTHER... BUT BOTH:


                    if (((System.Collections.ArrayList)item)[4] != null)
                    {
                        GetInvoicedLines(((System.Collections.ArrayList)item)[0].ToString(), ((System.Collections.ArrayList)item)[1].ToString(), (((System.Collections.ArrayList)item)[3] == null ? "" : ((System.Collections.ArrayList)item)[3].ToString().ToString()));

                    }
                    //else
                    //{
                    //MessageBox.Show("Whoa, finally got a \"NULLer\". Time to check Open Invoices for accuracy.");
                    string SONum = ((System.Collections.ArrayList)item)[0].ToString();

                    if (OrderStatus != "C")
                    {
                        GetOpenLines(SONum);
                    }

                    //}


*/







                    //now filter the orders?

                    //
                }
            }
            
        }
        private void FilterList()
        {
            try
            {
                int itemMin = int.Parse(saleItemCount.Text);
                foreach (ListViewItem lvi in reportLVI.Items)
                {
                    try
                    {
                        int qty = int.Parse(lvi.SubItems[7].Text);
                        if (qty < itemMin)
                        {
                            //remove from list
                            reportLVI.Items.Remove(lvi);
                        }
                    }
                    catch
                    {
                        //remove or keep?

                    }
                }


            }
            catch
            {
                MessageBox.Show("The value entered in Sale Item Count is invalid.");
                return;
            }
        }
        private void RetrieveSalesFromHistory()
        {


            string startDateYear = (startDate.Value.Year.ToString().Length == 4 ? startDate.Value.Year.ToString() : "20" + startDate.Value.Year.ToString());
            string startDateMonth = (startDate.Value.Month.ToString().Length == 2 ? startDate.Value.Month.ToString() : "0" + startDate.Value.Month.ToString());
            string startDateDay = (startDate.Value.Day.ToString().Length == 2 ? startDate.Value.Day.ToString() : "0" + startDate.Value.Day.ToString());
            string endDateYear = (endDate.Value.Year.ToString().Length == 4 ? endDate.Value.Year.ToString() : "20" + endDate.Value.Year.ToString());
            string endDateMonth = (endDate.Value.Month.ToString().Length == 2 ? endDate.Value.Month.ToString() : "0" + endDate.Value.Month.ToString());
            string endDateDay = (endDate.Value.Day.ToString().Length == 2 ? endDate.Value.Day.ToString() : "0" + endDate.Value.Day.ToString());
            //call #2... get sales history from between the two selected dates on the form...
            string salesFromHistory = Connectivity.MASCalls.Retrieve("SELECT TransactionNo,TransactionDate,BillToName,TransactionType,ReferenceNo,AmountDue FROM P2_TransactionHistoryHeader WHERE TransactionDate>={ d '" + startDateYear + "-" + startDateMonth + "-" + startDateDay + "'} AND TransactionDate<={ d '" + endDateYear + "-" + endDateMonth + "-" + endDateDay + "'} AND TransactionStatus<>'V' AND AmountDue>0");// AND ((UnitCost>0) OR (AmountDue>0))");
            if (salesFromHistory == null)
            {
                return;
            }
            ReportData.SalesFromHistory = JSONtoDictionary(salesFromHistory);
            //foreach (var item in ReportData.OrdersFromHistory["rows"])
            foreach(System.Collections.ArrayList salesLinesAL in (System.Collections.ArrayList)ReportData.SalesFromHistory["rows"])
            {
                
                List<SalesTransactions> tmpSalesReportData = new List<SalesTransactions>();

                
                //System.Collections.ArrayList headerSeq = salesLinesAL[0];               
                object TransactionNo = salesLinesAL[0];

                object TransactionDate = salesLinesAL[1];
                object BillToName = salesLinesAL[2];
                object TransactionType = salesLinesAL[3];

                object ReferenceNo = salesLinesAL[4];
                object AmountDue = salesLinesAL[5];
                //if (TransactionType.ToString() == "S" & ReferenceNo !=null)
                //{
                   
                //        break;
                   
                //}
                //if (ReferenceNo == null)
                //{
                //    break;
                //}

                
                string salesLines = Connectivity.MASCalls.Retrieve("SELECT ItemCode,ItemType,ItemCodeDesc,WarehouseCode,QuantityOrdered,UnitPrice,UnitCost,ExtensionAmt FROM P2_TransactionHistoryDetail WHERE TransactionNo='" + TransactionNo.ToString() + "' AND (ItemType='1' OR ItemType='3')");
                Dictionary<string, dynamic> salesLineDict = JSONtoDictionary(salesLines);
                if (salesLineDict != null)
                {
                    foreach (System.Collections.ArrayList invLinesAL in (System.Collections.ArrayList)salesLineDict["rows"])
                    {
                        //System.Collections.ArrayList salesArray = salesLineDict["rows"];
                        System.Collections.ArrayList headerSeq = (System.Collections.ArrayList)invLinesAL;

                        SalesTransactions tmpTx = new SalesTransactions();
                        if (headerSeq[1].ToString() == "1")
                        {
                            object SalesType = "pos";
                            object InvoiceNumber = "";
                            object InvoiceDate = "";
                            object LineType = headerSeq[1];
                            object ItemNumber = headerSeq[0];
                            object Description = headerSeq[2];
                            object Warehouse = headerSeq[3];
                            object QuantityOrdered = headerSeq[4];
                            object UnitPrice = headerSeq[5];
                            object UnitCost = headerSeq[6];
                            object ExtensionAmt = headerSeq[7];
                            


                            if (decimal.Parse(QuantityOrdered.ToString()) <= 0)
                            {
                            }

                            //orderList.Add(ItemNumber.ToString(), int.Parse(QuantityOrdered.ToString().Split('.')[0]));

                            if (SalesType == null)
                            {
                                SalesType = "";
                            }
                            if (InvoiceNumber == null)
                            {
                                InvoiceNumber = "";
                            }
                            if (InvoiceDate == null)
                            {
                                InvoiceDate = "";
                            }
                            if (LineType == null)
                            {
                                LineType = "";
                            }
                            if (ItemNumber == null)
                            {
                                ItemNumber = "";
                            }
                            if (Description == null)
                            {
                                Description = "";
                            }
                            if (Warehouse == null)
                            {
                                Warehouse = "";
                            }
                            if (QuantityOrdered == null)
                            {
                                QuantityOrdered = "";
                            }
                            if (UnitPrice == null)
                            {
                                UnitPrice = "";
                            }
                            if (UnitCost == null)
                            {
                                UnitCost = "";
                            }
                            if (ExtensionAmt == null)
                            {
                                ExtensionAmt = "";
                            }
                            if (AmountDue == null)
                            {
                                AmountDue = "0";
                            }

                            tmpTx.AmountDue = AmountDue.ToString();
                            tmpTx.SalesOrderNo = TransactionNo.ToString();
                            tmpTx.SalesOrderDate = TransactionDate.ToString();

                            tmpTx.SalesType = SalesType.ToString();
                            tmpTx.InvoiceNumber = InvoiceNumber.ToString();
                            tmpTx.InvoiceDate = InvoiceDate.ToString();
                            tmpTx.LineType = LineType.ToString();
                            tmpTx.ItemNumber = ItemNumber.ToString();
                            tmpTx.Description = Description.ToString();
                            tmpTx.Warehouse = Warehouse.ToString();
                            tmpTx.Quantity = QuantityOrdered.ToString();
                            tmpTx.UnitPrice = UnitPrice.ToString();
                            tmpTx.UnitCost = UnitCost.ToString();
                            tmpTx.Extension = ExtensionAmt.ToString();
                            if (tmpTx.SalesOrderNo == "0231847")
                            {
                            }
                        }
                        if (headerSeq[1].ToString() == "3")
                        {
                            object SalesType = "pos";
                            object InvoiceNumber = "";
                            object InvoiceDate = "";
                            object LineType = headerSeq[1];
                            object ChargeCode = headerSeq[0];
                            object Description = headerSeq[2];
                            object ExtensionAmt = headerSeq[7];
                            tmpTx.SalesType = SalesType.ToString();
                            tmpTx.InvoiceNumber = InvoiceNumber.ToString();
                            tmpTx.InvoiceDate = InvoiceDate.ToString();
                            tmpTx.LineType = LineType.ToString();
                            tmpTx.ChargeCode = ChargeCode.ToString();
                            tmpTx.Description = Description.ToString();
                            tmpTx.Extension = ExtensionAmt.ToString();
                        }

                        if (TransactionNo.ToString() == "0232492")
                        {
                        }


                        if (tmpTx.ItemNumber != null)
                        {
                            if (tmpTx.ItemNumber.Length > 0)
                            {
                                foreach (string itm in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                                {
                                    if (tmpTx.ItemNumber.Trim().ToUpper() == itm)
                                    {
                                        //tmpSalesReportData.Add(tmpTx);
                                        if (DateTime.Parse(tmpTx.SalesOrderDate) >= DateTime.Parse(startDate.Value.ToShortDateString()) & DateTime.Parse(tmpTx.SalesOrderDate) <= DateTime.Parse(endDate.Value.ToShortDateString()))
                                        {
                                            if (decimal.Parse(tmpTx.Quantity) > 0)
                                            {
                                                
                                                    SalesReportData.Add(tmpTx);
                              
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        //THIS IS WERE I LEFT OFF... UGH!



                        //MessageBox.Show(test.ToString());
                    }
                    int totQty = 0;
                    SalesTransactions currentTx = new SalesTransactions();
                    Dictionary<string, int> orderList = new Dictionary<string, int>();
                    

                    foreach (SalesTransactions st in tmpSalesReportData)
                    {                        
                        try { orderList.Add(st.ItemNumber,int.Parse(st.Quantity.Split('.')[0])); }
                        catch { orderList[st.ItemNumber] += int.Parse(st.Quantity.Split('.')[0]); }
                    }

                    foreach (SalesTransactions st in tmpSalesReportData)
                    {
                        if (orderList[st.ItemNumber] >= int.Parse(saleItemCount.Text))
                        {
                            SalesReportData.Add(st); 
                        }
                    }

                }

            }

            //now filter the orders?
        }
        private void RetrieveUnknownsFromHistory()
        {
            //call #3... get unknown history from between the two selected dates on the form...
            statusLbl.Text = "Status: Retrieving Unknowns From History Header Data From MAS200";
            statusLbl.Refresh();
            string unknownFromHistory = Connectivity.MASCalls.Retrieve("SELECT InvoiceNo,HeaderSeqNo,InvoiceDate,SalesOrderNo,OrderDate,BillToName FROM AR_InvoiceHistoryHeader WHERE InvoiceDate>={ d '" + startDate.Value.Year.ToString() + "-" + startDate.Value.Month.ToString() + "-" + startDate.Value.Day.ToString() + "'} AND OrderDate<={ d '" + endDate.Value.Year.ToString() + "-" + endDate.Value.Month.ToString() + "-" + endDate.Value.Day.ToString() + "'} AND InvoiceType<>'XD' AND SalesOrderNo IS NULL AND (SourceJournal<>'P2' OR InvoiceType<>'IN')");
            if (unknownFromHistory == null)
            { return; }
            ReportData.UnknownFromHistory = JSONtoDictionary(unknownFromHistory);

            //int counter = 0;
            foreach (System.Collections.ArrayList unknownLines in (System.Collections.ArrayList)ReportData.UnknownFromHistory["rows"])
            {
                //counter++;
                //statusLbl.Text = "Status: Processing Unknowns (" + counter.ToString() + "/" + ((System.Collections.ArrayList)ReportData.UnknownFromHistory["rows"]).Count.ToString() + ")";
                //statusLbl.Refresh();
                System.Collections.ArrayList headerSeq = (System.Collections.ArrayList)unknownLines[0];
                SalesTransactions tmpTx = new SalesTransactions();

                object SalesType = "unknown";
                object SalesOrderNo = "";
                object SalesOrderDate = headerSeq[2];
                object SalesOrderStatus = "";
                object CustomerName = headerSeq[5];
                object HasInvoice = "0";
                object HasOpen = "0";
                object Lines = "";
                object InvoiceNo = headerSeq[0];
                object HeaderSeqNo = headerSeq[1];

                //statusLbl.Text = "Status: Retrieving Invoiced Line From MAS200 (perhaps faster)";
                //statusLbl.Refresh();
                string invoiceHistoryDetail = Connectivity.MASCalls.Retrieve("SELECT ItemCode,ItemType,ItemCodeDesc,WarehouseCode,QuantityShipped,QuantityOrdered,QuantityBackordered,UnitPrice,UnitCost,ExtensionAmt FROM AR_InvoiceHistoryDetail WHERE InvoiceNo='" + InvoiceNo.ToString() + "' AND HeaderSeqNo='" + HeaderSeqNo.ToString() + "' AND (ItemType='1' OR ItemType='3')");
                Dictionary<string, dynamic> invHistory = JSONtoDictionary(invoiceHistoryDetail);
                if (invHistory != null)
                {
                    tmpTx.Lines = new List<SalesTransactions>();
                    //int counter2 = 0;
                    foreach (System.Collections.ArrayList unknownLines2 in (System.Collections.ArrayList)invHistory["rows"])
                    {
                        //counter++;
                        //statusLbl.Text = "Status: Processing Invoiced Lines (" + counter2.ToString() + "/" + ((System.Collections.ArrayList)invHistory["rows"]).Count.ToString() + ")";
                        //statusLbl.Refresh();
                        System.Collections.ArrayList unknownSeq = (System.Collections.ArrayList)unknownLines2[0];
                        SalesTransactions newTmpTx = new SalesTransactions();
                        if (unknownSeq[1].ToString() == "1")
                        {
                            newTmpTx.SalesType = "unknown";
                            newTmpTx.InvoiceNumber = "";
                            newTmpTx.InvoiceDate = "";
                            newTmpTx.LineType = unknownSeq[1].ToString();
                            newTmpTx.ItemNumber = unknownSeq[0].ToString();
                            newTmpTx.Description = unknownSeq[2].ToString();
                            newTmpTx.Warehouse = unknownSeq[3].ToString();
                            newTmpTx.Quantity = unknownSeq[5].ToString();
                            newTmpTx.UnitPrice = unknownSeq[7].ToString();
                            newTmpTx.UnitCost = unknownSeq[8].ToString();
                        }
                        if (unknownSeq[1].ToString() == "3")
                        {
                            newTmpTx.SalesType = "unknown";
                            newTmpTx.InvoiceNumber = "";
                            newTmpTx.InvoiceDate = "";
                            newTmpTx.LineType = unknownSeq[1].ToString();
                            newTmpTx.ChargeCode = unknownSeq[0].ToString();
                            newTmpTx.Description = unknownSeq[2].ToString();
                            newTmpTx.Extension = unknownSeq[9].ToString();
                        }

                        tmpTx.Lines.Add(newTmpTx);

                    }


                    if (tmpTx.ItemNumber.Length > 0)
                    {
                        foreach (string itm in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                        {
                            if (itm.Trim().ToUpper() == tmpTx.ItemNumber.Trim().ToUpper())
                            {
                                SalesReportData.Add(tmpTx);
                                break;
                            }
                        }

                    }
                }
              

            }

            

            //now filter the data.




            
        }



        private void WriteToLVI(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                ListViewItem lvi = new ListViewItem(row[0].ToString());
                for (int i = 1; i < dataTable.Columns.Count; i++)
                {
                    lvi.SubItems.Add(row[i].ToString());
                }
                reportLVI.Items.Add(lvi);
            }
        }
        private decimal getMAPP(string ItemNumber)
        {
            try
            {
                List<string> itm = Connectivity.SQLCalls.sqlQuery("SELECT MAPP FROM InventoryMaster WHERE ItemNumber='" + ItemNumber + "'");
                return decimal.Parse(itm[0].Split('|')[0]);
            }
            catch { return 0; }

        }
        private decimal getINETPrice(string ItemNumber)
        {
            try
            {
                List<string> itm = Connectivity.SQLCalls.sqlQuery("SELECT IDiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" + ItemNumber + "'");
                return decimal.Parse(itm[0].Split('|')[0]);
            }
            catch
            {
                return 0;
            }
        }
        private decimal getStorePrice(string ItemNumber)
        {
            try
            {
                List<string> itm = Connectivity.SQLCalls.sqlQuery("SELECT DiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" + ItemNumber + "'");
                return decimal.Parse(itm[0].Split('|')[0]);
            }
            catch { return 0; }
        }
        
        private Dictionary<string, dynamic> JSONtoDictionary(string json)
        {
            //statusLbl.Text = "Status: Parsing JSON...";
            //statusLbl.Refresh();
            if (json == null)
            {
                return new Dictionary<string, dynamic>();
            }
            var jss = new JavaScriptSerializer();
            var dict = jss.Deserialize<Dictionary<string, dynamic>>(json);
            return (Dictionary<string,dynamic>)dict;
        }

        private DataTable ParseResults(string results)
        {
            statusLbl.Text = "Status: Parsing Results";
            statusLbl.Refresh();
            DataTable tmpTable = new DataTable();
            string columnChunk = results.Split(new string[]{"{\"cols\":["},StringSplitOptions.None)[1].Split(']')[0];
            
            foreach (string column in columnChunk.Split(new string[] {"{\"name\":\""},StringSplitOptions.RemoveEmptyEntries))
            {
                string columnName = column.Split('"')[0];
                tmpTable.Columns.Add(columnName);
            }

            string rowsChunk = results.Split(new string[] { "rows\":[" }, StringSplitOptions.None)[1].Split(new string[] { "]}" }, StringSplitOptions.None)[0];

            foreach (string row in rowsChunk.Split(new string[] {"],"},StringSplitOptions.RemoveEmptyEntries))
            {
                int colCount = 0;
                
                object[] rowArray = new object[tmpTable.Columns.Count];
                foreach (DataColumn dc in tmpTable.Columns)
                {
                    rowArray[colCount] = row.Split(',')[colCount].Replace("[","").Replace("\"","");
                    colCount++;
                }
                tmpTable.Rows.Add(rowArray);
            }

            return tmpTable;
        }



        private void GetInvoicedLines(string SalesOrderNo, string OrderDate, string CustomerName)
        {
            //statusLbl.Text = "Status: Retrieving Invoiced Lines Header Data From MAS200 (perhaps a faster call)";
            //statusLbl.Refresh();
            if (SalesOrderNo == "0070769")
            {
            }
            string invoicedLines1 = Connectivity.MASCalls.Retrieve("SELECT InvoiceNo,HeaderSeqNo,InvoiceDate FROM AR_InvoiceHistoryHeader WHERE SalesOrderNo='" + SalesOrderNo + "'");
            Dictionary<string, dynamic> invHeader = JSONtoDictionary(invoicedLines1);
            List<SalesTransactions> tmpSalesReportData = new List<SalesTransactions>();
            //Dictionary<string, int> salesData = new Dictionary<string, int>();
            //need to capture HeaderSeqNo & InvoiceNo
            System.Collections.ArrayList HeaderArray = invHeader["rows"];
            System.Collections.ArrayList headerSeq = (System.Collections.ArrayList)HeaderArray[0];
            object header = headerSeq[1];
            object invoice = headerSeq[0];
            object invoicedate = headerSeq[2];
            string HeaderSeqNo = header.ToString();
            string InvoiceNo = invoice.ToString();
            string InvoiceDate = invoicedate.ToString();
            
            //MessageBox.Show("Finally damnit, the HeaderSeqNo=" + HeaderSeqNo + " and the InvoiceNo=" + InvoiceNo);
            //statusLbl.Text = "Status: Retrieving Invoiced Line Data From MAS200 (perhaps maybe even a bit faster)";
            //statusLbl.Refresh();
            string invoicedLines2 = Connectivity.MASCalls.Retrieve("SELECT ItemCode,ItemType,ItemCodeDesc,WarehouseCode,QuantityShipped,QuantityOrdered,QuantityBackordered,UnitPrice,UnitCost,ExtensionAmt FROM AR_InvoiceHistoryDetail WHERE InvoiceNo='" + InvoiceNo + "' AND HeaderSeqNo='" + HeaderSeqNo + "' AND (ItemType='1' OR ItemType='3')");
            Dictionary<string, dynamic> invLines = JSONtoDictionary(invoicedLines2);
            decimal invTotal = 0;
            //int counter = 0;
            foreach (System.Collections.ArrayList invLinesAL in (System.Collections.ArrayList)invLines["rows"])
            {
                //counter++;
                //statusLbl.Text = "Status: Processing Invoiced Line (" + counter.ToString() + "/" + ((System.Collections.ArrayList)invLines["rows"]).Count.ToString() + ")";
                //statusLbl.Refresh();
                object ItemCode = invLinesAL[0];
                object ItemType = invLinesAL[1];
                object ItemCodeDesc = invLinesAL[2];
                object WarehouseCode = invLinesAL[3];
                object QuantityShipped = invLinesAL[4];
                object QuantityOrdered = invLinesAL[5];
                object QuantityBackordered = invLinesAL[6];
                object UnitPrice = invLinesAL[7];
                object UnitCost = invLinesAL[8];
                object ExtensionAmt = invLinesAL[9];

                if (ItemCode == null)
                {
                    ItemCode = "";
                }
                if (ItemType == null)
                {
                    ItemType = "";
                }
                if (ItemCodeDesc == null)
                {
                    ItemType = "";
                }
                if (WarehouseCode == null)
                {
                    WarehouseCode = "";
                }
                if (QuantityShipped == null)
                {
                    QuantityShipped = "";
                }
                if (QuantityOrdered == null)
                {
                    QuantityOrdered = "";
                }
                if (QuantityBackordered == null)
                {
                    QuantityBackordered = "";
                }
                if (UnitPrice == null)
                {
                    UnitPrice = "";
                }
                if (UnitCost == null)
                {
                    UnitCost = "";
                }
                if (ExtensionAmt == null)
                {
                    ExtensionAmt = "";
                }
                SalesTransactions tmpTx = new SalesTransactions();
                tmpTx.ChargeCode = "";
                tmpTx.CustomerName = CustomerName;
                tmpTx.Description = "";
                tmpTx.Extension = "";
                tmpTx.HasInvoices = "";
                tmpTx.HasOpen = "";
                tmpTx.InvoiceDate = "";
                tmpTx.InvoiceNumber = "";
                tmpTx.ItemNumber = "";
                tmpTx.LineType = "";
                tmpTx.Quantity = "";
                tmpTx.SalesOrderDate = OrderDate;
                tmpTx.SalesOrderNo = SalesOrderNo;
                tmpTx.SalesOrderStatus = "";
                tmpTx.SalesType = "s/o (invoiced)";
                tmpTx.UnitCost = "";
                tmpTx.UnitPrice = "";
                tmpTx.Warehouse = "";


                if (ItemType.ToString() == "1")
                {
                    //tmpTx.SalesType = "invoiced";
                    tmpTx.InvoiceNumber = InvoiceNo + "-" + HeaderSeqNo;
                    tmpTx.InvoiceDate = InvoiceDate;
                    tmpTx.LineType = ItemType.ToString();
                    tmpTx.ItemNumber = ItemCode.ToString();
                    tmpTx.Description = ItemCodeDesc.ToString();
                    tmpTx.Warehouse = WarehouseCode.ToString();
                    //tmpTx.Quantity = (decimal.Parse(QuantityShipped.ToString())).ToString();
                    tmpTx.Quantity = QuantityOrdered.ToString();
                    //tmpTx.Quantity = ((int)((decimal.Parse(QuantityOrdered.ToString()) - decimal.Parse(QuantityBackordered.ToString())))).ToString();
                    tmpTx.UnitPrice = UnitPrice.ToString();
                    tmpTx.UnitCost = UnitCost.ToString();
                }
                if (ItemType.ToString() == "3")
                {
                    //tmpTx.SalesType = "invoiced";
                    tmpTx.InvoiceNumber = InvoiceNo + "-" + HeaderSeqNo;
                    tmpTx.InvoiceDate = InvoiceDate;
                    tmpTx.LineType = ItemType.ToString();
                    tmpTx.ChargeCode = ItemCode.ToString();
                    tmpTx.Description = ItemCodeDesc.ToString();
                    tmpTx.Extension = ExtensionAmt.ToString();
                }

                

                //THIS MAKES LINE ONLY SHOW IF ITEM WAS IN Item List
                if (tmpTx.SalesOrderNo == "0070769")
                {

                }
                
                if (tmpTx.ItemNumber.Length > 0)
                {
                    //so we have invoice number. the proper way to get prices is by exactly what they sold for! so pulling from inventorymaster is a no-no.
                    //we must pull from AR_InvoiceHistoryDetail, columns UnitCost and UnitPrice, respectably.




                    //List<string> pricing = Connectivity.SQLCalls.sqlQuery("SELECT IDiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" + tmpTx.ItemNumber + "'");
                    //if (pricing.Count > 0)
                    //{
                        //invTotal += decimal.Parse(tmpTx.Quantity) * decimal.Parse(tmpTx.UnitPrice);
                        if (decimal.Parse(tmpTx.Quantity) >= 0)
                        {
                            //invTotal += decimal.Parse(pricing[0].Split('|')[0]);
                            invTotal += decimal.Parse(tmpTx.UnitPrice);
                        }
                        else
                        {
                            //invTotal -= decimal.Parse(pricing[0].Split('|')[0]);
                            invTotal -= decimal.Parse(tmpTx.UnitPrice);
                        }
                    //}
                   
                    
                    foreach (string item in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                    {
                        if (tmpTx.ItemNumber.ToUpper().Replace("XXX","").Replace("ZZZ","") == item.ToUpper().Trim())
                        {
                            try
                            {
                                if (tmpTx.Quantity.StartsWith(".") == true)
                                {
                                    tmpTx.Quantity = "0.00";
                                }
                                //salesData.Add(tmpTx.ItemNumber, int.Parse(tmpTx.Quantity.Split('.')[0]));
                               
                            }
                            catch
                            {
                                //salesData[tmpTx.ItemNumber] += int.Parse(tmpTx.Quantity.Split('.')[0]);
                                MessageBox.Show("DAMNIT.");
                            }
                            if (decimal.Parse(tmpTx.Quantity) > 0)
                            {
                                tmpSalesReportData.Add(tmpTx);
                            }
                            
                            //SalesReportData.Add(tmpTx);
                            break;
                        }
                    }

                }
               



            }


            foreach (SalesTransactions st in tmpSalesReportData)
            {
                if (st.InvoiceDate == "")
                {
                    InvoicedLines(tmpSalesReportData, st);
                }
                else
                {
                    try
                    {
                        if ((DateTime.Parse(st.SalesOrderDate) <= DateTime.Parse(endDate.Value.ToShortDateString())))
                        {
                            if ((DateTime.Parse(st.SalesOrderDate) >= DateTime.Parse(startDate.Value.ToShortDateString())))
                            {
                                InvoicedLines(tmpSalesReportData, st);

                            }
                        }
                    }
                    catch { }
                }
                //tmpSalesReportData.Remove(st);
                

            }
            


        }

        private void InvoicedLines(List<SalesTransactions> tmpSalesReportData, SalesTransactions st)
        {
            SalesTransactions edt = st;
            decimal totPrice = 0;
            decimal totQty = 0;

            if (st.SalesOrderNo == "0071013")
            {
            }

            foreach (SalesTransactions st2 in tmpSalesReportData)
            {
                totQty += decimal.Parse(st2.Quantity);
                totPrice += decimal.Parse(st2.UnitPrice) * decimal.Parse(st2.Quantity);
            }
            if (totPrice >= decimal.Parse(orderMinimum.Text))
            {
                edt.AboveOrderMin = true;
            }
            else
            {
                edt.AboveOrderMin = false;
            }

            if (totQty > 0)
            //if (decimal.Parse(edt.Quantity) > 0)
            {
                //add to report
                SalesReportData.Add(edt);
            }
        }



        private void GetOpenLines(string SalesOrderNo)
        {
            if (SalesOrderNo == "Y356195" | SalesOrderNo == "Y356419")
            {
            }
            //statusLbl.Text = "Status: Retrieving Open Lines Header Data From MAS200 (perhaps as fast as the rest)";
            //statusLbl.Refresh();
            string headerData = Connectivity.MASCalls.Retrieve("SELECT OrderDate,BillToName FROM SO_SalesOrderHeader WHERE SalesOrderNo='" + SalesOrderNo + "' AND CurrentInvoiceNo IS NULL");
            //if (headerData == null)
            //{
             //   return;
            //}
            Dictionary<string, dynamic> CurrentHeaderInfo = JSONtoDictionary(headerData);
            System.Collections.ArrayList headerInfoArray = CurrentHeaderInfo["rows"];
            bool isHistorical = false;
            if (headerInfoArray.Count == 0)
            {
                //wait a second, we need info since we have said order in list.. try checking the history for non invoicers..
                //return;
                headerData = Connectivity.MASCalls.Retrieve("SELECT OrderDate,BillToName FROM SO_SalesOrderHistoryHeader WHERE SalesOrderNo='" + SalesOrderNo + "' AND LastInvoiceNo IS NULL AND OrderStatus='X'");
                if (headerData == null)
                {
                    return;
                }
                CurrentHeaderInfo = JSONtoDictionary(headerData);
                headerInfoArray = CurrentHeaderInfo["rows"];
                isHistorical = true;
            }
            if (headerInfoArray.Count == 0)
            {
                return;
            }

            string openLines = "";
            if (isHistorical == true)
            {
                return;
                openLines = Connectivity.MASCalls.Retrieve("SELECT ItemCode,ItemType,ItemCodeDesc,WarehouseCode,QuantityOrderedOriginal,QuantityShipped,QuantityBackordered,OriginalUnitPrice,UnitCost,LastExtensionAmt FROM SO_SalesOrderHistoryDetail WHERE SalesOrderNo='" + SalesOrderNo + "' AND (ItemType='1' OR ItemType='3')");
                if (SalesOrderNo == "Y356195" | SalesOrderNo == "Y356419")
                {
                }
            }
            else
            {
                openLines = Connectivity.MASCalls.Retrieve("SELECT ItemCode,ItemType,ItemCodeDesc,WarehouseCode,QuantityOrdered,QuantityShipped,QuantityBackordered,UnitPrice,UnitCost,ExtensionAmt FROM SO_SalesOrderDetail WHERE SalesOrderNo='" + SalesOrderNo + "' AND (ItemType='1' OR ItemType='3')");
            }
            Dictionary<string, dynamic> currentOpenLines = JSONtoDictionary(openLines);
            System.Collections.ArrayList HeaderArray = currentOpenLines["rows"];
            List<SalesTransactions> tmpSalesReportData = new List<SalesTransactions>();
            Dictionary<string, int> loadedLines = new Dictionary<string, int>();


            //int counter = 0;
            foreach (System.Collections.ArrayList headerSeq in (System.Collections.ArrayList)HeaderArray)
            {
                //statusLbl.Text = "Status: Processing Invoiced Lines Header Data (" + counter.ToString() + "/" + ((System.Collections.ArrayList)HeaderArray).Count.ToString() + ")";
                //statusLbl.Refresh();

               // counter++;
                object ItemCode = headerSeq[0];
                object ItemType = headerSeq[1];
                object ItemCodeDesc = headerSeq[2];
                object Warehousecode = headerSeq[3];
                object QuantityOrdered = headerSeq[4];
                object QuantityShipped = headerSeq[5];
                object QuantityBackordered = headerSeq[6];
                object UnitPrice = headerSeq[7];
                object UnitCost = headerSeq[8];
                object ExtensionAmt = headerSeq[9];

                SalesTransactions tmpTx = new SalesTransactions();
                tmpTx.ChargeCode = "";
                tmpTx.CustomerName = (((System.Collections.ArrayList)headerInfoArray[0])[1] == null?"N/A":((System.Collections.ArrayList)headerInfoArray[0])[1].ToString());
                tmpTx.Description = "";
                tmpTx.Extension = "";
                tmpTx.HasInvoices = "";
                tmpTx.HasOpen = "";
                tmpTx.InvoiceDate = "";
                tmpTx.InvoiceNumber = "";
                tmpTx.ItemNumber = "";
                tmpTx.LineType = "";
                tmpTx.Quantity = "";
                tmpTx.SalesOrderDate = (((System.Collections.ArrayList)headerInfoArray[0])[0] == null ? "N/A" : ((System.Collections.ArrayList)headerInfoArray[0])[0].ToString());
                tmpTx.SalesOrderNo = SalesOrderNo;
                tmpTx.SalesOrderStatus = "";                
                tmpTx.UnitCost = "";
                tmpTx.UnitPrice = "";
                tmpTx.Warehouse = "";
                tmpTx.SalesType = "s/o (open)";


                if (ItemType.ToString() == "1")
                {
                   // tmpTx.SalesType = "open";
                    tmpTx.InvoiceNumber = "";
                    tmpTx.InvoiceDate = "";
                    tmpTx.LineType = ItemType.ToString();
                    tmpTx.ItemNumber = ItemCode.ToString();
                    tmpTx.Description = ItemCodeDesc.ToString();
                    tmpTx.Warehouse = Warehousecode.ToString();
                    //tmpTx.Quantity = ((int)(decimal.Parse(QuantityOrdered.ToString()) - decimal.Parse(QuantityShipped.ToString()))).ToString();
                    //tmpTx.Quantity = ((int)(decimal.Parse(QuantityShipped.ToString()))).ToString();
                    tmpTx.Quantity = ((int)(decimal.Parse(QuantityOrdered.ToString()))).ToString();
                    tmpTx.UnitPrice = UnitPrice.ToString();
                    tmpTx.UnitCost = UnitCost.ToString();
                }
                if (ItemType.ToString() == "3")
                {
                    //tmpTx.SalesType = "open";
                    tmpTx.InvoiceNumber = "";
                    tmpTx.InvoiceDate = "";
                    tmpTx.LineType = ItemType.ToString();
                    tmpTx.ChargeCode = ItemCode.ToString();
                    tmpTx.Description = ItemCodeDesc.ToString();
                    tmpTx.Extension = ExtensionAmt.ToString();
                }
                if (tmpTx.Quantity.StartsWith(".")==true | tmpTx.Quantity == "")
                {
                    tmpTx.Quantity = "0.00";
                }

                try { loadedLines.Add(tmpTx.ItemNumber, int.Parse(tmpTx.Quantity.ToString().Split('.')[0])); }
                catch {loadedLines[tmpTx.ItemNumber] = int.Parse(tmpTx.Quantity.ToString().Split('.')[0]); }
                
                //original script leaves out backordered items.. we can add checkbox for this feature to turn it on...
                //tmpTx.Quantity = (decimal.Parse(tmpTx.Quantity) - decimal.Parse(QuantityBackordered.ToString())).ToString();
            




                decimal invTotal = 0;
                if (tmpTx.ItemNumber.Length > 0)
                {
                    //List<string> pricing = Connectivity.SQLCalls.sqlQuery("SELECT IDiscountMarkupPriceRate1 FROM InventoryMaster WHERE ItemNumber='" + tmpTx.ItemNumber + "'");
                    //if (pricing.Count > 0)
                   // {
                        // invTotal += decimal.Parse(tmpTx.Quantity) * decimal.Parse(tmpTx.UnitPrice);
                        if (decimal.Parse(tmpTx.Quantity) >= 0)
                        {
                            invTotal += decimal.Parse(tmpTx.UnitPrice.ToString());
                        }
                        else
                        {
                            invTotal -= decimal.Parse(tmpTx.UnitPrice.ToString());
                        }
                   // }


                    foreach (string item in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                    {
                        if (tmpTx.ItemNumber.ToUpper().Replace("XXX","").Replace("ZZZ","") == item.ToUpper().Trim())
                        {
                            try
                            {
                                if (tmpTx.Quantity.StartsWith(".") == true)
                                {
                                    tmpTx.Quantity = "0.00";
                                }
                                //salesData.Add(tmpTx.ItemNumber, int.Parse(tmpTx.Quantity.Split('.')[0]));

                            }
                            catch
                            {
                                //salesData[tmpTx.ItemNumber] += int.Parse(tmpTx.Quantity.Split('.')[0]);
                                MessageBox.Show("DAMNIT.");
                            }
                           
                                tmpSalesReportData.Add(tmpTx);
                            
                            //SalesReportData.Add(tmpTx);
                            break;
                        }
                    }

                }
                


            }
            //Need to add code here to filter out sos/pos from the above foreach right there... and use loadedLines...
            foreach (SalesTransactions st in tmpSalesReportData)
            {
                if (st.SalesOrderNo == "0070769")
                {
                }
                if (st.InvoiceDate == "")
                {
                    if (int.Parse(st.Quantity) > 0)
                    {
                        OpenLines(tmpSalesReportData, st);
                    }
                }
                else
                {
                    try
                    {
                        if ((DateTime.Parse(st.InvoiceDate) <= DateTime.Parse(endDate.Value.ToShortDateString())))
                        {
                            if ((DateTime.Parse(st.InvoiceDate) >= DateTime.Parse(startDate.Value.ToShortDateString())))
                            {
                                if (int.Parse(st.Quantity) > 0)
                                {
                                    OpenLines(tmpSalesReportData, st);
                                }
                            }
                        }
                    }
                    catch { }
                }

                
                        


            }
            

        }

        private void OpenLines(List<SalesTransactions> tmpSalesReportData, SalesTransactions st)
        {
            SalesTransactions edt = st;
            //tmpSalesReportData.Remove(st);
            if (st.SalesOrderNo == "Y356195" | st.SalesOrderNo == "Y356419")
            {
            }
            decimal totPrice = 0;
            decimal totQty = 0;
            foreach (SalesTransactions st2 in tmpSalesReportData)
            {
                totQty += decimal.Parse(st2.Quantity);
                totPrice += decimal.Parse(st2.UnitPrice) * decimal.Parse(st2.Quantity);
            }
            if (totPrice >= decimal.Parse(orderMinimum.Text))
            {
                edt.AboveOrderMin = true;
            }
            else
            {
                edt.AboveOrderMin = false;
            }
            //edt.SalesType += "(open)";
            if (totQty > 0)
            {
                //add to report
                SalesReportData.Add(edt);
            }
        }



        private void saveBtn_Click(object sender, EventArgs e)
        {
            LVItoCSV(@"c:\users\tomwestbrook\desktop\salesReportOutput.csv", reportLVI);
        }

        private void newReportBtn_Click(object sender, EventArgs e)
        {
            generateBtn.Enabled = true;
            newReportBtn.Visible = false;
        }

    }
}
