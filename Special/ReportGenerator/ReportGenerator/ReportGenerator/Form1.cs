using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace ReportGenerator
{
    public partial class Form1 : Form
    {
        public bool debugMode = true;

        public System.Threading.Thread thread;
        public System.Threading.Thread thread_POBO;
        public System.Threading.Thread thread_STX;
        public bool runThreadBeforeLogin = false;
        public bool ranThread = false;
        public bool ERPConnected = false;


        public string currentReportTitle = "";
        public string currentReportLogoText = "Tools Plus Inc.";

        public Form1()
        {
            InitializeComponent();
        }

        private void poBackorderRadio_CheckedChanged(object sender, EventArgs e)
        {
            ClearLVI();

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
        private void DisableCreateButtons()
        {
            poboCreateBtn.Enabled = false;
            salesTxCreateBtn.Enabled = false;
        }
        private void EnableCreateButtons()
        {
            poboCreateBtn.Enabled = true;
            salesTxCreateBtn.Enabled = true;
        }
        private void StartThread()
        {
            thread = new System.Threading.Thread(() =>
            {
                ERPConnected = Connectivity.MASCalls.preloadSpeedBoost();
            });
            thread.Start();
        }

        private void ResizeForm()
        {
            int BorderSpace = 10;
            int ReportListWidthPercentage = 27;
            int PreviewBarHeight = 24;
            int ReportTreeTopOffset = 10;

      

            reportTreeBox.Left = BorderSpace;
            reportTreeBox.Top = BorderSpace;
            reportTreeBox.Width = (int)((decimal)(this.ClientSize.Width) * (decimal)(0.01 * ReportListWidthPercentage));            
            reportTreeBox.Height = previewBtn.Text == "s" ? (this.ClientSize.Height - (BorderSpace * 2))-previewBtn.Height : this.ClientSize.Height / 2;
            reportTreeBox.Refresh();
            reportTree.Left = BorderSpace / 2;
            reportTree.Top = BorderSpace / 2 + ReportTreeTopOffset;
            reportTree.Width = reportTreeBox.Width - ((BorderSpace / 2) * 2);
            reportTree.Height = reportTreeBox.Height - ((BorderSpace / 2) * 2) - ReportTreeTopOffset;
            reportTree.Refresh();

            previewBtn.Top = reportTreeBox.Bottom + BorderSpace;
            previewBtn.Left = BorderSpace;
            previewBtn.Width = this.ClientSize.Width - (BorderSpace * 2);
            previewBtn.Height = PreviewBarHeight;
            previewBtn.Refresh();

            statusHeaderLbl.Left = reportTreeBox.Right + BorderSpace;
            statusHeaderLbl.Top = BorderSpace;
            statusHeaderLbl.AutoSize = true;
            statusHeaderLbl.Refresh();

            statusLbl.Left = statusHeaderLbl.Right;
            statusLbl.Top = BorderSpace;
            statusLbl.AutoSize = true;
            statusLbl.Refresh();

            /*reportTitleLbl.AutoSize = true;
            reportTitleLbl.Refresh();
            reportTitleLbl.Left = BorderSpace*5;
            reportTitleLbl.Top = previewBtn.Bottom + BorderSpace;
            reportTitleLbl.Refresh();

            reportLVI.Left = BorderSpace;
            reportLVI.Top = reportTitleLbl.Bottom + BorderSpace;
            reportLVI.Height = previewBtn.Text == "s" ? (this.ClientSize.Height / 2) - (BorderSpace * 2) : (this.ClientSize.Height / 2) - (BorderSpace * 2);
            reportLVI.Width = this.ClientSize.Width - (BorderSpace * 2);
            reportLVI.Refresh();

            exportPnl.Top = reportTitleLbl.Top;            
            exportPnl.AutoSize = true;
            exportPnl.Refresh();
            exportPnl.Left = reportLVI.Right - exportPnl.Width;
            exportPnl.Refresh();*/

            reportPnl.Top = previewBtn.Bottom + BorderSpace;
            reportPnl.Left = BorderSpace;
            reportPnl.Width = this.ClientSize.Width - (BorderSpace * 2);
            reportPnl.Height = (this.ClientSize.Height - BorderSpace) - (previewBtn.Bottom + BorderSpace);
            reportPnl.Refresh();
            mapPnl.Top = previewBtn.Bottom + BorderSpace;
            mapPnl.Left = BorderSpace;
            mapPnl.Width = this.ClientSize.Width - (BorderSpace * 2);
            mapPnl.Height = (this.ClientSize.Height - BorderSpace) - (previewBtn.Bottom + BorderSpace);
            mapPnl.Refresh();


            splashBox.Left = reportTreeBox.Right + BorderSpace;
            splashBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            splashBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            splashBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            splashBox.Refresh();

            orderDiscrepancyBox.Left = reportTreeBox.Right + BorderSpace;
            orderDiscrepancyBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            orderDiscrepancyBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            orderDiscrepancyBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            orderDiscrepancyBox.Refresh();

            salesTxBox.Left = reportTreeBox.Right + BorderSpace;
            salesTxBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            salesTxBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            salesTxBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            salesTxBox.Refresh();

            POBackorderBox.Left = reportTreeBox.Right + BorderSpace;
            POBackorderBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            POBackorderBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            POBackorderBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            POBackorderBox.Refresh();

            itemSalesGeoBox.Left = reportTreeBox.Right + BorderSpace;
            itemSalesGeoBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            itemSalesGeoBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            itemSalesGeoBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            itemSalesGeoBox.Refresh();

            this.ClientSize = new System.Drawing.Size(splashBox.Right + BorderSpace, previewBtn.Bottom + 3);
            this.Refresh();

            salesTxCreateBtn.Width = 100;
            salesTxCreateBtn.Height = 50;
            salesTxCreateBtn.Refresh();
            salesTxCreateBtn.Left = salesTxBox.Width - salesTxCreateBtn.Width - BorderSpace;
            salesTxCreateBtn.Top = salesTxBox.Height - salesTxCreateBtn.Height - BorderSpace;
            salesTxCreateBtn.Refresh();

            poboCreateBtn.Width = 100;
            poboCreateBtn.Height = 50;
            poboCreateBtn.Refresh();
            poboCreateBtn.Left = POBackorderBox.Width - poboCreateBtn.Width - BorderSpace;
            poboCreateBtn.Top = POBackorderBox.Height - poboCreateBtn.Height - BorderSpace;
            poboCreateBtn.Refresh();

            orderDiscrepancyCreateBtn.Width = 100;
            orderDiscrepancyCreateBtn.Height = 50;
            orderDiscrepancyCreateBtn.Refresh();
            orderDiscrepancyCreateBtn.Left = orderDiscrepancyBox.Width - orderDiscrepancyCreateBtn.Width - BorderSpace;
            orderDiscrepancyCreateBtn.Top = orderDiscrepancyBox.Height - orderDiscrepancyCreateBtn.Height - BorderSpace;
            orderDiscrepancyCreateBtn.Refresh();

            itemSalesGeoCreateBtn.Width = 100;
            itemSalesGeoCreateBtn.Height = 50;
            itemSalesGeoCreateBtn.Refresh();
            itemSalesGeoCreateBtn.Left = itemSalesGeoBox.Width - itemSalesGeoCreateBtn.Width - BorderSpace;
            itemSalesGeoCreateBtn.Top = itemSalesGeoBox.Height - itemSalesGeoCreateBtn.Height - BorderSpace;
            itemSalesGeoCreateBtn.Refresh();

            packagesShippedBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            packagesShippedBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            packagesShippedBox.Left = reportTreeBox.Right + BorderSpace;
            packagesShippedBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            packagesShippedBox.Refresh();

            p2LinecodeSalesInfoBox.Width = this.ClientSize.Width - splashBox.Left - BorderSpace;
            p2LinecodeSalesInfoBox.Height = previewBtn.Top - splashBox.Top - BorderSpace;
            p2LinecodeSalesInfoBox.Left = reportTreeBox.Right + BorderSpace;
            p2LinecodeSalesInfoBox.Top = statusHeaderLbl.Bottom + BorderSpace;
            p2LinecodeSalesInfoBox.Refresh();


        }

        private TreeView GenerateReportTree(TreeView currTree)
        {
            TreeView tmpTree = currTree;
            tmpTree.Nodes.Clear();
            tmpTree.Refresh();

            
            //Margie's Reports========================
            TreeNode margieNode = new TreeNode();
            margieNode.Text = "Margie's Reports";

            TreeNode poboNode = new TreeNode();
            poboNode.Text = "POs on Backorder";

            margieNode.Nodes.Add(poboNode);
            //////////////////////////////////////////

            //Joe's Reports===========================
            TreeNode joeNode = new TreeNode();
            joeNode.Text = "Joe's Reports";

            TreeNode salesTxNode = new TreeNode();
            salesTxNode.Text = "Sales Transaction";

            TreeNode orderDisNode = new TreeNode();
            orderDisNode.Text = "Order Discrepancy";

            TreeNode p2LinecodeSalesInfoNode = new TreeNode();
            p2LinecodeSalesInfoNode.Text = "P2 Linecode Sales Info";
           
            joeNode.Nodes.Add(salesTxNode);
            joeNode.Nodes.Add(orderDisNode);
            joeNode.Nodes.Add(p2LinecodeSalesInfoNode);
            /////////////////////////////////////////
            
            //Beta Reports===========================
            TreeNode betaNode = new TreeNode();
            betaNode.Text = "Beta Reports";

            TreeNode itemSaleGeoNode = new TreeNode();
            itemSaleGeoNode.Text = "Item Sales by Location";

            TreeNode starshipPackage = new TreeNode();
            starshipPackage.Text = "Starship Packages Shipped";

            betaNode.Nodes.Add(itemSaleGeoNode);
            betaNode.Nodes.Add(starshipPackage);

            tmpTree.Nodes.Add(margieNode);
            tmpTree.Nodes.Add(joeNode);
            tmpTree.Nodes.Add(betaNode);
            tmpTree.ExpandAll();

            return tmpTree;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ResizeForm();
            reportPnl.BringToFront();
            fixStupidWebBrowser();
            
            if (runThreadBeforeLogin == true)
            {
                StartThread();
            }

            if (debugMode == false)
            {
                UserSecurity.LoginFrm();
                this.SendToBack();
                this.Enabled = false;
            }
            
            DisableCreateButtons();
            //statusLbl.Text = "Connecting To ERP System...";
            statusLbl.Text = "User Signing In";
            statusLbl.ForeColor = Color.Red;
            statusLbl.Refresh();
            //thread = new System.Threading.Thread(new System.Threading.ThreadStart(Connectivity.MASCalls.preloadSpeedBoost));            
            reportTree = GenerateReportTree(reportTree);
            splashBox.BringToFront();
            SetToolTips();
        }

        void Form1_Resize(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            ResizeForm();
        }

        private void SetToolTips()
        {
            ToolTip tp = new ToolTip();

            //FORM TIPS
            //tp.SetToolTip(menuRadio, "Return to main menu.");
            //tp.SetToolTip(poBackorderRadio, "Margie's Purchase Orders on Back Order Report.");
            //tp.SetToolTip(salesTxRadio, "Joe Scovill's Sales Transaction History Report.");
            //tp.SetToolTip(previewBtn, "Expand / Retract built in Report Viewer.");


            //POBO REPORT TIPS
            tp.SetToolTip(poboHiddenReportTitleChk, "Check this so when exporting to Excel the column headers can be used to sort/filter results.");
            tp.SetToolTip(poboStartDatePicker, "Pick the date the report should begin on.");
            tp.SetToolTip(poboEndDatePicker, "Pick the date the report should end on.");
            tp.SetToolTip(poboLineCodeStart, "Select the line code to start from.");
            tp.SetToolTip(poboLineCodeEnd, "Select the line code to end on.");
            tp.SetToolTip(poboHiddenReportTitleChk, "Check this to show report title in generated report.");
            tp.SetToolTip(poboCreateBtn, "Generate a new report.");
            //tp.SetToolTip(poboToExcelBtn, "Export report to Excel.");
            
            
            
            
            //STX REPORT TIPS
            tp.SetToolTip(salesTxStartDatePicker, "Pick the date the sale started on.");
            tp.SetToolTip(salesTxEndDatePicker, "Pick the date the sale ended on.");
            tp.SetToolTip(salesTxCouponAmountTxt, "Value of the coupon to be applied.");
            tp.SetToolTip(salesTxOrderMinTxt, "Minimum order total to qualify for coupon.");
            tp.SetToolTip(salesTxMinSaleItemCountTxt, "Minimum amount of items purchased to qualify for coupon.");
            tp.SetToolTip(salesTxItemsTxt, "List all item numbers (not including XXX or ZZZ) that are part of this sale.");
            tp.SetToolTip(salesTxCheckSOChk, "Search through sales history in Sales Order / Accounts Receivable records.");
            tp.SetToolTip(salesTxCheckPOSChk, "Search through sales history in Point Of Sales records.");
            tp.SetToolTip(salesTxCheckOdditiesChk, "Search through sales history in other records.");
            tp.SetToolTip(salesTxIncludeCostChk, "Check off to include Unit Cost column.");
            tp.SetToolTip(salesTxOrdersBelowMinChk, "Check off to include orders less than the specified minimum.");
            tp.SetToolTip(salesTxShowMiscLinesChk, "Check off to include Misc lines.");
            tp.SetToolTip(salesTxShowMiscShipLinesChk, "Check off to include Ship lines (Must also have Misc Lines checked off as well).");
            tp.SetToolTip(salesTxSkipPartiallyInvoicedChk, "Check off to skip partially invoiced lines");
            tp.SetToolTip(salesTxBlankLineBetweenOrdersChk, "Check off to insert blank line between orders");
            tp.SetToolTip(salesTxCreateBtn, "Generate a new report.");
            


            tp.SetToolTip(exportCSVBtn, "Export report to a comma separated value file. This can be opened in Excel.");
            tp.SetToolTip(exportDOCXBtn, "Export report to a doc-x file. This can be opened in Word.");
            tp.SetToolTip(exportEMAILBtn, "Export report into Outlook.");
            tp.SetToolTip(exportEXCELBtn, "Export report into Excel.");
            tp.SetToolTip(exportHTMLBtn, "Export report to an HTML table. Could prove useful.");
            tp.SetToolTip(exportJSONBtn, "Export report to a JSON file. Probably not useful.");
            tp.SetToolTip(exportPDFBtn, "Export report to a PDF file. Requires Excel to be installed on computer to save and Adobe Reader to view.");
            tp.SetToolTip(exportTXTBtn, "Export report to a text file. This can be opened in Notepad or Word.");
            tp.SetToolTip(exportXLSBtn, "Export report to an XLSX File. Can be opened in Excel.");
            tp.SetToolTip(exportXMLBtn, "Export report to an XML file. Possibly useful in the future.");
            tp.SetToolTip(exportWORDBtn, "Export report into Word.");
            
        }

        private void ClearLVI()
        {
            reportLVI.Items.Clear();
            reportLVI.Columns.Clear();
            reportLVI.Clear();
            reportLVI.FullRowSelect = true;
            reportLVI.GridLines = true;
            reportLVI.Refresh();            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            thread_POBO = new System.Threading.Thread(new System.Threading.ThreadStart(ThreadReport_POBO));
            thread_POBO.Start();       
           
            /*ClearLVI();
            poboCreateBtn.Enabled = false;
            poboCreateBtn.Refresh();
            currentReportTitle = "PO's on Back Order";
            string lineStart = poboLineCodeStart.Text;
            string lineEnd = poboLineCodeEnd.Text;

            List<string> linecodesToCheck = new List<string>();
            bool foundStart = false;         
            foreach (string str in poboLineCodeStart.Items)
            {
                if (foundStart == true)
                {
                    if (str == lineEnd)
                    {
                        linecodesToCheck.Add(str);
                        break;
                    }
                    else
                    {
                        linecodesToCheck.Add(str);
                    }
                }
                else
                {
                    if (str == lineStart)
                    {
                        foundStart = true;
                        linecodesToCheck.Add(str);
                    }
                }
            }

            //string rawData = "";
            Dictionary<string, string> rawData = new Dictionary<string, string>();
            bool header = false;
            int lineCount = 1;
            foreach (string lineCode in linecodesToCheck)
            {
                string fuckyou = "SELECT ItemCode,PurchaseOrderNo,QuantityOrdered,QuantityReceived,QuantityBackordered,RequiredDate FROM PO_PurchaseOrderDetail WHERE ProductLine='" + lineCode + "' AND QuantityBackordered>0 AND RequiredDate>{d '" + startDate.Value.Year.ToString() + "-" + startDate.Value.Month.ToString("00") + "-" + startDate.Value.Day.ToString("00") + "'} AND RequiredDate<{d '" + endDate.Value.Year.ToString() + "-" + endDate.Value.Month.ToString("00") + "-" + endDate.Value.Day.ToString("00") + "'}";
                string json = Connectivity.MASCalls.Retrieve(fuckyou);

                if (header == false)
                {
                    if (json.Contains("rows\":") == true)
                    {
                        rawData.Add("HEADER", json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[0]);
                        header = true;
                    }
                }

               // if (json.Contains(":[]") == true)
               // {
               //     MessageBox.Show("Did not find nothing for linecode " + lineCode);
               // }
               // else
               // {
               //     MessageBox.Show("Holy Shit we got sumtin! (" + lineCode + ")");
               // }
                if (json.Contains("rows\":") == true)
                {
                    rawData.Add(lineCode,json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[1] + "\r\n");
                    //textBox1.Text += json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[1] +"\r\n\r\n";
                }
                statusLbl.Text = "Working on line " + lineCode + " (" + lineCount.ToString() + "/" + linecodesToCheck.Count.ToString() + ")";
                statusLbl.Refresh();
                lineCount++;

            }

            //process raw data.
            RawDataToLVI(rawData,linecodesToCheck);


            poboCreateBtn.Enabled = true;
            poboCreateBtn.Refresh();
            */

        }



        private void LVIToExcel(ListView myList,bool showTitle)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            ws.Columns.EntireColumn.ColumnWidth = 20;

            int i = 1;
            int i2 = 3;

            if (showTitle == true)
            {
                ws.Cells[1, 1] = currentReportTitle;
                ws.Cells[1, 3] = currentReportLogoText;
            }
            else
            {
                i2 = 2;
            }
            
            foreach (ListViewItem lvi in myList.Items)
            {
                i = 1;
                List<string> columnList = new List<string>();
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    columnList.Add(lvs.Text);       
                }
                for (int x = 0; x < columnList.Count; x++)
                {
                    foreach (ColumnHeader ch in myList.Columns)
                    {
                        if (ch.DisplayIndex == x)
                        {
                            //this should really be changed so we dont overwrite a header a million times over...
                            int y = 0;
                            if (showTitle == true)
                            {
                                y = 2;
                            }
                            else
                            {
                                y = 1;
                            }

                            ws.Cells[y, i].EntireRow.Font.Bold = true;
                            ws.Cells[y,i]=ch.Text;                         
                            ws.Cells[i2, i] = columnList[ch.Index];
                            i++;
                        }
                    }
                }
                i2++;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (reportLVI.Items.Count > 0)
            {
                LVIToExcel(reportLVI,poboHiddenReportTitleChk.Checked);
            }
            else
            {
                MessageBox.Show("There is no data to build a report from!");
            }
        }

        private void connectionTmr_Tick(object sender, EventArgs e)
        {

            if (UserSecurity.IsUserLoggedIn() == true | debugMode == true)
            {
                if (ranThread == false)
                {
                    this.Enabled = true;
                    if (debugMode == false)
                    {
                        UserSecurity.CloseLoginFrm();
                    }
                    statusLbl.Text = "Connecting To ERP Systems.."; 
                    statusLbl.Refresh();
                    if (runThreadBeforeLogin == false & ranThread == false)
                    {
                        ranThread = true;
                        StartThread();
                        
                    }
                   
                }
            }
            else { UserSecurity.LoginBringToFront(); }
            string report_POBO_Status = "";
            string report_STX_Status = "";
            string preload_Status = "";

            if (statusLbl.Text == "User Signing In")
            {
                return;
            }


            /*try { report_POBO_Status = thread_POBO.ThreadState == System.Threading.ThreadState.Running ? "Running POBO report." : ""; }
            catch { }*/
            report_POBO_Status = Report_PurchaseOrderBackOrder.ReturnReportStatus();
            /*try { report_STX_Status = thread_STX.ThreadState == System.Threading.ThreadState.Running ? "Running STX report." : ""; }
            catch { }*/
            report_STX_Status = Report_SalesTransaction.ReturnReportStatus();
            try { preload_Status = thread.ThreadState == System.Threading.ThreadState.Running ? "Connecting To ERP Systems." : ""; }
            catch { }
              
             
            if (preload_Status == "")
            {
                if (ERPConnected == true)
                {
                    statusLbl.ForeColor = Color.Green;
                    statusLbl.Text = "Connected To ERP System.";
    
                    EnableCreateButtons();
                }
                else
                {
                    statusLbl.ForeColor = Color.Red;
                    statusLbl.Text = "Could not connect to ERP System.";
              
                }
            }


            if (thread.ThreadState == System.Threading.ThreadState.Running)
            {
                if (statusLbl.Text == "Connecting To ERP System...")
                {
                    statusLbl.Text = "Connecting To ERP System.";
                    statusLbl.ForeColor = Color.Red;
                }
                else if (statusLbl.Text == "Connecting To ERP System.")
                {
                    statusLbl.Text = "Connecting To ERP System..";
                    statusLbl.ForeColor = Color.Red;
                }
                else if (statusLbl.Text == "Connecting To ERP System..")
                {
                    statusLbl.Text = "Connecting To ERP System...";
                    statusLbl.ForeColor = Color.Red;
                }
                else
                {
                    statusLbl.Text = "Connecting To ERP System...";
                    statusLbl.ForeColor = Color.Red;
                }
       
            }

            if (report_POBO_Status != "")
            {
                statusLbl.ForeColor = Color.Blue;
                statusLbl.Text = report_POBO_Status;          
            }
            if (report_STX_Status != "")
            {
                statusLbl.ForeColor = Color.Blue;
                statusLbl.Text = report_STX_Status;
            }


            //obviously:
            statusLbl.Refresh();


        }

        /*private void salesTxRadio_CheckedChanged(object sender, EventArgs e)
        {
            ClearLVI();
            if (salesTxRadio.Checked == true)
            {                
                salesTxBox.BringToFront();
            }
        }*/

       /* private void menuRadio_CheckedChanged(object sender, EventArgs e)
        {
            ClearLVI();
            if (menuRadio.Checked == true)
            {
                splashBox.BringToFront();
            }
        }*/

        private void previewBtn_Click(object sender, EventArgs e)
        {
           
            if (previewBtn.Text == "p")
            {
                this.ClientSize = new Size(this.ClientSize.Width, previewBtn.Bottom + 2);                
                previewBtn.Text = "q";
            }
            else
            {
                //this.Height = reportLVI.Bottom + 10;
                //this.ClientSize = new Size(this.ClientSize.Width, reportLVI.Bottom + 10);
                this.ClientSize = new Size(this.ClientSize.Width, reportPnl.Bottom + 10);   
                previewBtn.Text = "p";
            }
            

        }

        private void salesTxBox_Enter(object sender, EventArgs e)
        {

        }

        private void salesTxCreateBtn_Click(object sender, EventArgs e)
        {
            thread_STX = new System.Threading.Thread(new System.Threading.ThreadStart(ThreadReport_STX));
            thread_STX.Start();
            //salesTxCreateBtn.Enabled = false;
            //salesTxExcelBtn.Enabled = true;           
        }


        

        private void ThreadReport_POBO()
        {
            Report_PurchaseOrderBackOrder.PurchaseOrderBackOrderOptions options = PopulateOptions_POBO();
            Report_PurchaseOrderBackOrder.Generate(options);
            CopyDataToLVI(Report_PurchaseOrderBackOrder.ReturnReportInListView(), reportLVI);
            //reportLVI = ListViewTools.LVI_to_LVI(Report_PurchaseOrderBackOrder.ReturnReportInListView());
        }
        private void ThreadReport_STX()
        {
            Report_SalesTransaction.TransactionReportOptions options = PopulateOptions_STX();
            Report_SalesTransaction.GenerateReport(options);
            CopyDataToLVI(Report_SalesTransaction.ReturnReportInListView(), reportLVI);
            
        }
        
        private void CopyDataToLVI(ListView listViewData, ListView destinationView)
        {
            if (destinationView.InvokeRequired==true)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    destinationView.Items.Clear();
                    destinationView.Columns.Clear();
                    destinationView.Clear();

                    destinationView.FullRowSelect = listViewData.FullRowSelect;
                    destinationView.GridLines = listViewData.GridLines;
                    destinationView.View = listViewData.View;

                    destinationView.Refresh();

                    foreach (ColumnHeader col in listViewData.Columns)
                    {
                        destinationView.Columns.Add((ColumnHeader)col.Clone());
                    }
                    foreach (ListViewItem lvi in listViewData.Items)
                    {
                        destinationView.Items.Add((ListViewItem)lvi.Clone());
                    }

                    destinationView.Refresh();

                });
            }
            else{
            destinationView.Items.Clear();
            destinationView.Columns.Clear();
            destinationView.Clear();

            destinationView.FullRowSelect = listViewData.FullRowSelect;
            destinationView.GridLines = listViewData.GridLines;
            destinationView.View = listViewData.View;

            destinationView.Refresh();
            
            foreach (ColumnHeader col in listViewData.Columns)
            {
                destinationView.Columns.Add((ColumnHeader)col.Clone());
            }
            foreach (ListViewItem lvi in listViewData.Items)
            {
                destinationView.Items.Add((ListViewItem)lvi.Clone());
            }            

            destinationView.Refresh();
            }
        }



        private Report_PurchaseOrderBackOrder.PurchaseOrderBackOrderOptions PopulateOptions_POBO()
        {
            Report_PurchaseOrderBackOrder.PurchaseOrderBackOrderOptions options = new Report_PurchaseOrderBackOrder.PurchaseOrderBackOrderOptions();
            options.StartDate = salesTxStartDatePicker.Value;
            options.EndDate = salesTxEndDatePicker.Value;
            if (poboLineCodeStart.InvokeRequired == true)
            { this.Invoke((MethodInvoker)delegate{ options.LineCodeStart = poboLineCodeStart.Items[poboLineCodeStart.SelectedIndex].ToString(); }); }
            else
            {  options.LineCodeStart = poboLineCodeStart.Items[poboLineCodeStart.SelectedIndex].ToString(); }

            if (poboLineCodeStart.InvokeRequired == true)
            { this.Invoke((MethodInvoker)delegate { options.LineCodeEnd = poboLineCodeEnd.Items[poboLineCodeEnd.SelectedIndex].ToString(); }); }
            else
            { options.LineCodeEnd = poboLineCodeEnd.Items[poboLineCodeEnd.SelectedIndex].ToString(); }


            options.ReportTitle = "PO's on Backorder Report";
            options.ShowReportTitle = false;
            options.LineCodeList = new string[1024];
            poboLineCodeStart.Items.CopyTo(options.LineCodeList, 0);
            return options;
        }
        private Report_SalesTransaction.TransactionReportOptions PopulateOptions_STX()
        {
            Report_SalesTransaction.TransactionReportOptions options = new Report_SalesTransaction.TransactionReportOptions();
            options.CheckSOandAR = salesTxCheckSOChk.Checked;
            options.CheckPOS = salesTxCheckPOSChk.Checked;
            options.CheckOddities = salesTxCheckOdditiesChk.Checked;

            options.CouponAmount = decimal.Parse(salesTxCouponAmountTxt.Text);
            options.MinimumSaleItemCount = int.Parse(salesTxMinSaleItemCountTxt.Text);
            options.OrderMinimum = decimal.Parse(salesTxOrderMinTxt.Text);

            options.IncludeUnitCostColumn = salesTxIncludeCostChk.Checked;
            options.IncludeOrdersBelowOrderMinimum = salesTxOrdersBelowMinChk.Checked;
            options.DisplayMiscLines = salesTxShowMiscLinesChk.Checked;
            options.DisplayMiscShipLines = salesTxShowMiscShipLinesChk.Checked;
            options.SkipPartiallyInvoiced = salesTxSkipPartiallyInvoicedChk.Checked;
            options.BlankLineBetweenOrders = salesTxBlankLineBetweenOrdersChk.Checked;

            options.SaleStartDate = salesTxStartDatePicker.Value;
            options.SaleEndDate = salesTxEndDatePicker.Value;

            options.ShowIDColumn = salesTxIdColumnChk.Checked == true;

            options.QualifiedItemsList = salesTxItemsTxt.Text;

            return options;
        }

        private void salesTxExcelBtn_Click(object sender, EventArgs e)
        {

        }

        private void reportTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            LoadReportFromTreeNode(e.Node);
        }

        private void LoadReportFromTreeNode(TreeNode selectedNode)
        {
            if (selectedNode.Nodes.Count <= 0)
            {
                //selected child node, load report box
                if (selectedNode.Text == "POs on Backorder")
                {
                    //Margie's POBO report
                    POBackorderBox.BringToFront();
                    poboLineCodeStart.Items.Clear();
                    poboLineCodeEnd.Items.Clear();
                    poboLineCodeStart.Items.AddRange(Report_PurchaseOrderBackOrder.getLineCodes().ToArray());
                    object[] copyCmb = new object[poboLineCodeStart.Items.Count];
                    poboLineCodeStart.Items.CopyTo(copyCmb,0);
                    poboLineCodeEnd.Items.AddRange(copyCmb);
                    reportPnl.BringToFront();
                }
                if (selectedNode.Text == "Sales Transaction")
                {
                    //Joe's Sales Tx Report
                    salesTxBox.BringToFront();
                    reportPnl.BringToFront();
                }
                if (selectedNode.Text == "Order Discrepancy")
                {
                    //Joe's Order Discrepancy Report
                    orderDiscrepancyBox.BringToFront();
                    reportPnl.BringToFront();
                }
                if (selectedNode.Text == "P2 Linecode Sales Info")
                {
                    p2LinecodeSalesInfoBox.BringToFront();
                    reportPnl.BringToFront();
                }
                if (selectedNode.Text == "Item Sales by Location")
                {
                    mapBrowser.DocumentText = "";
                    itemSalesGeoBox.BringToFront();
                    mapPnl.BringToFront();
                }
                if (selectedNode.Text == "Starship Packages Shipped")
                {
                    reportPnl.BringToFront();
                    packagesShippedBox.BringToFront();
                }
            }
            else
            {
                //selected a parent node, do nothing.
            }
        }

        private void exportTXTBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To TXT...";
            fileSaveDlg.Filter = "Text File (*.txt)|*.txt";
            //fileSaveDlg.ShowDialog();
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_TXT(fileSaveDlg.FileName, reportLVI,'_','|');
            }
        }

        private void exportCSVBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To CSV...";
            fileSaveDlg.Filter = "Comma Seperated Value (*.csv)|*.csv";
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_CSV(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void exportXLSBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To XLS...";
            fileSaveDlg.Filter = "Excel Spreadsheet (*.xls)|*.xls";
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_XLS(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void exportHTMLBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To HTML...";
            fileSaveDlg.Filter = "HTML File (*.html)|*.html|HTM File (*.htm)|*.htm";
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_HTML(fileSaveDlg.FileName, reportLVI);
            }
            
        }

        private void exportXMLBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To XML...";
            fileSaveDlg.Filter = "XML File (*.xml)|*.xml";
            //fileSaveDlg.ShowDialog();
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_XML(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void exportJSONBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To JSON...";
            fileSaveDlg.Filter = "JSON File (*.json)|*.json";
            //fileSaveDlg.ShowDialog();
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_JSON(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void exportEXCELBtn_Click(object sender, EventArgs e)
        {
            if (reportLVI.Items.Count > 0)
            {
                //LVIToExcel(reportLVI, poboHiddenReportTitleChk.Checked);
                ListViewTools.LVI_to_EXCEL(reportLVI);
            }
            else
            {
                MessageBox.Show("There is no data to build a report from!");
            }
        }

        private void exportPDFBtn_Click(object sender, EventArgs e)
        {
            fileSaveDlg.Title = "Save Report To PDF...";
            fileSaveDlg.Filter = "PDF File (*.pdf)|*.pdf";
            //fileSaveDlg.ShowDialog();
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_PDF(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ResizeForm();
        }

        private void exportEMAILBtn_Click(object sender, EventArgs e)
        {
            ListViewTools.LVI_to_OUTLOOK(reportLVI);
        }

        private void exportDOCXBtn_Click(object sender, EventArgs e)
        {

            fileSaveDlg.Title = "Save Report To DOCX...";
            fileSaveDlg.Filter = "Word File (*.docx)|*.docx";
            //fileSaveDlg.ShowDialog();
            DialogResult res = fileSaveDlg.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
            {
                ListViewTools.LVI_to_DOCX_SLOW(fileSaveDlg.FileName, reportLVI);
            }
        }

        private void exportWORDBtn_Click(object sender, EventArgs e)
        {
            ListViewTools.LVI_to_WORD(reportLVI);
        }


        private void fixStupidWebBrowser()
        {
            try
            {
                bool is64 = Environment.Is64BitOperatingSystem;
                string KeyPath = "";
                if (is64 == true)
                {
                    KeyPath = "SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION";
                }
                else
                {
                    KeyPath = "SOFTWARE\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION";
                }
                Microsoft.Win32.RegistryKey regDM = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(KeyPath, false);
                object wtf = new object();
                if (regDM == null)
                {
                    regDM = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(KeyPath);
                }
                if (regDM != null)
                {
                    string location = System.Environment.GetCommandLineArgs()[0];
                    string appName = System.IO.Path.GetFileName(location);
                    wtf = regDM.GetValue(appName);
                    if (wtf == null)
                    {
                        regDM = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(KeyPath, true);
                        wtf = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(KeyPath, Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree);

                        Version OsVersion = System.Environment.OSVersion.Version;
                        if (OsVersion.Major == 6 & OsVersion.Minor == 1)
                        {
                            ((Microsoft.Win32.RegistryKey)wtf).SetValue(appName, 9000, Microsoft.Win32.RegistryValueKind.DWord);
                        }
                        else if (OsVersion.Major == 6 & OsVersion.Minor == 2)
                        {
                            ((Microsoft.Win32.RegistryKey)wtf).SetValue(appName, 10000, Microsoft.Win32.RegistryValueKind.DWord);
                        }
                        else if (OsVersion.Major == 5 & OsVersion.Minor == 1)
                        {
                            ((Microsoft.Win32.RegistryKey)wtf).SetValue(appName, 8000, Microsoft.Win32.RegistryValueKind.DWord);
                        }

                        ((Microsoft.Win32.RegistryKey)wtf).Close();

                    }


                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }


            
            
        }

        private void itemSalesGeoCreateBtn_Click(object sender, EventArgs e)
        {
            Report_ItemSalesGeo.ItemSalesGeoOptions options = new Report_ItemSalesGeo.ItemSalesGeoOptions();
            options.ItemNumber = itemSalesGeoItemNumberTxt.Text;
            options.StartDate = itemSalesGeoStartDatePicker.Value;
            options.EndDate = itemSalesGeoEndDatePicker.Value;
            string html = Report_ItemSalesGeo.getPageSource(options);
            
            //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\htmloutput.html", html);
            //MessageBox.Show("OK check desktop");


            //mapBrowser.Navigate("c:\\users\\tomwestbrook\\desktop\\htmloutput.html");
            mapBrowser.DocumentText = html;
        }

        private void packagesShippedCreateBtn_Click(object sender, EventArgs e)
        {
            reportLVI.Items.Clear();
            reportLVI.Columns.Clear();
            reportLVI.Clear();
            reportLVI.View = View.Details;
            reportLVI.FullRowSelect = true;
            reportLVI.GridLines = true;
            reportLVI.Refresh();
            
            

            bool DateRadio = packagesShippedDateRadio.Checked;
            if (packagesShippedLastDaysRadio.Checked == false & packagesShippedDateRadio.Checked == false)
            {
                MessageBox.Show("Please choose either Date Range or xx Days Radio Buttons.");
                return;
            }
            Report_StarshipShippedPackages.StarshipShippedPackagesReportSettings prs = new Report_StarshipShippedPackages.StarshipShippedPackagesReportSettings();
            prs.UseDateRange = DateRadio;
            prs.StartDate = packagesShippedStartDatePicker.Value;
            prs.EndDate = packagesShippedEndDatePicker.Value;
            try
            {
                prs.Days = int.Parse(packagesShippedXDaysTxt.Text);
            }
            catch 
            {
                if (DateRadio == false)
                {
                    MessageBox.Show("You need to enter a numerical value for Days to go back.");
                }
                else
                {
                    prs.Days = 0;
                }
            }

            prs.showBilledWeight = packagesShippedBilledWeightChk.Checked;
            prs.showDeliveryCharge = packagesShippedDeliveryChargeChk.Checked;
            prs.showDestinationAddress = packagesShippedDestinationAddressChk.Checked;
            prs.showPackageDimensions = packagesShippedPackageDimensionsChk.Checked;
            prs.showPackingUser = packagesShippedPackagingUserChk.Checked;
            prs.showPostalCode = packagesShippedPostalCodeChk.Checked;
            prs.showProcessDateColumn = packagesShippedProcessedDateChk.Checked;
            prs.showServiceName = packagesShippedServiceNameChk.Checked;
            prs.showTrackingNumber = packagingShippedTrackingIDChk.Checked;
            prs.showPackageType = packagesShippedPackageTypeChk.Checked;




            //Report_StarshipShippedPackages.StarShipShippedPackages result = Report_StarshipShippedPackages.CreateReport(prs);
            Report_StarshipShippedPackages.StarShipShippedPackages result = Report_StarshipShippedPackages.CreateReport(prs);
            foreach (string col in result.Columns)
            {
                ColumnHeader tmpCol = new ColumnHeader();
                tmpCol.Text = col;
                reportLVI.Columns.Add(col);
            }
            foreach (Report_StarshipShippedPackages.StarshipShippedPackagesReportLine line in result.StarshipShippedPackagesReport)
            {
                ListViewItem tmpItm = new ListViewItem();
                tmpItm.Text = line.ProcessDateTime.ToString();
                tmpItm.SubItems.Add(line.ServiceName);
                tmpItm.SubItems.Add(line.DeliveryCharge.ToString());
                tmpItm.SubItems.Add(line.BilledWeight.ToString());
                tmpItm.SubItems.Add(line.PostalCode.ToString());
                reportLVI.Items.Add(tmpItm);
            }

        }

        private void packagesShippedDateRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (packagesShippedDateRadio.Checked == true)
            {
                packagesShippedStartDateLbl.Enabled = true;
                packagesShippedEndDateLbl.Enabled = true;
                packagesShippedStartDatePicker.Enabled = true;
                packagesShippedEndDatePicker.Enabled = true;
                packagesShippedXDaysTxt.Enabled = false;
            }

        }

        private void packagesShippedLastDaysRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (packagesShippedLastDaysRadio.Checked == true)
            {
                packagesShippedStartDateLbl.Enabled = false;
                packagesShippedEndDateLbl.Enabled = false;
                packagesShippedStartDatePicker.Enabled = false;
                packagesShippedEndDatePicker.Enabled = false;
                packagesShippedXDaysTxt.Enabled = true;
            }
        }

        private void p2LinecodeSalesInfoCreateBtn_Click(object sender, EventArgs e)
        {
            Report_P2LinecodeSalesInfo.P2LinecodeSalesInfoSettings p2Settings = new Report_P2LinecodeSalesInfo.P2LinecodeSalesInfoSettings();
            p2Settings.Linecodes = new List<string>(p2LinecodesTxt.Lines);
            p2Settings.showCustomerNumber = p2LinecodeSalesInfoCustomerIDChk.Checked;
            p2Settings.showCustomerName = p2LinecodeSalesInfoNameChk.Checked;
            p2Settings.showCustomerAddress = p2LinecodeSalesInfoAddressChk.Checked;
            p2Settings.showCustomerPhone = p2LinecodeSalesInfoPhoneChk.Checked;
            p2Settings.showCustomerEmail = p2LinecodeSalesInfoEmailChk.Checked;
            p2Settings.showCustomerURL = p2LinecodeSalesInfoWebsiteChk.Checked;
            p2Settings.showSalesPersonNo = p2LinecodeSalesInfoSalespersonIDChk.Checked;
            p2Settings.showDateAccountCreated = p2LinecodeSalesInfoAccountCreationChk.Checked;
            p2Settings.showDateLastPurchased = p2LinecodeSalesInfoLastPurchaseChk.Checked;
            //p2Settings.showItemNumber = p2LinecodeSalesInfoItemChk.Checked;

 
            Report_P2LinecodeSalesInfo.P2LinecodeSalesInfoResults results = Report_P2LinecodeSalesInfo.CreateReport(p2Settings);
            reportLVI = ListViewTools.ColumnRowLists_to_LVI(results.Columns, results.Rows, ref reportLVI);
            

        }



        /*private void button1_Click_2(object sender, EventArgs e)
        {
            Connectivity.SQLCalls.sqlQuery("BULK INSERT GeoLocation FROM '\\\\toolsplus06\\c$\\wisertemp\\ziplatlong.csv' WITH (FIELDTERMINATOR = ',', ROWTERMINATOR = '\n')");

        }*/




        /*private void orderDiscrepancyRadio_CheckedChanged(object sender, EventArgs e)
        {
            ClearLVI();
            
            if (orderDiscrepancyRadio.Checked == true)
            {
                orderDiscrepancyBox.BringToFront();
            }
        }
         */

    }
}
