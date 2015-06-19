using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using System.Net.Http.Headers;
//using System.Net.Http;
using System.Web;


namespace JetDotComInterface
{
    public partial class Form1 : Form
    {
        ToolTip frmTips = new ToolTip();
        private const string VERSION = "0.0b";
        private const int apiDetailsHeight = 200;
        private string APIKey = "";
        private string MerchantID = "";
        private string Password = "";
        private string AuthBearer = "";
        public int TimerIntervalsPassed = 0;

        public Form1()
        {
            InitializeComponent();
            processedOrdersLVI.DoubleClick += new EventHandler(processedOrdersLVI_DoubleClick);
            itemsLVI.DoubleClick += new EventHandler(itemsLVI_DoubleClick);
        }

        void itemsLVI_DoubleClick(object sender, EventArgs e)
        {
           
            //System.IO.File.WriteAllText(@"c:\users\tomwestbrook\desktop\JetItemInfo.txt", itemInfo);
            //MessageBox.Show("Check ur desktop");
            ItemInfo currentItem = ProcessItem(itemsLVI.Items[itemsLVI.SelectedItems[0].Index].SubItems[0].Text);
            
            ItemViewer viewItem = new ItemViewer();            
            viewItem.Show();
            viewItem.PopulateInfo(currentItem);
            //throw new NotImplementedException();
        }

        void processedOrdersLVI_DoubleClick(object sender, EventArgs e)
        {
            OrderView viewOrder = new OrderView();
            viewOrder.Show();
            viewOrder.LoadCustomerData((JetSalesOrder)processedOrdersLVI.Items[processedOrdersLVI.SelectedItems[0].Index].Tag);
            viewOrder.CalculateOrderSummary();
        }
        private void SetToolTips()
        {
            frmTips.AutoPopDelay = 500;
            frmTips.InitialDelay = 750;
            frmTips.ReshowDelay = 500;
            frmTips.ShowAlways = true;

            
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            SetToolTips();

            titleLbl.Text = titleLbl.Text + " - Version " + VERSION;
            ClearLVI();                 
            ResetStatusLVI();
            ResetItemLVI();
            ResetSalesOrderLVI();
            ResetPendingOrderLVI();
            ImageListToStatusLVI();
            ResetProcessedOrdersLVI();
            ResetAcknowledgedOrdersToLVI();
            ResetInProgressOrdersToLVI();
            ResetCompletedOrdersToLVI();
            AddStatusEvent(true, "Loading Interface", "Completed.");
            SetCredentials();
            RefreshForm();

            UPS_API.GenerateLoginCredentialsXML();

            //GetAuthBearer();
        }
        private void GetAuthBearer()
        {
            AddStatusEvent(true, "GET Auth Bearer", "Retrieving Auth Bearer from Jet.com");
            string loginStr = "{\"user\":\"" + APIKey + "\",\"pass\":\"" + Password + "\"}";
            string ContentType = "Content-type: application/json; charset=utf-8";
            string ContentLength = "Content-length: " + loginStr.Length.ToString();
            string bearerData = Connectivity.HTTPCalls.HTTP_POST("https://merchant-api.jet.com/api/Token", loginStr);
            string IDToken = bearerData.Split(new string[] { "id_token\": \"" }, StringSplitOptions.None)[1].Split('"')[0];
            AuthBearer = IDToken;
            apiBearerTxt.Text = IDToken;
            
        }
        private void ImageListToStatusLVI()
        {
            ImageList newIL = new ImageList();
            newIL.Images.Add(Properties.Resources.check_3d);
            newIL.Images.Add(Properties.Resources.x_3d);
            statusLVI.SmallImageList = newIL;
            processedOrdersLVI.SmallImageList = newIL;
        }
        private void ResetProcessedOrdersLVI()
        {
            processedOrdersLVI.Items.Clear();
            processedOrdersLVI.Columns.Clear();
            processedOrdersLVI.Clear();
            ColumnHeader chkColumnCol = new ColumnHeader();
            ColumnHeader orderNumberCol = new ColumnHeader();
            ColumnHeader jetOrderNumberCol = new ColumnHeader();
            ColumnHeader orderStatusCol = new ColumnHeader();
            ColumnHeader customerName = new ColumnHeader();
            ColumnHeader orderedDate = new ColumnHeader();

            
            chkColumnCol.Width = 25;
            chkColumnCol.Text = "";
            processedOrdersLVI.Columns.Add(chkColumnCol);
            orderNumberCol.Width = processedOrdersLVI.Width / 7;
            orderNumberCol.Text = "Order Number";
            processedOrdersLVI.Columns.Add(orderNumberCol);
            jetOrderNumberCol.Width = processedOrdersLVI.Width / 5;
            jetOrderNumberCol.Text = "Jet Order Number";
            processedOrdersLVI.Columns.Add(jetOrderNumberCol);
            orderStatusCol.Width = processedOrdersLVI.Width / 7;
            orderStatusCol.Text = "Order Status";
            processedOrdersLVI.Columns.Add(orderStatusCol);
            customerName.Width = processedOrdersLVI.Width / 5;
            customerName.Text = "Customer Name";
            processedOrdersLVI.Columns.Add(customerName);
            orderedDate.Width = processedOrdersLVI.Width / 4;
            orderedDate.Text = "Ordered Date";
            processedOrdersLVI.Columns.Add(orderedDate);


        }
        private void ResetPendingOrderLVI()
        {
            pendingOrderLbl.Text = "Pending Orders";
            pendingOrderLVI.Items.Clear();
            pendingOrderLVI.Columns.Clear();
            pendingOrderLVI.Clear();
            ColumnHeader pendingOrdersCol = new ColumnHeader();
            pendingOrdersCol.Text = "Pending Order Number";
            pendingOrdersCol.TextAlign = HorizontalAlignment.Center;
            pendingOrdersCol.Width = pendingOrderLVI.Width - 25;
            pendingOrderLVI.Columns.Add(pendingOrdersCol);
        }
        private void ResetSalesOrderLVI()
        {
            readyOrdersLbl.Text = "Ready To Import Orders";
            readyOrderLVI.Items.Clear();
            readyOrderLVI.Columns.Clear();
            readyOrderLVI.Clear();
            ColumnHeader SONumber = new ColumnHeader();
            SONumber.Width = readyOrderLVI.Width - 25;
            SONumber.Text = "Sales Order Number";
            SONumber.TextAlign = HorizontalAlignment.Center;
            readyOrderLVI.Columns.Add(SONumber);

        }
        private void ResetItemLVI()
        {
            itemsOnJetLbl.Text = "Items On Jet.com";
            itemsLVI.Items.Clear();
            itemsLVI.Columns.Clear();
            itemsLVI.Clear();
            ColumnHeader itemCol = new ColumnHeader();
            itemCol.Text = "Item Number";
            itemCol.Width = itemsLVI.Width -25;
            itemCol.TextAlign = HorizontalAlignment.Center;
            itemsLVI.Columns.Add(itemCol);
        }
        private void ResetStatusLVI()
        {
            statusLVI.Items.Clear();
            statusLVI.Columns.Clear();
            statusLVI.Clear();
            ColumnHeader successCol = new ColumnHeader();
            successCol.Text = "";
            successCol.Width = 25;
            ColumnHeader statusText = new ColumnHeader();
            statusText.Text = "Status";
            statusText.TextAlign = HorizontalAlignment.Center;
            statusText.Width = statusLVI.Width / 4;
            ColumnHeader statusDescription = new ColumnHeader();
            statusDescription.Text = "Description";
            statusDescription.TextAlign = HorizontalAlignment.Center;
            statusDescription.Width = (int)(statusLVI.Width / 2.35);
            ColumnHeader timeStamp = new ColumnHeader();
            timeStamp.Text = "Event Time";
            timeStamp.TextAlign = HorizontalAlignment.Center;
            timeStamp.Width = (int)(statusLVI.Width / 3.95);
            statusLVI.Columns.Add(successCol);
            statusLVI.Columns.Add(statusText);
            statusLVI.Columns.Add(statusDescription);
            statusLVI.Columns.Add(timeStamp);



        }
        private void AddStatusEvent(bool success, string statusTitle, string statusDescription)
        {
            ListViewItem newStatusEvent = new ListViewItem();
            ListViewItem.ListViewSubItem StatusTitle = new ListViewItem.ListViewSubItem();
            ListViewItem.ListViewSubItem StatusDescription = new ListViewItem.ListViewSubItem();
            ListViewItem.ListViewSubItem DateTimeStamp = new ListViewItem.ListViewSubItem();
            StatusTitle.Text = statusTitle;            
            StatusDescription.Text = statusDescription;
            DateTimeStamp.Text = DateTime.Now.ToString();
            if (success == true)
            {                
                //newStatusEvent.ImageList.Images.Add(Properties.Resources.check_3d);
                newStatusEvent.ImageIndex = 0;
            }
            else
            {
                //newStatusEvent.ImageList.Images.Add(Properties.Resources.x_3d);
                newStatusEvent.ImageIndex = 1;
            }
            newStatusEvent.SubItems.Add(StatusTitle);
            newStatusEvent.SubItems.Add(StatusDescription);
            newStatusEvent.SubItems.Add(DateTimeStamp);
            statusLVI.Items.Add(newStatusEvent);
        }
        public void RefreshForm()
        {
            titleLbl.Top = 0;
            titleLbl.Left = 25;
            titleLbl.Width = this.Width;
            titleLbl.Height = 25;
            apiInfoBox.Top = titleLbl.Bottom + 10;
            //apiInfoBox.Left = 20;
            if (plusMinus1Lbl.Text == "+")
            {
                apiInfoBox.Height = 20;
            }
            else
            {
                apiInfoBox.Height = apiDetailsHeight;
            }
            apiInfoBox.Width = 486;
            currentTimeLbl.Left = 0;
            currentTimeLbl.Top = this.Height - currentTimeLbl.Height-28;
            currentTimeLbl.BringToFront();

            statusBox.Top = apiInfoBox.Bottom + 7;
            itemInfoBox.Top = statusBox.Bottom + 7;
            orderInfoBox.Top = itemInfoBox.Bottom + 7;
            controlsBtn.Top = apiInfoBox.Top;
            processedOrdersBox.Top = orderInfoBox.Bottom + 7;
            returnOrdersBox.Top = processedOrdersBox.Bottom + 7;
            updateIntervalLbl.Text = "Update every " + updateIntervalTrk.Value.ToString() + " mins.";
        }
        private void SetCredentials()
        {
           
            List<string> creds = Connectivity.SQLCalls.sqlQuery("SELECT APIKey,MerchantID,SecretCode FROM APIKeys WHERE APIName='Jet.com'");
            if (creds.Count > 0)
            {
                APIKey = creds[0].Split('|')[0];
                MerchantID = creds[0].Split('|')[1];
                Password = creds[0].Split('|')[2];
                apiUsernameTxt.Text = APIKey;
                apiPasswordTxt.Text = Password;
                apiMerchantTxt.Text = MerchantID;
                AddStatusEvent(true, "SQL GET Credentials", "Gets API codes to talk with Jet.com");
                
            }
            else
            {
                //MessageBox.Show("Could not pull credentials from SQL Database. Please check connectivity first before trying again.");
                AddStatusEvent(false, "SQL GET Credentials", "Could not pull credentials from SQL Database. Please check connectivity first before trying again.");
            }

        }


        private void ProcessSalesOrdersToLVI(string results)
        {
            ResetSalesOrderLVI();
            string actualResults = results.Split('[')[1].Split(']')[0];
            if (actualResults.Length > 0)
            {
                try
                {
                    foreach (string order in actualResults.Split(new string[] { "\"/" }, StringSplitOptions.None))
                    {
                        if (order.ToUpper().Contains("ORDERS") == true)
                        {
                            string realOrder = order.Split('/')[2].Split('"')[0];
                            ListViewItem orderItem = new ListViewItem();
                            orderItem.Text = realOrder;
                            readyOrderLVI.Items.Add(orderItem);
                            readyOrderLVI.Refresh();
                        }
                    }
                    readyOrdersLbl.Text = "Accepted Orders (" + readyOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    AddStatusEvent(true, "Processing Accepted Orders", "Preparing accepted orders from Jet.com");
                }
                catch (Exception wtf)
                {
                    AddStatusEvent(false, "Processing Accepted Orders", wtf.Message);
                }
            }
            else
            {
                readyOrdersLbl.Text = "Accepted Orders (0)\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Accepted Orders", "Completed OK but no orders returned.");
            }
        }

       /* public string HTTP_GET(string Url, string Data)
        {
            string Out = String.Empty;
            System.Net.WebRequest req = System.Net.WebRequest.Create(Url + "subscription-key);

            try
            {
                req.Method = "GET";
                req.Timeout = 100000;
                req.ContentType = "application/x-www-form-urlencoded";
                
                using (System.IO.Stream sendStream = req.GetRequestStream())
                {
                    sendStream.Write(sentData, 0, sentData.Length);
                    sendStream.Close();
                }
                System.Net.WebResponse res = req.GetResponse();
                System.IO.Stream ReceiveStream = res.GetResponseStream();
                string str = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8))
                {
                    Char[] read = new Char[256];
                    int count = sr.Read(read, 0, 256);

                    while (count > 0)
                    {
                        str = new String(read, 0, count);
                        Out += str;
                        count = sr.Read(read, 0, 256);
                    }
                }
                //MessageBox.Show("The Server Responded With: " + Out);

                webBrowser1.DocumentText = Out;

            }
            catch (ArgumentException ex)
            {
                Out = string.Format("HTTP_ERROR :: The second HttpWebRequest object has raised an Argument Exception as 'Connection' Property is set to 'Close' :: {0}", ex.Message);
            }
            catch (Exception ex)
            {
                Out = string.Format("HTTP_ERROR :: WebException raised! :: {0}", ex.Message);
            }
            catch
            {
                Out = string.Format("DAMNIT ERROR");
            }

            return Out;
        }*/

        public string HTTP_POST(string Url, string Data)
        {
            string Out = String.Empty;
            System.Net.WebRequest req = System.Net.WebRequest.Create(Url);

            try
            {
                req.Method = "POST";
                req.Timeout = 100000;
                req.ContentType = "application/x-www-form-urlencoded";
                byte[] sentData = Encoding.UTF8.GetBytes(Data);
                req.ContentLength = sentData.Length;
                using (System.IO.Stream sendStream = req.GetRequestStream())
                {
                    sendStream.Write(sentData, 0, sentData.Length);
                    sendStream.Close();
                }
                System.Net.WebResponse res = req.GetResponse();
                System.IO.Stream ReceiveStream = res.GetResponseStream();
                string str = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8))
                {
                    Char[] read = new Char[256];
                    int count = sr.Read(read, 0, 256);

                    while (count > 0)
                    {
                        str = new String(read, 0, count);
                        Out += str;
                        count = sr.Read(read, 0, 256);
                    }
                }
                //MessageBox.Show("The Server Responded With: " + Out);

               // webBrowser1.DocumentText = Out;

            }
            catch (ArgumentException ex)
            {
                Out = string.Format("HTTP_ERROR :: The second HttpWebRequest object has raised an Argument Exception as 'Connection' Property is set to 'Close' :: {0}", ex.Message);
            }
            catch (Exception ex)
            {
                Out = string.Format("HTTP_ERROR :: WebException raised! :: {0}", ex.Message);
            }
            catch
            {
                Out = string.Format("DAMNIT ERROR");
            }

            return Out;
        }


        private void ClearLVI()
        {
            itemsLVI.Items.Clear();
            itemsLVI.Columns.Clear();
            itemsLVI.Clear();

            ColumnHeader itemNumber = new ColumnHeader();
            itemNumber.Text="Item Number";
            itemNumber.Width = itemsLVI.Width/4;

            itemsLVI.Columns.Add(itemNumber);


        }

        private void plusMinus1Lbl_Click(object sender, EventArgs e)
        {
            if (plusMinus1Lbl.Text == "+")
            {
                plusMinus1Lbl.Text = "-";
                apiInfoBox.Height = apiDetailsHeight;
                apiInfoBox.BringToFront();
                plusMinus1Lbl.BringToFront();
            }
            else
            {
                plusMinus1Lbl.Text = "+";
                apiInfoBox.Height = 20;
                apiInfoBox.BringToFront();
                plusMinus1Lbl.BringToFront();
            }
            RefreshForm();
        }

        private void secondTmr_Tick(object sender, EventArgs e)
        {
            
            
            currentTimeLbl.Text = "Current Time: " + DateTime.Now.ToString();
            TimerIntervalsPassed++;
            if (TimerIntervalsPassed >= updateIntervalTrk.Value * 60)
            {
                //time to update!
               // getAllJetAcceptedOrdersBtn_Click(sender, e);
               // processNextOrderBtn_Click(sender, e);
                MessageBox.Show("Would have updated, but we disabled it for programming!");
                TimerIntervalsPassed = 0;
            }

            nextUpdateLbl.Text = "Next update in " + ((updateIntervalTrk.Value * 60) - TimerIntervalsPassed).ToString() + " seconds.";
        }

        private void getAllJetItemsBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/merchant-skus", Headers);
            System.IO.File.WriteAllText(@"c:\users\tomwestbrook\desktop\itemlist.txt", results);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET Accepted Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET Accepted Orders", "User attempt to get accepted orders from Jet.com.");
            }
            //MessageBox.Show(results);
            ProcessItemsToLVI(results);

        }
        private void ProcessItemsToLVI(string results)
        {
            ResetItemLVI();
            string innerResults = results.Split('[')[1].Split(']')[0];
            try
            {
                foreach (string item in innerResults.Split('/'))
                {
                    if (item.Split(',')[0].ToUpper().Contains("MERCHANT-SKUS") == false)
                    {
                        string actualItem = item.Split('"')[0];
                        ListViewItem newLVI = new ListViewItem();
                        newLVI.Text = actualItem;
                        
                        itemsLVI.Items.Add(newLVI);
                    }
                }
                itemsOnJetLbl.Text = "Items On Jet.com (" + itemsLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Items List", "Preparing the listing all Jet.com items");
            }
            catch (Exception wtf)
            {
                AddStatusEvent(false, "Processing Items List", wtf.Message);
            }
        }

        private void getAllJetPendingOrdersBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/created", Headers);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET Pending Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET Pending Orders", "User attempt to get pending orders from Jet.com.");
                ProcessPendingOrdersToLVI(results);
            }
            //MessageBox.Show(results);
           

        }
        private void ProcessPendingOrdersToLVI(string results)
        {
            ResetPendingOrderLVI();
            string actualResults = results.Split('[')[1].Split(']')[0];
            if (actualResults.Length > 0)
            {
                try
                {
                    foreach (string order in actualResults.Split(new string[] { "\"/" }, StringSplitOptions.None))
                    {
                        if (order.ToUpper().Contains("ORDERS") == true)
                        {
                            string realOrder = order.Split('/')[2].Split('"')[0];
                            ListViewItem orderItem = new ListViewItem();
                            orderItem.Text = realOrder;
                            pendingOrderLVI.Items.Add(orderItem);
                        }
                    }
                    pendingOrderLbl.Text = "Pending Orders (" + pendingOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    AddStatusEvent(true, "Processing Pending Orders", "Processes pulled pending orders from Jet.com");
                }
                catch (Exception wtf)
                {
                    AddStatusEvent(false, "Processing Pending Orders", wtf.Message);
                }
            }
            else
            {
                pendingOrderLbl.Text = "Pending Orders (0)\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Pending Orders", "Completed OK but no orders returned.");
            }

        }

        private void plusMinus2Lbl_Click(object sender, EventArgs e)
        {
            if (plusMinus2Lbl.Text == "+")
            {
                plusMinus2Lbl.Text = "-";
                statusBox.Height = statusLVI.Height + 45;
                statusBox.BringToFront();
                plusMinus2Lbl.BringToFront();
            }
            else
            {
                plusMinus2Lbl.Text = "+";
                statusBox.Height = 20;
            }
            RefreshForm();
        }

        private void plusMinus3Lbl_Click(object sender, EventArgs e)
        {
            if (plusMinus3Lbl.Text == "+")
            {
                plusMinus3Lbl.Text = "-";
                itemInfoBox.Height = itemsLVI.Height + this.itemsOnJetLbl.Height + 65;
                itemInfoBox.BringToFront();
                plusMinus3Lbl.BringToFront();
            }
            else
            {
                plusMinus3Lbl.Text = "+";
                itemInfoBox.Height = 20;
            }
            RefreshForm();
        }

        private void label5_Click(object sender, EventArgs e)
        {
            if (plusMinus4Lbl.Text == "+")
            {
                plusMinus4Lbl.Text = "-";
                orderInfoBox.Height = pendingOrderLVI.Height + pendingOrderLbl.Height + 75;
                orderInfoBox.BringToFront();
                plusMinus4Lbl.BringToFront();
            }
            else
            {
                plusMinus4Lbl.Text = "+";
                orderInfoBox.Height = 20;
            }
            RefreshForm();
        }

        private void gotoJetDashboardBtn_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://partner.jet.com/dashboard");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://developer.jet.com");
        }

        private void controlsBtn_Click(object sender, EventArgs e)
        {
            if (controlsBtn.Text == "Controls >" == true)
            {
                controlsBtn.Text = "Controls <";
                this.Width += 385;
            }
            else
            {
                controlsBtn.Text = "Controls >";
                this.Width -= 385;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:developer@jet.com");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://developer.jet.com/tos");
        }

        private void statusBox_Enter(object sender, EventArgs e)
        {

        }

        private void plusMinus5Lbl_Click(object sender, EventArgs e)
        {
            if (processedOrdersBox.Height == 20)
            {
                processedOrdersBox.Height = 200;
            }
            else
            {
                processedOrdersBox.Height = 20;
            }
            RefreshForm();
        }

        private void processNextOrderBtn_Click(object sender, EventArgs e)
        {
            ResetProcessedOrdersLVI();
            foreach (ListViewItem lvi in readyOrderLVI.Items)
            {
                JetSalesOrder orderInfo = ProcessSalesOrder(lvi.Text);
                ListViewItem newProcessedOrder = new ListViewItem();
                ListViewItem.ListViewSubItem JetOrderNo = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem TPOrderNo = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem OrderStatus = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem CustomerName = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem OrderDate = new ListViewItem.ListViewSubItem();
                
                
                JetOrderNo.Text = orderInfo.JetSalesOrderNumber;
                TPOrderNo.Text = orderInfo.TPSalesOrderNumber;
                OrderStatus.Text = orderInfo.SalesOrderStatus;
                CustomerName.Text = orderInfo.BuyerName;
                OrderDate.Text = orderInfo.OrderPlacedDate.ToString();

                newProcessedOrder.SubItems.Add(TPOrderNo);
                newProcessedOrder.SubItems.Add(JetOrderNo);
                newProcessedOrder.SubItems.Add(OrderStatus);
                newProcessedOrder.SubItems.Add(CustomerName);
                newProcessedOrder.SubItems.Add(OrderDate);

                if (orderInfo.SalesOrderStatus.ToUpper() == "READY")
                {
                    newProcessedOrder.ImageIndex = 0;
                }
                else
                {
                    newProcessedOrder.ImageIndex = 1;
                }
                newProcessedOrder.Tag = orderInfo; //Now all order info is stashed in the LVI. Gangsta!
                processedOrdersLVI.Items.Add(newProcessedOrder);

            }

            AcknowledgeOrders();
            
        }
        public bool AcknowledgeOrders()
        {
            AddStatusEvent(true, "Acknowledging Orders", "Updating Jet.com whether order is orderable.");
            foreach (ListViewItem order in processedOrdersLVI.Items)
            {                
                JetSalesOrder jetOrder = (JetSalesOrder)order.Tag;
            //JetSalesOrder jetOrder = (JetSalesOrder)processedOrdersLVI.Items[0].Tag;

                string JSON_OUT = "{ \"acknowledgement_status\": \"";
                //string RefOrderID = jetOrder.
                string OrderStatus = "";
                string itemJSON = "";
                //CHECK THESE FIRST
                //rejected- ship from location not available: The ship to locaiton is invalid
                //rejected- shipping method not supported: The address requested cannot be shipped to
                //rejected- unfulfillable address: The address requested cannot be shipped to

                //now to check ship to, we need to know countrycode


                //Returned Number:
                //-1: WarehouseID doesn't exist for this fulfillment center
                //-2: FulfillmentNodeID invalid
                //-3: Should never ever ever be hit.
                //+x: The warehouse id of given fulfillmentnode.
                int shipFromCheck = getWarehouseIDFromNodeID(jetOrder.FulfillmentNode);
                if (shipFromCheck < 0)
                {
                    OrderStatus = "rejected - ship from location not available";
                }
                else
                {
                    bool shipToCheck = checkShipToAddress(jetOrder);
                    if (shipToCheck == false)
                    {
                        OrderStatus = "rejected - unfulfillable address";
                    }
                    else
                    {
                        bool shipMethodCheck = checkShipMethod(jetOrder.OrderDetails);
                        if (shipMethodCheck == false)
                        {
                            OrderStatus = "rejected - shipping method not supported";
                        }
                    }
                }



                //rejected- item level error: The error occurred at the item level
                //accepted- all items in the order will be shipped
                bool fulfilled = true;
                foreach(OrderItem oi in jetOrder.OrderedItems)
                {
                    //int 0: nonfulfillable- invalid merchant SKU: This is not recognized as a valid value
                    //int 1: nonfulfillable- no inventory - No inventory for this product is available
                    //int 2: fulfillable- the item in the order can be shipped
                    int itemAccepted = checkItemAvailability(oi.TPItemNumber, oi.RequestedQty, shipFromCheck);
                    if (itemAccepted == 0)
                    {
                        itemJSON += "{ \"order_item_acknowledgement_status\": \"nonfulfillable - invalid merchant SKU\",\"order_item_id\": \"" + oi.JetItemID + "\", \"alt_order_item_id\": \"" + oi.TPItemNumber + "\"},";
                        fulfilled = false;
                    }
                    else if (itemAccepted == 1)
                    {
                        itemJSON += "{ \"order_item_acknowledgement_status\": \"nonfulfillable - no inventory\",\"order_item_id\": \"" + oi.JetItemID + "\", \"alt_order_item_id\": \"" + oi.TPItemNumber + "\"},";
                        fulfilled = false;
                    }
                    else if (itemAccepted == 2)
                    {
                        itemJSON += "{ \"order_item_acknowledgement_status\": \"fulfillable\",\"order_item_id\": \"" + oi.JetItemID + "\", \"alt_order_item_id\": \"" + oi.TPItemNumber + "\"},";  
                    }
                    else
                    {
                        throw new Exception("Item Accepted return value not valid. This can only happen if one updates the programming. Should never see this!");
                        AddStatusEvent(false, "Invalid Programmed Value", "A programming error has happened.");
                        return false;
                    }
                                          
                }


                if (fulfilled == true & OrderStatus == "")
                {
                    OrderStatus = "accepted";
                    SynchronizeToMAS(jetOrder);
                }
                if (fulfilled == false & OrderStatus == "")
                {
                    OrderStatus = "rejected - item level error";
                }
                if (OrderStatus == "")
                {
                    OrderStatus = "rejected - item level error";
                }


                JSON_OUT = JSON_OUT + OrderStatus + "\", \"reference_order_id\": \"" + jetOrder.JetSalesOrderNumber + "\", \"alt_order_id\": \"" + jetOrder.MerchantOrderID + "\", \"order_items\": [ ";

                itemJSON = itemJSON.Substring(0,itemJSON.Length - 1);
                JSON_OUT += itemJSON + "] }";
                System.IO.File.WriteAllText(@"c:\users\tomwestbrook\desktop\json_acknowledge.txt", JSON_OUT);
                //JSON_OUT += itemJSON + " ], \"shipments\": [ {";
                //JSON_OUT += " \"shipment_id\": \"" + "
                
                
                string PutContentType = "Content-type: application/json";
                string Authorization = "Authorization: Bearer " + AuthBearer;
                List<string> Headers = new List<string>();
                Headers.Add(PutContentType);
                Headers.Add(Authorization);
                               
                
                string results = Connectivity.HTTPCalls.HTTP_PUT("https://merchant-api.jet.com/api/orders/" + jetOrder.JetSalesOrderNumber + "/acknowledge",JSON_OUT, Headers);
                if (results == "")
                {
                    order.SubItems[3].Text = "OK";
                }
                else
                {
                    order.SubItems[3].Text = results;
                    //issue?
                }

               // MessageBox.Show(results);

            }
            return true;
        }

        public int getWarehouseIDFromNodeID(string FulfillmentNodeID)
        {
            //Returned Number:
            //-1: WarehouseID doesn't exist for this fulfillment center
            //-2: FulfillmentNodeID invalid
            //-3: Should never ever ever be hit.
            //+x: The warehouse id of given fulfillmentnode.

            List<string> nodes = Connectivity.SQLCalls.sqlQuery("SELECT WarehouseID FROM FulfillmentCenters WHERE JetFulfillmentNodeID='" + FulfillmentNodeID + "'");
            if (nodes.Count > 0)
            {
                string WarehouseID = nodes[0].Split('|')[0];
                WarehouseID = WarehouseID.Trim();
                if (WarehouseID.Length > 0)
                {
                    return int.Parse(nodes[0].Split('|')[0].Trim());
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                return -2;
            }
            return -3;
        }

/*        public bool checkShipFromAddress(string FulfillmentNodeID) //insert address struct
        {
            List<string> nodes = Connectivity.SQLCalls.sqlQuery("SELECT WarehouseID FROM FulfillmentCenters WHERE JetFulfillmentNodeID='" + FulfillmentNodeID + "'");
            if (nodes.Count > 0)
            {
                string WarehouseID = nodes[0].Split('|')[0];
                WarehouseID = WarehouseID.Trim();
                if (WarehouseID.Length > 0)
                {
                    //ok, so if a warehouseID is actually here, we know the fulfillment center/warehouse selected is available.

                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
            return false;
        }*/
        public bool checkShipToAddress(JetSalesOrder jetOrder) //insert address struct
        {
            List<string> AddressLines = new List<string>();
            if (jetOrder.ShippingToLocation.ShipToAddress.Address1 != "")
            {
                AddressLines.Add(jetOrder.ShippingToLocation.ShipToAddress.Address1);
            }
            if (jetOrder.ShippingToLocation.ShipToAddress.Address2 != "")
            {
                AddressLines.Add(jetOrder.ShippingToLocation.ShipToAddress.Address2);
            }
            string City = jetOrder.ShippingToLocation.ShipToAddress.City;
            string State = jetOrder.ShippingToLocation.ShipToAddress.State;
            string RealZipcode = jetOrder.ShippingToLocation.ShipToAddress.Zipcode;
            string tmpZipcode = "";
            string CountryCode = "";
            if (RealZipcode.Contains("-") == true)
            {
                tmpZipcode = RealZipcode.Split('-')[0];
            }
            else
            {
                tmpZipcode = RealZipcode;
            }

            if (int.Parse(tmpZipcode) < 1001)
            {
                CountryCode = "PR";
            }
            else
            {
                CountryCode = "US";
            }


            List<string> HeaderCollection = new List<string>();
            HeaderCollection.Add("Content-type: application/x-www-form-urlencoded");
            string XMLCredentialsHeader = UPS_API.API_Credentials.APIAccessRequestXML;
            string XMLAddressValidationRequest = UPS_API.GenerateAddressValidationXML(AddressLines, RealZipcode, CountryCode,"","",City,State);

            System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\upsAPIreq.txt", XMLCredentialsHeader + XMLAddressValidationRequest);
            
            string results = Connectivity.HTTPCalls.HTTP_POST(UPS_API.API_Credentials.APIProductionBaseURL + UPS_API.API_Credentials.APIStreetAddressValidationSuffix,XMLCredentialsHeader + XMLAddressValidationRequest,HeaderCollection);
            System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\upsAPIresp.txt", results);
            if (results != "")
            {
                if (results.Contains("<ResponseStatusDescription>") == true)
                {
                    if (results.Split(new string[] { "<ResponseStatusDescription>" }, StringSplitOptions.None)[1].Split('<')[0] == "Success")
                    {
                        //either address is valid or invalid...
                        if (results.Contains("<NoCandidatesIndicator/>") == true)
                        {
                            //invalid
                            MessageBox.Show("Address Doesn't Exist");
                            return false;
                        }
                        else
                        {
                            //valid
                            MessageBox.Show("Address Exists");
                            return true;
                        }
                    }
                    else
                    {
                        //error with xml data itself or server error...
                        MessageBox.Show("Data Level Error");
                        return false;
                    }
                }
            }
            return false;
        }
        public bool checkShipMethod(OrderDetail shippingDetails) //insert ship method struct
        {
            if (shippingDetails.RequestShippingCarrier.ToUpper() == "UPS")
            {
                if (shippingDetails.RequestServiceLevel.ToUpper() == "SCHEDULED")
                {
                    TimeSpan totShipTime = (shippingDetails.RequestDeliveryBy - shippingDetails.RequestShipBy);
                    int totShipDays = (int)totShipTime.TotalDays;
                    if (totShipDays > 4)
                    {
                        //ground shipping
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (shippingDetails.RequestServiceLevel.ToUpper() == "STANDARD")
                {
                    return true;
                }
                if (shippingDetails.RequestServiceLevel.ToUpper() == "GROUND")
                {
                    return true;
                }
            }
            if (shippingDetails.RequestShippingCarrier.ToUpper() == "USPS")
            {
                if (shippingDetails.RequestServiceLevel.ToUpper() == "SCHEDULED")
                {
                    TimeSpan totShipTime = (shippingDetails.RequestDeliveryBy - shippingDetails.RequestShipBy);
                    int totShipDays = (int)totShipTime.TotalDays;
                    if (totShipDays > 4)
                    {
                        //ground shipping
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (shippingDetails.RequestServiceLevel.ToUpper() == "STANDARD")
                {
                    return true;
                }
                if (shippingDetails.RequestServiceLevel.ToUpper() == "GROUND")
                {
                    return true;
                }
            }

            return false;
        }
        public int checkItemAvailability(string ItemNumber, int Qty, int WarehouseID)
        {
            //int 0: nonfulfillable- invalid merchant SKU: This is not recognized as a valid value
            //int 1: nonfulfillable- no inventory - No inventory for this product is available
            //int 2: fulfillable- the item in the order can be shipped

            List<string> itemCheck = Connectivity.SQLCalls.sqlQuery("SELECT QuantityOnHand FROM InventoryQuantities WHERE ItemNumber='" + ItemNumber + "' AND WhseCode='000'");
            if (itemCheck.Count > 0)
            {
                if (Qty <= int.Parse(itemCheck[0].Split('|')[0]))
                {
                    return 2;
                }
                else
                {
                    return 1;
                }
            }
            else
            {
                return 0;
            }
        }


        public FulfillmentInfo GetFulfillmentCenterInfo(string NodeID)
        {
            FulfillmentInfo tmpInfo = new FulfillmentInfo();
            tmpInfo.FulfillmentNodeID = NodeID;
            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/setup/fulfillmentNodes/", Headers);

            results = results.Split('[')[1].Split(']')[0];

            foreach (string resLine in results.Split('{'))
            {
                if (resLine.Replace("\r\n", "").Trim() != "")
                {
                    string res = resLine.Split('}')[0].Trim();
                    string nodeID = res.Split(new string[] { "\"jet_fulfillment_node_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                    if (nodeID == NodeID)
                    {
                        tmpInfo.FulfillmentNodeName = res.Split(new string[] { "\"location_name\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tmpInfo.FulfillmentNodeTimeZone = res.Split(new string[] { "\"time_zone\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tmpInfo.FulfillmentNodeIsDaylightSavingsEnabled = bool.Parse(res.Split(new string[] { "\"is_daylight_savings_enabled\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tmpInfo.FulfillmentNodeAddress1 = res.Split(new string[] { "\"address_1\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tmpInfo.FulfillmentNodeZipcode = res.Split(new string[] { "\"zip_code\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tmpInfo.FulfillmentNodeCity = res.Split(new string[] { "\"city\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tmpInfo.FulfillmentNodeState = res.Split(new string[] { "\"state\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        break;
                    }
                }

            }

            return tmpInfo;
        }
        public struct FulfillmentInfo
        {
            public string FulfillmentNodeID;
            public string FulfillmentNodeName;
            public string FulfillmentNodeTimeZone;
            public bool FulfillmentNodeIsDaylightSavingsEnabled;
            public string FulfillmentNodeAddress1;
            public string FulfillmentNodeZipcode;
            public string FulfillmentNodeCity;
            public string FulfillmentNodeState;
        }
        public static string GetMAPTypeDetails(string MAPType)
        {
            try
            {
                if (MAPType.ToLower() == "type0")
                {
                    return "A user is able to see sub MAP pricing for this item even if they are not logged in.";
                }
                if (MAPType.ToLower() == "type1")
                {
                    return "A user must be logged in to see the savings on this item and that saving may be represented as a lower than MAP price on this item.";
                }
                if (MAPType.ToLower() == "type2")
                {
                    return "A user must be logged in to see the savings and they need to be shown as Jet savings with the price being displayed as the MAP price.";
                }
                if (MAPType.ToLower() == "type3")
                {
                    return "A user will never see Jet savings associated with this item but if they are logged in and purchase other items savings may be allocated to those other items.";
                }
                if (MAPType.ToLower() == "type4")
                {
                    return "Never apply any savings associated with this item anywhere.";
                }
            }
            catch { }
            return "No MAP Type.";
            
        }
        private ItemInfo ProcessItem(string ItemNumber)
        {
            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string itemInfo = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/merchant-skus/" + ItemNumber , Headers);
            System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\itemtest.txt", itemInfo);
            ItemInfo tmpItem = new ItemInfo();
       

                try
                {
                    tmpItem.Brand = itemInfo.Split(new string[] { "\"brand\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { tmpItem.Brand = "N/A"; }
                try
                {
                    tmpItem.DisplayHeightInches = decimal.Parse(itemInfo.Split(new string[] { "\"display_height_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }catch{tmpItem.DisplayHeightInches = 0;}
                try{
                    tmpItem.DisplayLengthInches = decimal.Parse(itemInfo.Split(new string[] { "\"display_length_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }catch{tmpItem.DisplayLengthInches = 0;}
                try{
                    tmpItem.DisplayWidthInches = decimal.Parse(itemInfo.Split(new string[] { "\"display_width_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }catch{tmpItem.DisplayWidthInches = 0;}
                
                try
                {
                    tmpItem.ExcludeFromFeeAdjustments = bool.Parse(itemInfo.Split(new string[] { "\"exclude_from_fee_adjustments\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.FulfillmentTime = int.Parse(itemInfo.Split(new string[] { "\"fulfillment_time\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch {tmpItem.FulfillmentTime = 0; }
                try
                {
                    tmpItem.JetBrowseNodeID = itemInfo.Split(new string[] { "\"jet_browse_node_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch {tmpItem.JetBrowseNodeID="N/A"; }
                try
                {
                    tmpItem.JetRetailSKU = itemInfo.Split(new string[] { "\"jet_retail_sku\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch {  tmpItem.JetRetailSKU="N/A"; }
                try
                {
                    tmpItem.Manufacturer = itemInfo.Split(new string[] { "\"manufacturer\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { }
                try
                {
                    tmpItem.JetPrice = decimal.Parse(itemInfo.Split(new string[] { "\"price\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                }
                catch { }
                try
                {
                    tmpItem.MAPPrice = decimal.Parse(itemInfo.Split(new string[] { "\"map_price\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.MAPType = itemInfo.Split(new string[] { "\"map_implementation\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                }
                catch { tmpItem.MAPType = "N/A"; }
                try
                {
                    tmpItem.MerchantID = itemInfo.Split(new string[] { "\"merchant_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { tmpItem.MerchantID = "N/A"; }
                try
                {
                    tmpItem.MerchantSKU = itemInfo.Split(new string[] { "\"merchant_sku\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { tmpItem.MerchantSKU = "N/A"; }
                try
                {
                    tmpItem.MerchantSKUID = itemInfo.Split(new string[] { "\"merchant_sku_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { tmpItem.MerchantSKUID = "N/A"; }
                try
                {
                    tmpItem.MSRP = decimal.Parse(itemInfo.Split(new string[] { "\"msrp\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.MultipackQuantity = int.Parse(itemInfo.Split(new string[] { "\"multipack_quantity\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.NoReturnFeeAdjustment = decimal.Parse(itemInfo.Split(new string[] { "\"no_return_fee_adjustment\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.NumberOfUnitsForPricePerUnit = decimal.Parse(itemInfo.Split(new string[] { "\"number_units_for_price_per_unit\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.TypeOfUnitsForPricePerUnit = itemInfo.Split(new string[] { "\"type_of_unit_for_price_per_unit\":" }, StringSplitOptions.None)[1].Split('"')[1];
                }
                catch { }
                try
                {
                    tmpItem.PackageHeightInches = decimal.Parse(itemInfo.Split(new string[] { "\"package_height_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    tmpItem.PackageLengthInches = decimal.Parse(itemInfo.Split(new string[] { "\"package_length_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    tmpItem.PackageWidthInches = decimal.Parse(itemInfo.Split(new string[] { "\"package_width_inches\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.ProductDescription = itemInfo.Split(new string[] { "\"product_description\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { }
                try
                {
                    tmpItem.ProductTitle = itemInfo.Split(new string[] { "\"product_title\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { }
                try
                {
                    tmpItem.Proposition65 = bool.Parse(itemInfo.Split(new string[] { "\"prop_65\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.ShippingWeight = decimal.Parse(itemInfo.Split(new string[] { "\"shipping_weight_pounds\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.ShipsAlone = bool.Parse(itemInfo.Split(new string[] { "\"ships_alone\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                }
                catch { }
                try
                {
                    tmpItem.MainImageURL = itemInfo.Split(new string[] { "\"main_image_url\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                }
                catch { }
            
     
            try
            {
                tmpItem.Bullets = new List<string>();
                string bullets = itemInfo.Split(new string[] { "\"bullets\":" }, StringSplitOptions.None)[1].Split(']')[0].Replace("[", "").Trim();
                foreach (string bullet in bullets.Split(new string[] {"\","}, StringSplitOptions.None))
                {
                    tmpItem.Bullets.Add(bullet.Replace("\"", "").Trim());
                }
            }
            catch { }
            try
            {
                tmpItem.FulfillmentCenterInventories = new List<FulfillmentCenterInventory>();
                string fulfillmentInventory = itemInfo.Split(new string[] { "\"inventory_by_fulfillment_node\":" }, StringSplitOptions.None)[1].Split(']')[0].Replace("[", "").Trim();
                foreach (string fulInv in fulfillmentInventory.Split('{'))
                {
                    if (fulInv.Replace("\r\n", "").Trim() != "")
                    {
                        string node = fulInv.Split(new string[] { "\"fulfillment_node_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        string qty = fulInv.Split(new string[] { "\"quantity\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Replace("\r\n","").Replace("}","").Trim();
                        FulfillmentCenterInventory fci = new FulfillmentCenterInventory();
                        fci.FulfillmentCenterID = node;
                        fci.FulfillmentCenterQuantity = int.Parse(qty);
                        tmpItem.FulfillmentCenterInventories.Add(fci);
                    }
                    //fci.FulfillmentCenterName = GetFulfillmentCenterInfo(node).FulfillmentNodeName; //shouldn't be needed in two spots..


                }
            }
            catch { }
            try
            {
                tmpItem.FulfillmentCenterPrices = new List<FulfillmentCenterPrice>();
                string fulfillmentPricing = itemInfo.Split(new string[] { "\"price_by_fulfillment_node\":" }, StringSplitOptions.None)[1].Split(']')[0].Replace("[", "").Trim();
                foreach (string fulPrice in fulfillmentPricing.Split('{'))
                {
                    if (fulPrice.Replace("\r\n", "").Trim() != "")
                    {
                        string node = fulPrice.Split(new string[] { "\"fulfillment_node_id\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Replace("\r\n","").Trim();
                        string price = fulPrice.Split(new string[] { "\"fulfillment_node_price\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Replace("\r\n","").Replace("}","").Trim();
                        FulfillmentCenterPrice fcp = new FulfillmentCenterPrice();
                        fcp.FulfillmentCenterID = node;
                        fcp.FulfillmentCenterItemPrice = decimal.Parse(price);
                        fcp.FulfillmentCenterName = GetFulfillmentCenterInfo(node).FulfillmentNodeName;
                        tmpItem.FulfillmentCenterPrices.Add(fcp);
                    }
                }

            }
            catch { }






            try
            {

                string StdProductCodes = (itemInfo.Split(new string[] { "\"standard_product_codes\":" }, StringSplitOptions.None)[1].Split(']')[0] + "]").Trim();

                int codeCount = System.Text.RegularExpressions.Regex.Matches(StdProductCodes, "standard_product_code").Count;
                tmpItem.ProductBarcodes = new List<StandardProductCode>();
                for (int count = 1; count < codeCount; count++)
                {
                    StandardProductCode SPC = new StandardProductCode();
                    SPC.StandardProductCodeNumber = StdProductCodes.Split(new string[] { "\"standard_product_code\":" }, StringSplitOptions.None)[count].Split(',')[0].Trim().Replace("\"", "");
                    SPC.StandardProductCodeType = StdProductCodes.Split(new string[] { "\"standard_product_code_type\":" }, StringSplitOptions.None)[count].Split('}')[0].Replace("\"", "").Replace("\r\n","").Trim();
                    tmpItem.ProductBarcodes.Add(SPC);
                }
            }
            catch(Exception WTF2)
            {
                //handle this one 2.
            }


            return tmpItem;

        }
        private JetSalesOrder ProcessSalesOrder(string JetOrderNo)
        {
            
            //CLEAR A NEW SALES ORDER VAR...
            JetSalesOrder salesOrder = new JetSalesOrder();
            salesOrder.OrderedItems = new List<OrderItem>();
            salesOrder.OrderDetails = new OrderDetail();
            salesOrder.ShippingFromLocation = new ShipFromLocation();
            salesOrder.ShippingToLocation = new ShipToLocation();           
            /////////////////////////////////////////

            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string jetReturnedData = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/withoutShipmentDetail/" + JetOrderNo,Headers);
           // System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\jetOrderRequest.txt",jetReturnedData);
           // MessageBox.Show("Check your desktop for jetOrderRequest.txt");
           // salesOrder.SalesOrderStatus = "N/A";
           // salesOrder.TPSalesOrderNumber = "N/A";

           // try
            //{
                string BuyerField = jetReturnedData.Split(new string[]{"\"buyer\":"},StringSplitOptions.None)[1].Split(new string[] {"},"},StringSplitOptions.None)[0];
                salesOrder.BuyerName = BuyerField.Split(new string[] { "\"name\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"","");
                salesOrder.BuyerPhoneNumber = BuyerField.Split(new string[] { "\"phone_number\":" }, StringSplitOptions.None)[1].Split('}')[0].Trim().Replace("\"","");
                salesOrder.FulfillmentNode = jetReturnedData.Split(new string[] { "\"fulfillment_node\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"","");
                string HasShipments = jetReturnedData.Split(new string[] { "\"has_shipments\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"","");
                if (HasShipments.ToUpper() == "FALSE")
                {
                    salesOrder.HasShipments = false;
                }
                else
                {
                    salesOrder.HasShipments = true;
                }
                string JetRequestCancel = jetReturnedData.Split(new string[] { "\"jet_request_directed_cancel\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                if (JetRequestCancel.ToUpper() == "FALSE")
                {
                    salesOrder.JetRequestDirectedCancel = false;
                }
                else
                {
                    salesOrder.JetRequestDirectedCancel = true;
                }
                try
                {
                    salesOrder.MerchantID = jetReturnedData.Split(new string[] { "\"merchant_id\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "");
                }
                catch { }
                salesOrder.MerchantOrderID = jetReturnedData.Split(new string[] { "\"merchant_order_id\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "");
                salesOrder.OrderDetails.RequestShippingCarrier = jetReturnedData.Split(new string[] { "\"request_shipping_carrier\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "");
                salesOrder.OrderDetails.RequestServiceLevel = jetReturnedData.Split(new string[] { "\"request_service_level\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "");
                salesOrder.OrderDetails.RequestShipBy = DateTime.Parse(jetReturnedData.Split(new string[] { "\"request_ship_by\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "").Replace("T"," ").Replace("Z"," ").Substring(0,23));
                salesOrder.OrderDetails.RequestDeliveryBy = DateTime.Parse(jetReturnedData.Split(new string[] { "\"request_delivery_by\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0].Trim().Replace("\"", "").Replace("T", " ").Replace("Z", " ").Substring(0,23));

                //now the fun part.. iterating each item on order... UGH!
                int ItemCount = System.Text.RegularExpressions.Regex.Matches(jetReturnedData,"\"order_item_id\":").Count;

                for (int itemCounter = 1; itemCounter <= ItemCount; itemCounter++)
                {
                    string ItemInfo = jetReturnedData.Split(new string[] { "\"order_item_id\":" }, StringSplitOptions.None)[itemCounter];
                    OrderItem itm = new OrderItem();
                  
                    itm.ItemPricing = new ItemPrices();
                    itm.JetItemID = ItemInfo.Split(',')[0].Trim().Replace("\"", "");
                    itm.TPItemNumber = ItemInfo.Split(new string[] { "\"merchant_sku\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    itm.RequestedQty = int.Parse(ItemInfo.Split(new string[] { "\"request_order_quantity\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    itm.RequestedCancelQty = int.Parse(ItemInfo.Split(new string[] { "\"request_order_cancel_qty\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    itm.ItemTaxCode = ItemInfo.Split(new string[] { "\"item_tax_code\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");

                    ItemPrices itmPricing = new ItemPrices();
                    itmPricing.BasePrice = decimal.Parse(ItemInfo.Split(new string[] { "\"base_price\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    itmPricing.ItemTax = decimal.Parse(ItemInfo.Split(new string[] { "\"item_tax\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    itmPricing.ItemShippingCost = decimal.Parse(ItemInfo.Split(new string[] {"\"item_shipping_cost\":"},StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"",""));
                    itmPricing.ItemShippingTax = decimal.Parse(ItemInfo.Split(new string[] { "\"item_shipping_tax\":" }, StringSplitOptions.None)[1].Split('}')[0].Trim().Replace("\"", ""));
                    itm.ItemPricing = itmPricing;

                    itm.ItemFees = decimal.Parse(ItemInfo.Split(new string[]{"\"item_fees\":"},StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"",""));

                    //this section gonna suck.. need to iterate all fee adjustments... as it could be more than 1...
                    if (ItemInfo.Contains("\"fee_adjustments\":") == true)
                    {
                        string feeAdjustments = ItemInfo.Split(new string[] { "\"fee_adjustments\":" }, StringSplitOptions.None)[1].Split(']')[0];
                        int itemFeeAdjustmentCount = System.Text.RegularExpressions.Regex.Matches(feeAdjustments, "\"adjustment_name\":").Count;
                        itm.ItemFeeAdjustments = new List<ItemFeeAdjustments>();
                        for (int feeCounter = 1; feeCounter <= itemFeeAdjustmentCount; feeCounter++)
                        {
                            ItemFeeAdjustments newAdjustment = new ItemFeeAdjustments();
                            newAdjustment.AdjustmentName = feeAdjustments.Split(new string[] { "\"adjustment_name\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                            newAdjustment.AdjustmentType = feeAdjustments.Split(new string[] { "\"adjustment_type\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                            newAdjustment.AdjustmentValue = decimal.Parse(feeAdjustments.Split(new string[] { "\"value\":" }, StringSplitOptions.None)[1].Split('}')[0].Trim().Replace("\"", ""));
                            itm.ItemFeeAdjustments.Add(newAdjustment);

                        }
                    }

                    itm.RegulatoryFees = decimal.Parse(ItemInfo.Split(new string[] { "\"regulatory_fees\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    TaxInfo tax = new TaxInfo();
                    tax.NonTaxableAmount = new NonTaxableAmounts();
                    tax.TaxableAmount = new TaxableAmounts();
                    tax.TaxCollectedAmount = new TaxCollectedAmounts();
                    tax.TaxJurisdiction = new TaxJurisdictions();
                    tax.TaxRate = new TaxRates();
                    tax.ZeroRatedAmount = new ZeroRatedAmounts();

                    string taxInfo = "";
                    try
                    {
                        taxInfo = jetReturnedData.Split(new string[] { "\"tax_info\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"product_title\":" }, StringSplitOptions.None)[0];
                    }
                    catch
                    {
                        //tax info probably changed or some corruption..
                    }

                    string taxJurisdictions = "";
                    string taxableAmounts = "";
                    string nonTaxableAmounts = "";
                    string zeroRatedAmounts = "";
                    string taxCollectedAmounts = "";
                    string taxRates = "";

                    try { taxJurisdictions = taxInfo.Split(new string[] { "\"tax_jurisdictions\": {" }, StringSplitOptions.None)[1].Split('}')[0]; }  catch { }
                    try { taxableAmounts = taxInfo.Split(new string[] { "\"taxable_amounts\": {" }, StringSplitOptions.None)[1].Split('}')[0]; } catch { }
                    try { nonTaxableAmounts = taxInfo.Split(new string[] { "\"non_taxable_amounts\": {" }, StringSplitOptions.None)[1].Split('}')[0]; } catch { }
                    try { zeroRatedAmounts = taxInfo.Split(new string[] { "\"zero_rated_amounts\": {" }, StringSplitOptions.None)[1].Split('}')[0]; } catch { }
                    try { taxCollectedAmounts = taxInfo.Split(new string[] { "\"tax_collected_amounts\": {" }, StringSplitOptions.None)[1].Split('}')[0]; } catch { }
                    try { taxRates = taxInfo.Split(new string[] { "\"tax_rates\": {" }, StringSplitOptions.None)[1].Split('}')[0]; } catch { }

                    if (taxJurisdictions != "")
                    {
                        tax.TaxJurisdiction.TaxLocationCode = taxJurisdictions.Split(new string[] { "\"tax_location_code\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tax.TaxJurisdiction.TaxCity = taxJurisdictions.Split(new string[] { "\"tax_city\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tax.TaxJurisdiction.TaxCounty = taxJurisdictions.Split(new string[] { "\"tax_county\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim();
                        tax.TaxJurisdiction.TaxState = taxJurisdictions.Split(new string[] { "\"tax_state\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim();                       
                    }
                    if (taxableAmounts != "")
                    {
                        tax.TaxableAmount.DistrictAmount = decimal.Parse(taxableAmounts.Split(new string[] { "\"district_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.TaxableAmount.CityAmount = decimal.Parse(taxableAmounts.Split(new string[] {"\"city_amount\":"},StringSplitOptions.None)[1].Split(',')[0].Replace("\"","").Trim());
                        tax.TaxableAmount.CountyAmount = decimal.Parse(taxableAmounts.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.TaxableAmount.StateAmount = decimal.Parse(taxableAmounts.Split(new string[] { "\"state_amount\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim());
                    }
                    if (nonTaxableAmounts != "")
                    {
                        tax.NonTaxableAmount.DistrictAmount = decimal.Parse(nonTaxableAmounts.Split(new string[] { "\"district_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.NonTaxableAmount.CityAmount = decimal.Parse(nonTaxableAmounts.Split(new string[] { "\"city_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.NonTaxableAmount.CountyAmount = decimal.Parse(nonTaxableAmounts.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.NonTaxableAmount.StateAmount = decimal.Parse(nonTaxableAmounts.Split(new string[] {"\"state_amount\":"},StringSplitOptions.None)[1].Split('}')[0].Replace("\"","").Trim());                        
                    }
                    if (zeroRatedAmounts != "")
                    {
                        tax.ZeroRatedAmount.DistrictAmount = decimal.Parse(zeroRatedAmounts.Split(new string[] {"\"district_amount\":"},StringSplitOptions.None)[1].Split(',')[0].Replace("\"","").Trim());
                        tax.ZeroRatedAmount.CityAmount = decimal.Parse(zeroRatedAmounts.Split(new string[] { "\"city_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.ZeroRatedAmount.CountyAmount = decimal.Parse(zeroRatedAmounts.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.ZeroRatedAmount.StateAmount = decimal.Parse(zeroRatedAmounts.Split(new string[] { "\"state_amount\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim());                      

                    }
                    if (taxCollectedAmounts != "")
                    {
                        tax.TaxCollectedAmount.DistrictAmount = decimal.Parse(taxCollectedAmounts.Split(new string[]{"\"district_amount\":"},StringSplitOptions.None)[1].Split(',')[0].Replace("\"","").Trim());
                        tax.TaxCollectedAmount.CityAmount = decimal.Parse(taxCollectedAmounts.Split(new string[] { "\"city_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.TaxCollectedAmount.CountyAmount = decimal.Parse(taxCollectedAmounts.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.TaxCollectedAmount.StateAmount = decimal.Parse(taxCollectedAmounts.Split(new string[] { "\"state_amount\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim());
                    }
                    if (taxRates != "")
                    {
                        tax.TaxRate.DistrictAmount = decimal.Parse(taxRates.Split(new string[]{"\"district_amount\":"},StringSplitOptions.None)[1].Split(',')[0].Replace("\"","").Trim());
                        tax.TaxRate.CityAmount = decimal.Parse(taxRates.Split(new string[] {"\"city_amount\":"},StringSplitOptions.None)[1].Split(',')[0].Replace("\"","").Trim());
                        tax.TaxRate.CountyAmount = decimal.Parse(taxRates.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Replace("\"", "").Trim());
                        tax.TaxRate.StateAmount = decimal.Parse(taxRates.Split(new string[] { "\"state_amount\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim());
                    }
                    itm.ProductTaxInfo = tax;
                    /*TaxInfo taxes = new TaxInfo();
                    taxes.TaxLocation = new TaxJurisdictions();
                    taxes.TaxLocation.TaxState = new List<string>();                    
                    //again, we must iterate thru all tax jurisdictions (at least by 'tax_state')
                    string TaxInfo = ItemInfo.Split(new string[] { "\"tax_info\":" }, StringSplitOptions.None)[1].Split(new string[] { "\"product_title\":" }, StringSplitOptions.None)[0];
                    int taxJurisdictions = System.Text.RegularExpressions.Regex.Matches(TaxInfo, "\"tax_state\":").Count;
                    for (int taxCounter = 1; taxCounter <= taxJurisdictions; taxCounter++)
                    {
                        try
                        {
                            taxes.TaxLocation.TaxState.Add(TaxInfo.Split(new string[] { "\"tax_state\":" }, StringSplitOptions.None)[taxCounter].Split(',')[0].Trim().Replace("\"","").Replace("\r\n","").Replace("}","").Trim());
                        }
                        catch
                        {
                            taxes.TaxLocation.TaxState.Add(TaxInfo.Split(new string[] { "\"tax_state\":" }, StringSplitOptions.None)[taxCounter + 1].Split('}')[0].Trim().Replace("\"","").Replace("\r\n","").Replace("}","").Trim());
                        }
                    }
                    
                    

                    TaxCollectedAmounts collectedAmounts = new TaxCollectedAmounts();
                    string TaxCollectedInfo = ItemInfo.Split(new string[] { "tax_collected_amounts\":" }, StringSplitOptions.None)[1].Split(new string[] { "}," }, StringSplitOptions.None)[0];
                    collectedAmounts.DistrictAmount = decimal.Parse(TaxCollectedInfo.Split(new string[] { "\"district_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    collectedAmounts.CityAmount = decimal.Parse(TaxCollectedInfo.Split(new string[] { "\"city_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    collectedAmounts.CountyAmount = decimal.Parse(TaxCollectedInfo.Split(new string[] { "\"county_amount\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", ""));
                    collectedAmounts.StateAmount = decimal.Parse(TaxCollectedInfo.Split(new string[] { "\"state_amount\":" }, StringSplitOptions.None)[1].Split('}')[0].Trim().Replace("\"", ""));

                    taxes.TaxCollection = collectedAmounts;

                    itm.ProductTaxInfo = taxes;
                     */ 
                    itm.ProductTitle = ItemInfo.Split(new string[] { "\"product_title\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    itm.ProductURL = ItemInfo.Split(new string[] { "\"url\":" }, StringSplitOptions.None)[1].Split('}')[0].Trim().Replace("\"", "");
                    salesOrder.OrderedItems.Add(itm);

                    salesOrder.JetSalesOrderNumber = JetOrderNo;

                    salesOrder.SalesOrderStatus = jetReturnedData.Split(new string[] { "\"status\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Trim();
                    salesOrder.OrderPlacedDate = DateTime.Parse(jetReturnedData.Split(new string[] { "\"order_placed_date\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "").Replace("T", " ").Replace("Z", " ").Substring(0, 23));
                    salesOrder.OrderTransmissionDate = DateTime.Parse(jetReturnedData.Split(new string[] { "\"order_transmission_date\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "").Replace("T", " ").Replace("Z", " ").Substring(0, 23));

                    string shipToDetails = jetReturnedData.Split(new string[] {"\"shipping_to\":"},StringSplitOptions.None)[1].Split(new string[]{"\"status\":"},StringSplitOptions.None)[0];
                    salesOrder.ShippingToLocation.ShipTo.Name = shipToDetails.Split(new string[] { "\"name\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    salesOrder.ShippingToLocation.ShipTo.PhoneNumber = shipToDetails.Split(new string[] { "\"phone_number\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Replace("\r\n","").Trim();
                    salesOrder.ShippingToLocation.ShipToAddress.Address1 = shipToDetails.Split(new string[] { "\"address1\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    salesOrder.ShippingToLocation.ShipToAddress.Address2 = shipToDetails.Split(new string[] { "\"address2\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    salesOrder.ShippingToLocation.ShipToAddress.City = shipToDetails.Split(new string[] { "\"city\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    salesOrder.ShippingToLocation.ShipToAddress.State = shipToDetails.Split(new string[] { "\"state\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim().Replace("\"", "");
                    salesOrder.ShippingToLocation.ShipToAddress.Zipcode = shipToDetails.Split(new string[] { "\"zip_code\":" }, StringSplitOptions.None)[1].Split('}')[0].Replace("\"", "").Replace("\r\n", "").Trim();
                    
                }



               /* string testmsg = "Now for the list of data pulled to check...\r\n";
                testmsg += " Buyer Name: " + salesOrder.BuyerName + "\r\n";
                testmsg += " Buyer Phone: " + salesOrder.BuyerPhoneNumber + "\r\n";
                testmsg += " Fulfillment Node: " + salesOrder.FulfillmentNode + "\r\n";
                testmsg += " Has Shipments: " + salesOrder.HasShipments.ToString() + "\r\n";
                testmsg += " Jet Request Directed Cancel: " + salesOrder.JetRequestDirectedCancel.ToString() + "\r\n";
                testmsg += " Merchant ID: " + salesOrder.MerchantID + "\r\n";
                testmsg += " Merchant Order ID: " + salesOrder.MerchantOrderID +"\r\n";
                testmsg += " Requested Shipping Carrier: " + salesOrder.OrderDetails.RequestShippingCarrier + "\r\n";
                testmsg += " Requested Service Level: " + salesOrder.OrderDetails.RequestServiceLevel + "\r\n";
                testmsg += " Requested Ship By Date: " + salesOrder.OrderDetails.RequestShipBy.ToString() + "\r\n";
                testmsg += " Requested Delivery Date: " + salesOrder.OrderDetails.RequestDeliveryBy.ToString() + "\r\n";
                
                MessageBox.Show(testmsg);*/

         //   }
         //   catch (Exception WTF)
        //    {
       //         MessageBox.Show(WTF.Message);
        //    }

            return salesOrder;
        }

        public struct ItemInfo
        {
            public string Brand;
            public decimal DisplayHeightInches;
            public decimal DisplayWidthInches;
            public decimal DisplayLengthInches;
            public bool ExcludeFromFeeAdjustments;
            public int FulfillmentTime;
            public string JetBrowseNodeID;
            public string JetRetailSKU;
            public string Manufacturer;
            public string ManufacturerPartNumber;
            public decimal JetPrice;
            public string MAPType;
            public decimal MAPPrice;
            public string MerchantID;
            public string MerchantSKU;
            public string MerchantSKUID;
            public decimal MSRP;
            public int MultipackQuantity;
            public decimal NoReturnFeeAdjustment;
            public decimal NumberOfUnitsForPricePerUnit;
            public string TypeOfUnitsForPricePerUnit;
            public decimal PackageHeightInches;
            public decimal PackageWidthInches;
            public decimal PackageLengthInches;
            public string ProductDescription;
            public string ProductTitle;
            public bool Proposition65;
            public decimal ShippingWeight;
            public bool ShipsAlone;
            public List<StandardProductCode> ProductBarcodes;
            public string MainImageURL;
            public List<string> Bullets;
            public List<FulfillmentCenterInventory> FulfillmentCenterInventories;
            public List<FulfillmentCenterPrice> FulfillmentCenterPrices;

        }
        public struct FulfillmentCenterInventory
        {
            public string FulfillmentCenterName;
            public string FulfillmentCenterID;
            public int FulfillmentCenterQuantity;
        }
        public struct FulfillmentCenterPrice
        {
            public string FulfillmentCenterName;
            public string FulfillmentCenterID;
            public decimal FulfillmentCenterItemPrice;
        }
        public struct StandardProductCode
        {
            public string StandardProductCodeNumber;
            public string StandardProductCodeType;
        }


        public struct JetSalesOrder
        {
            public string JetSalesOrderNumber;
            public string TPSalesOrderNumber;
            public string SalesOrderStatus;
            
            public string BuyerName;
            public string BuyerPhoneNumber;
            public string FulfillmentNode;
            public bool HasShipments;
            public bool JetRequestDirectedCancel;
            public string MerchantID;
            public string MerchantOrderID;

            public OrderDetail OrderDetails;
            public List<OrderItem> OrderedItems;

            public DateTime OrderPlacedDate;
            public DateTime OrderTransmissionDate;
            public ShipFromLocation ShippingFromLocation;
            public ShipToLocation ShippingToLocation;
            
            
        }
        public struct ShipToLocation
        {
            public Recipient ShipTo;
            public Address ShipToAddress;
        }
        public struct Recipient
        {
            public string Name;
            public string PhoneNumber;
        }
        public struct Address
        {
            public string Address1;
            public string Address2;
            public string City;
            public string State;
            public string Zipcode;
        }
        public struct ShipFromLocation
        {
            public string FulfillmentCenterName;
            public string FulfillmentCenterURL;
        }
        public struct OrderDetail
        {
            public string RequestShippingCarrier;
            public string RequestServiceLevel;
            public DateTime RequestShipBy;
            public DateTime RequestDeliveryBy;
        }
        public struct OrderItem
        {
            public string JetItemID; //order_item_id
            public string TPItemNumber; //merchant_sku
            public int RequestedQty;
            public int RequestedCancelQty;
            public string ItemTaxCode;
            public ItemPrices ItemPricing;
            public decimal ItemFees;
            public List<ItemFeeAdjustments> ItemFeeAdjustments;
            public decimal RegulatoryFees;
            public TaxInfo ProductTaxInfo;
            public string ProductTitle;
            public string ProductURL;


        }
        public struct ItemPrices
        {
            public decimal BasePrice;
            public decimal ItemTax;
            public decimal ItemShippingCost;
            public decimal ItemShippingTax;
        }
        public struct ItemFeeAdjustments
        {
            public string AdjustmentName;
            public string AdjustmentType;
            public decimal AdjustmentValue;
        }
        public struct TaxInfo
        {
            public TaxJurisdictions TaxJurisdiction;
            public TaxableAmounts TaxableAmount;
            public NonTaxableAmounts NonTaxableAmount;
            public ZeroRatedAmounts ZeroRatedAmount;
            public TaxCollectedAmounts TaxCollectedAmount;
            public TaxRates TaxRate;
        }
        public struct TaxJurisdictions
        {
            public string TaxLocationCode;
            public string TaxCity;
            public string TaxCounty;
            public string TaxState;
        }
        public struct TaxableAmounts
        {
            public decimal DistrictAmount;
            public decimal CityAmount;
            public decimal CountyAmount;
            public decimal StateAmount;
        }
        public struct NonTaxableAmounts
        {
            public decimal DistrictAmount;
            public decimal CityAmount;
            public decimal CountyAmount;
            public decimal StateAmount;
        }
        public struct ZeroRatedAmounts
        {
            public decimal DistrictAmount;
            public decimal CityAmount;
            public decimal CountyAmount;
            public decimal StateAmount;
        }
        public struct TaxCollectedAmounts
        {
            public decimal DistrictAmount;
            public decimal CityAmount;
            public decimal CountyAmount;
            public decimal StateAmount;
        }
        public struct TaxRates
        {
            public decimal DistrictAmount;
            public decimal CityAmount;
            public decimal CountyAmount;
            public decimal StateAmount;
        }

        private void processedOrdersLVI_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            FulfillmentInfo fi = GetFulfillmentCenterInfo("b5035a1a17d342658760c2f402db425d");
            if (fi.FulfillmentNodeName == "")
            {
                MessageBox.Show("Unfortunately no information was returned for that node id.");
            }
            else
            {
                MessageBox.Show("The name of the fulfillment center is " + fi.FulfillmentNodeName + " and it is located in " + fi.FulfillmentNodeState);
            }
        }

        private void updateIntervalTrk_Scroll(object sender, EventArgs e)
        {
            updateIntervalLbl.Text = "Update every " + updateIntervalTrk.Value.ToString() + " mins.";
        }

        private void getAllJetAcceptedOrdersBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/ready", Headers);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET Accepted Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET Accepted Orders", "User attempt to get accepted orders from Jet.com.");
                ProcessSalesOrdersToLVI(results);
            }
            //MessageBox.Show(results);

        }

        private void getAckOrdersBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/acknowledged", Headers);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET Acknowledged Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET Acknowledged Orders", "User attempt to get acknowledged orders from Jet.com.");                
                ProcessAcknowledgedOrdersToLVI(results);
            }
            //MessageBox.Show(results);
           
        }
        private void ResetAcknowledgedOrdersToLVI()
        {
            acknowledgedOrderLbl.Text = "Acknowledged Order";
            acknowledgedOrderLVI.Items.Clear();
            acknowledgedOrderLVI.Columns.Clear();
            acknowledgedOrderLVI.Clear();

            ColumnHeader orderNo = new ColumnHeader();
            orderNo.Text = "Ack. Order Number";
            orderNo.TextAlign = HorizontalAlignment.Center;
            orderNo.Width = acknowledgedOrderLVI.Width - 25;
            acknowledgedOrderLVI.Columns.Add(orderNo);
        }
        private void ResetInProgressOrdersToLVI()
        {
            inProgressOrderLbl.Text = "In Progress Order";
            inProgressOrderLVI.Items.Clear();
            inProgressOrderLVI.Columns.Clear();
            inProgressOrderLVI.Clear();

            ColumnHeader orderNo = new ColumnHeader();
            orderNo.Text = "InProgress Order";
            orderNo.TextAlign = HorizontalAlignment.Center;
            orderNo.Width = inProgressOrderLVI.Width - 25;
            inProgressOrderLVI.Columns.Add(orderNo);
        }
        private void ResetCompletedOrdersToLVI()
        {
            completedOrderLbl.Text = "Completed Orders";
            completedOrderLVI.Items.Clear();
            completedOrderLVI.Columns.Clear();
            completedOrderLVI.Clear();

            ColumnHeader orderNo = new ColumnHeader();
            orderNo.Text = "Completed Order";
            orderNo.TextAlign = HorizontalAlignment.Center;
            orderNo.Width = completedOrderLVI.Width - 25;
            completedOrderLVI.Columns.Add(orderNo);
        }
        private void ProcessAcknowledgedOrdersToLVI(string results)
        {
            ResetAcknowledgedOrdersToLVI();
            string actualResults = results.Split('[')[1].Split(']')[0];
            if (actualResults.Length > 0)
            {
                try
                {
                    foreach (string order in actualResults.Split(new string[] { "\"/" }, StringSplitOptions.None))
                    {
                        if (order.ToUpper().Contains("ORDERS") == true)
                        {
                            string realOrder = order.Split('/')[2].Split('"')[0];
                            ListViewItem orderItem = new ListViewItem();
                            orderItem.Text = realOrder;
                            acknowledgedOrderLVI.Items.Add(orderItem);
                        }
                    }
                    acknowledgedOrderLbl.Text = "Acknowledged Orders (" + acknowledgedOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    AddStatusEvent(true, "Processing Acknowledged Orders", "Processes orders from Jet.com that we acknowledged.");
                }
                catch (Exception wtf)
                {
                    AddStatusEvent(false, "Processing Acknowledged Orders", wtf.Message);
                }
            }
            else
            {
                acknowledgedOrderLbl.Text = "Acknowledged Orders (0)\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Acknowledged Orders", "Completed OK but no orders returned.");
            }
        }
        private void ProcessInProgressOrdersToLVI(string results)
        {
            ResetInProgressOrdersToLVI();
            string actualResults = results.Split('[')[1].Split(']')[0];
            if (actualResults.Length > 0)
            {
                try
                {
                    foreach (string order in actualResults.Split(new string[] { "\"/" }, StringSplitOptions.None))
                    {
                        if (order.ToUpper().Contains("ORDERS") == true)
                        {
                            string realOrder = order.Split('/')[2].Split('"')[0];
                            ListViewItem orderItem = new ListViewItem();
                            orderItem.Text = realOrder;
                            inProgressOrderLVI.Items.Add(orderItem);
                        }
                    }
                    //pendingOrderLbl.Text = "Accepted Orders (" + pendingOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    inProgressOrderLbl.Text = "In Progress Orders (" + inProgressOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    AddStatusEvent(true, "Processing In Progress Orders", "Processes orders from Jet.com that are currently in progress.");
                }
                catch (Exception wtf)
                {
                    AddStatusEvent(false, "Processing In Progress Orders", wtf.Message);
                }
            }
            else
            {
                inProgressOrderLbl.Text = "In Progress Orders (0)\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Pending Orders", "Completed OK but no orders returned.");
            }
        }
        private void ProcessCompletedOrdersToLVI(string results)
        {
           
            ResetCompletedOrdersToLVI();
            string actualResults = results.Split('[')[1].Split(']')[0];
            if (actualResults.Length > 0)
            {
                try
                {
                    foreach (string order in actualResults.Split(new string[] { "\"/" }, StringSplitOptions.None))
                    {
                        if (order.ToUpper().Contains("ORDERS") == true)
                        {
                            string realOrder = order.Split('/')[2].Split('"')[0];
                            ListViewItem orderItem = new ListViewItem();
                            orderItem.Text = realOrder;
                            completedOrderLVI.Items.Add(orderItem);
                        }
                    }
                    //pendingOrderLbl.Text = "Accepted Orders (" + pendingOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    completedOrderLbl.Text = "Completed Orders (" + completedOrderLVI.Items.Count.ToString() + ")\r\nLast Updated: " + DateTime.Now.ToString();
                    AddStatusEvent(true, "Processing Completed Orders", "Processes orders from Jet.com that have been completed.");
                }
                catch (Exception wtf)
                {
                    AddStatusEvent(false, "Processing Completed Orders", wtf.Message);
                }
            }
            else
            {
                completedOrderLbl.Text = "Completed Orders (0)\r\nLast Updated: " + DateTime.Now.ToString();
                AddStatusEvent(true, "Processing Completed Orders", "Completed OK but no orders returned.");
            }
        }

        private void getProgressOrdersBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/inprogress", Headers);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET In Progress Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET In Progress Orders", "User attempt to get in progress orders from Jet.com.");
                ProcessInProgressOrdersToLVI(results);
            }
            //MessageBox.Show(results);
        }

        private void getCompletedOrdersBtn_Click(object sender, EventArgs e)
        {
            if (AuthBearer.Length <= 0)
            {
                GetAuthBearer();
            }


            string PutContentType = "Content-type: application/json";
            string Authorization = "Authorization: Bearer " + AuthBearer;
            List<string> Headers = new List<string>();
            Headers.Add(PutContentType);
            Headers.Add(Authorization);
            string results = Connectivity.HTTPCalls.HTTP_GET("https://merchant-api.jet.com/api/orders/complete", Headers);
            if (results.Contains("ERROR:") == true)
            {
                AddStatusEvent(false, "Manual GET Completed Orders", results);
            }
            else
            {
                AddStatusEvent(true, "Manual GET Completed Orders", "User attempt to get copmleted orders from Jet.com.");
                ProcessCompletedOrdersToLVI(results);
            }
            //MessageBox.Show(results);
        }

        private void plusMinus6Lbl_Click(object sender, EventArgs e)
        {
            if (plusMinus6Lbl.Text == "+")
            {
                plusMinus6Lbl.Text = "-";
                returnOrdersBox.Height += 150;
            }
            else
            {
                plusMinus6Lbl.Text = "+";
                returnOrdersBox.Height -= 150;
            }
            RefreshForm();
        }

        private void SynchronizeToMAS(JetSalesOrder SOI)
        {
           
            string JetOrderNumber = GetNextJetSalesOrderNumber();
            //we should have all necessary info to vi import this data into mas.
            

        }
        private string GetNextJetSalesOrderNumber()
        {
            List<string> orderNo = Connectivity.SQLCalls.sqlQuery("SELECT JetDotCom FROM NextSalesOrderNumber");
            if (orderNo.Count > 0)
            {
                //gets next available sales order number for jet.com. the pulled number (which is available) gets +1 added to it, and updated back in sql.
                string SO = orderNo[0].Split('|')[0];
                int soNum = int.Parse(SO.Split('J')[1]);
                string NextSO = "J" + (soNum + 1).ToString("000000");
                Connectivity.SQLCalls.sqlQuery("UPDATE NextSalesOrderNumber SET JetDotCom='" + NextSO + "' WHERE JetDotCom='" + SO + "'");
                return SO;
            }
            else
            {
                //fatal error.. cannot locate next sales order no...
                MessageBox.Show("Big issue.. we have Jet.com order(s) and cannot get the next order number. SQL Issue? Internet Issue? Tell Tom immediately as this is very serious.");
                return "";
            }
        }
    }
}
