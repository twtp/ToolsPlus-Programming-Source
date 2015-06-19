using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace PriceTwoSpyService
{
    public partial class guiFrm : Form
    {
        public const bool TestServer = false;

        public string getCurrentPricingDataRequest = "";
        public string getProductsRequest = "";
        public string getInsertProductRequest = "";
        public string getInsertURLRequest = "";
        public string findItemByItemNumber = "";
        public string automatchRequest = "";
        public decimal EURtoUSD = 1.28M;

        public struct ScrapeInfo
        {
            public string URLID;
            public string URL;
            public string SiteHumanName;
            public bool ShowInReports;
            public bool Active;
            public int CounterFailed;
            public string LastError;
            public DateTime LastChecked;
            public DateTime DateChecked;
            public decimal Amount;
            public string Currency;
            public bool Available;

        }
        public struct ItemData
        {
            
            public string ProductID;
            public string ProductName;
            public string ProductBrand;            
            public string CheckFrequency;
            public int CheckFrequencyInterval;
            public bool Active;
            public bool ShowInReports;
            public string SKU;
            public string InternalID; //TP ItemNumber....
            public string CategoryName;
            public int CategoryID;
            public decimal MinimumPriceAmount;
            public string MinimumPriceCurrency;
            public decimal MaximumPriceAmount;
            public string MaximumPriceCurrency;
            public DateTime LastCheckDateTime;
            public DateTime NextCheckedDateTime;
            public DateTime LastChangeDateTime;
            public List<ScrapeInfo> ScrapeData;

        }
        public struct ConnectionInfo
        {
            public string API_CurrentKey;
            public string API_BaseURL;
        }
        ConnectionInfo CurrentTestConnection = new ConnectionInfo();
        ConnectionInfo CurrentProductionConnection = new ConnectionInfo();
        List<ItemData> ItemsInformation = new List<ItemData>();
        public bool isScheduledTask = false;
        public guiFrm(int ServiceType)
        {
            //Currently only scheduled task and manual mode are programmed and maintained.

            //ServiceType 0 = manual mode (no args)
            //ServiceType 1 = scheduled task (/scheduledtask)
            //ServiceType 2 = service (/service)

            InitializeComponent();
            if (ServiceType == 1 | ServiceType == 2)
            {
                groupBox2.Left = 2;
                this.timeLbl.Left = 2;
                this.timeLbl.Top = 0;
                groupBox2.Top = timeLbl.Bottom;
                this.Width = groupBox2.Width +10;
                this.Height = groupBox2.Bottom+25;
                groupBox2.BringToFront();
                timeLbl.BringToFront();

            }
            if (ServiceType == 1)
            {
                isScheduledTask = true;
                bool result = StartAPIInterface();
                SetupLVI();
                button7_Click(null, null);
               
            }
            
        }

        private void guiFrm_Load(object sender, EventArgs e)
        {
            //lets get our money conversion rates first...

            if (isScheduledTask == true)
            {
                Application.Exit();
            }
            bool result = StartAPIInterface();
            SetupLVI();

        }

        private bool StartAPIInterface()
        {
            bool apiDataOK = GetAPIData();
            if (apiDataOK == true)
            {
                populateGetCurrentPricingDataRequest(TestServer);
                populateGetProductsRequest(TestServer);                                                
                return true;
            }
            else
            {
                return false;
            }
        }
        private bool GetAPIData()
        {
            //bool failure = false;
            List<string> apiDataResults = Connectivity.SQLCalls.sqlQuery("SELECT APIKey,APIUrl FROM APIKeys WHERE APIName='Price2Spy - Test Server'");
            if (apiDataResults.Count <= 0)
            {
                //email tom saying the stupid api key is missing...
                currentAPILbl.Text = "Test API Key: N/A (None Found)";
                //failure = true;
            }
            else if (apiDataResults.Count > 1)
            {
                //somehow we have more than one api key... notify tom via email
                currentAPILbl.Text = "Test API Key: N/A (Multiple Keys Found)";
                //failure = true;
            }
            else
            {
                //1 apikey returned, as expected...
                CurrentTestConnection.API_CurrentKey = apiDataResults[0].Split('|')[0];
                CurrentTestConnection.API_BaseURL = apiDataResults[0].Split('|')[1];
                currentAPILbl.Text = "Test API Key: " + CurrentTestConnection.API_CurrentKey;
                currentAPIUrlLbl.Text = "Test API URL: " + CurrentTestConnection.API_BaseURL;
               

            }
            apiDataResults = new List<string>();
            apiDataResults = Connectivity.SQLCalls.sqlQuery("SELECT APIKey,APIUrl FROM APIKeys WHERE APIName='Price2Spy - Production Server'");
            if (apiDataResults.Count <= 0)
            {
                //email tom saying the stupid api key is missing...
                this.productionAPIKeyLbl.Text = "Production API Key: N/A (None Found)";
                return false;
            }
            else if (apiDataResults.Count > 1)
            {
                //somehow we have more than one api key... notify tom via email
                this.productionAPIKeyLbl.Text = "Production API Key: N/A (Multiple Keys Found)";
                return false;
            }
            else
            {
                //1 apikey returned, as expected...
                CurrentProductionConnection.API_CurrentKey = apiDataResults[0].Split('|')[0];
                CurrentProductionConnection.API_BaseURL = apiDataResults[0].Split('|')[1];
                this.productionAPIKeyLbl.Text = "Production API Key: " + CurrentProductionConnection.API_CurrentKey;
                this.productionAPIUrlLbl.Text = "Production API URL: " + CurrentProductionConnection.API_BaseURL;
                return true;

            }

        }




        private void populateGetCurrentPricingDataRequest(bool TestServer)
        {
            getCurrentPricingDataRequest = "";
            getCurrentPricingDataRequest += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:sear=\"http://www.price2spy.com/api/search\">\r\n";
            getCurrentPricingDataRequest += " <soapenv:Header/>\r\n";
            getCurrentPricingDataRequest += "  <soapenv:Body>\r\n";
            getCurrentPricingDataRequest += "   <sear:getCurrentPricingDataRequest>\r\n";
            //getCurrentPricingDataRequest += "    <sear:search category=\"Tools\" />\r\n";
            getCurrentPricingDataRequest += "    <sear:search/>\r\n";
            if (TestServer == true)
            {
                getCurrentPricingDataRequest += "    <sear:apiKey>" + CurrentTestConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            else
            {
                getCurrentPricingDataRequest += "    <sear:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            getCurrentPricingDataRequest += "   </sear:getCurrentPricingDataRequest>\r\n";
            getCurrentPricingDataRequest += "  </soapenv:Body>\r\n";
            getCurrentPricingDataRequest += "</soapenv:Envelope>\r\n";


        }
        private void populateGetProductsRequest(bool TestServer)
        {
            getProductsRequest = "";
            getProductsRequest += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:sear=\"http://www.price2spy.com/api/search\">\r\n";
            getProductsRequest += " <soapenv:Header/>\r\n";
            getProductsRequest += "  <soapenv:Body>\r\n";
            getProductsRequest += "   <sear:getProductsRequest>\r\n";
            getProductsRequest += "    <sear:search/>\r\n";
            if (TestServer == true)
            {
                getProductsRequest += "    <sear:apiKey>" + CurrentTestConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            else
            {
                getProductsRequest += "    <sear:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            getProductsRequest += "   </sear:getProductsRequest>\r\n";
            getProductsRequest += "  </soapenv:Body>\r\n";
            getProductsRequest += "</soapenv:Envelope>\r\n";

        }
        private void populateInsertProductRequest(bool TestServer)
        {
            getInsertProductRequest = "";
            getInsertProductRequest += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:mod=\"http://www.price2spy.com/api/modify\">\r\n";
            getInsertProductRequest += " <soapenv:Header/>\r\n";
            getInsertProductRequest += "  <soapenv:Body>\r\n";
            getInsertProductRequest += "   <mod:insertProductRequest>\r\n";
            getInsertProductRequest += "    <mod:insert productName=\"" + "{TITLE}" + "\" checkFrequencyType=\"" + "{CHECKFREQUENCYTYPE}" + "\" checkFrequencyInterval=\"" + "{CHECKFREQUENCYINTERVAL}" + "\" active=\"" + "{ISACTIVE}" + "\" sku=\"" + "{SKU}" + "\" internalId=\"" + "{INTERNALID}" + "\" targetPrice=\"" + "{TARGETPRICE}" + "\" purchasePrice=\"" + "{PURCHASEPRICE}" + "\" categoryName=\"" + "{CATEGORYNAME}" + "\" brandName=\"" + "{BRANDNAME}" + "\" supplierName=\"" + "{SUPPLIERNAME}" + "\"/>\r\n";
            if (TestServer == true)
            {
                getInsertProductRequest += "    <mod:apiKey>" + CurrentTestConnection.API_CurrentKey + "</mod:apiKey>\r\n";
            }
            else
            {
                getInsertProductRequest += "    <mod:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</mod:apiKey>\r\n";
            }
            getInsertProductRequest += "   </mod:insertProductRequest>\r\n";
            getInsertProductRequest += "  </soapenv:Body>\r\n";
            getInsertProductRequest += "</soapenv:Envelope>\r\n";

        }
        private void populateInsertURLRequest(bool TestServer)
        {
            getInsertURLRequest = "";
            getInsertURLRequest += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:mod=\"http://www.price2spy.com/api/modify\">\r\n";
            getInsertURLRequest += " <soapenv:Header/>\r\n";
            getInsertURLRequest += " <soapenv:Body>\r\n";
            getInsertURLRequest += "  <mod:insertUrlRequest>\r\n";
            getInsertURLRequest += "   <mod:insert productId=\"" + "{PRODUCTID}" + "\" url=\"{URLTOADD}\" active=\"" + "{ISACTIVE}" + "\" showInReports=\"" + "{SHOWINREPORTS}" + "\"/>\r\n";
            if (TestServer == true)
            {
                getInsertURLRequest += "   <mod:apiKey>" + CurrentTestConnection.API_CurrentKey + "</mod:apiKey>\r\n";
            }
            else
            {
                getInsertURLRequest += "   <mod:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</mod:apiKey>\r\n";
            }
            getInsertURLRequest += "  </mod:insertUrlRequest>\r\n";
            getInsertURLRequest += " </soapenv:Body>\r\n";
            getInsertURLRequest += "</soap:Envelope>\r\n";

        }
        private void populateFindItemByItemNumberRequest(bool TestServer)
        {
            findItemByItemNumber = "";
            findItemByItemNumber += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:sear=\"http://www.price2spy.com/api/search\">\r\n";
            findItemByItemNumber += " <soapenv:Header/>\r\n";
            findItemByItemNumber += " <soapenv:Body>\r\n";
            findItemByItemNumber += "  <sear:getProductsRequest>\r\n";
            findItemByItemNumber += "   <sear:search internalId=\"{ITEMNUMBER}\"/>\r\n";
            if (TestServer == true)
            {
                findItemByItemNumber += "   <sear:apiKey>" + CurrentTestConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            else
            {
                findItemByItemNumber += "   <sear:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            }
            findItemByItemNumber += "  </sear:getProductsRequest>\r\n";
            findItemByItemNumber += " </soapenv:Body>\r\n";
            findItemByItemNumber += "</soapenv:Envelope>\r\n";

        }
        private void populateAutomatchRequest()
        {
            //unavailable on test server
            automatchRequest = "";
            automatchRequest += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:sear=\"http://www.price2spy.com/api/search\">\r\n";
            automatchRequest += " <soapenv:Header/>\r\n";
            automatchRequest += " <soapenv:Body>\r\n";
            automatchRequest += "  <sear:getAutomatchUrlsRequest>\r\n";
            automatchRequest += "   <sear:search id=\"{PRODUCTID}\" idMode=\"INTERNAL_ID\"/>\r\n";
            automatchRequest += "   <sear:apiKey>" + CurrentProductionConnection.API_CurrentKey + "</sear:apiKey>\r\n";
            automatchRequest += "  </sear:getAutomatchUrlsRequest>\r\n";
            automatchRequest += " </soapenv:Body>\r\n";
            automatchRequest += "</soapenv:Envelope>\r\n";
        }


        private void GetConversionRates()
        {
            statusLst.Items.Add("Status: Getting EUR to USD Conversion Factor...");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();
          /*  try
            {
                string conversionRate = Connectivity.HTTPCalls.HTTP_GET("http://rate-exchange.appspot.com/currency?from=EUR&to=USD");
                textBox1.Text = conversionRate;
                decimal rate = decimal.Parse(conversionRate.Split(new string[] { "\"rate\":" }, StringSplitOptions.None)[1].Split(',')[0].Trim());
                statusLst.Items.Add("Status: EUR -> USD = $" + rate.ToString());
                EURtoUSD = rate;
            }
            catch (Exception wtf)
            {
                statusLst.Items.Add("Status: Server not available. Using default.");
                statusLst.Items.Add("Status: EUR -> USD = $" + EURtoUSD.ToString());
                statusLst.Refresh();

            }
           
           */

            statusLst.Items.Add("Status: Server not available. Using default.");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Items.Add("Status: EUR -> USD = $" + EURtoUSD.ToString());
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();


        }

        private void PullDataFromServer(string ServerURL, string ServerAPIKey)
        {
            button1.Enabled = false;

            GetConversionRates();

            statusLst.Items.Add("Status: Sending API Call. Please Be Patient.");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();
            //string testresults = Connectivity.HTTPCalls.HTTP_POST(CurrentConnection.API_BaseURL, getProductsRequest);
            string testresults = Connectivity.HTTPCalls.HTTP_POST(ServerURL, getCurrentPricingDataRequest);
            FormatXMLForViewing(testresults);
            statusLst.Items.Add("Status: API Sent/Recieved");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();
            if (testresults.Contains("<ns3:get") == true)
            {
                //succesfull response
                statusLst.Items.Add("Status: Response OK");
                statusLst.SelectedIndex = statusLst.Items.Count - 1;
                statusLst.SelectedIndex = -1;
                statusLst.Refresh();
                string items = testresults.Split(new string[] { "<ns3:products>" }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { "</ns3:products>" }, StringSplitOptions.None)[0];
                foreach (string itm in items.Split(new string[] { "<ns3:product>" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string tmpItem = itm.Split(new string[] { "</ns3:product>" }, StringSplitOptions.None)[0];

                    ItemData nextItem = new ItemData();
                    nextItem.ProductID = tmpItem.Split(new string[] { "<ns3:productId>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:productId>" }, StringSplitOptions.None)[0];
                    nextItem.ProductName = tmpItem.Split(new string[] { "<ns3:productName>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:productName>" }, StringSplitOptions.None)[0];
                    try
                    {
                        nextItem.ProductBrand = tmpItem.Split(new string[] { "<ns3:productBrand>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:productBrand>" }, StringSplitOptions.None)[0];
                    }
                    catch { nextItem.ProductBrand = "N/A"; }
                    nextItem.CheckFrequency = tmpItem.Split(new string[] { "<ns3:checkFrequencyType>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:checkFrequencyType>" }, StringSplitOptions.None)[0];
                    nextItem.CheckFrequencyInterval = int.Parse(tmpItem.Split(new string[] { "<ns3:checkFrequencyInterval>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:checkFrequencyInterval>" }, StringSplitOptions.None)[0]);
                    nextItem.Active = Convert.ToBoolean(tmpItem.Split(new string[] { "<ns3:active>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:active>" }, StringSplitOptions.None)[0]);
                    nextItem.ShowInReports = Convert.ToBoolean(tmpItem.Split(new string[] { "<ns3:showInReports>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:showInReports>" }, StringSplitOptions.None)[0]);
                    try { nextItem.SKU = tmpItem.Split(new string[] { "<ns3:sku>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:sku>" }, StringSplitOptions.None)[0]; }
                    catch { nextItem.SKU = "N/A"; }
                    nextItem.InternalID = tmpItem.Split(new string[] { "<ns3:internalId>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:internalId>" }, StringSplitOptions.None)[0];
                    try
                    {
                        nextItem.CategoryName = tmpItem.Split(new string[] { "<ns3:categoryName>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:categoryName>" }, StringSplitOptions.None)[0];
                        nextItem.CategoryID = int.Parse(tmpItem.Split(new string[] { "<ns3:categoryId>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:categoryId>" }, StringSplitOptions.None)[0]);
                    }
                    catch { nextItem.CategoryID = -1; nextItem.CategoryName = "N/A"; }
                    string minPricings = "";
                    string maxPricings = "";
                    if (tmpItem.Contains("<ns3:minPrice/>") == false)
                    {
                        minPricings = tmpItem.Split(new string[] { "<ns3:minPrice>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:minPrice>" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        minPricings = "";
                    }
                    if (tmpItem.Contains("<ns3:maxPrice/>") == false)
                    {
                        maxPricings = tmpItem.Split(new string[] { "<ns3:maxPrice>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:maxPrice>" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        maxPricings = "";
                    }
                    if (minPricings != "")
                    {
                        nextItem.MinimumPriceAmount = decimal.Parse(minPricings.Split(new string[] { "<ns3:amount>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:amount>" }, StringSplitOptions.None)[0]);
                        nextItem.MinimumPriceCurrency = minPricings.Split(new string[] { "<ns3:currency>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:currency>" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        nextItem.MinimumPriceAmount = 0;
                        nextItem.MinimumPriceCurrency = "N/A";
                    }
                    if (maxPricings != "")
                    {
                        nextItem.MaximumPriceAmount = decimal.Parse(maxPricings.Split(new string[] { "<ns3:amount>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:amount>" }, StringSplitOptions.None)[0]);
                        nextItem.MaximumPriceCurrency = maxPricings.Split(new string[] { "<ns3:currency>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:currency>" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        nextItem.MaximumPriceAmount = 0;
                        nextItem.MaximumPriceCurrency = "N/A";
                    }
                    nextItem.LastCheckDateTime = DateTime.Parse(tmpItem.Split(new string[] { "<ns3:lastChecked>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:lastChecked>" }, StringSplitOptions.None)[0]);
                    nextItem.NextCheckedDateTime = DateTime.Parse(tmpItem.Split(new string[] { "<ns3:nextChecked>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:nextChecked>" }, StringSplitOptions.None)[0]);
                    if (tmpItem.Contains("<ns3:lastChange>") == true)
                    {
                        nextItem.LastChangeDateTime = DateTime.Parse(tmpItem.Split(new string[] { "<ns3:lastChange>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:lastChange>" }, StringSplitOptions.None)[0]);
                    }
                    else
                    {
                        nextItem.LastChangeDateTime = new DateTime();
                    }

                    List<ScrapeInfo> newScrapes = new List<ScrapeInfo>();



                    string scraperData = tmpItem.Split(new string[] { "<ns3:urls>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:urls>" }, StringSplitOptions.None)[0];


                    //need to find a way to decypher between <ns3:url> and <ns3:url>... wait.. that sounds difficult...
                    scraperData = scraperData.Replace("<ns3:url><ns3:urlId>", "<scrape><ns3:urlId>");


                    foreach (string scrape in scraperData.Split(new string[] { "<scrape>" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string tmpScraper = scrape;
                        ScrapeInfo thisScrape = new ScrapeInfo();

                        thisScrape.URLID = tmpScraper.Split(new string[] { "<ns3:urlId>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:urlId>" }, StringSplitOptions.None)[0];

                        thisScrape.URL = tmpScraper.Split(new string[] { "<ns3:url>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:url>" }, StringSplitOptions.None)[0];
                        thisScrape.SiteHumanName = tmpScraper.Split(new string[] { "<ns3:siteHumanName>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:siteHumanName>" }, StringSplitOptions.None)[0];
                        thisScrape.ShowInReports = Convert.ToBoolean(tmpScraper.Split(new string[] { "<ns3:showInReports>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:showInReports>" }, StringSplitOptions.None)[0]);
                        thisScrape.Active = Convert.ToBoolean(tmpScraper.Split(new string[] { "<ns3:active>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:active>" }, StringSplitOptions.None)[0]);
                        thisScrape.CounterFailed = int.Parse(tmpScraper.Split(new string[] { "<ns3:counterFailed>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:counterFailed>" }, StringSplitOptions.None)[0]);
                        if (tmpScraper.Contains("<ns3:lastError>") == true)
                        {
                            thisScrape.LastError = tmpScraper.Split(new string[] { "<ns3:lastError>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:lastError>" }, StringSplitOptions.None)[0];
                        }
                        else
                        {
                            thisScrape.LastError = "N/A";
                        }
                        if (tmpScraper.Contains("<ns3:lastChecked>") == true)
                        {
                            thisScrape.LastChecked = DateTime.Parse(tmpScraper.Split(new string[] { "<ns3:lastChecked>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:lastChecked>" }, StringSplitOptions.None)[0]);
                        }
                        else
                        {
                            thisScrape.LastChecked = new DateTime();
                        }
                        if (tmpScraper.Contains("<ns3:dateChecked>") == true)
                        {
                            thisScrape.DateChecked = DateTime.Parse(tmpScraper.Split(new string[] { "<ns3:dateChecked>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:dateChecked>" }, StringSplitOptions.None)[0]);
                        }
                        else
                        {
                            thisScrape.DateChecked = new DateTime();
                        }
                        if (tmpScraper.Contains("<ns3:amount>") == true)
                        {
                            thisScrape.Amount = decimal.Parse(tmpScraper.Split(new string[] { "<ns3:amount>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:amount>" }, StringSplitOptions.None)[0]);
                        }
                        else
                        {
                            thisScrape.Amount = 0;
                        }
                        if (tmpScraper.Contains("<ns3:currency>") == true)
                        {
                            thisScrape.Currency = tmpScraper.Split(new string[] { "<ns3:currency>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:currency>" }, StringSplitOptions.None)[0];
                        }
                        else
                        {
                            thisScrape.Currency = "N/A";
                        }
                        if (tmpScraper.Contains("<ns3:available>") == true)
                        {
                            thisScrape.Available = Convert.ToBoolean(tmpScraper.Split(new string[] { "<ns3:available>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:available>" }, StringSplitOptions.None)[0]);
                        }
                        else
                        {
                            thisScrape.Available = false;
                        }
                        if (thisScrape.SiteHumanName.ToUpper() != "TOOLS PLUS")
                        {
                            newScrapes.Add(thisScrape);
                        }


                    }
                    nextItem.ScrapeData = newScrapes;
                   
                    ItemsInformation.Add(nextItem);

                }
                //MessageBox.Show("This process has finished. We have " + ItemsInformation.Count.ToString() + " total products being scraped. The First Product's name is " + ItemsInformation[0].ProductName + " and it's first scraping url is " + ItemsInformation[0].ScrapeData[0].URL + ".");
                PopulateLVI();
                statusLst.Items.Add("Status: Download Completed.");
                statusLst.SelectedIndex = statusLst.Items.Count - 1;
                statusLst.SelectedIndex = -1;
                statusLst.Refresh();

            }
            else
            {
                //no bueno
                statusLst.Items.Add("Status: Response Error");
                statusLst.SelectedIndex = statusLst.Items.Count - 1;
                statusLst.SelectedIndex = -1;
                statusLst.Refresh();

            }
            textBox1.Text = testresults;
            UpdateDatabase();
            statusLst.Items.Add("Status: Process Complete.");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();
            button1.Enabled = true;
            //MessageBox.Show("Product ID for " + ((ItemData)productLVI.Items[0].Tag).ProductName + ": " + ((ItemData)productLVI.Items[0].Tag).ProductID);
        }

        private void button1_Click(object sender, EventArgs e)
        {

            PullDataFromServer(this.CurrentTestConnection.API_BaseURL, this.CurrentTestConnection.API_CurrentKey);
         
        }

        private void UpdateDatabase()
        {
            statusLst.Items.Add("Status: Updating Database...");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Refresh();

            foreach (ItemData id in ItemsInformation)
            {
                foreach (ScrapeInfo si in id.ScrapeData)
                {
                    //ItemNumber varchar(30) not null
                    //LastUpdated datetime not null
                    //Source varchar(15), not null
                    //StoreName varchar(50), not null (I decided it's best to do Source + "-" + StoreName for StoreName to ID Price2Spy in POINV
                    //URL varchar(255), not null
                    //FreeShipping bit not null
                    //Price money not null

                    string ItemNumber = id.InternalID;
                    string LastUpdated = si.LastChecked.ToString();
                    string Source = "Price2Spy";
                    string StoreName = si.SiteHumanName;
                    string URL = si.URL;
                    bool FreeShipping = false;
                    decimal Price = 0;
                    if (si.Currency == "USD")
                    {
                        Price = si.Amount;
                    }
                    if (si.Currency == "EUR")
                    {
                        Price = si.Amount * EURtoUSD;
                    }

                    //see if item exists in db by itemnumber and source and storename and url.... if so update, if not create

                    List<string> dbListing = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM PriceComparisons WHERE ItemNumber='" + ItemNumber + "' AND Source='" + Source + "' AND StoreName='" + "P" + StoreName + "' AND URL='" + URL + "'");
                    if (Price > 0)
                    {
                        if (dbListing.Count > 0)
                        {
                            //exists... update...
                            string IDnum = dbListing[0].Split('|')[0];
                            string updateQry = "UPDATE PriceComparisons SET LastUpdated='" + LastUpdated + "',Source='" + Source + "',StoreName='" + "P" + StoreName + "',URL='" + URL + "',FreeShipping=" + Convert.ToInt32(FreeShipping) + ",Price=" + Price.ToString("#.##") + " WHERE ID=" + IDnum;
                            //MessageBox.Show(updateQry);
                            Connectivity.SQLCalls.sqlQuery(updateQry);
                        }
                        else
                        {
                            //doesnt exist, add..
                            string insertQry = "INSERT INTO PriceComparisons (ItemNumber,LastUpdated,Source,StoreName,URL,FreeShipping,Price) VALUES('" + ItemNumber + "','" + LastUpdated + "','" + Source + "','" + "P" + StoreName + "','" + URL + "'," + Convert.ToInt32(FreeShipping) + "," + Price.ToString("#.##") + ")";
                            //MessageBox.Show(insertQry);
                            Connectivity.SQLCalls.sqlQuery(insertQry);


                        }
                    }
                    else
                    {
                        statusLst.Items.Add("Item " + ItemNumber + " has invalid data and won't be inserted/updated in the database.");
                        statusLst.SelectedIndex = statusLst.Items.Count - 1;
                        statusLst.SelectedIndex = -1;
                        statusLst.Refresh();
                    }


                }
            }

        }

        private void FormatXMLForViewing(string XML)
        {
            textBox1.Text = XML.Replace("> ", ">\r\n");


            System.Xml.Xsl.XslCompiledTransform xTrans = new System.Xml.Xsl.XslCompiledTransform();
            xTrans.Load("xmlread.cfg");
            StringReader sr = new StringReader(XML);
            XmlReader xReader = XmlReader.Create(sr);
            System.IO.MemoryStream ms = new MemoryStream();
            xTrans.Transform(xReader, null, ms);
            ms.Position = 0;
            webView.DocumentStream = ms;
            
           
        }

        private void SetupLVI()
        {
            productLVI.Items.Clear();
            productLVI.Columns.Clear();
            productLVI.Clear();
            productLVI.BackColor = Color.FromArgb(224, 224, 224);

            ColumnHeader LVIID = new ColumnHeader();
            ColumnHeader ItemNumber = new ColumnHeader();
            ColumnHeader ItemSKU = new ColumnHeader();
            ColumnHeader ItemCategory = new ColumnHeader();
            ColumnHeader ItemBrand = new ColumnHeader();
            ColumnHeader ItemName = new ColumnHeader();
            ColumnHeader ItemScrapedURL = new ColumnHeader();
            ColumnHeader ItemScrapedPrice = new ColumnHeader();
            ColumnHeader ItemScrapedCurrency = new ColumnHeader();
            ColumnHeader ItemLastScraped = new ColumnHeader();
            ColumnHeader ItemLastError = new ColumnHeader();
            ColumnHeader ItemSiteHumanName = new ColumnHeader();
            ColumnHeader ItemScraperActive = new ColumnHeader();


            LVIID.Text = "";
            ItemNumber.Text = "ItemNumber";
            ItemName.Text = "Item";
            ItemBrand.Text = "Brand";
            ItemCategory.Text = "Category";
            ItemSKU.Text = "SKU";    
            ItemSiteHumanName.Text = "Site Name";
            ItemScrapedURL.Text = "Site URL";
            ItemScrapedCurrency.Text = "Currency";
            ItemScrapedPrice.Text = "Price";
            ItemLastScraped.Text = "Last Scrape";
            ItemLastError.Text = "Last Error";
            ItemScraperActive.Text = "Active";

            productLVI.Columns.Add(LVIID);
            productLVI.Columns.Add(ItemNumber);
            productLVI.Columns.Add(ItemName);
            productLVI.Columns.Add(ItemBrand);
            productLVI.Columns.Add(ItemCategory);
            productLVI.Columns.Add(ItemSKU);
            productLVI.Columns.Add(ItemSiteHumanName);
            productLVI.Columns.Add(ItemScrapedURL);
            productLVI.Columns.Add(ItemScrapedCurrency);
            productLVI.Columns.Add(ItemScrapedPrice);
            productLVI.Columns.Add(ItemLastScraped);
            productLVI.Columns.Add(ItemLastError);
            productLVI.Columns.Add(ItemScraperActive);



        }


        private void PopulateLVI()
        {
            SetupLVI();
            ListViewItem tmpItm = new ListViewItem();
            ListViewItem.ListViewSubItem tmpSub = new ListViewItem.ListViewSubItem();
            long counter = 0;
            foreach (ItemData id in ItemsInformation)
            {
                foreach (ScrapeInfo si in id.ScrapeData)
                {
                    counter++;
                    //header stuff
                    tmpItm = new ListViewItem();
                    if (si.Amount <= 0)
                    {
                        tmpItm.BackColor = Color.DarkRed;
                        tmpItm.ForeColor = Color.White;
                    }
                    tmpItm.Text = counter.ToString();
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = id.InternalID;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = id.ProductName;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = id.ProductBrand;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = id.CategoryName;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = id.SKU;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();

                    //scrape stuff
                    tmpSub.Text = si.SiteHumanName;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.URL;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.Currency;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.Amount.ToString();
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.LastChecked.ToString();
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.LastError;
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpSub.Text = si.Active.ToString();
                    tmpItm.SubItems.Add(tmpSub);
                    tmpSub = new ListViewItem.ListViewSubItem();
                    tmpItm.Tag = id;


                    productLVI.Items.Add(tmpItm);
                    tmpItm = new ListViewItem();
                }
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Height = 720;
            button3.Enabled = false;
            button2.Enabled = true;
            button4.Enabled = true;
            
            button1.Visible = true;
            button9.Visible = true;
            button5.Visible = true;
            button8.Visible = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Height =435;
            button3.Enabled = true;
            button4.Enabled = true;
            button2.Enabled = false;

            button1.Visible = false;
            button9.Visible = false;
            button5.Visible = false;
            button8.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Height = 209;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = false;
            button1.Visible = false;
            button9.Visible = false;
            button5.Visible = false;
            button8.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string conversionRate = Connectivity.HTTPCalls.HTTP_GET("http://rate-exchange.appspot.com/currency?from=EUR&to=USD");
            textBox1.Text = conversionRate;
        }

        private void button5_Click_1(object sender, EventArgs e)
        {

            populateInsertProductRequest(TestServer);
           
            //to upload an item we can use:
            //1. productName - name of item (our description)
            //2. checkFrequencyType - DAYS or HOURS
            //3. checkFrequencyInterval - how many hrs or days before re-scraping..
            //4. Active - True or False
            //5. SKU - The TP itemnumber minus the first three characters
            //6. internalId - our TP ItemNumber
            //7. targetPrice - If we want to look for MAPP violations, this is the field to give value to.
            //8. purchasePrice - Price at which we received said item. However I would leave this blank.
            //9. categoryName - In our case, unless we want separate batches of items to be scraped, should have a single category for all items.
            //10. brandName - The brand of said item.
            //11. supplierName - The name of the people who sell us this item.

            //We will use #1,#4,#5,#6, and #10 from the above list...

            //now for some test code
            List<string> testItem = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber,ItemDescription,ProductLine FROM InventoryMaster WHERE ItemNumber='ADJ1415-C'");
            if (testItem.Count > 0)
            {
                bool Active = true; //#4
                string ItemNumber = testItem[0].Split('|')[0]; //#6
                string ItemDescription = testItem[0].Split('|')[1]; //#1
                string ProductLineCode = testItem[0].Split('|')[2];
                string ProductLine = ""; //#10
                string SKU = ItemNumber.Substring(3); //#5
                List<string> productLine = Connectivity.SQLCalls.sqlQuery("SELECT ManufFullName FROM ProductLine WHERE ProductLine.ProductLine='" + ProductLineCode + "'");
                if (productLine.Count > 0)
                {
                    ProductLine = productLine[0].Split('|')[0];
                }
                else
                {
                    MessageBox.Show("So the test is half way then didn't find the product line in ProductLine.ProductLine...");
                    return;
                }
                List<string> categoryID = Connectivity.SQLCalls.sqlQuery("SELECT EBayCategoryID FROM vWebOffloadEBay WHERE ItemNumber='" + ItemNumber + "'");
                if (categoryID.Count > 0)
                {
                    List<string> category = Connectivity.SQLCalls.sqlQuery("SELECT CategoryName FROM EBayCategory WHERE CategoryID=" + categoryID[0].Split('|')[0]);
                    if (category.Count > 0)
                    {
                        string productCategory = category[0].Split('|')[0];
                        productCategory = productCategory.Replace("&", "&amp;");
                        productCategory = productCategory.Replace(">", "&gt;");
                        
                        getInsertProductRequest = getInsertProductRequest.Replace("{CATEGORYNAME}", productCategory);
                        //getInsertProductRequest = getInsertProductRequest.Replace("{CATEGORYNAME}", "Tools");
                    }
                    else
                    {
                        getInsertProductRequest = getInsertProductRequest.Replace("{CATEGORYNAME}", "Tools");
                    }
                }
                else
                {
                    getInsertProductRequest = getInsertProductRequest.Replace("{CATEGORYNAME}", "Tools");
                }
                //all data is made, lets put it to use....

                getInsertProductRequest = getInsertProductRequest.Replace("{TITLE}", ItemDescription);
                getInsertProductRequest = getInsertProductRequest.Replace("{CHECKFREQUENCYTYPE}", "DAYS");
                getInsertProductRequest = getInsertProductRequest.Replace("{CHECKFREQUENCYINTERVAL}", "1");
                getInsertProductRequest = getInsertProductRequest.Replace("{ISACTIVE}", "True");
                getInsertProductRequest = getInsertProductRequest.Replace("{SKU}", SKU);
                getInsertProductRequest = getInsertProductRequest.Replace("{INTERNALID}", ItemNumber);
                getInsertProductRequest = getInsertProductRequest.Replace("{TARGETPRICE}", "");
                getInsertProductRequest = getInsertProductRequest.Replace("{PURCHASEPRICE}", "");
             
                getInsertProductRequest = getInsertProductRequest.Replace("{BRANDNAME}", ProductLine);
                getInsertProductRequest = getInsertProductRequest.Replace("{SUPPLIERNAME}", ProductLine);

                //MessageBox.Show(getInsertProductRequest);
                File.WriteAllText(@"c:\users\tomwestbrook\desktop\insertProductP2S_1.txt", getInsertProductRequest);
                string response = Connectivity.HTTPCalls.HTTP_POST(this.CurrentTestConnection.API_BaseURL, getInsertProductRequest);
                //string response = "test";
                
                
                if (response.Contains("<insertProductResponse") == true)             
                {
                    string itemID = response.Split(new string[] { "modify\">" }, StringSplitOptions.None)[1].Split(new string[] { "</insertProductResponse>" }, StringSplitOptions.None)[0];
                    long tmplng = 0;
                    bool testID = long.TryParse(itemID,out tmplng);
                    if (testID == true)
                    {
                        //use id to upload the default urls to this item...
                        populateInsertURLRequest(TestServer);
                        getInsertURLRequest = getInsertURLRequest.Replace("{PRODUCTID}", itemID);
                        string url = GetToolsPlusURLFromItemNumber(ItemNumber);
                        if (url.Contains("Not Found") == false)
                        {
                            getInsertURLRequest = getInsertURLRequest.Replace("{URLTOADD}", url);
                        }
                        else 
                        {
                            statusLst.Items.Add("Couldn't find URL for " + ItemNumber);
                            statusLst.SelectedIndex = statusLst.Items.Count - 1;
                            statusLst.SelectedIndex = -1;
                            statusLst.Refresh();
                        }
                        getInsertURLRequest = getInsertURLRequest.Replace("{ISACTIVE}", "True");
                        getInsertURLRequest = getInsertURLRequest.Replace("{SHOWINREPORTS}", "True");
                        //MessageBox.Show(getInsertURLRequest);
                        File.WriteAllText(@"c:\users\tomwestbrook\desktop\insertItemP2S_2.txt", getInsertURLRequest);
                        string urlResponse = Connectivity.HTTPCalls.HTTP_POST(this.CurrentTestConnection.API_BaseURL, getInsertURLRequest);
                        if (urlResponse.Contains("<insertUrlResponse") == true)
                        {
                            FormatXMLForViewing(urlResponse);
                            statusLst.Items.Add(ItemNumber + " added to Price2Spy.");
                            statusLst.SelectedIndex = statusLst.Items.Count - 1;
                            statusLst.SelectedIndex = -1;
                            statusLst.Refresh();
                        }

                        //{ISACTIVE}
                        //{SHOWINREPORTS}
                    }
                    else
                    {
                        //error: invalid id
                    }
                    
                }
                else
                {
                    //error: bad request
                    textBox1.Text = response;
                    if (response.Split('(')[1].Split(')')[0] == "500")
                    {
                        statusLst.Items.Add(ItemNumber + " already exists on server or connection issue... remove n' add..");
                        statusLst.SelectedIndex = statusLst.Items.Count - 1;
                        statusLst.SelectedIndex = -1;
                        statusLst.Refresh();
                        string productId = GetPriceTwoSpyIDNumber(ItemNumber);                      
                        if (productId != "N/A")
                        {
                            statusLst.Items.Add("Found " + ItemNumber + " on Price2Spy. ProductID: " + productId);
                            statusLst.SelectedIndex = statusLst.Items.Count - 1;
                            statusLst.SelectedIndex = -1;
                            statusLst.Refresh();
                            bool isDeleted = DeleteItemFromPrice2Spy(productId);
                            if (isDeleted == true)
                            {
                               
                                statusLst.Items.Add("Re-running the Insert Product Process for " + ItemNumber + "(" + productId + ")...");
                                statusLst.SelectedIndex = statusLst.Items.Count - 1;
                                statusLst.SelectedIndex = -1;
                                statusLst.Refresh();
                                button5_Click_1(sender, e);
                            }
                            else
                            {
                                statusLst.Items.Add("Couldn't remove " + ItemNumber + " (" + productId + ") from Price2Spy.");
                                statusLst.SelectedIndex = statusLst.Items.Count - 1;
                                statusLst.SelectedIndex = -1;
                                statusLst.Refresh();
                            }
                        }
                        else
                        {
                            //product didnt exist?
                            statusLst.Items.Add("Could not find item " + ItemNumber + " on Price2Spy.");
                            statusLst.SelectedIndex = statusLst.Items.Count - 1;
                            statusLst.SelectedIndex = -1;
                            statusLst.Refresh();
                        }

                    }
                    else
                    {
                        statusLst.Items.Add("Server: " + response.Split('(')[1].Split(')')[0] + " error on " + ItemNumber);
                        statusLst.SelectedIndex = statusLst.Items.Count - 1;
                        statusLst.SelectedIndex = -1;
                        statusLst.Refresh();
                    }
                }

            }
            else
            {
                MessageBox.Show("This test did not go as planned. Hell it didn't even run!");
            }


        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        public string GetToolsPlusURLFromItemNumber(string ItemNumber)
        {
            List<string> fullItemName = Connectivity.SQLCalls.sqlQuery("SELECT ManufFullNameWeb FROM ProductLine WHERE ProductLine='" + ItemNumber.Substring(0,3) + "'");
            if (fullItemName.Count > 0)
            {
                string url = "http://www.tools-plus.com/" + fullItemName[0].Split('|')[0].Replace(" ", "-") + "-" + ItemNumber.Substring(3) + ".html";

                return url.ToLower();
            }
            else
            {
                return "Not Found.";
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://office.oaza.rs:2280/price2spy");

        }
        public string GetPriceTwoSpyIDNumber(string ItemNumber)
        {
            populateFindItemByItemNumberRequest(TestServer);
            findItemByItemNumber = findItemByItemNumber.Replace("{ITEMNUMBER}", ItemNumber);
            string results = Connectivity.HTTPCalls.HTTP_POST(this.CurrentTestConnection.API_BaseURL, findItemByItemNumber);

            textBox1.Text = results;
            //MessageBox.Show(results);
            if (results.Contains("<ns3:productId>") == true)
            {
                return results.Split(new string[] { "<ns3:productId>" }, StringSplitOptions.None)[1].Split(new string[] { "</ns3:productId>" }, StringSplitOptions.None)[0].Trim();
            }
            return "";

        }
        public bool DeleteItemFromPrice2Spy(string Price2SpyIDNumber)
        {
            string JSON_OUT = "";
            JSON_OUT += "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:mod=\"http://www.price2spy.com/api/modify\">\r\n";
            JSON_OUT += " <soapenv:Header/>\r\n";
            JSON_OUT += " <soapenv:Body>\r\n";
            JSON_OUT += "  <mod:deleteProductRequest>\r\n";
            JSON_OUT += "   <mod:productId>" + Price2SpyIDNumber + "</mod:productId>\r\n";
            JSON_OUT += "   <mod:apiKey>" + this.CurrentTestConnection.API_CurrentKey + "</mod:apiKey>\r\n";
            JSON_OUT += "  </mod:deleteProductRequest>\r\n";
            JSON_OUT += " </soapenv:Body>\r\n";
            JSON_OUT += "</soapenv:Envelope>\r\n";
            string results = Connectivity.HTTPCalls.HTTP_POST(this.CurrentTestConnection.API_BaseURL, JSON_OUT);
            //if (results.Trim() == "<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/\" ><soap:Body/></soap:Envelope>")
            if (results.Trim() == "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"><soap:Body/></soap:Envelope>")
            {
                //alright, now attempt the process again...
                return true;
            }
            else
            {
                //error - perhaps item wasnt there?
                MessageBox.Show(results);
                return false;
            }
            return false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(CurrentTestConnection.API_CurrentKey);

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(CurrentTestConnection.API_BaseURL);
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(CurrentProductionConnection.API_CurrentKey);
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Windows.Forms.Clipboard.SetText(CurrentProductionConnection.API_BaseURL);
        }

        private void secondTmr_Tick(object sender, EventArgs e)
        {
            timeLbl.Text = DateTime.Now.ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            PullDataFromServer(CurrentProductionConnection.API_BaseURL, CurrentProductionConnection.API_CurrentKey);
            int errorCount = 0;
            int resCount = 0;
            string emailMessage = "<html><body><table style=\"Width:90%\" border=\"1\" rules=\"rows\"><tr><td><b>Product ID</b></td><td><b>URL</b></td><td><b>Price</b></td><td><b>Last Checked</b></td><td><b>Last Error Description</b></td></tr>";
           
          
           //Email.Email.SendEmail(emailMessage, true);
            foreach (ListViewItem lvi in productLVI.Items)
            {
                resCount++;
                if (lvi.BackColor == Color.DarkRed)
                {
                    errorCount++;
                    //UNSCANNABLEEEE!!!                    
                    //emailMessage += "<tr><td>test item</td><td><a href=\"http://www.google.com\">Site Link</a></td><td>0.00</td><td>11-25-2014 9:19:00 AM</td><td>N/A</td></tr>";
                    emailMessage += "<tr><td>" + lvi.SubItems[1].Text + "</td><td><a href=\"" + lvi.SubItems[7].Text + "\">"+ lvi.SubItems[6].Text  +"</a></td><td>" + lvi.SubItems[9].Text + "</td><td>" + lvi.SubItems[10].Text + "</td><td>" + lvi.SubItems[11].Text + "</td></tr>\r\n";
                }
            }
            emailMessage += "</table><br>There are " + errorCount.ToString() + " errors from Price2Spy. Total scrapes was " + resCount.ToString() + "</body></html>";
            //now inform me and eventually others there are items not being scraped...

           
            Email.Email.SendEmail(emailMessage, false);
            statusLst.Items.Add(errorCount.ToString() + "/" + resCount.ToString() + " results have errors that should be fixed.");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
            statusLst.Items.Add("Completed Processing Downloaded Items.");
            statusLst.SelectedIndex = statusLst.Items.Count - 1;
            statusLst.SelectedIndex = -1;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            populateAutomatchRequest();
            automatchRequest = automatchRequest.Replace("{PRODUCTID}", "D-HDWHT47048");
            //MessageBox.Show(automatchRequest);
            string result = Connectivity.HTTPCalls.HTTP_POST(CurrentProductionConnection.API_BaseURL, automatchRequest);
            textBox1.Text = result;
            //MessageBox.Show("COMPLETE");

        }

        private void button9_Click(object sender, EventArgs e)
        {
            ItemData testItem = (ItemData)(productLVI.Items[1].Tag);
            populateAutomatchRequest();
            automatchRequest = automatchRequest.Replace("{PRODUCTID}", testItem.InternalID);
            System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\damnp2s.txt", automatchRequest);
            string results = Connectivity.HTTPCalls.HTTP_POST(CurrentProductionConnection.API_BaseURL, automatchRequest);
            textBox1.Text = results;
            MessageBox.Show("COMPLETE");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string emailMessage = "<html><body><table style=\"Width:90%\" border=\"1\" rules=\"rows\"><tr><td><b>Product ID</b></td><td><b>URL</b></td><td><b>Price</b></td><td><b>Last Checked</b></td><td><b>Last Error Description</b></td></tr>";
            emailMessage += "<tr><td>test item</td><td><a href=\"http://www.google.com\">Site Link</a></td><td>0.00</td><td>11-25-2014 9:19:00 AM</td><td>N/A</td></tr>";
            emailMessage += "</table></body></html>";
            Email.Email.SendEmail(emailMessage, true);
            

        }

        private void statusLst_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timeLbl_Click(object sender, EventArgs e)
        {

        }



    }
}
