using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WiserProductFeedV2
{
    public partial class Form1 : Form
    {
        public struct WiserCreds
        {
            public string WiserDomainURL;
            public string WiserDomainPort;
            public string WiserDomainUsername;
            public string WiserDomainPassword;
            public void PopulateCredentials()
            {
                List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT APIKey,APIName,APIUrl FROM APIKeys WHERE APIUrl LIKE '%wiser.com%'");
                if (results.Count > 0)
                {
                    WiserDomainPassword = results[0].Split('|')[0];
                    WiserDomainUsername = results[0].Split('|')[1];
                    WiserDomainURL = "ftp:" + results[0].Split('|')[2].Split(':')[1];
                    WiserDomainPort = results[0].Split('|')[2].Split(':')[2];
                    WiserDomainPort = WiserDomainPort.Replace("/","");
                }
            }
        }
        public WiserCreds ServerCredentials = new WiserCreds();




  

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Shown += new EventHandler(Form1_Shown);
            ServerCredentials.PopulateCredentials();

        }

        void Form1_Shown(object sender, EventArgs e)
        {            
            this.Refresh();
            statusLbl.BackColor = Color.FromArgb(0, 113, 189);
            statusLbl.ForeColor = Color.White;
            ServerCredentials.PopulateCredentials();
            //RunDownloadAndInsertProcess();
            RunFullProcess();
           
            Application.Exit();
        }

        //public const string FilePath_CompetitorPrices = @"c:\users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv";
        //public const string FilePath_UploadFile = @"c:\users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv";
        //public const string FilePath_FormattedFeedLocalized = @"c:\inetpub\wwwroot\formattedPrices.csv";
        //public const string FilePath_FormattedFeedRemote = @"\\ToolsPlus06\c$\inetpub\wwwroot\formattedPrices.csv";

        public const string FilePath_CompetitorPricesGoogle = @"c:\WiserTemp\CompetitorPricesGoogle.csv"; //local to computer
        public const string FilePath_CompetitorPricesEBay = @"c:\WiserTemp\CompetitorPricesEBay.csv";
        public const string FilePath_CompetitorPricesOther1 = @"c:\WiserTemp\CompetitorPricesOther1.csv"; //local to computer
        public const string FilePath_CompetitorPricesOther2 = @"c:\WiserTemp\CompetitorPricesOther2.csv";
        public const string FilePath_CompetitorPricesOther3 = @"c:\WiserTemp\CompetitorPricesOther3.csv"; //local to computer

        public const string FilePath_UploadFile = @"c:\WiserTemp\WiserFeed.csv"; //local to computer
        public const string FilePath_FormattedFeedLocalized = @"c:\WiserTemp\formattedPrices.csv"; //local to sqlsvr.
        public const string FilePath_FormattedFeedRemote = @"c:\WiserTemp\formattedPrices.csv"; //network path of file (if run remotely).


        public const bool missingASINChk = false;
        public struct WiserFeedData
        {
            public string UPC;
            public string ASIN;
            public string Brand;
            public string Model;
            public string MPN;
            public string InventoryNumber;
            public string ProductName;
            public decimal ProductPrice;
            public decimal MinimumPrice;
            public decimal Cost;
            public decimal ShippingPrice;
            public int InStock;
            public string ProductURL;
            public string ImageURL;
            public string Label;
            public string StrSearchKeywords;

        }

        private void CreateUploadFile()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Bold);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();
            int repoAdditionCount = 0;
            setupLVI();
            feedLVI.Visible = false;
            feedLVI.Refresh();
            //UPC, ASIN, Brand, Model
            //MPN, Inventory Number, Product Name, Product Price
            //Minimum Price, Cost, Shipping Price, In Stock
            //Product URL, Image URL, Label, str_search_keywords



            //Original (No to-be-publisheds): List<string> exampleProducts = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT InventoryMaster.ItemID,InventoryMaster.ItemNumber,ProductLine,IDiscountMarkupPriceRate1,StdCost,ShippingType FROM InventoryMaster,PartNumbers,InventoryQuantities WHERE IsMASKit=0 AND WebPublished=1 AND WebOrderable=1 AND InventoryQuantities.ItemNumber=InventoryMaster.ItemNumber AND PartNumbers.ItemNumber = InventoryMaster.ItemNumber  AND InventoryMaster.ItemNumber NOT IN (SELECT ItemNumberPointer FROM PartNumberGroupMaster)");
            List<string> exampleProducts = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT InventoryMaster.ItemID,InventoryMaster.ItemNumber,ProductLine,IDiscountMarkupPriceRate1,StdCost,ShippingType FROM InventoryMaster,PartNumbers,InventoryQuantities WHERE IsMASKit=0  AND (WebPublished=1 OR ((PartNumbers.EBayToBePublished=1 OR PartNumbers.WebToBePublished=1)) AND PartNumbers.TBPubPriority > 1) AND WebOrderable=1 AND InventoryQuantities.ItemNumber=InventoryMaster.ItemNumber AND PartNumbers.ItemNumber = InventoryMaster.ItemNumber AND InventoryMaster.ItemNumber NOT IN (SELECT ItemNumberPointer FROM PartNumberGroupMaster)");
            if (exampleProducts.Count > 0)
            {
                int counter = 1;
                foreach (string itm in exampleProducts)
                {
                    if (itm.Split('|')[1] != "ADJCLAMP-C")
                    {
                        statusLbl.Text = "Status: Parsing " + itm.Split('|')[1] + " (" + counter.ToString() + "/" + exampleProducts.Count.ToString() + ")";
                        statusLbl.Refresh();
                        counter++;
                        string UPC = "", ASIN = "", Brand = "", Model = "", MPN = "", InventoryNumber = "", ProductName = "",
                            ProductPrice = "", MinimumPrice = "", Cost = "", ShippingPrice = "", InStock = "",
                            ProductURL = "", ImageURL = "", sLabel = "", StrSearchKeywords = "";

                        List<string> componentMap = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + itm.Split('|')[1] + "'");
                        string componentID = "";
                        if (componentMap.Count > 0)
                        {
                            componentID = componentMap[0].Split('|')[0];
                        }
                        List<string> upcList = Connectivity.SQLCalls.sqlQuery("SELECT Barcode FROM InventoryComponentBarcodes WHERE ComponentID=" + componentID);
                        if (upcList.Count > 0)
                        {
                            if (upcList[0].Split('|')[0].StartsWith("TP"))
                            {
                                UPC = "";
                            }
                            else
                            {
                                UPC = upcList[0].Split('|')[0];
                                if (UPC.Length > 12)
                                {
                                    UPC = "";
                                }
                                while (UPC.Length < 12 & UPC != "")
                                {
                                    UPC = "0" + UPC;
                                }

                            }
                        }
                        InventoryNumber = itm.Split('|')[1];
                        Model = itm.Split('|')[1];

                        List<string> tryASIN = Connectivity.SQLCalls.sqlQuery("SELECT AmazonASIN FROM UPCtoASIN WHERE ItemNumber='" + InventoryNumber + "'");
                        if (tryASIN.Count > 0)
                        {
                            //pull from db...
                            ASIN = tryASIN[0].Split('|')[0];
                            if (ASIN == "UPCNOTFOUND")
                            {
                                ASIN = "";
                            }
                        }
                        else
                        {
                            if (UPC.StartsWith("TP") == false & missingASINChk == true)
                            {
                                ASIN = Connectivity.HTTPCalls.HTTP_GET("http://upctoasin.com/" + UPC);
                                if (ASIN.Contains("<") == true)
                                {
                                    ASIN = "";
                                }
                                else
                                {
                                    //add to db...
                                    Connectivity.SQLCalls.sqlQuery("INSERT INTO UPCtoASIN (ItemNumber,UPC,AmazonASIN) VALUES('" + InventoryNumber + "','" + UPC + "','" + ASIN + "')");
                                    repoAdditionCount++;
                                }
                            }
                            else
                            {
                                ASIN = "";
                            }
                        }



                        List<string> brandList = Connectivity.SQLCalls.sqlQuery("SELECT ManufFullNameWeb FROM ProductLine WHERE ProductLine='" + itm.Split('|')[2] + "'");
                        if (brandList.Count > 0)
                        {
                            Brand = brandList[0].Split('|')[0];
                        }
                        List<string> descriptions = Connectivity.SQLCalls.sqlQuery("SELECT Desc2,Desc1,Desc3 FROM PartNumbers WHERE ItemNumber='" + InventoryNumber + "'");
                        if (descriptions.Count > 0)
                        {
                            ProductName = Brand + " " + descriptions[0].Split('|')[0] + " " + descriptions[0].Split('|')[1] + " " + descriptions[0].Split('|')[2];
                            MPN = descriptions[0].Split('|')[0];
                        }

                        List<string> QOH = Connectivity.SQLCalls.sqlQuery("SELECT QuantityOnHand FROM InventoryQuantities WHERE ItemNumber='" + InventoryNumber + "'");
                        if (QOH.Count > 0)
                        {
                            InStock = QOH[0].Split('|')[0];
                        }


                        ProductURL = ("http://www.tools-plus.com/" + Brand.Replace(" ", "-") + "-" + MPN + ".html").ToLower();
                        ImageURL = ("http://p.hostingprod.com/%40tools-plus.com/itemimage/noauth-nobg/" + Brand.Replace(" ", "-") + "-" + MPN + ".jpg").ToLower();
                        ProductPrice = itm.Split('|')[3];
                        //Cost = itm.Split('|')[4];
                        Cost = ProductPrice;
                        string ShippingType = itm.Split('|')[5];

                        if (ShippingType != "3")
                        {
                            if (decimal.Parse(ProductPrice) > 50)
                            {
                                ShippingPrice = "0.00";
                            }
                            else
                            {
                                ShippingPrice = "8.50";
                            }
                        }
                        else
                        {

                            ShippingPrice = "0.00";

                        }
                        MinimumPrice = itm.Split('|')[3];
                        sLabel = "";
                        StrSearchKeywords = "";
                        //so far we have UPC, InventoryNumber, ASIN, Brand, Model,ProductName,MPN,InStock,ProductURL,ImageURL,ProdutcPrice,Cost,MinimumPrice, ShippingPrice
                        //we still need  Label, StrSearchKeywords, and to possibly re-adjust shipping later on....

                        WiserFeedData wfd = new WiserFeedData();
                        wfd.UPC = UPC;
                        wfd.ASIN = ASIN;
                        wfd.Brand = Brand;
                        wfd.Model = Model;
                        wfd.MPN = MPN;
                        wfd.InventoryNumber = InventoryNumber;
                        wfd.ProductName = ProductName;
                        wfd.ProductPrice = decimal.Parse(ProductPrice);
                        wfd.MinimumPrice = decimal.Parse(MinimumPrice);
                        wfd.Cost = decimal.Parse(Cost);
                        wfd.ShippingPrice = decimal.Parse(ShippingPrice);
                        wfd.InStock = int.Parse(InStock);
                        wfd.ProductURL = ProductURL;
                        wfd.ImageURL = ImageURL;
                        wfd.Label = sLabel;
                        wfd.StrSearchKeywords = StrSearchKeywords;
                        AddItemToLVI(wfd);
                        feedLVI.Items[feedLVI.Items.Count - 1].EnsureVisible();
                        feedLVI.Refresh();

                    }
                }

            }
            else
            {
                MessageBox.Show("The built in example query failed to produce a few items. Stopping...");
                return;
            }

            statusLbl.Text = "Status: Completed. Added " + repoAdditionCount.ToString() + " new items to our ASIN repo.";

            WriteLVItoCSV();
            feedLVI.Visible = true;
            feedLVI.Refresh();
        }
        private void AddItemToLVI(WiserFeedData wfd)
        {
            ListViewItem newLVI = new ListViewItem();
            newLVI.Text = wfd.UPC;
            newLVI.SubItems.Add(wfd.ASIN);
            newLVI.SubItems.Add(wfd.Brand);
            newLVI.SubItems.Add(wfd.Model);
            newLVI.SubItems.Add(wfd.MPN);
            newLVI.SubItems.Add(wfd.InventoryNumber);
            newLVI.SubItems.Add(wfd.ProductName);
            newLVI.SubItems.Add(string.Format("{0:C}", wfd.ProductPrice));
            newLVI.SubItems.Add(string.Format("{0:C}", wfd.MinimumPrice));
            newLVI.SubItems.Add(string.Format("{0:C}", wfd.Cost));
            newLVI.SubItems.Add(string.Format("{0:C}", wfd.ShippingPrice));
            newLVI.SubItems.Add(wfd.InStock.ToString());
            newLVI.SubItems.Add(wfd.ProductURL);
            newLVI.SubItems.Add(wfd.ImageURL);
            newLVI.SubItems.Add(wfd.Label);
            newLVI.SubItems.Add(wfd.StrSearchKeywords);
            feedLVI.Items.Add(newLVI);
        }
        private void setupLVI()
        {
            ColumnHeader upcCol = new ColumnHeader();
            ColumnHeader asinCol = new ColumnHeader();
            ColumnHeader brandCol = new ColumnHeader();
            ColumnHeader modelCol = new ColumnHeader();

            ColumnHeader mpnCol = new ColumnHeader();
            ColumnHeader inventorynumberCol = new ColumnHeader();
            ColumnHeader productnameCol = new ColumnHeader();
            ColumnHeader productpriceCol = new ColumnHeader();

            ColumnHeader minimumpriceCol = new ColumnHeader();
            ColumnHeader costCol = new ColumnHeader();
            ColumnHeader shippingpriceCol = new ColumnHeader();
            ColumnHeader instockCol = new ColumnHeader();

            ColumnHeader producturlCol = new ColumnHeader();
            ColumnHeader imageurlCol = new ColumnHeader();
            ColumnHeader labelCol = new ColumnHeader();
            ColumnHeader str_search_keywordsCol = new ColumnHeader();

            upcCol.Width = feedLVI.Width / 16;
            asinCol.Width = feedLVI.Width / 16;
            brandCol.Width = feedLVI.Width / 16;
            modelCol.Width = feedLVI.Width / 16;

            mpnCol.Width = feedLVI.Width / 16;
            inventorynumberCol.Width = feedLVI.Width / 16;
            productnameCol.Width = feedLVI.Width / 16;
            productpriceCol.Width = feedLVI.Width / 16;

            minimumpriceCol.Width = feedLVI.Width / 16;
            costCol.Width = feedLVI.Width / 16;
            shippingpriceCol.Width = feedLVI.Width / 16;
            instockCol.Width = feedLVI.Width / 16;

            producturlCol.Width = feedLVI.Width / 16;
            imageurlCol.Width = feedLVI.Width / 16;
            labelCol.Width = feedLVI.Width / 16;
            str_search_keywordsCol.Width = feedLVI.Width / 16;

            upcCol.Text = "UPC";
            asinCol.Text = "ASIN";
            brandCol.Text = "Brand";
            modelCol.Text = "Model";

            mpnCol.Text = "MPN";
            inventorynumberCol.Text = "Inventory Number";
            productnameCol.Text = "Product Name";
            productpriceCol.Text = "Product Price";

            minimumpriceCol.Text = "Minimum Price";
            costCol.Text = "Cost";
            shippingpriceCol.Text = "Shipping Price";
            instockCol.Text = "In Stock";

            producturlCol.Text = "Product URL";
            imageurlCol.Text = "Image URL";
            labelCol.Text = "Label";
            str_search_keywordsCol.Text = "str_search_keywords";

            feedLVI.Columns.Clear();
            feedLVI.Items.Clear();
            feedLVI.Clear();

            feedLVI.Columns.Add(upcCol);
            feedLVI.Columns.Add(asinCol);
            feedLVI.Columns.Add(brandCol);
            feedLVI.Columns.Add(modelCol);

            feedLVI.Columns.Add(mpnCol);
            feedLVI.Columns.Add(inventorynumberCol);
            feedLVI.Columns.Add(productnameCol);
            feedLVI.Columns.Add(productpriceCol);

            feedLVI.Columns.Add(minimumpriceCol);
            feedLVI.Columns.Add(costCol);
            feedLVI.Columns.Add(shippingpriceCol);
            feedLVI.Columns.Add(instockCol);

            feedLVI.Columns.Add(producturlCol);
            feedLVI.Columns.Add(imageurlCol);
            feedLVI.Columns.Add(labelCol);
            feedLVI.Columns.Add(str_search_keywordsCol);


        }
        private void WriteLVItoCSV()
        {
            //string tmpCSVPath = @"C:\users\tomwestbrook\documents\wiserfeed\WiserFeed.csv";
            using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(FilePath_UploadFile))
            {



                //string tmpCSV = "";
                //tmpCSV += "UPC/EAN,ASIN,Brand,Model,MPN,Inventory Number,Product Name,Product Price,Minimum Price,Cost,Shipping Price,In Stock,Product URL, Image URL, Label, str_search_keywords,\r\n";
                outfile.Write("UPC/EAN,ASIN,Brand,Model,MPN,Inventory Number,Product Name,Product Price,Minimum Price,Cost,Shipping Price,In Stock,Product URL, Image URL, Label, str_search_keywords,\r\n");

                int counter = 0;
                foreach (ListViewItem lvi in feedLVI.Items)
                {
                    counter++;
                    statusLbl.Text = "Status: Preparing to write to file #" + counter.ToString() + "/" + feedLVI.Items.Count.ToString();
                    statusLbl.Refresh();
                    foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                    {
                        //tmpCSV += "\"" + lvsi.Text.Replace("$", "").Replace("\"", "\"\"") + "\",";
                        outfile.Write("\"" + lvsi.Text.Replace("$", "").Replace("\"", "\"\"") + "\",");
                    }
                    //tmpCSV += "\r\n";
                    outfile.Write("\r\n");
                }
                outfile.Close();
            }

            /*     try
                 {
                     if (System.IO.File.Exists(@"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv") == true)
                     {
                         System.IO.File.Copy(@"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv", @"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed - " + DateTime.Now.ToShortDateString().Replace("/", "-") + ".csv");
                         System.IO.File.Delete(@"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv");
                     }
                 }
                 catch { }

                 System.IO.File.WriteAllText(@"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv", tmpCSV);*/
            statusLbl.Text = "Status: Finished writing file.";
            statusLbl.Refresh();
        }
        private void UploadToWiser()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();



            statusLbl.Text = "Status: Transferring Data...";
            statusLbl.Refresh();

            //MessageBox.Show(ServerCredentials.WiserDomainUsername);
            //MessageBox.Show(ServerCredentials.WiserDomainPassword);
            //MessageBox.Show(ServerCredentials.WiserDomainURL);
            //MessageBox.Show(ServerCredentials.WiserDomainPort);

            //string result = Connectivity.FTPCalls.FTP_UPLOAD("ftp://ftp.wiser.com", "/upload/WP_Inv.csv", FilePath_UploadFile, "User4079", "U5sAPZfkCm7v", 2221);
            string result = Connectivity.FTPCalls.FTP_UPLOAD(ServerCredentials.WiserDomainURL,"/upload/WP_Inv.csv",FilePath_UploadFile,ServerCredentials.WiserDomainUsername,ServerCredentials.WiserDomainPassword,int.Parse(ServerCredentials.WiserDomainPort));
            if (result.ToUpper().Contains("TRANSFER COMPLETE") == true)
            {
                statusLbl.Text = "Status: Transfer Complete!";
                statusLbl.Refresh();
            }
        }


        private void DownloadFromWiser()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();
            statusLbl.Text = "Status: Downloading Feed...0%";
            statusLbl.Refresh();
            string ReturnedFeedGoogle = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1256108.csv", ServerCredentials.WiserDomainUsername, ServerCredentials.WiserDomainPassword, int.Parse(ServerCredentials.WiserDomainPort));
            System.IO.File.WriteAllText(FilePath_CompetitorPricesGoogle, ReturnedFeedGoogle);
            statusLbl.Text = "Status: Downloading Feed... 50%";
            statusLbl.Refresh();
            string ReturnedFeedEBay = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1161993133.csv", "User4079", "U5sAPZfkCm7v", 2221);
            System.IO.File.WriteAllText(FilePath_CompetitorPricesEBay, ReturnedFeedEBay);
            statusLbl.Text = "Status: Downloading Feed...60%.";
            statusLbl.Refresh();
            string ReturnedFeedOther1 = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1162037709.csv", "User4079", "U5sAPZfkCm7v", 2221);
            System.IO.File.WriteAllText(FilePath_CompetitorPricesOther1, ReturnedFeedOther1);
            statusLbl.Text = "Status: Downloading Feed...80%.";
            statusLbl.Refresh();
            string ReturnedFeedOther2 = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1162064327.csv", "User4079", "U5sAPZfkCm7v", 2221);
            System.IO.File.WriteAllText(FilePath_CompetitorPricesOther2, ReturnedFeedOther2);
            statusLbl.Text = "Status: Downloading Feed...90%.";
            statusLbl.Refresh();
            string ReturnedFeedOther3 = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1162064336.csv", "User4079", "U5sAPZfkCm7v", 2221);
            System.IO.File.WriteAllText(FilePath_CompetitorPricesOther3, ReturnedFeedOther3);
            statusLbl.Text = "Status: Downloading Feed...Complete.";
            statusLbl.Refresh();
        }
        private void FormatWiserFeedForBulkInsertion()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();
            //DELETE FROM CSVTest
            //BULK
            //INSERT CSVTest
            //FROM 'c:\inetpub\wwwroot\competitorprices.csv'
            //WITH
            //(
            //    FIELDTERMINATOR = ',',
            //    ROWTERMINATOR = '\n'
            //)

            List<string> formattedLines = new List<string>();

            statusLbl.Text = "Status: Loading Feed...Please Wait";
            statusLbl.Refresh();
            //string[] dataFeed = System.IO.File.ReadAllLines(@"C:\Users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv");
            string[] dataFeedGoogle = System.IO.File.ReadAllLines(FilePath_CompetitorPricesGoogle);
            string[] dataFeedEBay = System.IO.File.ReadAllLines(FilePath_CompetitorPricesEBay);
            string[] dataFeedOther1 = System.IO.File.ReadAllLines(FilePath_CompetitorPricesOther1);
            string[] dataFeedOther2 = System.IO.File.ReadAllLines(FilePath_CompetitorPricesOther2);
            string[] dataFeedOther3 = System.IO.File.ReadAllLines(FilePath_CompetitorPricesOther3);
            string[] dataFeed = { };
            int counter = 0;
            int linecount = 1;
            for (int x = 0; x <= 4; x++)
            {
                counter = 0;
                if (x == 0)
                {
                    dataFeed = dataFeedGoogle;
                }
                //else
                //{
                //    dataFeed = dataFeedEBay;
               // }
                if (x == 1)
                {
                    dataFeed = dataFeedEBay;
                }
                if (x == 2)
                {
                    dataFeed = dataFeedOther1;
                }
                if (x == 3)
                {
                    dataFeed = dataFeedOther2;
                }
                if (x == 4)
                {
                    dataFeed = dataFeedOther3;
                }

                foreach (string line in dataFeed)
                {
                    counter++;
                    if (x == 0)
                    {
                        statusLbl.Text = "Status: Processing Google line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                        statusLbl.Refresh();
                    }
                    if (x==1)
                    {
                        statusLbl.Text = "Status: Processing EBay line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                        statusLbl.Refresh();
                    }
                    if (x == 2)
                    {
                        statusLbl.Text = "Status: Processing Other1 line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                        statusLbl.Refresh();
                    }
                    if (x == 3)
                    {
                        statusLbl.Text = "Status: Processing Other2 line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                        statusLbl.Refresh();
                    }
                    if (x == 4)
                    {
                        statusLbl.Text = "Status: Processing Other3 line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                        statusLbl.Refresh();
                    }
                    string newCSVLine = "";
                    string[] csvTokens = System.Text.RegularExpressions.Regex.Split(line.Substring(0, line.Length - 1), @"""?\s*,\s*""?");
                    int tokencount = 0;
                    foreach (string token in csvTokens)
                    {
                        tokencount++;
                        if (tokencount == 1)
                        {   //this just checks for the header... if so, skip this line by breaking out from the inside of the foreach loop
                            if (token.ToUpper().Contains("PRODUCT NAME") == true | token.ToUpper().Contains("NVENTORY") == true)
                            {
                                break;
                            }
                        }
                        if (tokencount == 2)
                        {
                            if (token == "")
                            {
                                break;
                            }
                            //first get line number. this token references the Item Number. It then references the current time (with correct format). this inserts three fields.
                            newCSVLine += linecount.ToString() + ",";
                            newCSVLine += token.Replace(",", " ") + ",";
                            newCSVLine += DateTime.Now.ToString("yyyy-M-d HH:mm:ss") + ".000,";
                        }
                        if (tokencount == 3)
                        {
                            if (token == "")
                            {
                                break;
                            }
                            //this references the Source data column.
                            if (x == 0)
                            {
                                newCSVLine += "Wiser Google,";
                            }
                            if (x==1)
                            {
                                newCSVLine += "Wiser EBay,";
                            }
                            if (x == 2)
                            {
                                newCSVLine += "Wiser Other Feed 1,";
                            }
                            if (x == 3)
                            {
                                newCSVLine += "Wiser Other Feed 2,";
                            }
                            if (x == 4)
                            {
                                newCSVLine += "Wiser Other Feed 3,";
                            }
                        }
                        if (tokencount == 19)
                        {
                            if (token == "")
                            {
                                break;
                            }
                            //this references the StoreName (aka the competitor). We dont use URL so adding a space to a second field here.
                            newCSVLine += token.Replace(",", " ") + ",";
                            newCSVLine += " ,";
                        }
                        if (tokencount == 20)
                        {
                            try { decimal.Parse(token); }
                            catch { break; }

                            if (token == "")
                            {
                                break;
                            }
                            //this references the competitor price
                            newCSVLine += token.Replace(",", " ") + ",";
                        }
                        if (tokencount == 21)
                        {
                            try { decimal.Parse(token); }
                            catch { break; }
                            if (token == "")
                            {
                                break;
                            }
                            //this references the competitor shipping. Then add it to the main list.
                            newCSVLine += token.Replace(",", " ");
                            formattedLines.Add(newCSVLine);
                            linecount++;
                        }




                    }

                }
            }
            statusLbl.Text = "Status: Preparing to write formatted file...";
            statusLbl.Refresh();

            //System.IO.File.WriteAllLines(@"C:\users\tomwestbrook\Documents\WiserFeed\formattedPrices.csv", formattedLines);
            System.IO.File.WriteAllLines(FilePath_FormattedFeedRemote, formattedLines);
            statusLbl.Text = "Status: Finished creating formatted file.";
            statusLbl.Refresh();

        }
        private void BulkInsertWiserFeed()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();
            statusLbl.Text = "Status: Bulk Inserting Items to SQL... Please Wait.";
            statusLbl.Refresh();
            Connectivity.SQLCalls.sqlQuery("DELETE FROM PriceComparisons2");
            Connectivity.SQLCalls.sqlQuery("BULK INSERT PriceComparisons2 FROM '" + FilePath_FormattedFeedLocalized + "' WITH (FIELDTERMINATOR = ',', ROWTERMINATOR = '\n')");
            statusLbl.Text = "Status: Bulk Insertion Complete.";
            statusLbl.Refresh();
        }
        private void Complete()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            createLbl.Refresh();
            uploadLbl.Refresh();
            downloadLbl.Refresh();
            formatLbl.Refresh();
            insertLbl.Refresh();
            updateLbl.Refresh();
            statusLbl.Text = "Status: Complete!";
            statusLbl.Refresh();
            endTmr.Enabled = true;
        }
        private void endTmr_Tick(object sender, EventArgs e)
        {
            Application.Exit();
            endTmr.Enabled = false;
        }

        private void createBtn_Click(object sender, EventArgs e)
        {
            CreateUploadFile();
        }
        private void uploadBtn_Click(object sender, EventArgs e)
        {            
            UploadToWiser();
        }
        private void downloadBtn_Click(object sender, EventArgs e)
        {
            DownloadFromWiser();
        }
        private void formatBtn_Click(object sender, EventArgs e)
        {
            FormatWiserFeedForBulkInsertion();
        }
        private void bulkBtn_Click(object sender, EventArgs e)
        {
            BulkInsertWiserFeed();
        }
        private void UpdateCompetitorList()
        {
            createLbl.Font = new System.Drawing.Font(createLbl.Font, FontStyle.Regular);
            uploadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            downloadLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            formatLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            insertLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Regular);
            updateLbl.Font = new System.Drawing.Font(uploadLbl.Font, FontStyle.Bold);
            updateLbl.Refresh();
            statusLbl.Text = "Status: Updating Competitor List...";
            statusLbl.Refresh();
            List<string> competitors = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT StoreName FROM PriceComparisons2");
            int counter = 0;
            foreach (string comp in competitors)
            {
                string cp = comp.Split('|')[0];
                counter++;
                statusLbl.Text = "Status: Updating Competitor List (" + counter.ToString() + "/" + competitors.Count.ToString() + ")";
                statusLbl.Refresh();
                string query = "SELECT CompetitorName FROM Competitors WHERE CompetitorName='" + cp.Replace("'", "''") + "'";
                List<string> exists = Connectivity.SQLCalls.sqlQuery(query);
                if (exists.Count <= 0)
                {
                    string query2 = "INSERT INTO Competitors (CompetitorName,Enabled,Blacklisted,Whitelisted,AddTax) VALUES('" + cp.Replace("'", "''") + "',1,0,1,0)";
                    Connectivity.SQLCalls.sqlQuery(query2);
                }
                else
                {
                }
            }

        }
        private void RunFullProcess()
        {
            CreateUploadFile();
            UploadToWiser();
            DownloadFromWiser();
            FormatWiserFeedForBulkInsertion();
            BulkInsertWiserFeed();
            UpdateCompetitorList();
            Complete();
        }
        private void RunDownloadAndInsertProcess()
        {           
            DownloadFromWiser();
            FormatWiserFeedForBulkInsertion();
            BulkInsertWiserFeed();
            UpdateCompetitorList();
            Complete();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }





    }
}
