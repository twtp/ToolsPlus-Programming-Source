using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace WiserProductFeed
{
    public partial class WiserMainFrm : Form
    {
        public const string FilePath_CompetitorPrices = @"c:\users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv";
        public const string FilePath_FormattedFeedLocalized = @"c:\inetpub\wwwroot\formattedPrices.csv";
        public const string FilePath_FormattedFeedRemote = @"\\ToolsPlus06\c$\inetpub\wwwroot\formattedPrices.csv";

        public string[] args;
        System.Threading.Thread newThread;
        public List<string> BlacklistedCompetitors = new List<string>();
        public ListView.ListViewItemCollection competitorInfo = new ListView.ListViewItemCollection(new ListView());
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

        public WiserMainFrm()
        {
            InitializeComponent();
        }

        private void setupIncommingLVI()
        {
            incommingLVI.Items.Clear();
            incommingLVI.Columns.Clear();
            incommingLVI.Clear();

            ColumnHeader itemNumberCol = new ColumnHeader();
            ColumnHeader competitorNameCol = new ColumnHeader();
            ColumnHeader competitorPriceCol = new ColumnHeader();
            ColumnHeader sourceCol = new ColumnHeader();
           // ColumnHeader URL = new ColumnHeader();
            ColumnHeader competitorFreeShipping = new ColumnHeader();
            ColumnHeader totalCompetitorsCol = new ColumnHeader();
            ColumnHeader lastUpdatedCol = new ColumnHeader();

            itemNumberCol.Text = "Item Number";
            itemNumberCol.Width = incommingLVI.Width / 10;
            competitorNameCol.Text = "Competitor";
            competitorNameCol.Width = incommingLVI.Width / 8;
            competitorPriceCol.Text = "Price";
            competitorPriceCol.Width = incommingLVI.Width / 8;
            sourceCol.Text = "Source";
            sourceCol.Width = incommingLVI.Width / 7;
         //   URL.Text = "URL";
         //   URL.Width = incommingLVI.Width / 8;
            competitorFreeShipping.Text = "Shipping";
            competitorFreeShipping.Width = incommingLVI.Width / 8;
            totalCompetitorsCol.Text = "Competitors";
            totalCompetitorsCol.Width = incommingLVI.Width / 8;
            lastUpdatedCol.Text = "Detail:";
            lastUpdatedCol.Width = incommingLVI.Width / 8;

            incommingLVI.Columns.Add(itemNumberCol);
            incommingLVI.Columns.Add(sourceCol);
            incommingLVI.Columns.Add(competitorNameCol);
         //   incommingLVI.Columns.Add(URL);
            incommingLVI.Columns.Add(competitorPriceCol);
            incommingLVI.Columns.Add(competitorFreeShipping);
            incommingLVI.Columns.Add(totalCompetitorsCol);
            incommingLVI.Columns.Add(lastUpdatedCol);


        }

        private void WiserMainFrm_Load(object sender, EventArgs e)
        {
            args = Environment.GetCommandLineArgs();
            if (args != null)
            {
                if (args.GetLength(0) > 1)
                {

                    if (args[1].ToLower() == "/sch")
                    {
                        this.Visible = true;
                        this.Size = new Size(657, 71);
                        //pictureBox1.Location = new Point(592, 0);
                        //pictureBox1.Refresh();

                        this.Refresh();
                        //button1_Click(sender, e);
                        this.Refresh();
                        //button2_Click(sender, e);
                        this.Refresh();
                        //button4_Click(sender, e);
                        this.Refresh();
                        button6_Click(sender, e);
                        this.Refresh();
                        button3_Click(sender, e);
                        this.Refresh();
                        while (newThread.IsAlive == true)
                        {
                            Application.DoEvents();
                        }

                        Application.Exit();
                    }
                    else
                    {
                        delayLbl.Text = "Delay Setting: " + delayTrk.Value.ToString();
                        setupLVI();
                        setupIncommingLVI();
                    }
                }
                else
                {
                    delayLbl.Text = "Delay Setting: " + delayTrk.Value.ToString();
                    setupLVI();
                    setupIncommingLVI();
                }
            }
            else
            {
                delayLbl.Text = "Delay Setting: " + delayTrk.Value.ToString();
                setupLVI();
                setupIncommingLVI();
            }
           
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

        private void button1_Click(object sender, EventArgs e)
        {
            int repoAdditionCount = 0;
            setupLVI();
            feedLVI.Visible = false;
            feedLVI.Refresh();
            //UPC, ASIN, Brand, Model
            //MPN, Inventory Number, Product Name, Product Price
            //Minimum Price, Cost, Shipping Price, In Stock
            //Product URL, Image URL, Label, str_search_keywords



            //List<string> exampleProducts = Connectivity.SQLCalls.sqlQuery("SELECT TOP 10 ItemID,ItemNumber,ProductLine,IDiscountMarkupPriceRate1,StdCost,ShippingType FROM InventoryMaster WHERE EBayPrice > 150 AND ItemNumber NOT LIKE 'AAATEST'");
            List<string> exampleProducts = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT InventoryMaster.ItemID,InventoryMaster.ItemNumber,ProductLine,IDiscountMarkupPriceRate1,StdCost,ShippingType FROM InventoryMaster,PartNumbers,InventoryQuantities WHERE IsMASKit=0 AND WebPublished=1 AND WebOrderable=1 AND InventoryQuantities.ItemNumber=InventoryMaster.ItemNumber AND PartNumbers.ItemNumber = InventoryMaster.ItemNumber  AND InventoryMaster.ItemNumber NOT IN (SELECT ItemNumberPointer FROM PartNumberGroupMaster)");
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
                            if (UPC.StartsWith("TP") == false & missingASINChk.Checked == true)
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
            newLVI.SubItems.Add(string.Format("{0:C}",wfd.ProductPrice));
            newLVI.SubItems.Add(string.Format("{0:C}",wfd.MinimumPrice));
            newLVI.SubItems.Add(string.Format("{0:C}",wfd.Cost));
            newLVI.SubItems.Add(string.Format("{0:C}",wfd.ShippingPrice));
            newLVI.SubItems.Add(wfd.InStock.ToString());
            newLVI.SubItems.Add(wfd.ProductURL);
            newLVI.SubItems.Add(wfd.ImageURL);
            newLVI.SubItems.Add(wfd.Label);
            newLVI.SubItems.Add(wfd.StrSearchKeywords);
            feedLVI.Items.Add(newLVI);
        }

        private void WriteLVItoCSV()
        {
            string tmpCSVPath = @"C:\users\tomwestbrook\documents\wiserfeed\WiserFeed.csv";
            using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(tmpCSVPath))
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
                        outfile.Write("\"" + lvsi.Text.Replace("$","").Replace("\"","\"\"") + "\",");
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

        private void button2_Click(object sender, EventArgs e)
        {
            statusLbl.Text = "Status: Transferring Data...";
            statusLbl.Refresh();
            string result = Connectivity.FTPCalls.FTP_UPLOAD("ftp://ftp.wiser.com", "/upload/WP_Inv.csv", @"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv", "User4079", "U5sAPZfkCm7v", 2221);
            if (result.ToUpper().Contains("TRANSFER COMPLETE") == true)
            {
                statusLbl.Text = "Status: Transfer Complete!";
                statusLbl.Refresh();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            statusLbl.Text = "Status: Downloading Feed...";
            statusLbl.Refresh();
            string ReturnedFeed = Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com", "/Download/wp_inv_1256108.csv", "User4079", "U5sAPZfkCm7v", 2221);
            System.IO.File.WriteAllText(@"C:\Users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv", ReturnedFeed);
            statusLbl.Text = "Status: Downloading Feed...Complete.";
            statusLbl.Refresh();
            //MessageBox.Show(@"Check C:\Users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv.");
            
        }
        private void BulkInsertWiserFeed()
        {
            Connectivity.SQLCalls.sqlQuery("DELETE FROM CSVTest");
            Connectivity.SQLCalls.sqlQuery("BULK INSERT CSVTest FROM '" + FilePath_FormattedFeedLocalized +"' WITH (FIELDTERMINATOR = ',', ROWTERMINATOR = '\n')");
        }
        private void FormatWiserFeedForBulkInsertion()
        {
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
            string[] dataFeed = System.IO.File.ReadAllLines(FilePath_CompetitorPrices);
            int counter = 0;
            int linecount = 1;
            foreach (string line in dataFeed)
            {
                counter++;
                statusLbl.Text = "Status: Processing line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                statusLbl.Refresh();

                string newCSVLine = "";
                string[] csvTokens = System.Text.RegularExpressions.Regex.Split(line.Substring(0, line.Length - 1), @"""?\s*,\s*""?");
                int tokencount = 0;
                
                foreach (string token in csvTokens)
                {
                    tokencount++;
                    
                    if (tokencount == 1)
                    {
                        //this just checks for the header... if so, skip this line by breaking out from the inside of the foreach loop
                        if (token.ToUpper().Contains("PRODUCT NAME")==true | token.ToUpper().Contains("NVENTORY") == true)
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
                        newCSVLine += "Wiser.com,";
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

            statusLbl.Text = "Status: Preparing to write formatted file...";
            statusLbl.Refresh();

            //System.IO.File.WriteAllLines(@"C:\users\tomwestbrook\Documents\WiserFeed\formattedPrices.csv", formattedLines);
            System.IO.File.WriteAllLines(FilePath_FormattedFeedRemote,formattedLines);
            statusLbl.Text = "Status: Finished creating formatted file.";
            statusLbl.Refresh();

        }


        private void DisplayLatestFeed()
        {
            //incommingLVI.Items.Clear();
            competitorInfo.Clear();
            string[] dataFeed = System.IO.File.ReadAllLines(@"C:\Users\tomwestbrook\Documents\WiserFeed\CompetitorPrices.csv");
            int counter = 0;
            foreach (string line in dataFeed)
            {
                counter++;
                statusLbl.Text = "Status: Working on line #" + counter.ToString() + "/" + dataFeed.GetLength(0);
                statusLbl.Refresh();

                ListViewItem competitorScrape = new ListViewItem();

                string[] csvTokens = System.Text.RegularExpressions.Regex.Split(line.Substring(0,line.Length - 1), @"""?\s*,\s*""?");
                int tokencount = 0;
                foreach (string token in csvTokens)
                {
                    tokencount++;
                    if (tokencount == 1)
                    {
                        if (token.ToUpper() == "PRODUCT NAME" | token.ToUpper().Contains("NVENTORY") == true)
                        {
                            break;
                        }
                        else
                        {
                            competitorScrape.Text = token;
                        }
             
                    }
                    else
                    {
                        if (tokencount == 3)
                        {
                            //competitorScrape.Text = token;
                            competitorScrape.SubItems.Add("Wiser.com");
                            competitorScrape.SubItems.Add(token);   
                        }
                        if (tokencount == 19)//19)
                        {
                            //competitorScrape.SubItems.Add(token);                            
                        }
                        if (tokencount == 20)//20)
                        {
                            competitorScrape.SubItems.Add(token);
                        }
                        if (tokencount == 21)//21)
                        {
                            competitorScrape.SubItems.Add(token);
                        }
                        if (tokencount == 22) //22)
                        {
                            competitorScrape.SubItems.Add(token);
                        }                        
                        
                    }
                   // MessageBox.Show("Line #" + counter.ToString() + ": " + token);
                   
                }
                if (competitorScrape.Text != "")
                {
                    competitorScrape.SubItems.Add(DateTime.Now.ToString());
                    //incommingLVI.Items.Add(competitorScrape);
                    competitorInfo.Add(competitorScrape);
                    //incommingLVI.Refresh();
                }
            }
        }

        private void FilterBlackList()
        {
            this.BlacklistedCompetitors = Connectivity.SQLCalls.sqlQuery("SELECT CompetitorName FROM Competitors WHERE Blacklisted=1");
            
        }
        private void DelayWriteCheck(int delay)
        {
            if (delayWriteChk.Checked == true)
            {
                System.Threading.Thread.Sleep(delay);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //System.Threading.Thread work = new System.Threading.Thread(WorkThread);
            //work.Start();
            newThread = new System.Threading.Thread(WorkThread);

            //ListView.ListViewItemCollection damnit = new ListView.ListViewItemCollection(new ListView());
            

           // foreach (ListViewItem lvi in incommingLVI.Items)
           // {
           //     damnit.Add((ListViewItem)lvi.Clone());
           // }
            statusLbl.Text = "Status: Updating Database...";
            statusLbl.Refresh();

            newThread.Start(competitorInfo);          
            
        }
        public delegate ListView.ListViewItemCollection GetFileList();
        //public ListView.ListViewItemCollection GetFiles()
        //{
        //    return incommingLVI.Items;
        //}
        public void WorkThread(object lviData)
        {
            //GetFileList getFileList = new GetFileList(GetFiles);
            int totInserts = 0;
            int totUpdates = 0;
            int totSkipped = 0;
            int totBlacklisted = 0;
            int counter = 0;

            if (useBlacklistChk.Checked == true)
            {
                FilterBlackList();
            }

            

            //this.Invoke(new MethodInvoker(delegate { threadLVI = incommingLVI; }));
            
            ListView fakeView = new ListView();
            fakeView.View = View.Details;
            fakeView.FullRowSelect = true;
            fakeView.Columns.Add("1");
            fakeView.Columns.Add("2");
            fakeView.Columns.Add("3");
            fakeView.Columns.Add("4");
            fakeView.Columns.Add("5");
            fakeView.Columns.Add("6");
            fakeView.Columns.Add("7");

           // this.Invoke((MethodInvoker)(() => threadLVI.Items.AddRange(incommingLVI.Items)));

            DateTime startTime = DateTime.Now;
            DateTime currentProgress = DateTime.Now;
            //string TimeRemaining = "";

            foreach (ListViewItem lvi in (ListView.ListViewItemCollection)lviData)
            {
                counter++;
                bool blacklistFlag = false;
                foreach (string blacklisted in BlacklistedCompetitors)
                {
                    if (lvi.SubItems[2].Text == blacklisted.Split('|')[0])
                    {
                        totBlacklisted++;
                        blacklistFlag = true;
                        break;
                    }
                }
                if (blacklistFlag == false)
                {

                    currentProgress = DateTime.Now;
                    decimal percentDone = (decimal)(((decimal)(counter)) / ((decimal)(((ListView.ListViewItemCollection)lviData).Count)) * 100);
                    TimeSpan totTime = currentProgress - startTime;
                    decimal totDone = counter;
                    decimal totItems = ((ListView.ListViewItemCollection)lviData).Count;

                    decimal totTimes = totItems / totDone;
                    TimeSpan finTime = new TimeSpan();
                    for (int x = 0; x < totTimes; x++)
                    {
                        finTime = finTime.Add(totTime);
                    }
                    //MessageBox.Show(finTime.Duration().ToString());
                    string TimeRemaining = string.Format("{0:H:mm:ss}", new DateTime(finTime.Ticks));
                    

                                     

                    string ItemNumber = lvi.Text;
                    List<string> result = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber='" + ItemNumber + "'");
                    this.Invoke((MethodInvoker)(()=>DelayWriteCheck(delayTrk.Value)));
                    if (result.Count > 0)
                    {
                        string SourceVal = lvi.SubItems[1].Text;
                        string CompetitorVal = lvi.SubItems[2].Text;
                        string PriceVal = lvi.SubItems[3].Text;
                        string ShippingVal = lvi.SubItems[4].Text;
                        // string CompetitorCountVal = lvi.SubItems[5].Text;
                        string LastUpdateVal = lvi.SubItems[6].Text;

                        //ShippingVal = ShippingVal == "0" ? "1" : "0"; --was for support for original PriceComparisons table.
                        try
                        {
                            decimal.Parse(PriceVal);
                            decimal.Parse(ShippingVal);
                            DateTime.Parse(LastUpdateVal);

                            //now we check to see if a price comparison exists already for this item and it's competitor.
                            List<string> competitorExists = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM PriceComparisons2 WHERE ItemNumber='" + ItemNumber + "' AND Source='" + SourceVal + "' AND StoreName='" + CompetitorVal + "'");
                            if (competitorExists.Count > 0)
                            {
                                string query = "UPDATE PriceComparisons2 SET Price=" + PriceVal + ",Shipping=" + ShippingVal + ",LastUpdated='" + LastUpdateVal + "' WHERE ItemNumber='" + ItemNumber + "' AND StoreName='" + CompetitorVal + "' AND Source='" + SourceVal + "'";
                                Connectivity.SQLCalls.sqlQuery(query);
                                this.Invoke((MethodInvoker)(() => DelayWriteCheck(delayTrk.Value)));
                                totUpdates++;
                                this.Invoke((MethodInvoker)(() => lastUpdateLbl.Text = "Detail: " + totUpdates.ToString() + " updated / " + totInserts.ToString() + " new / " + totSkipped.ToString() + " skipped(errorous) / " + totBlacklisted.ToString() + " blacklisted / " + percentDone.ToString("0.0000") + "% done / Time Remaining: " + TimeRemaining.ToString()));
                                this.Invoke((MethodInvoker)(() => lastUpdateLbl.Refresh()));

                            }
                            else
                            {
                                string query = "INSERT INTO PriceComparisons2 (ItemNumber,LastUpdated,Source,StoreName,URL,Shipping,Price) VALUES('" + ItemNumber + "','" + LastUpdateVal + "','" + SourceVal + "','" + CompetitorVal + "','N/A'," + ShippingVal + "," + PriceVal + ")";
                                Connectivity.SQLCalls.sqlQuery(query);
                                this.Invoke((MethodInvoker)(() => DelayWriteCheck(delayTrk.Value)));
                                totInserts++;
                                this.Invoke((MethodInvoker)(() => lastUpdateLbl.Text = "Detail: " + totUpdates.ToString() + " updated / " + totInserts.ToString() + " new / " + totSkipped.ToString() + " skipped(errorous) / " + totBlacklisted.ToString() + " blacklisted / " + percentDone.ToString("0.0000") + "% done / Time Remaining: " + TimeRemaining.ToString()));
                                this.Invoke((MethodInvoker)(() => lastUpdateLbl.Refresh()));
                            }
                        }
                        catch { }

                    }
                    else
                    {
                        totSkipped++;
                        this.Invoke((MethodInvoker)(() => lastUpdateLbl.Text = "Detail: " + totUpdates.ToString() + " updated / " + totInserts.ToString() + " new / " + totSkipped.ToString() + " skipped(errorous) / " + totBlacklisted.ToString() + " blacklisted / " + percentDone.ToString("0.0000") + "% done / Time Remaining: " + TimeRemaining.ToString()));
                                this.Invoke((MethodInvoker)(() => lastUpdateLbl.Refresh()));
                    }
                }
                //this.Invoke((MethodInvoker)(()=>statusLbl.Text = "Status: Finished updating competitor information."));
                //this.Invoke((MethodInvoker)(()=>statusLbl.Refresh()));


            }
           
        }



       
        private void button6_Click(object sender, EventArgs e)
        {
            incommingLVI.Visible = false;
            incommingLVI.Refresh();
            DisplayLatestFeed();
            incommingLVI.Visible = true;
            

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void delayTrk_Scroll(object sender, EventArgs e)
        {
            delayLbl.Text = "Delay Setting: " + delayTrk.Value.ToString();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            FormatWiserFeedForBulkInsertion();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            statusLbl.Text = "Status: Bulk Inserting Items to SQL... Please Wait.";
            statusLbl.Refresh();
            BulkInsertWiserFeed();
            statusLbl.Text = "Status: Bulk Insertion Complete.";
            statusLbl.Refresh();
        }

    }
}
