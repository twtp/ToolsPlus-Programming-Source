using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace WisePricerCSVCreator
{
    public partial class CSVDesigner : Form
    {
        public CSVDesigner()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            statusLbl.Text = "Status: Starting Product Search...";
            csvView.Clear();
            CreateCSVViewHeader();
            textBox2.Text = "";
            textBox2.Text = "";
            List<string> ItemNumbers = new List<string>();

            foreach (string item in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                ItemNumbers.Add(item);
            }

            //MessageBox.Show("I now have " + ItemNumbers.Count.ToString() + " items and the second one is " + ItemNumbers[1]);
         
            foreach (string itemNumber in ItemNumbers)
            {
                string UPC = "";   //UPC is our Barcode. All Set
                string ASIN = "";  //ASIN is manufacturer name + (itemnumber minus linecode)
                string Brand = ""; //Brand is (NOT SURE EXACTLY) convert line code to full name
                string Model = ""; //itemnumber minus linecode
                string MPN = ""; // item number minus linecode
                string InventoryNumber = ""; //linecode plus itemnumber
                string ProductName = "";  // itemdescription
                string ProductPrice = ""; //webprice
                string MinimumPrice = ""; //webprice
                string Cost = ""; //stdcost
                string ShippingPrice = ""; //tricky.. free shipping = 0, not free shippin if price over 50 then free shipping
                string InStock = ""; //simply qoh
                string ProductURL = ""; //yahoo url
                string ImageURL = ""; //yahoo image url
                string Label = ""; //not exactly used
                string str_search_keywords = ""; //keywords we dont yet have???

                string ComponentID = "";


                List<string> asinID = SQLCalls.sqlQuery("SELECT ManufFullNameWeb FROM ProductLine,InventoryMaster WHERE PrimaryVendorNumber=PrimaryVendor AND ItemNumber='" + itemNumber + "' AND ProductLine.ProductLine=InventoryMaster.ProductLine");
                if (asinID.Count > 0)
                {
                    ASIN = (asinID[0].Split('|')[0] + "-" + itemNumber.Substring(3)).ToLower();
                }
                List<string> resBrand = SQLCalls.sqlQuery("SELECT ManufFullName FROM ProductLine WHERE ProductLine.ProductLine ='" + itemNumber[0].ToString() + itemNumber[1].ToString() + itemNumber[2].ToString() + "'");
                if (resBrand.Count > 0)
                {
                    Brand = resBrand[0].Split('|')[0];
                }
                List<string> resComponentID = SQLCalls.sqlQuery("SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + itemNumber + "'");
                if (resComponentID.Count > 0)
                {
                    ComponentID = resComponentID[0].Split('|')[0];
                }
                List<string> resUPC = SQLCalls.sqlQuery("SELECT Barcode FROM InventoryComponentBarcodes WHERE ComponentID='" + ComponentID + "'");
                if (resUPC.Count > 0)
                {
                    UPC = resUPC[0].Split('|')[0];
                }
                Model = itemNumber.Substring(3);
                MPN = itemNumber;
                InventoryNumber = itemNumber;

                List<string> resProductName = SQLCalls.sqlQuery("SELECT Desc3 FROM PartNumbers WHERE ItemNumber='" + itemNumber + "'");
                if (resProductName.Count > 0)
                {
                    ProductName = resProductName[0].Split('|')[0];
                }

                List<string> resProductPrice = SQLCalls.sqlQuery("SELECT EBayPrice From InventoryMaster WHERE ItemNumber='" + itemNumber + "'");
                if (resProductPrice.Count > 0)
                {
                    ProductPrice = decimal.Parse(resProductPrice[0].Split('|')[0]).ToString("0.00");
                }
                List<string> resMinimumPrice = SQLCalls.sqlQuery("SELECT MAPP From InventoryMaster WHERE ItemNumber='" + itemNumber + "'");
                if (resMinimumPrice.Count > 0)
                {
                    //MinimumPrice = resMinimumPrice[0].Split('|')[0];
                    //guess we just gonna match product price
                    MinimumPrice = ProductPrice;
                }
                List<string> resCost = SQLCalls.sqlQuery("SELECT StdPrice FROM InventoryMaster WHERE ItemNumber='" + itemNumber + "'");
                if (resCost.Count > 0)
                {
                    Cost = decimal.Parse(resCost[0].Split('|')[0].ToString()).ToString("0.00");
                }
                List<string> resInStock = SQLCalls.sqlQuery("SELECT QuantityOnHand FROM InventoryQuantities WHERE ItemNumber='" + itemNumber + "'");
                if (resInStock.Count > 0)
                {
                    InStock = resInStock[0].Split('|')[0];
                }

                string LineCode = itemNumber[0].ToString() + itemNumber[1].ToString() + itemNumber[2].ToString();
                string itemCode = itemNumber.Substring(3, itemNumber.Length - 3);
                string urlBase = "http://www.tools-plus.com/";
                string imgUrlBase = "http://ep.yimg.com/ay/toolsplus/";
                List<string> getManufacturerName = SQLCalls.sqlQuery("SELECT ManufFullNameWeb FROM ProductLine WHERE ProductLine.ProductLine='" + LineCode + "'");
                if (getManufacturerName.Count > 0)
                {
                    string desc = ProductName;
                    desc = desc.Replace(" ", "-");
                    desc = desc.Replace("/", "-");
                    desc = desc.Replace("\"", "quot-");
                    ImageURL = (imgUrlBase + getManufacturerName[0].Split('|')[0] + "-" + itemCode + "-" + desc + "-").ToLower(); //".jpg").ToLower();

                    string prodPage = (urlBase + getManufacturerName[0].Split('|')[0] + "-" + itemCode + ".html").ToLower().Replace(" ","-");
                    ImageURL = GetLatestImageUploaded(prodPage, itemNumber);
                    if (ImageURL == "N/A")
                    {
                        //no image found
                    }
                    else
                    {
                        //image found.
                    }
                }
                ProductURL = (urlBase + getManufacturerName[0].Split('|')[0] + "-" + itemCode + ".html").ToLower().Replace(" ","-");

                //Shipping part...
                List<string> shippingType = SQLCalls.sqlQuery("SELECT ShippingType FROM InventoryMaster WHERE ItemNumber='" + itemNumber + "'");
                if (shippingType.Count > 0)
                {
                    string shipping = shippingType[0];
                    if (shipping.Split('|')[0] == "3")
                    {
                        //free shipping
                        ShippingPrice = ProductPrice;
                    }
                    else
                    {
                        //if webprice > 50, free ship..
                        List<string> resProdPrice = SQLCalls.sqlQuery("SELECT EBayPrice From InventoryMaster WHERE ItemNumber='" + itemNumber + "'");
                        if (resProdPrice.Count > 0)
                        {
                            decimal prodPrice = decimal.Parse(resProductPrice[0].Split('|')[0]);
                            if (prodPrice > 50)
                            {
                                //free shipping
                                ShippingPrice = ProductPrice;
                            }
                            else
                            {
                                //50 and below...
                                ShippingPrice = (decimal.Parse(ProductPrice) + 8.50M).ToString("0.00");
                            }
                        }

                    }



                }
                ////////////////////



                //remove once figured out...               
               // ShippingPrice = "N/A";             
                Label = "N/A";
                str_search_keywords = "N/A";



                bool missingParamFlag = false;
                if (UPC.Length == 0)
                    missingParamFlag = true;
                if (ASIN.Length == 0)
                    missingParamFlag = true;
                if (Brand.Length == 0)
                    missingParamFlag = true;
                if (Model.Length == 0)
                    missingParamFlag = true;
                if (MPN.Length == 0)
                    missingParamFlag = true;
                if (InventoryNumber.Length == 0)
                    missingParamFlag = true;
                if (ProductName.Length == 0)
                    missingParamFlag = true;
                if (ProductPrice.Length == 0)
                    missingParamFlag = true;
                if (MinimumPrice.Length == 0)
                    missingParamFlag = true;
                if (Cost.Length == 0)
                    missingParamFlag = true;
                if (ShippingPrice.Length == 0)
                    missingParamFlag = true;
                if (InStock.Length == 0)
                    missingParamFlag = true;
                if (ProductURL.Length == 0)
                    missingParamFlag = true;
                if (ImageURL.Length == 0)
                    missingParamFlag = true;
                if (Label.Length == 0)
                    missingParamFlag = true;
                if (str_search_keywords.Length == 0)
                    missingParamFlag = true;






                textBox2.Text += UPC + "," + ASIN + "," + Brand + "," + Model + "," + MPN + "," + InventoryNumber + "," + ProductName + "," +
                    ProductPrice + "," + MinimumPrice + "," + Cost + "," + ShippingPrice + "," + InStock + "," + ProductURL + "," + ImageURL + "," +
                    Label + "," + str_search_keywords + ",\r\n";



                if (missingParamFlag == true)
                {
                    AddLineToCSVView(true, UPC, ASIN, Brand, Model, MPN, InventoryNumber, ProductName, ProductPrice, MinimumPrice, Cost, ShippingPrice, InStock, ProductURL, ImageURL, Label, str_search_keywords);
                }
                else
                {
                    AddLineToCSVView(false, UPC, ASIN, Brand, Model, MPN, InventoryNumber, ProductName, ProductPrice, MinimumPrice, Cost, ShippingPrice, InStock, ProductURL, ImageURL, Label, str_search_keywords);
                }


            }


            int badCount = 0;
            foreach (ListViewItem lvi in csvView.Items)
            {
                if (lvi.BackColor == Color.Tomato)
                {
                    badCount++;
                }
            }

            statusLbl.Text = "Status: Complete. " + csvView.Items.Count.ToString() + " items loaded. " + badCount.ToString() + " items missing data.";

        }

        private void AddLineToCSVView(bool highlight_line,string UPC, string ASIN, string Brand, string Model, string MPN, string InventoryNumber, string ProductName,
            string ProductPrice, string MinimumPrice, string Cost, string ShippingPrice, string InStock, string ProductURL, string ImageURL,
            string Label, string StringSearchKeywords)
        {
            ListViewItem addLVI = new ListViewItem();
            addLVI.Text = UPC;
            addLVI.SubItems.Add(ASIN);
            addLVI.SubItems.Add(Brand);
            addLVI.SubItems.Add(Model);
            addLVI.SubItems.Add(MPN);
            addLVI.SubItems.Add(InventoryNumber);
            addLVI.SubItems.Add(ProductName);
            addLVI.SubItems.Add(ProductPrice);
            addLVI.SubItems.Add(MinimumPrice);
            addLVI.SubItems.Add(Cost);
            addLVI.SubItems.Add(ShippingPrice);
            addLVI.SubItems.Add(InStock);
            addLVI.SubItems.Add(ProductURL);
            addLVI.SubItems.Add(ImageURL);
            addLVI.SubItems.Add(Label);
            addLVI.SubItems.Add(StringSearchKeywords);


            int lineCount = csvView.Items.Count;

            bool colored = lineCount % 2 != 0;

            if (colored == true)
            {
                addLVI.BackColor = Color.FromArgb(150,190,190);
            }
            else
            {
                addLVI.BackColor = Color.FromArgb(120,160,160);
            }

            if (highlight_line == true)
            {
                addLVI.BackColor = Color.Tomato;
            }
            csvView.Items.Add(addLVI);
            csvView.Refresh();
            System.Threading.Thread.Sleep(50);




        }
        private void CreateCSVViewHeader()
        {
            ColumnHeader csvHeader = new ColumnHeader();
            csvHeader.Text = "UPC";
            ColumnHeader csvHeader2 = new ColumnHeader();
            csvHeader2.Text = "ASIN";
            ColumnHeader csvHeader3 = new ColumnHeader();
            csvHeader3.Text = "Brand";
            ColumnHeader csvHeader4 = new ColumnHeader();
            csvHeader4.Text = "Model";
            ColumnHeader csvHeader5 = new ColumnHeader();
            csvHeader5.Text = "MPN";
            ColumnHeader csvHeader6 = new ColumnHeader();
            csvHeader6.Text = "Item Number";
            ColumnHeader csvHeader7 = new ColumnHeader();
            csvHeader7.Text = "Product Name";
            ColumnHeader csvHeader8 = new ColumnHeader();
            csvHeader8.Text = "Product Price:";
            ColumnHeader csvHeader9 = new ColumnHeader();
            csvHeader9.Text = "Min. Price";
            ColumnHeader csvHeader10 = new ColumnHeader();
            csvHeader10.Text = "Cost";
            ColumnHeader csvHeader11 = new ColumnHeader();
            csvHeader11.Text = "Ship Price";
            ColumnHeader csvHeader12 = new ColumnHeader();
            csvHeader12.Text = "In Stock";
            ColumnHeader csvHeader13 = new ColumnHeader();
            csvHeader13.Text = "Product URL";
            ColumnHeader csvHeader14 = new ColumnHeader();
            csvHeader14.Text = "Image URL";
            ColumnHeader csvHeader15 = new ColumnHeader();
            csvHeader15.Text = "Label";
            ColumnHeader csvHeader16 = new ColumnHeader();
            csvHeader16.Text = "String Search Keywords";


            csvView.Columns.Add(csvHeader);
            csvView.Columns.Add(csvHeader2);
            csvView.Columns.Add(csvHeader3);
            csvView.Columns.Add(csvHeader4);
            csvView.Columns.Add(csvHeader5);
            csvView.Columns.Add(csvHeader6);
            csvView.Columns.Add(csvHeader7);
            csvView.Columns.Add(csvHeader8);
            csvView.Columns.Add(csvHeader9);
            csvView.Columns.Add(csvHeader10);
            csvView.Columns.Add(csvHeader11);
            csvView.Columns.Add(csvHeader12);
            csvView.Columns.Add(csvHeader13);
            csvView.Columns.Add(csvHeader14);
            csvView.Columns.Add(csvHeader15);
            csvView.Columns.Add(csvHeader16);

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //string[] test = Environment.GetCommandLineArgs();
            //if (test != null)
            //{
            //    MessageBox.Show(test[1]);
            //}
            csvView.Clear();
            CreateCSVViewHeader();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "CSV File|*.csv";
            
            
            saveFileDialog1.ShowDialog();
            string saveFile = saveFileDialog1.FileName;
            if (saveFile == "")
            {
                return;
            }

            try
            {
                System.IO.File.WriteAllText(saveFile, textBox2.Text);
                MessageBox.Show("Saved to " + saveFile);
            }
            catch (Exception WTF)
            {
                MessageBox.Show("There was an issue trying to save. The error: " + WTF.Message);

            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void importItemListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                openFileDialog1.ShowDialog();
                string filepath = openFileDialog1.FileName;

                string GetData = System.IO.File.ReadAllText(filepath);

                string[] allItems = GetData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                textBox1.Lines = allItems;

                MessageBox.Show("Successfully Loaded " + allItems.GetLength(0).ToString() + " items.");
            }
            catch (Exception WTF)
            {
                if (openFileDialog1.FileName != "")
                {
                    MessageBox.Show("Error trying to open file. " + WTF.Message);
                }
            }


        }


        private string GetLatestImageUploaded(string itemPageURL, string itemNumber)
        {
            statusLbl.Text = "Status: Finding Yahoo Image For " + itemNumber;
            statusLbl.Refresh();
            string htmlCode = "";
            try
            {
                using (WebClient client = new WebClient())
                {
                    htmlCode = client.DownloadString(itemPageURL.ToLower());
                }
            }
            catch
            {
                //most likely a 404, return n/a
                statusLbl.Text = "Status: Yahoo Page Not Found.";
                statusLbl.Refresh();
                return "N/A";
            }

            //ok if we made it here, we can do some code parsing...

            try
            {
                string ImageURL = htmlCode.Split(new string[] { "<a rel=\"lightbox-itemimg\" href=\"" }, StringSplitOptions.None)[1].Split(new string[] { ".jpg" }, StringSplitOptions.None)[0] + ".jpg";
                statusLbl.Text = "Status: Yahoo Image Found.";
                statusLbl.Refresh();

                return ImageURL;
            }
            catch
            {
                try
                {
                    string ImageURL = htmlCode.Split(new string[] { "<a href=\"http://ep.yimg.com/ay/toolsplus/" }, StringSplitOptions.None)[1].Split(new string[] { ".jpg" }, StringSplitOptions.None)[0] + ".jpg";
                    statusLbl.Text = "Status: Yahoo Image Was Found.";
                    statusLbl.Refresh();
                    return ImageURL;
                }
                catch
                {
                    try
                    {
                        string ImageURL = htmlCode.Split(new string[] { "<a rel=\"lightbox\" href=\"" }, StringSplitOptions.None)[1].Split(new string[] { ".jpg" }, StringSplitOptions.None)[0] + ".jpg";
                        statusLbl.Text = "Status: Yahoo Image Finally Found.";
                        statusLbl.Refresh();
                        return ImageURL;
                        //<a rel="lightbox" href="
                    }
                    catch
                    {
                        //damnit... stop changing the pages on me.... for now u get nothing
                        statusLbl.Text = "Status: Yahoo Image Could Not Be Found.";
                        statusLbl.Refresh();
                        return "N/A";
                    }
                }
            }

            
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

    }
}
