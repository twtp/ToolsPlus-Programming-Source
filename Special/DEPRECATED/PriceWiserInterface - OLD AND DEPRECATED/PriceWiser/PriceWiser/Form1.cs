using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WisePricer
{
    public partial class Form1 : Form
    {
        public string ttImport = "9:04:00 AM";
        public DateTime TimeToImport = DateTime.Now;       
        public DateTime LastImportDate = new DateTime();
        public Form1()
        {
            InitializeComponent();
        }
        private void resetStatusLVI()
        {
            statusLVI.Items.Clear();
            statusLVI.Columns.Clear();
            statusLVI.Clear();
            ColumnHeader statusSymbol = new ColumnHeader();
            statusSymbol.Width = 25;
            statusSymbol.Text = "";
            statusLVI.Columns.Add(statusSymbol);
            ColumnHeader statusText = new ColumnHeader();
            statusText.Width = statusLVI.Width / 4;
            statusText.Text = "Status";
            statusText.TextAlign = HorizontalAlignment.Center;
            statusLVI.Columns.Add(statusText);
            ColumnHeader statusDesc = new ColumnHeader();
            statusDesc.Width = (int)(statusLVI.Width / 2.55);
            statusDesc.Text = "Description";
            statusDesc.TextAlign = HorizontalAlignment.Center;
            statusLVI.Columns.Add(statusDesc);
            ColumnHeader statusTimeStamp = new ColumnHeader();
            statusTimeStamp.Width = statusLVI.Width / 4;
            statusTimeStamp.Text = "Time";
            statusTimeStamp.TextAlign = HorizontalAlignment.Center;
            statusLVI.Columns.Add(statusTimeStamp);

            ImageList myList = new ImageList();
            myList.Images.Add(Properties.Resources.check_3d);
            myList.Images.Add(Properties.Resources.x_3d);
            statusLVI.SmallImageList = myList;

        }
        private void resetPriceWiserLVI()
        {
            priceWiserLVI.Items.Clear();
            priceWiserLVI.Columns.Clear();
            priceWiserLVI.Clear();
            ColumnHeader statusCol = new ColumnHeader();
            statusCol.Width = 25;
            statusCol.Text = "";
            priceWiserLVI.Columns.Add(statusCol);
            ColumnHeader itemNumberCol = new ColumnHeader();
            itemNumberCol.Width = (int)(priceWiserLVI.Width / 4.15);
            itemNumberCol.Text = "Item Number";
            itemNumberCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(itemNumberCol);
            ColumnHeader competitorsCol = new ColumnHeader();
            competitorsCol.Width = (priceWiserLVI.Width / 8);
            competitorsCol.Text = "Competitors";
            competitorsCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(competitorsCol);
            ColumnHeader minPriceCol = new ColumnHeader();
            minPriceCol.Width = (int)(priceWiserLVI.Width / 7.5);
            minPriceCol.Text = "Min. Price";
            minPriceCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(minPriceCol);
            ColumnHeader maxPriceCol = new ColumnHeader();
            maxPriceCol.Width = (int)(priceWiserLVI.Width / 7.5);
            maxPriceCol.Text = "Max. Price";
            maxPriceCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(maxPriceCol);
            ColumnHeader currentPriceCol = new ColumnHeader();
            currentPriceCol.Width = (int)(priceWiserLVI.Width / 7.5);
            currentPriceCol.Text = "Price";
            currentPriceCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(currentPriceCol);
            ColumnHeader newPriceCol = new ColumnHeader();
            newPriceCol.Width  = (int)(priceWiserLVI.Width / 7.5);
            newPriceCol.Text = "New Price";
            newPriceCol.TextAlign = HorizontalAlignment.Center;
            priceWiserLVI.Columns.Add(newPriceCol);

            ImageList myList = new ImageList();
            myList.Images.Add(Properties.Resources.check_3d);
            myList.Images.Add(Properties.Resources.x_3d);
            priceWiserLVI.SmallImageList = myList;
            
        }
        private void RefreshForm()
        {
            TimeToImport = DateTime.Parse(DateTime.Now.Date.ToShortDateString() + " " + ttImport);
            timeLbl.Parent = topImg;
            processingLbl.Parent = topImg;
            processingLbl.Visible = false;
            topImg.Top = 0;
            topImg.Left = 0;
            topImg.Height = 43;
            topImg.Width = this.Width;
            topImg.SendToBack();

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            resetStatusLVI();
            resetPriceWiserLVI();
            RefreshForm();
            addStatusMessage(StatusMessageSuccess.Successful, "Booted Up.", "PriceWiser SMC has been started!");
        }

        private string PullCSVFromWiser()
        {
            return Connectivity.FTPCalls.FTP_GET("ftp://ftp.wiser.com","/Download/wp_inv_1256108.csv","User4079", "U5sAPZfkCm7v",2221);

            //try
            //{
                //connect to ftp backend... pull our export..
                //for testing purposes though...
            //    return System.IO.File.ReadAllText(@"c:\users\tomwestbrook\desktop\wiserpricer.csv");
            //}
            //catch
            //{
            //    return "";
            //}
            
        }
        private void addItemToWiserLVI(string ItemNumber, int Competitors, decimal MinPrice, decimal MaxPrice, decimal Price, decimal NewPrice)
        {
            ListViewItem newItem = new ListViewItem();
            if (MinPrice > 0 & NewPrice > 0)
            {
                newItem.ImageIndex = StatusMessageSuccess.Successful;
            }
            else
            {
                newItem.ImageIndex = StatusMessageSuccess.Error;
            }
            newItem.SubItems.Add(ItemNumber);
            newItem.SubItems.Add(Competitors.ToString());
            newItem.SubItems.Add(MinPrice.ToString("$#.##;-$#.##;$0.00"));
            newItem.SubItems.Add(MaxPrice.ToString("$#.##;-$#.##;$0.00"));
            newItem.SubItems.Add(Price.ToString("$#.##;-$#.##;$0.00"));
            newItem.SubItems.Add(NewPrice.ToString("$#.##;-$#.##;$0.00"));
            priceWiserLVI.Items.Add(newItem);
            priceWiserLVI.Refresh();

        }
        private bool ProcessCSVFromWiser(string CSVFile)
        {
            processingLbl.Text = "";
            processingLbl.Visible = true;
            processingLbl.Refresh();
            try
            {
                string[] CSVLines = CSVFile.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                int failedLines = 0;
                int tryCount = 0;
                foreach (string line in CSVLines)
                {
                    tryCount++;
                    processingLbl.Text = "Processing " + tryCount.ToString() + "/" + CSVLines.GetLength(0).ToString();
                    processingLbl.Refresh();
                    if (line.Contains("Product Name") == false)
                    {
                        string partNumber = line.Split(',')[6];
                        //List<string> itemNumber = Connectivity.SQLCalls.sqlQuery("SELECT ITEMNUMBER FROM vInventorySpreadsheet WHERE PartNumber='" + partNumber + "'");
                        List<string> itemNumber = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber LIKE '___" + partNumber + "'");
                        if (itemNumber.Count > 0)
                        {

                            string ItemNumber = itemNumber[0].Split('|')[0];
                            int Competitors = 0;
                            decimal MinPrice = 0;
                            decimal MaxPrice = 0;
                            decimal Price = 0;
                            decimal NewPrice = 0;

                            if (int.TryParse(line.Split(',')[33], out Competitors) == true)
                            {
                                if (decimal.TryParse(line.Split(',')[9], out MinPrice) == true)
                                {
                                    if (decimal.TryParse(line.Split(',')[10], out MaxPrice) == true)
                                    {
                                        if (decimal.TryParse(line.Split(',')[8], out Price) == true)
                                        {
                                            if (decimal.TryParse(line.Split(',')[15], out NewPrice) == true)
                                            {
                                                addItemToWiserLVI(ItemNumber, Competitors, MinPrice, MaxPrice, Price, NewPrice);
                                                List<string> location = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM PriceWiserItems WHERE ItemNumber='" + ItemNumber + "'");
                                                if (location.Count > 0)
                                                {
                                                    string itemID = location[0].Split('|')[0];
                                                    Connectivity.SQLCalls.sqlQuery("UPDATE PriceWiserItems SET MinPrice=" + MinPrice + ", MaxPrice=" + MaxPrice + ", Price=" + Price + ", NewPrice=" + NewPrice + ", LastUpdate='" + DateTime.Now.ToString() + "' WHERE ID=" + itemID);

                                                }
                                                else
                                                {
                                                    string query = "INSERT INTO PriceWiserItems (ItemNumber,MinPrice,MaxPrice,Price,NewPrice,LastUpdate) VALUES('" + ItemNumber + "'," + MinPrice + "," + MaxPrice + "," + Price + "," + NewPrice + ",'" + DateTime.Now.ToString() + "')";
                                                    //MessageBox.Show(query);
                                                    Connectivity.SQLCalls.sqlQuery(query);
                                                    
                                                }
                                            }
                                            else
                                            {
                                                if (line.Split(',')[15] == "")
                                                {
                                                    NewPrice = 0;
                                                    addItemToWiserLVI(ItemNumber, Competitors, MinPrice, MaxPrice, Price, NewPrice);
                                                    List<string> location = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM PriceWiserItems WHERE ItemNumber='" + itemNumber + "'");
                                                    if (location.Count > 0)
                                                    {
                                                        string itemID = location[0].Split('|')[0];
                                                        Connectivity.SQLCalls.sqlQuery("UPDATE PriceWiserItems SET MinPrice=" + MinPrice + ", MaxPrice=" + MaxPrice + ", Price=" + Price + ", NewPrice=" + NewPrice + ", LastUpdate='" + DateTime.Now.ToString() + "' WHERE ID=" + itemID);

                                                    }
                                                    else
                                                    {
                                                        Connectivity.SQLCalls.sqlQuery("INSERT INTO PriceWiserItems (ItemNumber,MinPrice,MaxPrice,Price,NewPrice,LastUpdate) VALUES('" + itemNumber + "'," + MinPrice + "," + MaxPrice + "," + Price + "," + NewPrice + ",'" + DateTime.Now.ToString() + "')");

                                                    }
                                                }
                                                else
                                                {
                                                    failedLines++;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            failedLines++;
                                        }
                                    }
                                    else
                                    {
                                        failedLines++;
                                    }
                                }
                                else
                                {
                                    failedLines++;
                                }
                            }
                            else
                            {
                                failedLines++;
                            }



                        }
                    }
                }
            }
            catch
            {
                processingLbl.Text = "";
                processingLbl.Visible = false;
                return false;
            }
            processingLbl.Text = "";
            processingLbl.Visible = false;
            return true;
        }

        

        private void secondTmr_Tick(object sender, EventArgs e)
        {
            TimeSpan ts = TimeToImport - DateTime.Now;
            timeLbl.Text = "Time: " + DateTime.Now.ToString() + "\r\nNext Download In: " + ts.TotalSeconds.ToString("#") + " secs.";
            timeLbl.Refresh();
            if (DateTime.Now.TimeOfDay >= TimeToImport.TimeOfDay & LastImportDate.Day != DateTime.Now.Day)
            {
                TimeToImport = DateTime.Parse(DateTime.Now.Date.AddDays(1).ToShortDateString() + " " + ttImport);
                addStatusMessage(StatusMessageSuccess.Successful, "CSV Pulling Time", "Preparing to pull the CSV file from Wiser.");
                LastImportDate = DateTime.Now;
                string csv = PullCSVFromWiser();
                if (csv != "")
                {
                    addStatusMessage(StatusMessageSuccess.Successful, "CSV Pulled", "Pulled the csv file from Wiser ok.");
                    bool ok = ProcessCSVFromWiser(csv);
                    if (ok == true)
                    {
                        addStatusMessage(StatusMessageSuccess.Successful, "Processed CSV", "Processed CSV data from Wiser.");
                    }
                    else
                    {
                        addStatusMessage(StatusMessageSuccess.Error, "Processed CSV Failure", "Processing the CSV from Wiser has failed.");
                    }
                }
                else
                {
                    addStatusMessage(StatusMessageSuccess.Error, "CSV Pull failed", "Pulling the csv file from Wiser failed.");
                }
               
            }
        }

        internal class StatusMessageSuccess
        {
            public const int Successful = 0;
            public const int Error = 1;
            public const int Warning = 2;
            public const int Unknown = 3;
            public const int Other = 4;
        }

        private void addStatusMessage(int SuccessType, string StatusMessage, string StatusDescription)
        {
                   
            if (StatusMessageSuccess.Warning == SuccessType)
            {
                MessageBox.Show("Status Messages with status type of 'Warning' is not available yet!");
                return;
            }
            if (StatusMessageSuccess.Unknown == SuccessType)
            {
                MessageBox.Show("Status Messages with status type of 'Unknown' is not available yet!");
                return;
            }
            if (StatusMessageSuccess.Other == SuccessType)
            {
                MessageBox.Show("Status Messages with status type of 'Other' is not available yet!");
                return;
            }

            ListViewItem newSuccessMessage = new ListViewItem();
            newSuccessMessage.ImageIndex = SuccessType;
            newSuccessMessage.SubItems.Add(StatusMessage);
            newSuccessMessage.SubItems.Add(StatusDescription);
            newSuccessMessage.SubItems.Add(DateTime.Now.ToString());
            statusLVI.Items.Add(newSuccessMessage);
            statusLVI.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string result = Connectivity.FTPCalls.FTP_GET("ftp://ftp.fu-berlin.de/doc/iso/iso3166-countrycodes.txt");
           MessageBox.Show(result);

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            string result = Connectivity.FTPCalls.FTP_LIST("ftp://ftp.wiser.com","/Download", "User4079", "U5sAPZfkCm7v", 2221);
            MessageBox.Show(result);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Connectivity.FTPCalls.FTP_UPLOAD("ftp://ftp.wiser.com", "/upload/WP_Inv.csv", @"C:\Users\tomwestbrook\Documents\WiserFeed\WiserFeed.csv", "User4079", "U5sAPZfkCm7v", 2221));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            TimeToImport = DateTime.Parse(DateTime.Now.Date.AddDays(1).ToShortDateString() + " " + ttImport);
            addStatusMessage(StatusMessageSuccess.Successful, "CSV Pulling Time", "Preparing to pull the CSV file from Wiser.");
            LastImportDate = DateTime.Now;
            string csv = PullCSVFromWiser();
            if (csv == "")
            {
                bool processFlag = ProcessCSVFromWiser(csv);
                if (processFlag == true)
                {
                    addStatusMessage(StatusMessageSuccess.Successful, "Processed CSV", "Processed CSV data from Wiser.");
                }
                else
                {
                    addStatusMessage(StatusMessageSuccess.Error, "Processed CSV Failure", "Processing the CSV from Wiser has failed.");
                }
            }
            else
            {
                addStatusMessage(StatusMessageSuccess.Error, "CSV Pull failed", "Pulling the csv file from Wiser failed.");
            }

        }

    }
}
