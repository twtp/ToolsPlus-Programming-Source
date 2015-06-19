using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace PriceSnagger
{
    public partial class Form1 : Form
    {
        public const string VERSION = "1.0.0";

        Form2.ConfigParameters ConfigParams = new Form2.ConfigParameters();
        bool RanTodaysCheck = false;
        DateTime LastCheckTime = new DateTime();
        bool isRunning = false;
        List<ItemPricings> ItemList = new List<ItemPricings>();
        bool isPaused = false;

        public struct ItemPricings
        {
            public string ItemNumber;
            public string ItemPrefix;
            public string ItemSerial;
        }



        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            versionLbl.Text = "Version: " + VERSION;
            
            Form2 config = new Form2();
            ConfigParams = config.GetConfig();
            ResetWork();
        }

        private void ResetWork()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();

            ColumnHeader ItemNumber = new ColumnHeader();
            ItemNumber.Text = "Item Number:";
            ItemNumber.TextAlign = HorizontalAlignment.Center;
            ItemNumber.Width = listView1.Width / 5;

            ColumnHeader ItemPrefix = new ColumnHeader();
            ItemPrefix.Text = "Prefix:";
            ItemNumber.TextAlign = HorizontalAlignment.Center;
            ItemNumber.Width = listView1.Width / 5;
            
            ColumnHeader ItemSerial = new ColumnHeader();
            ItemSerial.Text = "Serial:";
            ItemSerial.TextAlign = HorizontalAlignment.Center;
            ItemSerial.Width = listView1.Width / 5;

            ColumnHeader ToolUpPrice = new ColumnHeader();
            ToolUpPrice.Text = "ToolUp.com:";
            ToolUpPrice.TextAlign = HorizontalAlignment.Center;
            ToolUpPrice.Width = listView1.Width / 5;

            ColumnHeader CPOPrice = new ColumnHeader();
            CPOPrice.Text = "CPOoutlets.com:";
            CPOPrice.TextAlign = HorizontalAlignment.Center;
            CPOPrice.Width = listView1.Width / 5;

            listView1.Columns.Add(ItemNumber);
            listView1.Columns.Add(ItemPrefix);
            listView1.Columns.Add(ItemSerial);
            listView1.Columns.Add(ToolUpPrice);
            listView1.Columns.Add(CPOPrice);



        }


        
        private void button1_Click(object sender, EventArgs e)
        {
            isRunning = true;
            //LastCheckTime = DateTime.Now;
            statusLbl.Text = "Status: Gathering all items weborderable, published, and higher than red priority.";

            PriceSnagger.SQLCalls sqlQuery = new SQLCalls();
            string SQLQuery = "SELECT ItemNumber, WebOrderable, TBPubPriority FROM PartNumbers WHERE (WebOrderable=1 AND WebPublished=1  AND ItemNumber LIKE 'MLW0%') OR (WebToBePublished=1 AND TBPubPriority>2  AND ItemNumber LIKE 'MLW0%')";
            List<string> QueryList = sqlQuery.sqlQuery(SQLQuery);

            //now we have a list of weborderable & published (since they are in the list) items
            foreach (string itemLine in QueryList)
            {
                ItemPricings IP = new ItemPricings();
                IP.ItemNumber = itemLine.Split('|')[0];
                IP.ItemPrefix = IP.ItemNumber.Substring(0,3);
                IP.ItemSerial = IP.ItemNumber.Substring(3, IP.ItemNumber.Length - 3);
                

                if (IP.ItemSerial != "TEST")
                {

                    ItemList.Add(IP);
                }
                                
            }

            //now before going further, let's add our current info to the listview...

            foreach (ItemPricings ip in ItemList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = ip.ItemNumber;
                ListViewItem.ListViewSubItem lvsi = new ListViewItem.ListViewSubItem();
                lvsi.Text = ip.ItemPrefix;
                ListViewItem.ListViewSubItem lvsi2 = new ListViewItem.ListViewSubItem();
                lvsi2.Text = ip.ItemSerial;
                ListViewItem.ListViewSubItem lvsi3 = new ListViewItem.ListViewSubItem();
                lvsi3.Text = "";
                ListViewItem.ListViewSubItem lvsi4 = new ListViewItem.ListViewSubItem();
                lvsi4.Text = "";
                lvi.SubItems.Add(lvsi);
                lvi.SubItems.Add(lvsi2);
                lvi.SubItems.Add(lvsi3);
                lvi.SubItems.Add(lvsi4);
                listView1.Items.Add(lvi);
            }



            statusLbl.Text = "Status: Gathering Items Complete.";

            button2.Enabled = true;

        }


        private void UpdatePriceComparisons(ListViewItem lvi)
        {
            //do toolsup.com first...
            
                SQLCalls updateDB = new SQLCalls();
                
                //step 1. Try to find this item in PriceComparisons with source being ToolsUp.com
                List<string> ResultsToolUp = updateDB.sqlQuery("SELECT Source FROM PriceComparisons WHERE Source='ToolUp.com' AND ItemNumber='" + lvi.Text + "'");
                if (ResultsToolUp.Count > 0)
                {
                    //if the item did exist, remove it!
                    updateDB.sqlQuery("DELETE FROM PriceComparisons WHERE ItemNumber='" + lvi.Text + "' AND Source='ToolUp.com'");
                    System.Threading.Thread.Sleep(200);
                }
                List<string> ResultsCPO = updateDB.sqlQuery("SELECT Source FROM PriceComparisons WHERE Source='CPOoutlets.com' AND ItemNumber='" + lvi.Text + "'");
                if (ResultsCPO.Count > 0)
                {
                    //if the item did exist, remove it!
                    updateDB.sqlQuery("DELETE FROM PriceComparisons WHERE ItemNumber='" + lvi.Text + "' AND Source='CPOoutlets.com'");
                    System.Threading.Thread.Sleep(200);
                }

                //step 2. Now that the sql table is cleaned up of this item for this particular source.. add it!
               
                if (lvi.SubItems[3].Text.ToUpper() != "N/A")
                {
                    //ToolsUp.com does free shipping on orders over $199.00.. excludes oversized and freight but let's not get too picky here lol!
                    string FreeShipping = "";
                    //MessageBox.Show("IS FREE SHIP? " + decimal.Parse((lvi.SubItems[3].Text.Trim('$')).Trim()).ToString() + " is greater than $199?");
                    if (decimal.Parse((lvi.SubItems[3].Text.Trim('$')).Trim()) >= 199)
                    {
                        FreeShipping = "1";
                    }
                    else
                    {
                        FreeShipping = "0";
                    }
                    updateDB.sqlQuery("INSERT INTO PriceComparisons VALUES ('" + lvi.Text.ToUpper() + "','" + DateTime.Now.ToString() + "','ToolUp.com','ToolUp.com','http://search.toolup.com/search?w=" + lvi.Text + "'," + FreeShipping + "," + lvi.SubItems[3].Text + ")");
                    
                }
                if (lvi.SubItems[4].Text.ToUpper() != "N/A")
                {
                    string FreeShipping = "0";
                    updateDB.sqlQuery("INSERT INTO PriceComparisons VALUES ('" + lvi.Text.ToUpper() + "','" + DateTime.Now.ToString() + "','CPOoutlets.com','CPOoutlets.com','http://www.cpooutlets.com/on/demandware.store/Sites-cpooutlets-Site/default/Search-Show?q=" + lvi.Text + "&prefn1=condition&prefv1=new&srule=price-low-to-high&start=0&sz=24'," + FreeShipping + "," + lvi.SubItems[4].Text + ")");
                }

            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button3.Enabled = true;
            button3.Text = "Pause";
            isPaused = false;


            int ItemCount = ItemList.Count;
            if (ItemList.Count != listView1.Items.Count)
            {
                MessageBox.Show("Hey, buddy... theres a difference in records here... Item List Has " + ItemList.Count.ToString() + " and the listview has " + listView1.Items.Count.ToString());
            }
            for (int counter = 0; counter < ItemCount; counter++)
            {
                if (isPaused == false)
                {
                    
                    statusLbl.Text = "Status: Processing Item #" + ItemList[counter].ItemNumber + ". This is item " + counter.ToString() + "/" + ItemCount.ToString();
                    string ToolUp_SearchQuery = "http://search.toolup.com/search?w=" + ItemList[counter].ItemSerial;
                    string CPO_SearchQuery = "http://www.cpooutlets.com/on/demandware.store/Sites-cpooutlets-Site/default/Search-Show?q=" + ItemList[counter].ItemSerial + "&prefn1=condition&prefv1=new&srule=price-low-to-high&start=0&sz=24";
                    string ZORO_SearchQuery = "http://www.zoro.com/s/?q=&pn=&mn=&mp=" + ItemList[counter].ItemSerial + "&bn=&c=&pf=&pt=&SEARCH=SEARCH";


                    string ToolUp_results = "";
                    string CPO_Results = "";
                    string ZORO_Results = "";

                    using (WebClient client = new WebClient())
                    {
                        ToolUp_results = client.DownloadString(ToolUp_SearchQuery);
                        
                    }

                    try
                    {
                        ToolUp_results = ToolUp_results.Split(new string[] { "<span class=\"price\" id=\"" }, StringSplitOptions.None)[1].Split('>')[1].Split('<')[0];
                        listView1.Items[counter].SubItems[3].Text = ToolUp_results.Replace(",","");
                       
                    }
                    catch
                    {
                        listView1.Items[counter].SubItems[3].Text = "N/A";
                    }


                    using (WebClient client = new WebClient())
                    {
                        CPO_Results = client.DownloadString(CPO_SearchQuery);
                    }
                    try
                    {
                        if (CPO_Results.ToUpper().Split(new string[] { "<TITLE>" }, StringSplitOptions.None)[1].Split(new string[] { "</TITLE>" }, StringSplitOptions.None)[0] == "SITES-CPOOUTLETS-SITE")
                        {
                            //File.WriteAllText("C:\\users\\tomwestbrook\\documents\\CPO_Result.txt", CPO_Results);
                            //MessageBox.Show("$" + CPO_Results.Split(new string[] { "standard:" }, StringSplitOptions.None)[1]); //.Split(new string[] { "sale: \"" }, StringSplitOptions.None)[1]);
                            CPO_Results = CPO_Results.Split(new string[] { "<div class=\"salesprice\">" }, StringSplitOptions.None)[1].Split(new string[] { "</div>" }, StringSplitOptions.None)[0];
                            listView1.Items[counter].SubItems[4].Text = CPO_Results.Replace(",", "");
                        }
                        else
                        {
                                                  
                            listView1.Items[counter].SubItems[4].Text = "N/A";
                        }
                        
                    }
                    catch
                    {
                        listView1.Items[counter].SubItems[4].Text = "N/A";
                    }



                    //LEFT OFF HERE..
                    using (WebClient client = new WebClient())
                    {
                        ZORO_Results = client.DownloadString(ZORO_SearchQuery);
                    }
                    try
                    {

                    }
                    catch
                    {

                    }
                    //////////////////////////




                    ListViewItem lastItem = listView1.Items[counter];

                    UpdatePriceComparisons(lastItem);

                    listView1.Refresh();
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    counter--;
                    //listView1.Refresh();
                    System.Threading.Thread.Sleep(30);
                    Application.DoEvents();
                }
            }
            statusLbl.Text = "Status: Processing Items Completed!";
            isRunning = false;
        }




        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
           // try
           // {
           //     MessageBox.Show("Part Index: " + listView1.SelectedIndices[0].ToString());
           // }
           // catch
           // {
           // }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (button3.Text == "Pause")
            {
                button3.Text = "Resume";
                isPaused = true;
            }
            else
            {
                button3.Text = "Pause";
                isPaused = false;
            }
        }

        private void statusLbl_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 settingsFrm = new Form2();
            settingsFrm.Show();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //MessageBox.Show(((int)((DateTime.Now - LastCheckTime).TotalHours)).ToString());
            //MessageBox.Show((DateTime.Now - LastCheckTime).TotalHours.ToString());
           
            if ((int)((DateTime.Now - LastCheckTime).TotalHours) > 20)
            {
                RanTodaysCheck = false;
            }
            if (RanTodaysCheck == false)
            {
                LastCheckTime = DateTime.Now;
                string Today = DateTime.Now.DayOfWeek.ToString();

                if ((Today.ToUpper() == "MONDAY" & ConfigParams.Monday_Run == true) | (Today.ToUpper() == "TUESDAY" & ConfigParams.Tuesday_Run == true) | (Today.ToUpper() == "WEDNESDAY" & ConfigParams.Wednesday_Run == true) | (Today.ToUpper() == "THURSDAY" & ConfigParams.Thursday_Run == true) | (Today.ToUpper() == "FRIDAY" & ConfigParams.Friday_Run ==true) | (Today.ToUpper() == "SATURDAY" & ConfigParams.Saturday_Run == true) | (Today.ToUpper() == "SUNDAY" & ConfigParams.Sunday_Run == true))
                {
                    statusLbl.Text = "Status: Waiting for " + ConfigParams.TimeToRun.ToString() + " to run the check today.";

                    int runCheck = TimeSpan.Compare(DateTime.Now.TimeOfDay, ConfigParams.TimeToRun.TimeOfDay);
                    if (runCheck > 0)
                    {
                        //time to run that check...
                        RanTodaysCheck = true;
                        button1_Click(sender, e);
                        button2_Click(sender, e);
                    }

                }
                else
                {
                    statusLbl.Text = "Status: No check is taking place today.";
                }
            }
            else
            {
                if (isRunning == false)
                {
                    statusLbl.Text = "Status: Check ran already today.";
                }
            }
        }
    }
}
