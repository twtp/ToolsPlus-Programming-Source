using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;
using mshtml;


namespace LivePriceComparer
{



    public partial class Form1 : Form
    {







        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        public static int GWL_STYLE = -16;
        public static int WS_CHILD = 0x40000000;






        public const string Version = "0.3.0b";
        private string LastSearch = "";
        private bool isNavigating = false;
        private bool dontSearch = false;
        private string ItemNumber = "";

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            versionLbl.Text = "v. " + Version;
            RefreshSizes();
            //googleBrowser.Navigate("https://www.google.com/shopping");
            CompetitorsData = GetCompetitorData();

            ebayBrowser.Navigate("about:blank");
            googleBrowser.Navigate("about:blank");
            yahooBrowser.Navigate("about:blank");


            
            tabControl1.SelectedTab = tabControl1.TabPages[1];
            tabControl1.SelectedTab = tabControl1.TabPages[2];
            tabControl1.SelectedTab = tabControl1.TabPages[3];
            tabControl1.SelectedTab = tabControl1.TabPages[4];
            tabControl1.SelectedTab = tabControl1.TabPages[0];


        }
        private void FindBtn_Click(object sender, EventArgs e)
        {
            ebayStatus.Text = "eNo";
            googleStatus.Text = "gNo";
            yahooStatus.Text = "yNo";
            if (itemTxt.Text == LastSearch)
            {
                SetStatus("Not re-searching the same item as last time!");
                return;
            }
            FindItem();
            FindItem2();
            FindItem3();
        }
        private void itemTxt_TextChanged(object sender, EventArgs e)
        {
            if (itemTxt.Text.Contains("<consume>"))
            {
                InjectIntoPOINV();
                itemTxt.Text = "";
                return;
            }
            if (itemTxt.Text.Contains("<puke>"))
            {
                ExpelFromPOINV();
                itemTxt.Text = "";
                return;
            }


            if (dontSearch == true)
            {
                dontSearch = false;
                return;
            }
            //MessageBox.Show("i changed!");
            if (itemTxt.Text.Contains("<FIND>") == true)
            {

                if (itemTxt.Text.Contains("<ITEMNUMBER>")==true)
                {
                    ItemNumber = itemTxt.Text.Split(new string[] { "<ITEMNUMBER>" }, StringSplitOptions.None)[1].Split(new string[] { "</ITEMNUMBER>" }, StringSplitOptions.None)[0];
                    //itemTxt.Text = itemTxt.Text.Split(new string[] { "</ITEMNUMBER>" }, StringSplitOptions.None)[1];
                    dontSearch = true;
                }
                else
                {
                    ItemNumber = "";
                }
                if (dontSearch == true)
                {
                    itemTxt.Text = itemTxt.Text.Replace("<FIND>", "").Split(new string[] { "</ITEMNUMBER>" }, StringSplitOptions.None)[1];
                }
                else
                {
                    itemTxt.Text = itemTxt.Text.Replace("<FIND>", "");
                }
                itemTxt.Refresh();
                FindBtn_Click(sender, e);

            }
        }
        private async Task FindItem()
        {
            
            clearAllBrowsers();
            SetStatus("Preparing to search...");

            if (EBayUseEBayAPI == true)
            {

                await Task.Run(() =>
                {
                    if (processEBayChk.Checked == true)
                    {
                        SetStatus("Searching EBay...");
                        FindAndProcessEBayResults();
                    }
                    // SetStatus("Finished Processing EBay");
                });
            }

            if (EBayUseRawData == true)
            {
                await Task.Run(() =>
                {
                    if (processEBayChk.Checked == true)
                    {
                        SetStatus("Searching EBay (Raw)...");
                        FindAndProcessEBayResultsRaw();
                    }
                    // SetStatus("Finished Processing EBay");
                });
            }




           // await Task.Run(() =>
           // {
           //     if (processGoogleChk.Checked == true)
           //     {
           //         SetStatus("Searching Google Shopping...");
           //         FindAndProcessGoogleShoppingResults();
           //     }
           // });
           // await Task.Run(() =>
           // {
           //     if (processYahooChk.Checked == true)
           //     {
           //         SetStatus("Searching Yahoo Shopping...");
           //         FindAndProcessYahooShoppingResults();
           //     }
           // });
                LastSearch = itemTxt.Text;

            
            //SetStatus("Finished processing search results!");
        }
        private async Task FindItem2()
        {
            
            await Task.Run(() =>
            {
                if (processGoogleChk.Checked == true)
                {
                    SetStatus("Searching Google Shopping...");
                    FindAndProcessGoogleShoppingResults();
                }
               // SetStatus("Finished Processing Goole");
            });         


            //SetStatus("Finished processing search results!");
        }
        private async Task FindItem3()
        {
           

            await Task.Run(() =>
            {
                if (processYahooChk.Checked == true)
                {
                    SetStatus("Searching Yahoo Shopping...");
                    FindAndProcessYahooShoppingResults();
                   // SetStatus("Finished Processing Yahoo");
                }
            });
            //LastSearch = itemTxt.Text;


            //SetStatus("Finished processing search results!");
        }
        private void SetStatus(string cMessage)
        {
            if (statusLbl.InvokeRequired == true)
            {
                
               Invoke(new MethodInvoker(() =>{statusLbl.Text = "Status: " + cMessage; statusLbl.Refresh();}));

           
            }
            else
            {
                statusLbl.Text = "Status: " + cMessage;
                statusLbl.Refresh();
            }
        }
        private void SetDebug(string text)
        {
            if (debugTxt.InvokeRequired == true)
            {
                debugTxt.Invoke(new MethodInvoker(() =>
                {
                    debugTxt.Text = text;
                    debugTxt.Refresh();
                }));
            }
            else
            {
                debugTxt.Text = text;
                debugTxt.Refresh();
            }
        }
        private void debugDisplayBtn_Click(object sender, EventArgs e)
        {
            debugBrowser.DocumentText = debugTxt.Text;
        }
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            RefreshSizes();
        }
        private void RefreshSizes()
        {
            //Take care of Widths here...
            tabControl1.Width = this.ClientSize.Width - 16;
            tabControl1.Left = 8;
            ebayBrowser.Width = tabControl1.ClientSize.Width - 16;
            ebayBrowser.Left = 8;


            //Take care of Heights here...
            tabControl1.Height = (int)((float)(this.ClientSize.Height) * 0.9f);
            tabControl1.Top = itemTxt.Bottom + 10;
            ebayBrowser.Height = (int)((float)(tabControl1.ClientSize.Height) * 0.9f);
            ebayBrowser.Top = 1;

            statusLbl.Top = this.ClientSize.Height - statusLbl.Height;
            statusLbl.Left = 0;
        }




        private void debugBrowser2_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (EBayUseRawData == true)
            {
                rawFlag = false;
            }
        }

        private void showOnlyNewItemsChk_CheckedChanged(object sender, EventArgs e)
        {
            ShowOnlyNewItems = showOnlyNewItemsChk.Checked;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            debugBrowser2.Navigate("javascript:void(document.getElementById('custLink').click());");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            debugBrowser2.Navigate("javascript:void(document.getElementById('e1-13').checked=true);");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            debugBrowser2.Navigate("javascript:void(document.getElementById('e1-3').click());");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            ebayStatus.Refresh();
            googleStatus.Refresh();
            yahooStatus.Refresh();
        }

        public void ExpelFromPOINV()
        {
            SetParent(this.Handle, IntPtr.Zero);
            this.BringToFront();
        }
        public void InjectIntoPOINV()
        {
            //this.Hide();
            System.Diagnostics.Process hostProcess = System.Diagnostics.Process.GetProcessesByName("poinv-b5").FirstOrDefault();
            if (hostProcess != null)
            {

                //this.Hide();
                //this.FormBorderStyle = FormBorderStyle.None;
                //this.SetBounds(0, 0, 0, 0, BoundsSpecified.Location);
                IntPtr hostHandle = hostProcess.MainWindowHandle;
                IntPtr guestHandle = this.Handle;
                SetWindowLong(guestHandle, GWL_STYLE, GetWindowLong(guestHandle, GWL_STYLE) | WS_CHILD);
                SetParent(guestHandle, hostHandle);

                //this.Show();
            }
        }

































        


    }
}
