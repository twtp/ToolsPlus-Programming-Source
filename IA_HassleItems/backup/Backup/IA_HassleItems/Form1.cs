using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace IA_HassleItems
{
    public partial class IA_HasslesFrm : Form
    {
        public IA_HasslesFrm()
        {
            InitializeComponent();
        }

        private void IA_HasslesFrm_Load(object sender, EventArgs e)
        {
            listView1.Columns.Add("Item Name");
            listView1.Columns.Add("# of Adjustments");
            listView1.Columns[0].Width = listView1.Width / 2;
            listView1.Columns[1].Width = listView1.Width / 2;
            

        }

        public void button1_Click(object sender, EventArgs e)
        {



            //string sql = "SELECT TransactionType,TransactionReference,TransactionTime,'|',InventoryComponentMap.ItemNumber,'|',WarehouseTransactionLog.ComponentID,'|' FROM WarehouseTransactionLog,InventoryComponentMap WHERE TransactionReference LIKE 'IA%' AND InventoryComponentMap.ComponentID = WarehouseTransactionLog.ComponentID ORDER BY WarehouseTransactionLog.ComponentID";
            //string sql = "http://10.0.50.16/whse/arbitraryMotorola.plex?SELECT LocationTypeID FROM LocationMaster WHERE WarehouseID=5";
            //string sql = "http://10.0.50.16/whse/arbitraryMotorola.plex?SELECT LocationTypeID FROM LocationMaster WHERE WarehouseID=5";
            string sql = "http://10.0.50.16/whse/arbitraryMotorola.plex?SELECT%20LocationTypeID%20FROM%20LocationMaster%20WHERE%20WarehouseID=5";
            //System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("http://10.0.50.16/whse/arbitraryMotorola.plex?" + sql);
            
               // using (System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
               // {
               //     using (System.IO.StreamReader r = new System.IO.StreamReader(resp.GetResponseStream()))
               //     {
            webBrowser1.Navigate(sql);
            webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);
            //System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(sql);

            //System.Net.HttpWebResponse response = (System.Net.HttpWebResponse)request.GetResponse();

            //System.IO.Stream resStream = response.GetResponseStream();

                        //string text = r.ReadToEnd();
                        //Text = Text.Replace("\r\n", "\n");
                        /*foreach (string row in Text.Split('\n'))
                        {
                            bool ItemExists = false;
                            for (int xCounter = 0; xCounter < listView1.Items.Count; xCounter++)
                            {
                                if (row.Split('|')[1] == listView1.Items[xCounter].Text)
                                {
                                    int IA_count = int.Parse(listView1.Items[xCounter].SubItems[0].Text);
                                    IA_count++;
                                    listView1.Items[xCounter].SubItems[0].Text = IA_count.ToString();
                                    ItemExists = true;
                                }
                            }
                            if (ItemExists == false)
                            {
                                ListViewItem newLVI = new ListViewItem();
                                newLVI.Text = row.Split('|')[1];
                                newLVI.SubItems[0].Text = "1";
                                listView1.Items.Add(newLVI);
                            }
                        }
                        //MessageBox.Show(text);
                        resp.Close();
                          
                         */
            //        }
            //    }
           /*     StringBuilder sb = new StringBuilder();
                byte[] buf = new byte[8192];

                        string tempString = null;
                        int count = 0;

                        do
                        {
                            count = resStream.Read(buf, 0, buf.Length);
                            if (count != 0)
                            {
                                tempString = Encoding.ASCII.GetString(buf, 0, count);
                                sb.Append(tempString);
                            }
                        }
                        while (count > 0);

                        string data = sb.ToString();

            */
        }

        void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            
        }
    }
}