using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EBaySubtitleDefaulter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Maximum = 100;
            progressBar1.Minimum = 0;
            SetStatus("Loaded. Press Green Button To Run.",Color.White,0);
        }
        private void SetStatus(string message, Color TextColor, int progressPercent)
        {
            if (statusLbl.InvokeRequired == true)
            {

                statusLbl.Invoke(new MethodInvoker(()=>
                {
                    statusLbl.Text = message;
                    statusLbl.ForeColor = TextColor;
                    statusLbl.Refresh();
                }));

            }
            else
            {
                statusLbl.Text = message;
                statusLbl.ForeColor = TextColor;
                statusLbl.Refresh();
            }

            if (progressBar1.InvokeRequired == true)
            {
                progressBar1.Invoke(new MethodInvoker(() =>
                    {

                        progressBar1.Value = progressPercent;
                        progressBar1.Refresh();                        

                    }));
            }
            else
            {
                progressBar1.Value = progressPercent;
                progressBar1.Refresh();
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            RunTask();
        }

        private async Task RunTask()
        {
            await Task.Run(() =>
                {
                    SubtitleTask();
                });
        }

        private async Task<string> SubtitleTask()
        {
            Connectivity.SQLCalls.QueryResults dst = Connectivity.SQLCalls.sqlQuery("SELECT Subtitle FROM Subtitles WHERE ID=1");
            if (dst.Rows.Count > 0)
            {
                string subtitle = dst.Rows[0].Split('|')[0];

                Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT PartNumbers.ItemNumber FROM PartNumbers WHERE EBayPublished=1 AND PartNumbers.ItemNumber NOT IN (SELECT ItemNumber FROM SubtitledItems)");
                List<string> Items = new List<string>();

                int progress = 0;

                if (qr.Rows.Count > 0)
                {
                    int totcount = qr.Rows.Count;
                    int count = 1;

                    foreach (string row in qr.Rows)
                    {
                        progress = (int)((float)((float)count / ((float)totcount * 2.00f) * 100.00f));
                        string item = row.Split('|')[0].Trim();
                        Items.Add(item);
                        SetStatus("Adding " + item + " to list. (" + count.ToString() + "/" + totcount.ToString() + ")", Color.Green,progress);
                        System.Threading.Thread.Sleep(50);
                        count++;
                    }

                    count = 1;
                    foreach (string itm in Items)
                    {
                        progress = 50 + (int)((float)((float)count / ((float)totcount * 2.00f) * 100.00f));
                        SetStatus("Subtitling item " + itm + " (" + count.ToString() + "/" + Items.Count.ToString() + ")", Color.Green,progress);
                        Connectivity.SQLCalls.sqlQuery("INSERT INTO SubtitledItems (ItemNumber,SubtitleID) VALUES('" + itm + "',1)");
                        System.Threading.Thread.Sleep(50);
                        count++;
                    }
                    SetStatus("Completed @ " + DateTime.Now.ToShortTimeString() + "!", Color.White, 100);
                }
                else
                {
                    SetStatus("Found 0 items to default subtitle on.", Color.Orange,0);
                }
            }
            else
            {
                SetStatus("Could not find default subtitle.", Color.Red,0);
            }
            return null;
        }
    }
}
