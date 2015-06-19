using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PackingTime
{
    public partial class Form1 : Form
    {
        public struct UserPackingInfo
        {
            public string Username;
            public int UserID;
            public int TotalPackages;
            public int AvgTotalPackageTimeInSeconds;
            public int ShortestPackingTimeInSeconds;
            public int LongestPackingTimeInSeconds;
            public int AvgLastTenPackagesTimeInSeconds;


        }
        public struct PackingInfo
        {
            public DateTime StartDate;
            public DateTime StopDate;
            public List<UserPackingInfo> Packers;

        }
        public PackingInfo PackingStatus = new PackingInfo();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetupLVI();
        }

        private void SetupLVI()
        {
            listView1.Clear();
            listView1.Items.Clear();
            listView1.Columns.Clear();

            ColumnHeader packerName = new ColumnHeader();
            ColumnHeader totalPacked = new ColumnHeader();
            ColumnHeader avgPackTime = new ColumnHeader();
            ColumnHeader shortestPackTime = new ColumnHeader();
            ColumnHeader longestPackTime = new ColumnHeader();
            ColumnHeader lastTenPackagesTime = new ColumnHeader();
            ColumnHeader commentColumn = new ColumnHeader();

            packerName.Text = "Packer";
            packerName.Width = listView1.Width / 7;
            totalPacked.Text = "Tot. Packages";
            totalPacked.Width = listView1.Width / 8;
            avgPackTime.Text = "Avg. Time";
            avgPackTime.Width = listView1.Width / 7;
            shortestPackTime.Text = "Fastest Time";
            shortestPackTime.Width = listView1.Width /7;
            longestPackTime.Text = "Slowest Time";
            longestPackTime.Width = listView1.Width / 7;
            lastTenPackagesTime.Text = "Last 10 Avg.";
            lastTenPackagesTime.Width = listView1.Width / 6;
            commentColumn.Text = "Comments";
            commentColumn.Width = listView1.Width / 6;
            listView1.Columns.Add(packerName);
            listView1.Columns.Add(totalPacked);
            listView1.Columns.Add(avgPackTime);
            listView1.Columns.Add(shortestPackTime);
            listView1.Columns.Add(longestPackTime);
            listView1.Columns.Add(lastTenPackagesTime);
            listView1.Columns.Add(commentColumn);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            PackingStatus.StartDate = dateTimePicker1.Value;
            PackingStatus.StopDate = dateTimePicker2.Value;
            PackingStatus.Packers = new List<UserPackingInfo>();

            List<string> PackerIDList = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT UserID FROM totePackingTime");

            foreach (string packer in PackerIDList)
            {
                UserPackingInfo upi = new UserPackingInfo();
                //upi.ShortestPackingTimeInSeconds = int.MaxValue;
                //first make sure user is not 'deleted'
                List<string> userInfo = Connectivity.SQLCalls.sqlQuery("SELECT NTUsername,ShortName FROM Users WHERE ID=" + packer.Split('|')[0] + " AND Deleted=0");
                if (userInfo.Count > 0)
                {
                    upi.UserID = int.Parse(packer.Split('|')[0]);
                    upi.Username = userInfo[0].Split('|')[0];                    

                    List<string> packingData = Connectivity.SQLCalls.sqlQuery("SELECT ID,ReferenceNumber,StartTime,EndTime FROM totePackingTime WHERE StartTime>'" + PackingStatus.StartDate.ToString() + "' AND StartTime<'" + PackingStatus.StopDate.ToString() + "' AND UserID=" + upi.UserID + " AND EndTime > 0");
                    List<string> lastTen = Connectivity.SQLCalls.sqlQuery("SELECT TOP 10 ID,ReferenceNumber,StartTime,EndTime FROM totePackingTime WHERE UserID=" + upi.UserID + " AND EndTime > 0 ORDER BY StartTime");

                    if (lastTen.Count > 0)
                    {
                        TimeSpan tenTime = new TimeSpan();
                        int itemCount = 0;
                        foreach (string time in lastTen)
                        {
                            DateTime StartTime = DateTime.Parse(time.Split('|')[2]);
                            DateTime EndTime = DateTime.Parse(time.Split('|')[3]);


                            TimeSpan dumbTime = new TimeSpan();
                            dumbTime = EndTime - StartTime;
                            if (dumbTime.TotalSeconds < 3600)
                            {
                                tenTime += EndTime - StartTime;
                                itemCount++;
                            }
                            else
                            {

                            }

                        }
                        try
                        {
                            upi.AvgLastTenPackagesTimeInSeconds = (int)tenTime.TotalSeconds / itemCount;
                        }
                        catch
                        {
                            upi.AvgLastTenPackagesTimeInSeconds = 0;
                        }
                    }
                    else
                    {
                        upi.AvgLastTenPackagesTimeInSeconds = 0;
                    }
                    TimeSpan totTime = new TimeSpan();
                    
                    foreach (string package in packingData)
                    {
                        
                        string ID = package.Split('|')[0];
                        string ReferenceNumber = package.Split('|')[1];
                        DateTime StartTime = DateTime.Parse(package.Split('|')[2]);
                        DateTime EndTime = DateTime.Parse(package.Split('|')[3]);
                        TimeSpan Duration = EndTime-StartTime;
                        if (Duration.TotalSeconds < 3600)
                        {
                            upi.TotalPackages++;
                            totTime += Duration;
                        
                            if (upi.ShortestPackingTimeInSeconds == 0)
                            {
                                upi.ShortestPackingTimeInSeconds = (int)Duration.TotalSeconds;
                            }
                            if (upi.LongestPackingTimeInSeconds == 0)
                            {
                                upi.LongestPackingTimeInSeconds = (int)Duration.TotalSeconds;
                            }

                            if (Duration.TotalSeconds < (double)upi.ShortestPackingTimeInSeconds)
                            {
                                upi.ShortestPackingTimeInSeconds = (int)Duration.TotalSeconds;
                            }
                            if (Duration.TotalSeconds > (double)upi.LongestPackingTimeInSeconds)
                            {
                                upi.LongestPackingTimeInSeconds = (int)Duration.TotalSeconds;
                            }
                        }

                    }
                    if (upi.TotalPackages > 0)
                    {
                        upi.AvgTotalPackageTimeInSeconds = (int)totTime.TotalSeconds / upi.TotalPackages;
                    }
                    else
                    {
                        upi.AvgTotalPackageTimeInSeconds = 0;
                    }
                    PackingStatus.Packers.Add(upi);

                }
            }

            FillListView();
        }


        private void FillListView()
        {
            SetupLVI();
            foreach (UserPackingInfo upi in PackingStatus.Packers)
            {
                ListViewItem newPacker = new ListViewItem();

                TimeSpan avgTotalPackageTime = new TimeSpan(0,0,upi.AvgTotalPackageTimeInSeconds);
                TimeSpan avgLastTenPackagesTime = new TimeSpan(0, 0, upi.AvgLastTenPackagesTimeInSeconds);
                TimeSpan shortestPackageTime = new TimeSpan(0, 0, upi.ShortestPackingTimeInSeconds);
                TimeSpan longestPackageTime = new TimeSpan(0, 0, upi.LongestPackingTimeInSeconds);

                newPacker.Text = upi.Username;
                newPacker.SubItems.Add(upi.TotalPackages.ToString());
                newPacker.SubItems.Add(avgTotalPackageTime.ToString(@"mm\:ss"));
                newPacker.SubItems.Add(shortestPackageTime.ToString(@"mm\:ss"));
                newPacker.SubItems.Add(longestPackageTime.ToString(@"mm\:ss"));
                newPacker.SubItems.Add(avgLastTenPackagesTime.ToString(@"mm\:ss"));


                if (avgLastTenPackagesTime.TotalSeconds > avgTotalPackageTime.TotalSeconds)
                {
                    string datastr = (avgLastTenPackagesTime.TotalSeconds / avgTotalPackageTime.TotalSeconds).ToString("##.##").Substring(2) + "%";
                    if (datastr.Contains("finit") == false)
                    {
                        newPacker.SubItems.Add("Last 10 Packages slower by " + datastr);
                    }

                }
                if (avgLastTenPackagesTime.TotalSeconds < avgTotalPackageTime.TotalSeconds)
                {
                    string datastr = (avgTotalPackageTime.TotalSeconds / avgLastTenPackagesTime.TotalSeconds).ToString("##.##").Substring(2) + "%";
                    if (datastr.Contains("finit") == false)
                    {
                        newPacker.SubItems.Add("Last 10 Packages faster by " + datastr);
                    }
                }

                listView1.Items.Add(newPacker);

            }
        }

    }
}
