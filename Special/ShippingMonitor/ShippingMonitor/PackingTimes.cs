using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ShippingMonitor
{    

    public static class PackingTimes
    {        
        public static List<string> packingData = new List<string>();
        public static List<string> lastTen = new List<string>();
        public static List<string> PackerIDList = new List<string>();
        public static List<string> userInfo = new List<string>();
        public static string packerIdListSQL = "";
        public static string userInfoSQL = "";
        public static string lastTenSQL = "";
        public static string packingDataSQL = "";


        public struct UserPackingInfo
        {
            public string Username;
            public int UserID;
            public int TotalPackages;
            public int AvgTotalPackageTimeInSeconds;
            public int ShortestPackingTimeInSeconds;
            public int LongestPackingTimeInSeconds;
            public int AvgLastTenPackagesTimeInSeconds;
            public int AvgTimeBetweenPacksInSeconds;


        }
        public struct PackingInfo
        {
            public DateTime StartDate;
            public DateTime StopDate;
            public List<UserPackingInfo> Packers;

        }
        public static PackingInfo PackingStatus = new PackingInfo();



        public static void ResetListView(ref ListView listView1, int formWidth)
        {
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.None);
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();

            ColumnHeader packerName = new ColumnHeader();
            packerName.Text = "Packer";
            packerName.TextAlign = HorizontalAlignment.Center;
            packerName.Width = formWidth / 6;
            ColumnHeader fastestTime = new ColumnHeader();
            fastestTime.Text = "Total";
            fastestTime.TextAlign = HorizontalAlignment.Center;
            fastestTime.Width = formWidth / 7;
            ColumnHeader slowestTime = new ColumnHeader();
            slowestTime.Text = "Fastest";
            slowestTime.TextAlign = HorizontalAlignment.Center;
            slowestTime.Width = formWidth/ 5;
            ColumnHeader avgTimeBetweenPacks = new ColumnHeader();
            avgTimeBetweenPacks.Text = "Between Avg";
            avgTimeBetweenPacks.TextAlign = HorizontalAlignment.Center;
            avgTimeBetweenPacks.Width = formWidth / 4;


            ColumnHeader avgTime = new ColumnHeader();
            avgTime.Text = "Average";
            avgTime.TextAlign = HorizontalAlignment.Center;
            avgTime.Width = formWidth / 4;

            listView1.Columns.Add(packerName);
            listView1.Columns.Add(fastestTime);
            listView1.Columns.Add(slowestTime);
            listView1.Columns.Add(avgTimeBetweenPacks);
            listView1.Columns.Add(avgTime);
        }

        public static void TopPackerInfo()
        {

        }

        public static void UpdatePackersGruntWork()
        {
        }
        public static void UpdatePackers(ref ListView listView1, int formWidth)
        {
            TimeSpan totTimeBetweenPackings = new TimeSpan();
            int packageCount = 0;
            DateTime lastDateEndTime = new DateTime();

            PackingStatus.StartDate = DateTime.Now;//.AddDays(-1);//dateTimePicker1.Value;
            PackingStatus.StopDate = DateTime.Now;//.AddDays(1); //dateTimePicker2.Value;
            PackingStatus.Packers = new List<UserPackingInfo>();

            packerIdListSQL = "SELECT DISTINCT TOP 5 UserID FROM totePackingTime WHERE StartTime < '" + DateTime.Now.ToString() + "'";
            PackerIDList = Connectivity.SQLCalls.sqlQuery(packerIdListSQL);

            foreach (string packer in PackerIDList)
            {
                UserPackingInfo upi = new UserPackingInfo();
                //upi.ShortestPackingTimeInSeconds = int.MaxValue;
                //first make sure user is not 'deleted'
                userInfoSQL = "SELECT NTUsername,ShortName FROM Users WHERE ID=" + packer.Split('|')[0] + " AND Deleted=0";
                userInfo = Connectivity.SQLCalls.sqlQuery(userInfoSQL);
                if (userInfo.Count > 0)
                {
                    upi.UserID = int.Parse(packer.Split('|')[0]);
                    upi.Username = userInfo[0].Split('|')[0];

                    packingDataSQL = "SELECT ID,ReferenceNumber,StartTime,EndTime FROM totePackingTime WHERE StartTime>'" + PackingStatus.StartDate.ToShortDateString() + "' AND UserID=" + upi.UserID + " AND EndTime > 0";
                    packingData = Connectivity.SQLCalls.sqlQuery(packingDataSQL);

                    lastTenSQL = "SELECT TOP 10 ID,ReferenceNumber,StartTime,EndTime FROM totePackingTime WHERE UserID=" + upi.UserID + " AND EndTime > 0 ORDER BY StartTime DESC";
                    lastTen = Connectivity.SQLCalls.sqlQuery(lastTenSQL);
                    TimeSpan duration = new TimeSpan();
                    int x = 0;
                    for (x = 0; x < lastTen.Count-1; x++)
                    {
                        //TimeSpan ts1 = TimeSpan.FromTicks(DateTime.Parse(lastTen[x + 1].Split('|')[2]).Ticks);
                        //TimeSpan ts2 = TimeSpan.FromTicks(DateTime.Parse(lastTen[x].Split('|')[3]).Ticks);
                        duration = DateTime.Parse(lastTen[x].Split('|')[3]).Subtract(DateTime.Parse(lastTen[x+1].Split('|')[2]));

                        totTimeBetweenPackings.Add(duration);                       
                        
                    }
                    int avgTimeBetween = (int)duration.TotalSeconds / x;
                    upi.AvgTimeBetweenPacksInSeconds = avgTimeBetween;

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
                        TimeSpan Duration = EndTime - StartTime;

                       
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
                    //need to add upi.AvgTimeBetweenPacks...
                   
                    PackingStatus.Packers.Add(upi);

                }
            }
            
            FillListView(ref listView1, formWidth);
        }



        public static void FillListView(ref ListView listView1, int formWidth)
        {
            ResetListView(ref listView1, formWidth);
            foreach (UserPackingInfo upi in PackingStatus.Packers)
            {
                ListViewItem newPacker = new ListViewItem();

                TimeSpan avgTotalPackageTime = new TimeSpan(0, 0, upi.AvgTotalPackageTimeInSeconds);
                TimeSpan avgLastTenPackagesTime = new TimeSpan(0, 0, upi.AvgLastTenPackagesTimeInSeconds);
                TimeSpan shortestPackageTime = new TimeSpan(0, 0, upi.ShortestPackingTimeInSeconds);
                TimeSpan longestPackageTime = new TimeSpan(0, 0, upi.LongestPackingTimeInSeconds);

                newPacker.Text = upi.Username;
                newPacker.SubItems.Add(upi.TotalPackages.ToString());                
                newPacker.SubItems.Add(shortestPackageTime.ToString());
                newPacker.SubItems.Add(new TimeSpan(0,0,upi.AvgTimeBetweenPacksInSeconds).Duration().ToString());
                newPacker.SubItems.Add(avgTotalPackageTime.ToString());

                //newPacker.SubItems.Add(longestPackageTime.ToString());
                //newPacker.SubItems.Add(avgLastTenPackagesTime.ToString());


               /* if (avgLastTenPackagesTime.TotalSeconds > avgTotalPackageTime.TotalSeconds)
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
                */
                listView1.Items.Add(newPacker);

            }
        }



    }
}
