using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace ReportGenerator
{
    public static class Report_PurchaseOrderBackOrder
    {
        public struct PurchaseOrderBackOrderOptions
        {
            public string LineCodeStart;
            public string LineCodeEnd;
            public string[] LineCodeList;
            public string ReportTitle;
            public bool ShowReportTitle;
            public DateTime StartDate;
            public DateTime EndDate;
        }

        private static string ReportStatus = "";

        //RETURN TYPES================================
        private static ListView reportLVI = new ListView();
        private static string reportCSV = "";
        //============================================
        public static ListView ReturnReportInListView()
        {
            return reportLVI;
        }
        public static string ReturnReportInCSV()
        {
            reportCSV = LVItoCSV(reportLVI);
            return reportCSV;
        }
        public static string ReturnReportStatus()
        {
            return ReportStatus;
        }
        //============================================

        private static string LVItoCSV(ListView ListViewControl)
        {
            string tmpFile = "";
            int Cols = ListViewControl.Columns.Count;

            foreach (ColumnHeader ch in ListViewControl.Columns)
            {
                if (ch.Text.Contains("\"") == true | ch.Text.Contains(",") == true)
                {
                    tmpFile += "\"" + ch.Text + "\",";
                }
                else
                {
                    tmpFile += ch.Text + ",";
                }
            }
            tmpFile += "\r\n";


            foreach (ListViewItem lvi in ListViewControl.Items)
            {
                for (int x = 0; x < Cols; x++)
                {
                    if (lvi.SubItems[x].Text.Contains("\"") == true | lvi.SubItems[x].Text.Contains(",") == true)
                    {
                        tmpFile += "\"" + lvi.SubItems[x].Text + "\",";
                    }
                    else
                    {
                        tmpFile += lvi.SubItems[x].Text + ",";
                    }
                }
                tmpFile += "\r\n";
            }

            return tmpFile;
            //statusLbl.Text = "File saved to \"" + FilePath + "\"!";
        }

        private static void ResetLVI()
        {
            reportLVI.Items.Clear();
            reportLVI.Columns.Clear();
            reportLVI.Clear();
            reportLVI.GridLines = true;
            reportLVI.FullRowSelect = true;
            reportLVI.View = View.Details;
            reportLVI.Refresh();
        }
        public static void Generate(PurchaseOrderBackOrderOptions options)
        {
            
            ResetLVI();
            //poboCreateBtn.Enabled = false;
            //poboCreateBtn.Refresh();
            DateTimePicker startDate = new DateTimePicker();
            DateTimePicker endDate = new DateTimePicker();
            startDate.Value = options.StartDate;
            endDate.Value = options.EndDate;

            string currentReportTitle = options.ReportTitle;
            string lineStart = options.LineCodeStart;
            string lineEnd = options.LineCodeEnd;

            List<string> linecodesToCheck = new List<string>();
            bool foundStart = false;
            foreach (string str in options.LineCodeList)
            {
                if (foundStart == true)
                {
                    if (str == lineEnd)
                    {
                        linecodesToCheck.Add(str);
                        break;
                    }
                    else
                    {
                        linecodesToCheck.Add(str);
                    }
                }
                else
                {
                    if (str == lineStart)
                    {
                        foundStart = true;
                        linecodesToCheck.Add(str);
                    }
                }
            }

            //string rawData = "";
            Dictionary<string, string> rawData = new Dictionary<string, string>();
            bool header = false;
            int lineCount = 1;
           
            foreach (string lineCode in linecodesToCheck)
            {
                ReportStatus = "Running POBO Report... Checking LineCode '" + lineCode + "' (" + lineCount.ToString() + "/" + linecodesToCheck.Count.ToString() + ")";

                string fuckyou = "SELECT ItemCode,PurchaseOrderNo,QuantityOrdered,QuantityReceived,QuantityBackordered,RequiredDate FROM PO_PurchaseOrderDetail WHERE ProductLine='" + lineCode + "' AND QuantityBackordered>0 AND RequiredDate>{d '" + startDate.Value.Year.ToString() + "-" + startDate.Value.Month.ToString("00") + "-" + startDate.Value.Day.ToString("00") + "'} AND RequiredDate<{d '" + endDate.Value.Year.ToString() + "-" + endDate.Value.Month.ToString("00") + "-" + endDate.Value.Day.ToString("00") + "'}";
                string json = Connectivity.MASCalls.Retrieve(fuckyou);

                if (header == false)
                {
                    if (json.Contains("rows\":") == true)
                    {
                        rawData.Add("HEADER", json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[0]);
                        header = true;
                    }
                }

                // if (json.Contains(":[]") == true)
                // {
                //     MessageBox.Show("Did not find nothing for linecode " + lineCode);
                // }
                // else
                // {
                //     MessageBox.Show("Holy Shit we got sumtin! (" + lineCode + ")");
                // }
                if (json.Contains("rows\":") == true)
                {
                    rawData.Add(lineCode, json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[1] + "\r\n");
                    //textBox1.Text += json.Split(new string[] { "rows\":" }, StringSplitOptions.None)[1] +"\r\n\r\n";
                }
                //statusLbl.Text = "Working on line " + lineCode + " (" + lineCount.ToString() + "/" + linecodesToCheck.Count.ToString() + ")";
                //statusLbl.Refresh();
                lineCount++;

            }

            //process raw data.
            RawDataToLVI(rawData, linecodesToCheck);

            ReportStatus = "";
            //poboCreateBtn.Enabled = true;
            //poboCreateBtn.Refresh();
            
        }

        private static void RawDataToLVI(Dictionary<string, string> data, List<string> LinesToCheck)
        {
            ReportStatus = "Running POBO Report... Processing Aquired Data...";

            reportLVI.Items.Clear();
            reportLVI.Columns.Clear();
            reportLVI.Clear();
            reportLVI.FullRowSelect = true;
            reportLVI.GridLines = true;
            reportLVI.View = View.Details;
            reportLVI.Refresh();

            string header = data["HEADER"];
            if (header.Length > 0)
            {
                string cols = header.Split(new string[] { "cols\":[" }, StringSplitOptions.None)[1].Split(']')[0];
                string[] columns = cols.Split(new string[] { "}," }, StringSplitOptions.None);
                foreach (string col in columns)
                {
                    ColumnHeader tmpCol = new ColumnHeader();
                    string ColumnName = col.Split(new string[] { "name\":\"" }, StringSplitOptions.None)[1].Split(new string[] { "\"," }, StringSplitOptions.None)[0];
                    string ColumnType = col.Split(new string[] { "typeID\":" }, StringSplitOptions.None)[1].Split(',')[0];
                    string ColumnSize = col.Split(new string[] { "size\":" }, StringSplitOptions.None)[1].Split(',')[0];

                    tmpCol.Text = ColumnName;
                    tmpCol.Tag = "{" + ColumnName + "}{" + ColumnType + "}{" + ColumnSize + "}";
                    reportLVI.Columns.Add(tmpCol);
                }

                foreach (string linecode in LinesToCheck)
                {
                    if (linecode == "BEN")
                    {
                    }

                    try
                    {
                        string rows = data[linecode];
                        string lineRows = "[" + rows.Split(new string[] { "[[" }, StringSplitOptions.None)[1].Split(new string[] { "]]" }, StringSplitOptions.None)[0] + "]";
                        string[] lines = lineRows.Split(new string[] { "[" }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string line in lines)
                        {
                            ListViewItem tmpRow = new ListViewItem();
                            string[] lineColumnData = line.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string colData in lineColumnData)
                            {

                                string cData = colData.Split('"')[1];
                                if (tmpRow.Text.Length == 0)
                                {
                                    tmpRow.Text = cData;
                                }
                                else
                                {
                                    tmpRow.SubItems.Add(cData);
                                }

                            }
                            reportLVI.Items.Add(tmpRow);

                        }

                    }
                    catch { }
                }


            }
            else
            {
                MessageBox.Show("No header returned!");
            }


        }        

        public static List<string> getLineCodes()
        {/*
            lineCodeEnd.Items.Clear();
            lineCodeStart.Items.Clear();
            lineCodeEnd.Refresh();
            lineCodeStart.Refresh();
            List<string> LineCodes = Connectivity.SQLCalls.sqlQuery("SELECT ProductLine FROM ProductLine");
            foreach (string line in LineCodes)
            {
                lineCodeStart.Items.Add(line.Split('|')[0]);
                lineCodeEnd.Items.Add(line.Split('|')[0]);
            }
            */
            List<string> final = new List<string>();
            List<string> tmp= Connectivity.SQLCalls.sqlQuery("SELECT ProductLine FROM ProductLine");
            foreach (string t in tmp)
            {
                final.Add(t.Split('|')[0]);
            }
            return final;
        }




    }
}
