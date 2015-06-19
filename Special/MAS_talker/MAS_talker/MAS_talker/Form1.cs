using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MAS_talker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime startTime = DateTime.Now;            
            string results = SQLCalls.MASCalls.MAS200.Retrieve(textBox1.Text);
            DateTime endTime = DateTime.Now;
            TimeSpan ts = endTime - startTime;

            statusLbl.Text = "Status: Took MAS " + ts.TotalSeconds.ToString("0.00") + " seconds to retrieve this data.";
            statusLbl.Refresh();
            textBox2.Text = results;            
            MAStoLVI(results, listView1);
        }

        public struct MASColumnHeader
        {
            public string Name;
            public int TypeID;
            public int Size;
        }              
        public void MAStoLVI(string responseData, ListView outputLV)
        {
            string ColumnData = "";
            string RowData = "";

            try
            {
                ColumnData = responseData.Split('[')[1].Split(']')[0];
                RowData = responseData.Split(new string[] { "rows\":[" }, StringSplitOptions.None)[1].Split(new string[] { "]}" }, StringSplitOptions.None)[0];
            }
            catch { MessageBox.Show("The last query errored as it returned nothing."); return; }
            string[] Columns = ColumnData.Split(new string[] { "{\"name\":\"" }, StringSplitOptions.RemoveEmptyEntries);
            for (int x = 0; x < Columns.GetLength(0); x++)
            {
                Columns[x] = Columns[x].Split('}')[0];
            }

            string[] Rows = RowData.Split(new string[]{"["},StringSplitOptions.RemoveEmptyEntries);
            for (int y = 0; y < Rows.GetLength(0); y++)
            {
                Rows[y] = Rows[y].Split(']')[0];
            }

            List<MASColumnHeader> finalColumns = new List<MASColumnHeader>();
            List<List<string>> finalRows = new List<List<string>>();
            foreach (string data in Columns)
            {
                MASColumnHeader mch = new MASColumnHeader();
                string Name, TypeID, Size = "";
                Name = data.Split('"')[0];
                TypeID = data.Split(new string[] { "typeID\":" }, StringSplitOptions.None)[1].Split(',')[0];
                Size = data.Split(new string[] { "size\":" }, StringSplitOptions.None)[1].Split(',')[0];
                mch.Name = Name;
                mch.TypeID = int.Parse(TypeID);
                mch.Size = int.Parse(Size);
                finalColumns.Add(mch);
            }
            foreach (string data in Rows)
            {
                List<string> fields = new List<string>();
                bool splitOffset = false;
                foreach (string field in data.Split(','))
                {
                    if (splitOffset == false)
                    {
                        splitOffset = true;
                        fields.Add(field.Replace("\"", ""));
                    }
                    else
                    {
                        fields.Add(field.Replace("\"", ""));
                    }
                }
                finalRows.Add(fields);
            }

            List<ListViewItem> resultsRows = new List<ListViewItem>();
            List<ColumnHeader> resultsColumns = new List<ColumnHeader>();

            outputLV.Items.Clear();
            outputLV.Columns.Clear();
            outputLV.Clear();
            outputLV.Refresh();
            outputLV.GridLines = true;
            outputLV.FullRowSelect = true;
            outputLV.View = View.Details;

            foreach (MASColumnHeader mch in finalColumns)
            {
                ColumnHeader tmpCol = new ColumnHeader();
                tmpCol.Text = mch.Name;
                tmpCol.Tag = mch;
                tmpCol.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                resultsColumns.Add(tmpCol);
            }
            foreach (List<string> fieldList in finalRows)
            {
                ListViewItem tmpLVI = new ListViewItem();
                bool firstCol = true;
                foreach(string fld in fieldList)
                {
                    if (firstCol == false)
                    {
                        tmpLVI.SubItems.Add(fld);
                    }
                    else
                    {
                        tmpLVI.Text = fld;
                        firstCol = false;
                    }

                }
                resultsRows.Add(tmpLVI);
            }

            foreach (ColumnHeader ch in resultsColumns)
            {
                outputLV.Columns.Add(ch);
            }
            foreach (ListViewItem lvi in resultsRows)
            {
                outputLV.Items.Add(lvi);
            }


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void WriteLVIToExcel(ListView myList, bool showTitle)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            ws.Columns.EntireColumn.ColumnWidth = 20;

            int i = 1;
            int i2 = 3;

            //if (showTitle == true)
            //{
            //    ws.Cells[1, 1] = currentReportTitle;
            //    ws.Cells[1, 3] = currentReportLogoText;
            //}
            //else
            //{
                i2 = 2;
            //}

            foreach (ListViewItem lvi in myList.Items)
            {
                i = 1;
                List<string> columnList = new List<string>();
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    columnList.Add(lvs.Text);
                }
                for (int x = 0; x < columnList.Count; x++)
                {
                    foreach (ColumnHeader ch in myList.Columns)
                    {
                        if (ch.DisplayIndex == x)
                        {
                            //this should really be changed so we dont overwrite a header a million times over...
                            int y = 0;
                            if (showTitle == true)
                            {
                                y = 2;
                            }
                            else
                            {
                                y = 1;
                            }

                            ws.Cells[y, i].EntireRow.Font.Bold = true;
                            ws.Cells[y, i] = ch.Text;
                            ws.Cells[i2, i] = columnList[ch.Index];
                            i++;
                        }
                    }
                }
                i2++;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            GC.Collect();
        }
        private void WriteLVItoCSV(string filepath)
        {
            if (filepath == "")
            {
                filepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\mas_talker_csv.csv";       
            }
            
            using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(filepath))
            {

                //outfile.Write("UPC/EAN,ASIN,Brand,Model,MPN,Inventory Number,Product Name,Product Price,Minimum Price,Cost,Shipping Price,In Stock,Product URL, Image URL, Label, str_search_keywords,\r\n");
                foreach (ColumnHeader col in listView1.Columns)
                {
                    outfile.Write(col.Text + ",");
                }
                outfile.Write("\r\n");
                int counter = 0;
                foreach (ListViewItem lvi in listView1.Items)
                {
                    counter++;
                    statusLbl.Text = "Status: Preparing to write to file #" + counter.ToString() + "/" + listView1.Items.Count.ToString();
                    statusLbl.Refresh();
                    foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                    {                        
                        outfile.Write("\"" + lvsi.Text.Replace("$", "").Replace("\"", "\"\"") + "\",");
                    }                    
                    outfile.Write("\r\n");
                }
                outfile.Close();
            }

            statusLbl.Text = "Status: Finished writing file.";
            statusLbl.Refresh();
            MessageBox.Show("File saved to " + filepath);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WriteLVItoCSV("");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            WriteLVIToExcel(listView1, false);
        }
    }
}
