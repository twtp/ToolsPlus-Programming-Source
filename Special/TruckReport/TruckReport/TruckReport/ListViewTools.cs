using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace ReportGenerator
{
    public static class ListViewTools
    {
        //API CALLS
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out IntPtr ProcessId);
        private static IntPtr GetWindowThreadProcessId(IntPtr hWnd)
        {
            IntPtr processId;
            IntPtr returnResult = GetWindowThreadProcessId(hWnd, out processId);

            return processId;
        }

        //SAVE AS
        public static void LVI_to_TXT(string filepath, ListView listview, char horizontalTableLine='=', char verticalTableLine='=')
        {
            string finalTable = "";
            string tableRow = "";
            string tableColumns = "";
            List<string> tableLines = new List<string>();
            int[] columns = new int[listview.Columns.Count];
            int tableLength = 0;
            for (int x = 0; x < listview.Columns.Count; x++)
            {
                columns[x] = listview.Columns[x].Text.Length;
            }
            for (int y = 0; y < listview.Items.Count; y++)
            {
                for (int z = 0; z < listview.Items[y].SubItems.Count; z++)
                {
                    if (listview.Items[y].SubItems[z].Text.Length > columns[z])
                    {
                        columns[z] = listview.Items[y].SubItems[z].Text.Length;
                    }
                }
            }

            for (int x = 0; x < columns.GetLength(0); x++)
            {
                tableLength += columns[x] + 1; // +1 for the '=' before the column
            }
            tableLength++; //+1 for the end '='
            tableRow = tableRow.PadRight(tableLength,horizontalTableLine) + "\r\n";

            for (int x = 0; x < columns.GetLength(0); x++)
            {
                tableColumns += verticalTableLine;
                tableColumns += listview.Columns[x].Text.PadRight(columns[x]);
            }
            tableColumns += verticalTableLine + "\r\n";

            finalTable += tableRow + tableColumns + tableRow;

            for (int x = 0; x < listview.Items.Count; x++)
            {
                string newLine = "";
                for (int y = 0; y < listview.Items[x].SubItems.Count; y++)
                {
                    newLine += verticalTableLine + listview.Items[x].SubItems[y].Text.PadRight(columns[y]);
                }
                newLine += verticalTableLine + "\r\n";
                tableLines.Add(newLine);
            }

            foreach (string line in tableLines)
            {
                finalTable += line;
            }
            finalTable += tableRow;

            try
            {
                System.IO.File.WriteAllText(filepath, finalTable);
            }
            catch
            {
                MessageBox.Show("Error trying to save. Make sure this file isn't already open or try saving to a different filename/location.");
            }


        }
        public static void LVI_to_CSV(string filepath, ListView listview)
        {
            string Columns = "";
            foreach (ColumnHeader col in listview.Columns)
            {
                Columns += col.Text + ",";
            }
            Columns += "\r\n";
            List<string> Rows = new List<string>();
            foreach (ListViewItem lvi in listview.Items)
            {
                string tmpRow = "";
                foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                {
                    tmpRow += lvsi.Text + ",";
                }
                tmpRow += "\r\n";
                Rows.Add(tmpRow);
            }
            try
            {
                using (System.IO.StreamWriter writer = new System.IO.StreamWriter(filepath))
                {
                    writer.Write(Columns);
                    foreach (string row in Rows)
                    {
                        writer.Write(row);
                    }
                    writer.Close();
                }
            }
            catch
            {
                MessageBox.Show("Error trying to save. Make sure this file isn't already open or try saving to a different filename/location.");
            }
        }
        public static void LVI_to_XLS(string filepath, ListView listview)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            ws.Columns.EntireColumn.ColumnWidth = 20;

            int i = 1;
            int i2 = 3;

            i2 = 2;
      

            foreach (ListViewItem lvi in listview.Items)
            {
                i = 1;
                List<string> columnList = new List<string>();
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    columnList.Add(lvs.Text);
                }
                for (int x = 0; x < columnList.Count; x++)
                {
                    foreach (ColumnHeader ch in listview.Columns)
                    {
                        if (ch.DisplayIndex == x)
                        {
                            //this should really be changed so we dont overwrite a header a million times over...
                            int y = 0;
                            //if (showTitle == true)
                            //{
                            //    y = 2;
                            //}
                            //else
                            //{
                            y = 1;
                            //}

                            ws.Cells[y, i].EntireRow.Font.Bold = true;
                            ws.Cells[y, i] = ch.Text;
                            ws.Cells[i2, i] = columnList[ch.Index];
                            i++;
                        }
                    }
                }
                i2++;
            }           
            
            wb.SaveAs(filepath);



            IntPtr excelPointer = new IntPtr();
            GetWindowThreadProcessId((IntPtr)app.Hwnd, out excelPointer);
            System.Diagnostics.Process.GetProcessById((int)excelPointer).Kill();
            System.Diagnostics.Process.GetProcessById((int)excelPointer).WaitForExit();
        }
        public static void LVI_to_XML(string filepath, ListView listview)
        {
            string XMLFinal = "";

            XMLFinal += "<?xml version=\"1.0\" encoding=\"UTF-8\">\r\n";
            XMLFinal += " <ReportTable>\r\n";
            List<string> columnStartTag = new List<string>();
            List<string> columnEndTag = new List<string>();
            foreach (ColumnHeader col in listview.Columns)
            {
                columnStartTag.Add("<" + col.Text + ">");
                columnEndTag.Add("</" + col.Text + ">");
            }

            for (int x = 0; x < listview.Items.Count; x++)
            {
                XMLFinal += "  " + "<TableRow>\r\n";
                for (int y = 0; y < listview.Items[x].SubItems.Count; y++)
                {
                    XMLFinal += "   " + columnStartTag[y] + listview.Items[x].SubItems[y].Text + columnEndTag[y] + "\r\n";
                }
                XMLFinal += "  " + "</TableRow>\r\n";
            }

            XMLFinal += " </ReportTable>\r\n";

            System.IO.File.WriteAllText(filepath, XMLFinal);

        }
        public static void LVI_to_HTML(string filepath, ListView listview)
        {
            string HTMLoutput = "";
            HTMLoutput += "<html>\r\n<body>\r\n<table>\r\n\r\n<tr>\r\n";
            foreach (ColumnHeader col in listview.Columns)
            {
                HTMLoutput += "<td>" + col.Text + "</td>\r\n";
            }
            HTMLoutput += "</tr>\r\n";

            for (int x = 0; x < listview.Items.Count; x++)
            {
                HTMLoutput += "<tr>\r\n";
                for (int y = 0; y < listview.Items[x].SubItems.Count; y++)
                {
                    HTMLoutput += "<td>" + listview.Items[x].SubItems[y].Text + "</td>\r\n";
                }
                HTMLoutput += "</tr>\r\n";
            }
            HTMLoutput += "</table>\r\n";

            HTMLoutput += "</body>\r\n</html>";
            System.IO.File.WriteAllText(filepath, HTMLoutput);

        }
        public static void LVI_to_JSON(string filepath, ListView listview)
        {
            string jo = "{\"cols\":[";

            foreach (ColumnHeader ch in listview.Columns)
            {
                jo += "{\"name\":\"" + ch.Text + "\"},";
            }
            jo = jo.TrimEnd(',');
            jo += "],\"rows\":[";
            foreach (ListViewItem lvi in listview.Items)
            {
                jo += "[";
                foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                {
                    jo += "\"" + lvsi.Text + "\",";
                }
                jo = jo.TrimEnd(',');
                jo += "],";
            }
            jo = jo.TrimEnd(',');
            jo += "]}";

            System.IO.File.WriteAllText(filepath, jo);

        }
        public static void LVI_to_PDF(string filepath, ListView listview)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            ws.Columns.EntireColumn.ColumnWidth = 20;

            int i = 1;
            int i2 = 3;

            i2 = 2;
    

            foreach (ListViewItem lvi in listview.Items)
            {
                i = 1;
                List<string> columnList = new List<string>();
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    columnList.Add(lvs.Text);
                }
                for (int x = 0; x < columnList.Count; x++)
                {
                    foreach (ColumnHeader ch in listview.Columns)
                    {
                        if (ch.DisplayIndex == x)
                        {
                            //this should really be changed so we dont overwrite a header a million times over...
                            int y = 0;
                            //if (showTitle == true)
                            //{
                            //    y = 2;
                            //}
                            //else
                            //{
                            y = 1;
                            //}

                            ws.Cells[y, i].EntireRow.Font.Bold = true;
                            ws.Cells[y, i] = ch.Text;
                            ws.Cells[i2, i] = columnList[ch.Index];
                            i++;
                        }
                    }
                }
                i2++;
            }
            wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, filepath);



            IntPtr excelPointer = new IntPtr();
            GetWindowThreadProcessId((IntPtr)app.Hwnd, out excelPointer);
            System.Diagnostics.Process.GetProcessById((int)excelPointer).Kill();
            System.Diagnostics.Process.GetProcessById((int)excelPointer).WaitForExit();
        }
        public static void LVI_to_DOCX_SLOW(string filepath, ListView listview)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = false;
            app.ScreenUpdating = false;

            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();
            Microsoft.Office.Interop.Word.Paragraph para1 = doc.Content.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Table dataTable = doc.Tables.Add(para1.Range, listview.Items.Count + 1, listview.Columns.Count);
            dataTable.Borders.Enable = 1;

            for (int a = 0; a < listview.Columns.Count; a++)
            {
                dataTable.Rows[1].Cells[a + 1].Range.Bold = 1;
                dataTable.Rows[1].Cells[a + 1].Range.Text = listview.Columns[a].Text;
            }

            for (int x = 0; x < listview.Items.Count; x++)
            {
                for (int y = 0; y < listview.Columns.Count; y++)
                {
                    dataTable.Rows[x + 2].Cells[y + 1].Range.Text = listview.Items[x].SubItems[y].Text;
                }
            }
            doc.SaveAs(filepath);
            doc.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(para1);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            app = null;
            doc = null;
            para1 = null;

            //IntPtr wordPointer = new IntPtr();
            //GetWindowThreadProcessId((IntPtr).Hwnd, out wordPointer);
            //System.Diagnostics.Process.GetProcessById((int)wordPointer).Kill();
            //System.Diagnostics.Process.GetProcessById((int)wordPointer).WaitForExit();

        }
        public static void LVI_to_DOCX(string filepath, ListView listview)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = false;
            app.ScreenUpdating = false;
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();
            Microsoft.Office.Interop.Word.Paragraph para1 = doc.Content.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Table dataTable = doc.Tables.Add(para1.Range, listview.Items.Count + 1, listview.Columns.Count);
            dataTable.Borders.Enable = 1;
            
            var data = new object[listview.Items.Count + 1, listview.Columns.Count];
            for (var row = 1; row <= listview.Items.Count + 1; row++)
            {
                for (var column = 1; column <= listview.Columns.Count; column++)
                {
                    data[row, column] = listview.Items[row].SubItems[column].Text;
                }
            }
            
           // dataTable.Rows = (Microsoft.Office.Interop.Word.Rows)data;


            for (int a = 0; a < listview.Columns.Count; a++)
            {
                dataTable.Rows[1].Cells[a + 1].Range.Bold = 1;
                dataTable.Rows[1].Cells[a + 1].Range.Text = listview.Columns[a].Text;
            }

            for (int x = 0; x < listview.Items.Count; x++)
            {
                for (int y = 0; y < listview.Columns.Count; y++)
                {
                    dataTable.Rows[x + 2].Cells[y + 1].Range.Text = listview.Items[x].SubItems[y].Text;
                }
            }
            doc.SaveAs(filepath);
            doc.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(para1);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            app = null;
            doc = null;
            para1 = null;

            //IntPtr wordPointer = new IntPtr();
            //GetWindowThreadProcessId((IntPtr).Hwnd, out wordPointer);
            //System.Diagnostics.Process.GetProcessById((int)wordPointer).Kill();
            //System.Diagnostics.Process.GetProcessById((int)wordPointer).WaitForExit();

        }

        //OPEN WITH
        public static void LVI_to_WORD(ListView listview)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;

            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();
            Microsoft.Office.Interop.Word.Paragraph para1 = doc.Content.Paragraphs.Add();                
            Microsoft.Office.Interop.Word.Table dataTable = doc.Tables.Add(para1.Range, listview.Items.Count + 1, listview.Columns.Count);
            dataTable.Borders.Enable = 1;

            for (int a = 0; a < listview.Columns.Count; a++)
            {
                dataTable.Rows[1].Cells[a + 1].Range.Bold = 1;
                dataTable.Rows[1].Cells[a+1].Range.Text = listview.Columns[a].Text;
            }

            for (int x = 0; x < listview.Items.Count; x++)
            {
                for (int y = 0; y < listview.Columns.Count; y++)
                {
                    dataTable.Rows[x+2].Cells[y+1].Range.Text = listview.Items[x].SubItems[y].Text;
                }
            }


        }
        public static void LVI_to_OUTLOOK(ListView listview, string subject="Generated Report")
        {
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem newMail = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            newMail.Subject = subject;

            string HTMLoutput = "";
            HTMLoutput += "<html>\r\n<body>\r\n<table>\r\n\r\n<tr>\r\n";
            foreach (ColumnHeader col in listview.Columns)
            {
                HTMLoutput += "<td>" + col.Text + "</td>\r\n";
            }
            HTMLoutput += "</tr>\r\n";

            for (int x = 0; x < listview.Items.Count; x++)
            {
                HTMLoutput += "<tr>\r\n";
                for (int y = 0; y < listview.Items[x].SubItems.Count; y++)
                {
                    HTMLoutput += "<td>" + listview.Items[x].SubItems[y].Text + "</td>\r\n";
                }
                HTMLoutput += "</tr>\r\n";
            }
            HTMLoutput += "</table>\r\n";

            HTMLoutput += "</body>\r\n</html>";
            newMail.HTMLBody += HTMLoutput;
            newMail.Display();
           // newMail = 
        }
        public static void LVI_to_EXCEL(ListView listview)
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

            foreach (ListViewItem lvi in listview.Items)
            {
                i = 1;
                List<string> columnList = new List<string>();
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    columnList.Add(lvs.Text);
                }
                for (int x = 0; x < columnList.Count; x++)
                {
                    foreach (ColumnHeader ch in listview.Columns)
                    {
                        if (ch.DisplayIndex == x)
                        {
                            //this should really be changed so we dont overwrite a header a million times over...
                            int y = 0;
                            //if (showTitle == true)
                            //{
                            //    y = 2;
                            //}
                            //else
                            //{
                            y = 1;
                            //}

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

        //INTERNALS THAT DONT WORK YET...
        public static ListView LVI_to_LVI(ListView listViewData)
        {
            ListView destinationView = new ListView();
            destinationView = listViewData;
            
                    destinationView.Items.Clear();
                    destinationView.Columns.Clear();
                    destinationView.Clear();

                    destinationView.FullRowSelect = listViewData.FullRowSelect;
                    destinationView.GridLines = listViewData.GridLines;
                    destinationView.View = listViewData.View;

                    destinationView.Refresh();

                    foreach (ColumnHeader col in listViewData.Columns)
                    {
                        destinationView.Columns.Add((ColumnHeader)col.Clone());
                    }
                    foreach (ListViewItem lvi in listViewData.Items)
                    {
                        destinationView.Items.Add((ListViewItem)lvi.Clone());
                    }

                    destinationView.Refresh();

                    return destinationView;
            
        }


        //TO LVI
        private struct MASColumnHeader
        {
            public string Name;
            public int TypeID;
            public int Size;
        }
        public static ListView MAS_to_LVI(string responseData)
        {
            System.Windows.Forms.ListView outputLV = new System.Windows.Forms.ListView();
            outputLV.GridLines = true;
            outputLV.FullRowSelect = true;
            outputLV.View = System.Windows.Forms.View.Details;

            string ColumnData = "";
            string RowData = "";

            try
            {
                ColumnData = responseData.Split('[')[1].Split(']')[0];
                RowData = responseData.Split(new string[] { "rows\":[" }, StringSplitOptions.None)[1].Split(new string[] { "]}" }, StringSplitOptions.None)[0];
            }
            catch { MessageBox.Show("The last query errored as it returned nothing."); return new ListView(); }
            string[] Columns = ColumnData.Split(new string[] { "{\"name\":\"" }, StringSplitOptions.RemoveEmptyEntries);
            for (int x = 0; x < Columns.GetLength(0); x++)
            {
                Columns[x] = Columns[x].Split('}')[0];
            }

            string[] Rows = RowData.Split(new string[] { "[" }, StringSplitOptions.RemoveEmptyEntries);
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
                foreach (string fld in fieldList)
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

            return outputLV;
        }
        public static ListView TSV_to_LVI(string tsvFile)
        {
            ListView newView = new ListView();
            newView.Items.Clear();
            newView.Columns.Clear();
            newView.Clear();
            newView.FullRowSelect = true;
            newView.GridLines = true;
            newView.View = View.Details;
            
            
            foreach (string Row in tsvFile.Split('\n'))
            {
                ListViewItem newItem = new ListViewItem();
                foreach (string Col in Row.Split('\t'))
                {
                    ListViewItem.ListViewSubItem newSubItem = new ListViewItem.ListViewSubItem();
                    newSubItem.Text = Col;
                    newItem.SubItems.Add(newSubItem);
                }
                newView.Items.Add(newItem);
            }
            return newView;
        }

    }
}
