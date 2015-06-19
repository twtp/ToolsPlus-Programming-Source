using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EBayItemSpecifics
{
    public partial class Form1 : Form
    {
        
        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        public List<System.Diagnostics.Process> excelProcesses = new List<System.Diagnostics.Process>();
        public List<System.Diagnostics.Process> currExcelProcesses = new List<System.Diagnostics.Process>();
        public System.Diagnostics.Process thisExcelProcess = new System.Diagnostics.Process();
        public List<string> UniqueSpecifics = new List<string>();

        public struct EBayCategorySpecifics
        {
            public string CategoryID;
            public string CategoryName;
            public List<EBayCategorySpecificsDetail> Specifics;
        }
        public struct EBayCategorySpecificsDetail
        {
            public string SpecificName;
            public List<string> SpecificValues;
        }

        public List<EBayCategorySpecifics> EBaySpecifics = new List<EBayCategorySpecifics>();

        public void getExcelProcesses(bool beforeLoad)
        {
            if (beforeLoad == true)
            {
                excelProcesses.Clear();
                System.Diagnostics.Process[] processesBefore = System.Diagnostics.Process.GetProcessesByName("excel");
                foreach (System.Diagnostics.Process proc in processesBefore)
                {
                    excelProcesses.Add(proc);                   
                }                
            }
            else
            {
                currExcelProcesses.Clear();
                System.Diagnostics.Process[] processesAfter = System.Diagnostics.Process.GetProcessesByName("excel");
                foreach (System.Diagnostics.Process proc in processesAfter)
                {
                    currExcelProcesses.Add(proc);
                }

                foreach (System.Diagnostics.Process pB in excelProcesses)
                {
                    bool flagged = false;
                    foreach (System.Diagnostics.Process pA in currExcelProcesses)
                    {
                        if (pA == pB)
                        {
                            flagged = true;
                            break;                            
                        }
                    }
                    if (flagged == false)
                    {
                        thisExcelProcess = pB;
                    }
                }

            }

        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            getExcelProcesses(true);
            setupLVI();
        }
        private void setupLVI()
        {
            ebayLVI.Items.Clear();
            ebayLVI.Columns.Clear();
            ebayLVI.Clear();
            ebayLVI.View = View.Details;
            ebayLVI.FullRowSelect = true;
            ebayLVI.GridLines = true;
            ColumnHeader catID = new ColumnHeader();
            ColumnHeader catName = new ColumnHeader();
            ColumnHeader specificName = new ColumnHeader();
            ColumnHeader specificValue = new ColumnHeader();
            catID.Text = "Category ID";
            catName.Text = "Category Name";
            specificName.Text = "Specifics Name";
            specificValue.Text = "Possible Values";
            catID.Width = ebayLVI.Width / 10;
            catName.Width = ebayLVI.Width / 2;
            specificName.Width = ebayLVI.Width / 5;
            specificValue.Width = ebayLVI.Width / 5;
            ebayLVI.Columns.Add(catID);
            ebayLVI.Columns.Add(catName);
            ebayLVI.Columns.Add(specificName);
            ebayLVI.Columns.Add(specificValue);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            setupLVI();
            var nullVal = System.Type.Missing;
            if (xlApp == null)
            {
                MessageBox.Show("This computer doesn't have Excel installed or has too old of a version.");
                return;
            }
            else
            {
                 xlWorkbook = xlApp.Workbooks.Open(@"C:\ebayspecifics.txt", nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal);
                //xlWorkbook = xlApp.Workbooks.Open(@"C:\test.txt", nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal, nullVal);
                
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                //Excel.Range er = xlWorksheet.get_Range("B:B", nullVal);
                //er.Cells[1,1] = 

                getExcelProcesses(false);

                var cellRange = xlWorksheet.get_Range("A1", System.Type.Missing).CurrentRegion;

                int CellEnd = cellRange.End[Excel.XlDirection.xlDown].Row;

               
                long counter = 1;
                while (((Excel.Range)xlWorksheet.Cells[counter, 1]).Text != "")
                {
                    counter++;

                    statusLbl.Text = "Status: Working on Category #" + counter.ToString() + " of " + CellEnd.ToString();
                    statusLbl.Refresh();

                        string catID = ((Excel.Range)xlWorksheet.Cells[counter, 1]).Text;
                        bool flagged = false;
                        //foreach (EBayCategorySpecifics ecs in EBaySpecifics)
                        for (int x = 0; x < EBaySpecifics.Count; x++)
                        {
                            

                            if (EBaySpecifics[x].CategoryID == catID)
                            {
                                //add info to this element
                                EBayCategorySpecificsDetail ecds = new EBayCategorySpecificsDetail();
                                ecds.SpecificValues = new List<string>();
                                ecds.SpecificName = ((Excel.Range)xlWorksheet.Cells[counter, 3]).Text;
                                //EBaySpecifics[x].Specifics.Add(ecds);
                                bool specificFlag = false;
                                //foreach (EBayCategorySpecificsDetail det in EBaySpecifics[x].Specifics)
                                for (int y = 0; y < EBaySpecifics[x].Specifics.Count; y++)
                                {
                                    if (EBaySpecifics[x].Specifics[y].SpecificName == ecds.SpecificName)
                                    {
                                        //add info to this element
                                        EBaySpecifics[x].Specifics[y].SpecificValues.Add(((Excel.Range)xlWorksheet.Cells[counter, 4]).Text);                                        
                                        //ecs.Specifics.Add(det);
                                        specificFlag = true;
                                    }
                                }
                                if (specificFlag == false)
                                {
                                    //add info to new element...
                                    EBayCategorySpecificsDetail ecds2 = new EBayCategorySpecificsDetail();
                                    ecds2.SpecificName = ((Excel.Range)xlWorksheet.Cells[counter, 3]).Text;
                                    ecds2.SpecificValues = new List<string>();
                                    ecds2.SpecificValues.Add(((Excel.Range)xlWorksheet.Cells[counter, 4]).Text);
                                    //ecs.Specifics.Add(ecds2);
                                    EBaySpecifics[x].Specifics.Add(ecds2);
                                }


                                flagged = true;
                            }
                        }
                        if (flagged == false)
                        {
                            //add info to new element
                            EBayCategorySpecifics newCatSpec = new EBayCategorySpecifics();
                            newCatSpec.Specifics = new List<EBayCategorySpecificsDetail>();
                            EBayCategorySpecificsDetail newCatSpecDet = new EBayCategorySpecificsDetail();
                            newCatSpecDet.SpecificValues = new List<string>();
                            newCatSpec.CategoryID = catID;
                            newCatSpec.CategoryName = ((Excel.Range)xlWorksheet.Cells[counter, 2]).Text;
                            newCatSpecDet.SpecificName = ((Excel.Range)xlWorksheet.Cells[counter, 3]).Text;
                            newCatSpecDet.SpecificValues.Add(((Excel.Range)xlWorksheet.Cells[counter, 4]).Text);
                            newCatSpec.Specifics.Add(newCatSpecDet);
                            EBaySpecifics.Add(newCatSpec);
                        }
               }
               
                    
                        //hit eof... kill it!
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        xlWorkbook.Close(false, nullVal, nullVal);
                        xlApp.Application.Quit();
                        xlApp.Quit();
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook) != 0) { };
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet) != 0) { };
                        while (System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) != 0) { };
                        xlWorkbook = null;
                        xlWorksheet = null;
                        xlApp = null;

                        thisExcelProcess.Kill();


                        MessageBox.Show("alright, should be good...");
                


                //MessageBox.Show(((Excel.Range)xlWorksheet.Cells[1, 1]).Text);
                //        DisplayInLVI();
                        GatherUniqueSpecifics();
            }
            
        }
        public void GatherUniqueSpecifics()
        {
            for (int x = 0; x < EBaySpecifics.Count; x++)
            {
                statusLbl.Text = "Status: Gathering Specifics " + x.ToString() + " of " + EBaySpecifics.Count.ToString();
                statusLbl.Refresh();
                for (int y = 0; y < EBaySpecifics[x].Specifics.Count; y++)
                {
                    bool flagged = false;
                    for (int z = 0; z < UniqueSpecifics.Count; z++)
                    {
                        if (EBaySpecifics[x].Specifics[y].SpecificName == UniqueSpecifics[z])
                        {
                            flagged = true;
                            break;
                        }
                    }
                    if (flagged == false)
                    {
                        UniqueSpecifics.Add(EBaySpecifics[x].Specifics[y].SpecificName);
                    }
                }
            }
            textBox1.Text = "";
            foreach (string itm in UniqueSpecifics)
            {
                textBox1.Text += itm + "\r\n";
            }

        }
        public void DisplayInLVI()
        {
            setupLVI();
            foreach (EBayCategorySpecifics ecs in EBaySpecifics)
            {
                ListViewItem MainItem = new ListViewItem();
                
                MainItem.Text = ecs.CategoryID;
                MainItem.SubItems.Add(ecs.CategoryName);
                ebayLVI.Items.Add(MainItem);
                foreach (EBayCategorySpecificsDetail ecds in ecs.Specifics)
                {
                    ListViewItem Names = new ListViewItem();
                    Names.Text = "";
                    Names.SubItems.Add("");
                    Names.SubItems.Add(ecds.SpecificName);
                    ebayLVI.Items.Add(Names);
                    foreach (string eVal in ecds.SpecificValues)
                    {
                        ListViewItem MainValues = new ListViewItem();
                        MainValues.Text = "";
                        MainValues.SubItems.Add("");
                        MainValues.SubItems.Add("");
                        MainValues.SubItems.Add(eVal);
                        ebayLVI.Items.Add(MainValues);
                    }
                }
                
            }
        }
    }
}
