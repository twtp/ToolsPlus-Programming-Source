using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlgorithmGenerator
{
    public struct Algorithm
    {
        public string AlgorithmName;
        public decimal PercentModifier;
        public List<AlgorithmDetail> AlgorithmChart;
    }
    public struct AlgorithmDetail
    {
        public decimal DollarAmount;
        public decimal MaxPercentAbove;
        public decimal MaxPercentBelow;
    }
    public partial class Form1 : Form
    {
        Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        Algorithm loadAlgorithm = new Algorithm();

        public object misValue = System.Reflection.Missing.Value;



        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            if (xlApp == null)
            {
                MessageBox.Show("This computer doesn't have Excel installed or has too old of a version.");
                return;
            }
            else
            {
                
                xlWorkbook = xlApp.Workbooks.Add(misValue);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
                Excel.Range er = xlWorksheet.get_Range("A:A", misValue);
                er.EntireColumn.ColumnWidth = 22;
                er = xlWorksheet.get_Range("B1","B1");
                er.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightBlue);
                er.Cells.Borders.LineStyle = BorderStyle.FixedSingle;
                er = xlWorksheet.get_Range("B2", "B2");
                er.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightSlateGray);
                er.Cells.Borders.LineStyle = BorderStyle.FixedSingle;
                er = xlWorksheet.get_Range("B:B", misValue);
                er.EntireColumn.ColumnWidth = 18;
                er = xlWorksheet.get_Range("C:C", misValue);
                er.EntireColumn.ColumnWidth = 18;
                xlWorksheet.Cells[1, 1] = "Algorithm Name:";
                xlWorksheet.Cells[2, 1] = "Difference Percentage:";                
                xlWorksheet.Cells[3, 1] = "Static Sell Price";
                xlWorksheet.Cells[3, 2] = "% Over Sell Price";
                xlWorksheet.Cells[3, 3] = "% Under Sell Price";

                for (int x = 1; x < 501; x++)
                {
                    xlWorksheet.Cells[x + 3, 1] = x.ToString();                   
                }
                er = xlWorksheet.get_Range("B4", "B503");
                er.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
                er.Cells.Borders.LineStyle = BorderStyle.FixedSingle;
                er = xlWorksheet.get_Range("C4", "C503");
                er.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.LightCoral);
                er.Cells.Borders.LineStyle = BorderStyle.FixedSingle;
                xlApp.Visible = true;
                
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            button2.Enabled = false;
            xlApp.Visible = false;
            System.Threading.Thread.Sleep(2000);
            loadAlgorithm.AlgorithmChart = new List<AlgorithmDetail>();

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

            var aName = (Excel.Range)xlWorksheet.Cells[1, 2];
            var aPerc = (Excel.Range)xlWorksheet.Cells[2, 2];
            try
            {
                loadAlgorithm.AlgorithmName = aName.Text;
                loadAlgorithm.PercentModifier = decimal.Parse(aPerc.Text);
                for (int fu = 4; fu < 504; fu++)
                {
                    AlgorithmDetail newDetail = new AlgorithmDetail();

                    var dollarAmount = (Excel.Range)xlWorksheet.Cells[fu, 1];
                    newDetail.DollarAmount = decimal.Parse(dollarAmount.Text);                    
                    var maxOver = (Excel.Range)xlWorksheet.Cells[fu, 2];
                    newDetail.MaxPercentAbove = decimal.Parse(maxOver.Text);
                    var maxUnder = (Excel.Range)xlWorksheet.Cells[fu, 3];
                    newDetail.MaxPercentBelow = decimal.Parse(maxUnder.Text);
                    loadAlgorithm.AlgorithmChart.Add(newDetail);
                }
                InsertIntoDatabase(loadAlgorithm);
                MessageBox.Show("Great! All info has been parsed sucessfully!");
                MessageBox.Show("As a test for algorithm \"" + loadAlgorithm.AlgorithmName + "\"... if cost was $" + loadAlgorithm.AlgorithmChart[400].DollarAmount.ToString() + " then max percentage above would be " + loadAlgorithm.AlgorithmChart[400].MaxPercentAbove.ToString() + "% and the max percentage below would be " + loadAlgorithm.AlgorithmChart[400].MaxPercentBelow.ToString() + "%. Also, this algorithm prices the item at " + Math.Abs(loadAlgorithm.PercentModifier).ToString() + ((loadAlgorithm.PercentModifier >= 0) ? ("% above competitor price") : ("% below competitor price")));
            }
            catch (Exception wtf)
            {
                MessageBox.Show("It appears you have entered invalid data. See below error.\r\n\r\n" + wtf.Message);
            }
            button1.Enabled = true;
            button2.Enabled = true;
        }
        public void InsertIntoDatabase(Algorithm curAlgorithm)
        {
            Connectivity.SQLCalls.sqlQuery("INSERT INTO DynamicPricingAlgorithmHeader (AlgorithmName,DifferentialPercentage) VALUES('" + curAlgorithm.AlgorithmName + "'," + curAlgorithm.PercentModifier.ToString() + ")");
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM DynamicPricingAlgorithmHeader WHERE AlgorithmName='" + curAlgorithm.AlgorithmName + "'");
            if (results.Count > 0)
            {
                if (results.Count > 1)
                {
                    //error
                }
                else
                {
                    int algID = 0;
                    try
                    {
                        algID = int.Parse(results[0].Split('|')[0]);

                    }
                    catch
                    {
                        MessageBox.Show("The ID returned is bogus (or tom's code is...)");
                    }
                    foreach (AlgorithmDetail aDetail in curAlgorithm.AlgorithmChart)
                    {
                        Connectivity.SQLCalls.sqlQuery("INSERT INTO DynamicPricingAlgorithmLines (HeaderID,DollarAmount,MaxPercentAbove,MaxPercentBelow) VALUES(" + algID.ToString() + "," + aDetail.DollarAmount.ToString() + "," + aDetail.MaxPercentAbove.ToString() + "," + aDetail.MaxPercentBelow.ToString() + ")");

                    }
                }
            }
            else
            {
                //error
            }



        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
