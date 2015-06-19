using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MassiveWarehouseVarianceFixer
{
    public partial class Form1 : Form
    {
        public static string ImportFilePath = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void browseBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select a file to import...";
            openFileDialog1.Filter = "*.* (All Files)|*.*";
            DialogResult res = openFileDialog1.ShowDialog();
            if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
            {
                ImportFilePath = openFileDialog1.FileName;
                filePathLbl.Text = ImportFilePath;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetupLV();
        }
        private void SetupLV()
        {
            varianceLV.Items.Clear();
            varianceLV.Columns.Clear();
            varianceLV.Clear();
            varianceLV.View = View.Details;
            varianceLV.GridLines = true;
            varianceLV.FullRowSelect = true;

            ColumnHeader itmCol = new ColumnHeader();
            ColumnHeader masQtyCol = new ColumnHeader();
            ColumnHeader whseQtyCol= new ColumnHeader();
            ColumnHeader varianceCol = new ColumnHeader();
            ColumnHeader processCol = new ColumnHeader();
            ColumnHeader resultsCol = new ColumnHeader();

            itmCol.Text = "Item Number";
            itmCol.Width = 150;
            masQtyCol.Text = "MAS Qty";
            masQtyCol.Width = 70;
            whseQtyCol.Text = "WHSE Qty";
            whseQtyCol.Width = 70;
            varianceCol.Text = "Variance";
            varianceCol.Width = 70;
            processCol.Text = "Process";
            processCol.Width = 70;
            resultsCol.Text = "Result";
            resultsCol.Width = 200;

            varianceLV.Columns.Add(itmCol);
            varianceLV.Columns.Add(masQtyCol);
            varianceLV.Columns.Add(whseQtyCol);
            varianceLV.Columns.Add(varianceCol);
            varianceLV.Columns.Add(processCol);
            varianceLV.Columns.Add(resultsCol);
        }
        private void LoadFile()
        {
            SetupLV();
            string VarianceFile = System.IO.File.ReadAllText(filePathLbl.Text);

            foreach (string line in VarianceFile.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string ItemNumber = line.Split('\t')[0];
                int MASQty = int.Parse(line.Split('\t')[1]);
                int WHSEQty = int.Parse(line.Split('\t')[2]);
                int Variance = int.Parse(line.Split('\t')[3]);

                ListViewItem lvi = new ListViewItem();
                lvi.Text = ItemNumber;
                lvi.SubItems.Add(MASQty.ToString());
                lvi.SubItems.Add(WHSEQty.ToString());
                lvi.SubItems.Add(Variance.ToString());
                lvi.SubItems.Add((Variance > 0 ? "YES" : ""));
                lvi.SubItems.Add("");
                varianceLV.Items.Add(lvi);
            }

        }

        private void RunFixEngine()
        {
            //This is the fix "algorithm"
            // 1.) Load list view line
            // 2.) Take item number and check qty in shipping staging.
            // 3.) If qty in shipping staging is equal to variance qty then delete that db row
            // 4.) If qty in shipping staging is less than variance qty then do nothing and inform user of anomaly.
            // 5.) If qty in shipping staging is greater than variance qty then prompt user asking to update record in db
            // 6.) If user input is yes then update row
            // 7.) If user input is no then skip item but report after running

            int counter = 0;
            foreach (ListViewItem lvi in varianceLV.Items)
            {
                string ItemNumber = lvi.Text;
                int MASQty = int.Parse(lvi.SubItems[1].Text);
                int WHSEQty = int.Parse(lvi.SubItems[2].Text);
                int Variance = int.Parse(lvi.SubItems[3].Text);

                if (lvi.SubItems[4].Text == "YES")
                {
                    Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT LocationContents.Quantity,LocationContents.ComponentID FROM LocationContents INNER JOIN InventoryComponentMap ON InventoryComponentMap.ComponentID=LocationContents.ComponentID WHERE LocationContents.LocationID = 10488 AND InventoryComponentMap.ItemNumber='" + ItemNumber + "'");
                    if (qr.Rows.Count > 0)
                    {
                        if (qr.Rows.Count == 1)
                        {
                            int dbQty = int.Parse(qr.Rows[0].Split('|')[0]);
                            string ComponentID = qr.Rows[0].Split('|')[1];
                            if (dbQty == Variance)
                            {
                                Connectivity.SQLCalls.sqlQuery("DELETE FROM LocationContents WHERE LocationID=10488 AND ComponentID=" + ComponentID + " AND Quantity=" + Variance.ToString());
                                varianceLV.Items[counter].SubItems[5].Text = "Processed OK.";
                            }
                            else if (dbQty < Variance)
                            {
                                varianceLV.Items[counter].SubItems[5].Text = "Qty in location is less than variance. Aborting.";
                            }
                            else
                            {
                                DialogResult res = MessageBox.Show("Item " + ItemNumber + " has a variance of " + Variance.ToString() + ", but shipping staging has " + dbQty.ToString() + ". Would you like to update this line accordingly? The new qty in shipping staging will be " + (dbQty - Variance).ToString(), "Double Check", MessageBoxButtons.YesNo);
                                if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
                                {
                                    Connectivity.SQLCalls.sqlQuery("UPDATE LocationContents SET Quantity=" + (dbQty - Variance).ToString() + " WHERE LocationID=10488 AND ComponentID=" + ComponentID + " AND Quantity=" + dbQty.ToString());
                                    varianceLV.Items[counter].SubItems[5].Text = "Updated shipping staging to reflect the correct qty.";
                                }
                                else
                                {
                                    varianceLV.Items[counter].SubItems[5].Text = "Aborted By User. DB has more qty than variance.";
                                }
                            }
                        }
                        else
                        {
                            varianceLV.Items[counter].SubItems[5].Text = "Multiple rows in shipping staging? Aborting.";
                        }
                    }
                    else
                    {
                        varianceLV.Items[counter].SubItems[5].Text = "Item not found in shipping staging. Aborted.";
                    }
                }
                counter++;
            }



        }

        private void startBtn_Click(object sender, EventArgs e)
        {
            LoadFile();
            
        }

        private void runBtn_Click(object sender, EventArgs e)
        {
            RunFixEngine();
        }

    }
}
