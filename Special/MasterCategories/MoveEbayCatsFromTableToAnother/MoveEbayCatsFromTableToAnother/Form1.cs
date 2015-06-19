using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MoveEbayCatsFromTableToAnother
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ResetLVI();

            LoadDataToNewTable();
        }

        private void LoadDataToNewTable()
        {
            List<string> data = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID,CategoryName FROM EBayHomeGardenConversionID");
            if (data.Count > 0)
            {
                foreach (string d in data)
                {
                    ListViewItem tmpLVI = new ListViewItem();
                    tmpLVI.Text = d.Split('|')[0];
                    tmpLVI.SubItems.Add(d.Split('|')[1]);
                    mainLVI.Items.Add(tmpLVI);
                }

                DialogResult res = MessageBox.Show("Are you sure you want to write to the database?", "Warning", MessageBoxButtons.YesNo);
                if (res== System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (ListViewItem lvi in mainLVI.Items)
                    {
                        string tmpID = lvi.Text;
                        string tmpName = lvi.SubItems[1].Text;
                        Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayCategories (ID,Name) VALUES(" + tmpID + ",'" + tmpName + "')");
                    }
                }


            }
        }
        private void ResetLVI()
        {
            mainLVI.Items.Clear();
            mainLVI.Columns.Clear();
            mainLVI.Clear();
            mainLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader nameCol = new ColumnHeader();
            idCol.Text = "ID";
            nameCol.Text = "Name";
            idCol.Width = mainLVI.Width / 10;
            nameCol.Width = mainLVI.Width - (mainLVI.Width / 10);
            mainLVI.Columns.Add(idCol);
            mainLVI.Columns.Add(nameCol);
        }
    }
}
