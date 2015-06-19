using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RenamingSearsCategories
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> allRows = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM SearsCategories");

            foreach (string row in allRows)
            {
                string id = row.Split('|')[0];
                List<string> rowData = Connectivity.SQLCalls.sqlQuery("SELECT Name FROM SearsCategories WHERE ID=" + id);
                string Name = rowData[0];
                Name = Name.Replace("|", " > ");
                Name = Name.Substring(0, Name.Length - 3).Trim();
                Connectivity.SQLCalls.sqlQuery("UPDATE SearsCategories SET Name='" + Name + "' WHERE ID=" + id);


            }


        }
    }
}
