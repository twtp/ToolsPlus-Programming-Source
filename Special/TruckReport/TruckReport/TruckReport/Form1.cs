using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TruckReport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public struct TruckReportItem
        {
            public string ItemNumber;
            public int Quantity;
            public DateTime LastInventoriedDate;
        }

        public List<TruckReportItem> TruckReport = new List<TruckReportItem>();
        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> truckItems = Connectivity.SQLCalls.sqlQuery("SELECT InventoryComponentMap.ItemNumber,LocationContents.Quantity,LastInventoriedDate FROM LocationContents INNER JOIN InventoryComponentMap ON LocationContents.ComponentID = InventoryComponentMap.ComponentID INNER JOIN InventoryMaster ON InventoryMaster.ItemNumber=InventoryComponentMap.ItemNumber WHERE LocationID=10498 AND InventoryComponentMap.ItemNumber NOT LIKE 'ZZZ%' AND IsMasKit=0");
            foreach (string itm in truckItems)
            {
                string itemNumber = itm.Split('|')[0];
                int quantity = int.Parse(itm.Split('|')[1]);
                DateTime lastInventoriedDate = DateTime.Parse(itm.Split('|')[2]);
                TruckReportItem tri = new TruckReportItem();
                tri.ItemNumber = itemNumber;
                tri.Quantity = quantity;
                tri.LastInventoriedDate = lastInventoriedDate;
                TruckReport.Add(tri);
            }

            PopulateListView();
            Application.Exit();

        }
        public void PopulateListView()
        {
            SetupListView();
            foreach (TruckReportItem tri in TruckReport)
            {
                ListViewItem tmpLVI = new ListViewItem();
                tmpLVI.Text = tri.ItemNumber;
                tmpLVI.SubItems.Add(tri.Quantity.ToString());
                tmpLVI.SubItems.Add(tri.LastInventoriedDate.ToString());
                listView1.Items.Add(tmpLVI);
            }

            //string results = LVI_to_TXT(listView1);
            string results = LVI_to_HTML(listView1);
            Connectivity.EmailCalls.EmailAddresses = new string[] { "tom@tools-plus.com", "jeanne@tools-plus.com", "john@tools-plus.com"};
           // Connectivity.EmailCalls.EmailAddresses = new string[] { "tom@tools-plus.com"};
            Connectivity.EmailCalls.SendEmail2("Truck Inventory Report",results,true);

        }

        


        public void SetupListView()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.FullRowSelect = true;
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.Refresh();
            ColumnHeader itmNumber = new ColumnHeader();
            ColumnHeader qty = new ColumnHeader();
            ColumnHeader date = new ColumnHeader();
            itmNumber.Text = "Item Number:";
            qty.Text = "Quantity:";
            date.Text = "Last Inventoried";
            listView1.Columns.Add(itmNumber);
            listView1.Columns.Add(qty);
            listView1.Columns.Add(date);
            listView1.Refresh();

        }


        public static string LVI_to_TXT(ListView listview, char horizontalTableLine = '=', char verticalTableLine = '=')
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
            tableRow = tableRow.PadRight(tableLength, horizontalTableLine) + "\r\n";

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

           

            return finalTable;

        }
        public static string LVI_to_HTML(ListView listview)
        {
            string HTMLoutput = "";
            HTMLoutput += "<html>\r\n<body>\r\n<table border=\"1\">\r\n\r\n<tr>\r\n";
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
            return HTMLoutput;

        }
    }
}
