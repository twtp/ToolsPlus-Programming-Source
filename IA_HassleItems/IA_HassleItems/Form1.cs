using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace IA_HassleItems
{
    public partial class IA_HasslesFrm : Form
    {

        public int minIA = 3;

        public IA_HasslesFrm()
        {
            InitializeComponent();
        }

        private void IA_HasslesFrm_Load(object sender, EventArgs e)
        {
            listView1.Columns.Add("Item Name");
            listView1.Columns.Add("# of Adjustments");
            listView1.Columns[0].Width = listView1.Width / 2;
            listView1.Columns[1].Width = listView1.Width / 2;
            processingPnl.Visible = false;
            processingPnl.SendToBack();

        }

        public void processingMessage(string Message)
        {
            processLbl.Text = Message;
            processLbl.Refresh();
            label2.Refresh();
        }
        public void showProcessingPanel()
        {
            processingPnl.Visible = true;
            processingPnl.BringToFront();
        }
        public void hideProcessingPanel()
        {
            processingPnl.Visible = false;
            processingPnl.SendToBack();
        }


        public void button1_Click(object sender, EventArgs e)
        {
            showProcessingPanel();
            processingMessage("Querying SQL Table...");
            try
            {
                minIA = int.Parse(textBox1.Text);
            }
            catch
            {
                MessageBox.Show("THAT WASNT A NUMBER!!!");
                return;
            }
            //string connectionString = "Data Source=127.0.0.1;Initial Catalog=TestDatabase;User ID=sa;Password=hacker1121;";
            string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";
            string queryString = "SELECT InventoryComponentMap.ItemNumber FROM WarehouseTransactionLog,InventoryComponentMap WHERE TransactionReference LIKE 'IA%' AND InventoryComponentMap.ComponentID = WarehouseTransactionLog.ComponentID ORDER BY WarehouseTransactionLog.ComponentID";
            
            //string queryString = "SELECT TransactionType,TransactionReference,TransactionTime,'|',InventoryComponentMap.ItemNumber,'|',WarehouseTransactionLog.ComponentID,'|' FROM WarehouseTransactionLog,InventoryComponentMap WHERE TransactionReference LIKE 'IA%' AND InventoryComponentMap.ComponentID = WarehouseTransactionLog.ComponentID ORDER BY WarehouseTransactionLog.ComponentID";
            //string queryString = "SELECT NTUsername FROM Users";


           List<string> ItemsWithIA = new List<string>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(queryString, connection))
                {
                    
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                       
                        while (reader.Read())
                        {

                            ItemsWithIA.Add(reader.GetString(0));
                        }
                        reader.Close();
                    }
                    
                }
                connection.Close();
            }
            int yCount = 0;
            foreach (string ItemIA in ItemsWithIA)
            {
                yCount++;
                bool foundItem = false;
                int xCount = 0;
                for (xCount = 0; xCount < listView1.Items.Count; xCount++)
                {
                    
                    if (ItemIA == listView1.Items[xCount].Text)
                    {
                        foundItem = true;
                        break;
                    }

                }
                if (foundItem == true)
                {
                    int IACount = int.Parse(listView1.Items[xCount].SubItems[1].Text);
                    IACount++;
                    listView1.Items[xCount].SubItems[1].Text = IACount.ToString();
                }
                else
                {
                    ListViewItem newListItem = new ListViewItem();
                    newListItem.Text = ItemIA;
                    newListItem.SubItems.Add("1");
                    listView1.Items.Add(newListItem);
                }

                processingMessage("Processing " + yCount.ToString() + "/" + ItemsWithIA.Count.ToString()); 

            }


            //now to filter by minIA...
            
            for (int z = 0; z < listView1.Items.Count; z++)
            {
                processingMessage("Filtering Results " + z.ToString() + "/" + listView1.Items.Count.ToString());
                if (int.Parse(listView1.Items[z].SubItems[1].Text.ToString()) < minIA)
                {
                    listView1.Items[z].Remove();
                    z--;
                    
                }
            }
            listView1.Refresh();

            //now order the items by subitem[1]
            processingMessage("Sorting Items...");
            SortListByIA(false);

            

            hideProcessingPanel();
            
        }

        public void SortListByIA(bool Ascending)
        {
            ListViewSorter Sorter = new ListViewSorter();
            listView1.ListViewItemSorter = Sorter;
            if (!(listView1.ListViewItemSorter is ListViewSorter))
                return;
            Sorter = (ListViewSorter)listView1.ListViewItemSorter;

            if (Ascending == false)
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.Descending;
            }
            else
            {
                listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            }
            
            Sorter.ByColumn = 1;

            listView1.Sort();

        }


        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();

            xla.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);

            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;

            int i = 1;

            int j = 1;

            ws.Cells[i, j] = "Item Name";
            j++;
            ws.Cells[i, j] = "Adjustments";
            i++;
            j = 1;

            foreach (ListViewItem comp in listView1.Items)
            {

                ws.Cells[i, j] = comp.Text.ToString();

                //MessageBox.Show(comp.Text.ToString());

                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {

                    ws.Cells[i, j] = drv.Text.ToString();

                    j++;

                }

                j = 1;

                i++;

            }

        }





        public class ListViewSorter : System.Collections.IComparer
        {
            public int Compare(object o1, object o2)
            {
                if (!(o1 is ListViewItem))
                    return (0);
                if (!(o2 is ListViewItem))
                    return (0);

                ListViewItem lvi1 = (ListViewItem)o2;
                string str1 = lvi1.SubItems[ByColumn].Text;
                ListViewItem lvi2 = (ListViewItem)o1;
                string str2 = lvi2.SubItems[ByColumn].Text;

                int result;
                if (lvi1.ListView.Sorting == System.Windows.Forms.SortOrder.Ascending)
                    result = String.Compare(str1, str2);
                else
                    result = String.Compare(str2, str1);

                LastSort = ByColumn;

                return (result);
            }


            public int ByColumn
            {
                get { return Column; }
                set { Column = value; }
            }
            int Column = 0;

            public int LastSort
            {
                get { return LastColumn; }
                set { LastColumn = value; }
            }
            int LastColumn = 0;
        }   




  
    }
}