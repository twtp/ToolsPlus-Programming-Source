using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EBayCategoryDBInserter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void SetupLVI()
        {
            catLVI.Items.Clear();
            catLVI.Columns.Clear();
            catLVI.Clear();
            catLVI.Refresh();

            ColumnHeader CategoryName = new ColumnHeader();
            ColumnHeader CategoryID = new ColumnHeader();
            CategoryName.Width = catLVI.Width - (catLVI.Width / 5);
            CategoryID.Width = catLVI.Width / 5;
            CategoryName.Text = "Category";
            CategoryID.Text = "ID";
            catLVI.Columns.Add(CategoryName);
            catLVI.Columns.Add(CategoryID);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SetupLVI();
            ProcessFile(@"C:\Users\tomwestbrook\Documents\ToolsPlus Programming Source\ebay categories\ebayCategories.csv");
        }
        private string[] GetCSVFields(string csvLine)
        {
            return System.Text.RegularExpressions.Regex.Split(csvLine.Substring(0, csvLine.Length - 1), @"""?\s*,\s*""?");
        }
        private void ProcessFile(string FilePath)
        {
            string[] CategoriesInfo = System.IO.File.ReadAllLines(FilePath);

            if (CategoriesInfo.GetLength(0) > 0)
            {
                

                foreach (string category in CategoriesInfo)
                {
                    string[] tokens = GetCSVFields(category);
                    if (tokens[0].ToUpper().Contains("CATEGORY ID") == true)
                    {
                        //do nothing
                    }
                    else
                    {
                        try { int.Parse(tokens[0].Trim()); }
                        catch { break; }

                        int CategoryID = int.Parse(tokens[0].Trim());
                        string CategoryName = "";

                        foreach (string token in tokens)
                        {
                            
                            if (token.Length > 0)
                            {
                                try { int.Parse(token); }
                                catch { CategoryName += token + " > "; }
                            }
                        }
                        CategoryName = CategoryName.Trim();
                        CategoryName = CategoryName.Trim('>');
                        CategoryName = CategoryName.Trim();
                        //MessageBox.Show(CategoryName);
                        ListViewItem newCategory = new ListViewItem();
                        newCategory.Text = CategoryName;
                        newCategory.SubItems.Add(CategoryID.ToString());
                        catLVI.Items.Add(newCategory);


                    }

                }

                MessageBox.Show("All set...");
     
                
            }
            else
            {
                MessageBox.Show("Nothing to load!");
            }

        }

        private void addToDBBtn_Click(object sender, EventArgs e)
        {
            int counter = 0;
            foreach (ListViewItem lvi in catLVI.Items)
            {
                counter++;
                statusLbl.Text = "Status: Processing " + counter.ToString() + "/" + catLVI.Items.Count.ToString();
                statusLbl.Refresh();
                Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayHomeGardenConversionID (CategoryName,CategoryID) VALUES('" + lvi.Text + "'," + lvi.SubItems[1].Text + ")");
                System.Threading.Thread.Sleep(20);
            }
        }
    }
}
