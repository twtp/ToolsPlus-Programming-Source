using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SearsAttributes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            SetupLVI();
            SetupLVI2();

        }
        public void BetterParser()
        {
            this.Refresh();
            
            SetupLVI();
            SetupLVI2();
            //string[] mapping = System.IO.File.ReadAllLines(@"C:\users\tomwestbrook\desktop\sears-mapping.csv");
            int counter = 0;


                    Microsoft.VisualBasic.FileIO.TextFieldParser TFP = new Microsoft.VisualBasic.FileIO.TextFieldParser(@"c:\users\tomwestbrook\desktop\sears-mapping.csv");
                    TFP.HasFieldsEnclosedInQuotes = true;
                    TFP.SetDelimiters(",");
                    bool foundGroup = false;
                    string[] Fields;
                    while (!TFP.EndOfData)
                    {
                        Fields = TFP.ReadFields();

                       
                            bool existsFlag = false;
                            foreach (ListViewItem itm in listView2.Items)
                            {                               
                                if (Fields[1] == itm.SubItems[1].Text)
                                {
                                    existsFlag = true;
                                    break;
                                }
                            }
                            if (existsFlag == false & Fields[1].Contains("|")==true)
                            {
                                //listBox1.Items.Add(Fields[1]);
                                ListViewItem newLVI = new ListViewItem();
                                newLVI.Text = Fields[0];
                                newLVI.SubItems.Add(Fields[1]);
                                listView2.Items.Add(newLVI);
                            }
                        
                       

                       
                    }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            BetterParser();

        }
        private void SetupLVI()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();

            ColumnHeader AttributeCol = new ColumnHeader();
            ColumnHeader AttributeTypeCol = new ColumnHeader();
            ColumnHeader AttributeFreeText = new ColumnHeader();
            ColumnHeader MultipleChoice = new ColumnHeader();
            ColumnHeader Trademark = new ColumnHeader();
            ColumnHeader AttributeValues = new ColumnHeader();

            AttributeCol.Text = "Attribute Name";
            AttributeCol.Width = listView1.Width / 6;

            AttributeTypeCol.Text = "Attribute Type";
            AttributeTypeCol.Width = listView1.Width / 7;

            AttributeFreeText.Text = "Accepts Custom Value";
            AttributeFreeText.Width = listView1.Width / 6;

            MultipleChoice.Text = "Multiple Choice";
            MultipleChoice.Width = listView1.Width / 8;

            Trademark.Text = "Trademark";
            Trademark.Width = listView1.Width / 12;

            AttributeValues.Text = "Accepted Values";
            AttributeValues.Width = listView1.Width / 6;

            listView1.Columns.Add(AttributeCol);
            listView1.Columns.Add(AttributeTypeCol);
            listView1.Columns.Add(AttributeFreeText);
            listView1.Columns.Add(MultipleChoice);
            listView1.Columns.Add(Trademark);
            listView1.Columns.Add(AttributeValues);

        }

        private void SetupLVI2()
        {
            listView2.Items.Clear();
            listView2.Columns.Clear();
            listView2.Clear();

            ColumnHeader CatIDCol = new ColumnHeader();
            ColumnHeader CatNameCol = new ColumnHeader();
            

            CatIDCol.Text = "Category ID";
            CatIDCol.Width = listView2.Width / 6;

            CatNameCol.Text = "Category Name";
            CatNameCol.Width = listView2.Width - (listView2.Width/7);
           

            listView2.Columns.Add(CatIDCol);
            listView2.Columns.Add(CatNameCol);
           

        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in listView2.Items)
            {
                string CatID = lvi.Text;
                string CatName = lvi.SubItems[1].Text;

                Connectivity.SQLCalls.sqlQuery("INSERT INTO SearsCategoriesHeader (CategoryID,CategoryName) VALUES(" + CatID + ",'" + CatName.Replace("'","''") + "')");


            }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView2.SelectedIndices.Count == 0)
            {
                return;
            }
            SetupLVI();

            string category = listView2.Items[listView2.SelectedIndices[0]].SubItems[1].Text;




            Microsoft.VisualBasic.FileIO.TextFieldParser TFP = new Microsoft.VisualBasic.FileIO.TextFieldParser(@"c:\users\tomwestbrook\desktop\sears-mapping.csv");
            TFP.HasFieldsEnclosedInQuotes = true;
            TFP.SetDelimiters(",");
            bool foundGroup = false;
            string[] Fields;
            while (!TFP.EndOfData)
            {
                Fields = TFP.ReadFields();
                try
                {
                    int.Parse(Fields[0]);

                    if (Fields[1] == category)
                    {
                        string Attribute = Fields[2];
                        string AttributeType = Fields[3];
                        string AttributeFreeText = Fields[4];
                        string MultipleChoice = Fields[5];
                        string Trademark = Fields[6];
                        string AttributeValues = Fields[7];

                        ListViewItem newLVI = new ListViewItem();
                        newLVI.Text = Attribute;
                        newLVI.SubItems.Add(AttributeType);
                        newLVI.SubItems.Add(AttributeFreeText);
                        newLVI.SubItems.Add(MultipleChoice);
                        newLVI.SubItems.Add(Trademark);
                        newLVI.SubItems.Add(AttributeValues);

                        listView1.Items.Add(newLVI);
                        foundGroup = true;
                    }
                    else
                    {
                        if (foundGroup == true)
                        {
                            break;
                        }
                    }

                }
                catch { }







            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem catLVI in listView2.Items)
            {               
                catLVI.Selected = true;
                System.Threading.Thread.Sleep(500);

                foreach (ListViewItem lvi in listView1.Items)
                {
                    statusLbl.Text = "Status: Processing Category " + catLVI.Index.ToString() + "/" + listView2.Items.Count.ToString() + "... adding Attribute " + lvi.Index.ToString() + "/" + listView1.Items.Count.ToString();
                    statusLbl.Refresh();
                    string CategoryID = listView2.Items[listView2.SelectedIndices[0]].Text;
                    string AttributeName = lvi.Text;
                    string AttributeType = lvi.SubItems[1].Text;
                    string AcceptsCustomValues = lvi.SubItems[2].Text.ToUpper() == "YES" ? "1" : "0";
                    string MultipleChoice = lvi.SubItems[3].Text.ToUpper() == "YES" ? "1" : "0";
                    string Trademark = lvi.SubItems[4].Text.ToUpper() == "YES" ? "1" : "0";
                    string AcceptedValues = lvi.SubItems[5].Text;
                   

                    Connectivity.SQLCalls.sqlQuery("INSERT INTO SearsCategoriesLines (CategoryID,AttributeName,AttributeType,AcceptsCustomValue,MultipleChoice,Trademark,AcceptedValues) VALUES('" + CategoryID + "','" + AttributeName + "','" + AttributeType + "'," + AcceptsCustomValues + "," + MultipleChoice + "," + Trademark + ",'" + AcceptedValues + "')");

                }
                catLVI.Selected = false;

            }

           


        }
    }
}
