using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FeedGenerator
{
    public partial class ModifyLine : Form
    {
        public ModifyLine()
        {
            InitializeComponent();
        }
        public ListViewItem currentItem = new ListViewItem();
        public ListViewItem modifiedItem = new ListViewItem();
        public int CurrentOffset = -1;
        public Form1 currentForm = new Form1();
        private void ModifyLine_Load(object sender, EventArgs e)
        {

        }

        public void CreateFormArray(ListView.ColumnHeaderCollection Columns, ListViewItem LineData, int Offset, Form1 CurrentForm)
        {
            currentItem = LineData;
            CurrentOffset = Offset;
            currentForm = CurrentForm;
            int xCoord = 10;
            int yCoord = 10;
            List<int> ColumnIndexMap = new List<int>();

            int indexCol = 0;



            for (int i = 0; i < Columns.Count; i++)
            {
                foreach (ColumnHeader col in Columns)
                {
                    if (col.Index == i)
                    {
                        ColumnIndexMap.Add(col.DisplayIndex);
                    }
                }


            }








            /*
            for (int z = 0; z < Columns.Count; z++)
            {
                if (ColumnIndexMap.Count == 0)
                {
                    for (int y = 0; y < Columns.Count; y++)
                    {
                        if (Columns[y].Index == 0)
                        {
                            //ColumnIndexMap.Add(c.DisplayIndex);
                            indexCol = Columns[y].DisplayIndex;
                        }
                    }
                    //foreach (ColumnHeader c in Columns)

                }
                if (Columns[z].Index != 0)
                {
                    ColumnIndexMap.Add(Columns[z].DisplayIndex);
                }
            }*/

            for (int y = 0; y < ColumnIndexMap.Count; y++)
            {
                for (int z = 0; z < ColumnIndexMap.Count; z++)
                {
                    if (ColumnIndexMap[z]==y)
                    {
                        if (y > 0)
                        {
                            Label lbl = new Label();
                            lbl.Text = Columns[z].Text;
                            //lbl.Width = 50;
                            lbl.AutoSize = true;
                            lbl.Height = 20;
                            if (lbl.Width < 150)
                            {
                                lbl.AutoSize = false;
                                lbl.Width = 150;
                            }
                            lbl.Top = yCoord;
                            lbl.Left = xCoord;
                            lbl.Name = "Label" + ColumnIndexMap[z].ToString();
                            lbl.ForeColor = Color.White;
                            lbl.Parent = itemPanel;
                            lbl.Visible = true;

                            TextBox tbox = new TextBox();
                            tbox.Name = "TextBox" + z.ToString();
                            tbox.Text = LineData.SubItems[z].Text;
                            tbox.Height = 20;
                            tbox.Width = lbl.Width;
                            tbox.Top = lbl.Bottom + 5;
                            tbox.Left = lbl.Left;
                            tbox.Parent = itemPanel;
                            tbox.Visible = true;
                            xCoord += lbl.Width;
                        }
                    }
                }
            }

            /*
            int count = 1;
            foreach (int i in ColumnIndexMap)
            {
                if (Columns[i].Index != 0)
                {
                    Label lbl = new Label();
                    lbl.Text = Columns[i].Text;
                    //lbl.Width = 50;
                    lbl.AutoSize = true;
                    lbl.Height = 20;
                    if (lbl.Width < 150)
                    {
                        lbl.AutoSize = false;
                        lbl.Width = 150;
                    }
                    lbl.Top = yCoord;
                    lbl.Left = xCoord;
                    lbl.Name = "Label" + i.ToString();
                    lbl.Parent = itemPanel;
                    lbl.Visible = true;

                    TextBox tbox = new TextBox();
                    tbox.Name = "TextBox" + i.ToString();
                    tbox.Text = LineData.SubItems[i].Text;
                    tbox.Height = 20;
                    tbox.Width = lbl.Width;
                    tbox.Top = lbl.Bottom + 5;
                    tbox.Left = lbl.Left;
                    tbox.Parent = itemPanel;
                    tbox.Visible = true;
                    xCoord += lbl.Width;
                    count++;
                }
            }*/

           /* for (int x = 1; x < Columns.Count;x++)
            {
                Label lbl = new Label();                
                lbl.Text = Columns[ColumnIndexMap[x]].Text;
                //lbl.Width = 50;
                lbl.AutoSize = true;
                lbl.Height = 20;
                if (lbl.Width < 150)
                {
                    lbl.AutoSize = false;
                    lbl.Width = 150;
                }
                lbl.Top = yCoord;
                lbl.Left = xCoord;
                lbl.Name = "Label" + x.ToString();
                lbl.Parent = itemPanel;
                lbl.Visible = true;

                TextBox tbox = new TextBox();
                tbox.Name = "TextBox" + x.ToString();
                tbox.Text = LineData.SubItems[ColumnIndexMap[x]].Text;
                tbox.Height = 20;
                tbox.Width = lbl.Width;
                tbox.Top = lbl.Bottom + 5;
                tbox.Left = lbl.Left;
                tbox.Parent = itemPanel;
                tbox.Visible = true;
                xCoord += lbl.Width;
            }*/
        }

        private void modifyOnceBtn_Click(object sender, EventArgs e)
        {
            ListViewItem tmpItem = new ListViewItem();
            tmpItem.Text = currentItem.Text;
            for (int x = 0; x < (this.itemPanel.Controls.Count/2)+1; x++)
            {
                foreach (Control c in this.itemPanel.Controls)
                {
                    //MessageBox.Show(c.GetType().ToString());
                    if (c.GetType().ToString() == "System.Windows.Forms.TextBox")
                    {
                        if (c.Name == "TextBox" + x.ToString())
                        {
                            tmpItem.SubItems.Add(c.Text);
                        }
                    }
                }
            }
            modifiedItem = tmpItem;
            currentForm.previewLVI.Items.Remove(currentForm.previewLVI.SelectedItems[0]);
            currentForm.previewLVI.Items.Insert(CurrentOffset, modifiedItem);
            //previewLVI.Items.Remove(previewLVI.SelectedItems[0]);
            //previewLVI.Items.Insert(offset, ml.modifiedItem);
            this.Close();

        }

    }
}
