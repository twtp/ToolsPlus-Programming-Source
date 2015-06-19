using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Renci.SshNet;
using System.Text.RegularExpressions;
namespace VPS_Talker
{
    public partial class Form1 : Form
    {

        public SshClient client = new SshClient("50.23.169.36",9042,"esavelle","summ3r10");


        public Form1()
        {
            InitializeComponent();
        }

        private void ClearLVI()
        {
            listView1.Columns.Clear();
            listView1.Items.Clear();
            listView1.Clear();
            listView1.Refresh();

            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.View = View.Details;
        }
        private void ThrowDataInLVI(string returnedData)
        {
            ClearLVI();   
            //each field is denoted by \t, each line is denoted by \n...

            string[] Lines = returnedData.Split('\n');
            
            int counter = 0;
            foreach (string line in Lines)
            {
                if (counter == 0)
                {
                    //header... 
                    string[] headerColumns = line.Split('\t');
                    
                    listBox1.Refresh();
                    listBox1.Items.Add("Found " + headerColumns.GetLength(0).ToString() + " Columns.");
                    foreach (string col in headerColumns)
                    {
                        ColumnHeader tmpCol = new ColumnHeader();
                        tmpCol.Text = col;
                        tmpCol.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                        listView1.Columns.Add(tmpCol);
                    }
                }
                else
                {
                    ListViewItem tmpItm = new ListViewItem();
                    string[] rowFields = line.Split('\t');
                    int counter2 = 0;
                    foreach (string field in rowFields)
                    {
                        if (counter2 == 0)
                        {
                            tmpItm.Text = field;
                        }
                        else
                        {
                            tmpItm.SubItems.Add(field);
                        }
                        counter2++;
                    }
                    listView1.Items.Add(tmpItm);
                }
                counter++;
            }
            
            listBox1.Items.Add("Pulled " + (Lines.GetLength(0) - 1).ToString() + " Lines.");
            listBox1.Refresh();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormClosed += new FormClosedEventHandler(Form1_FormClosed);
            client.ErrorOccurred += new EventHandler<Renci.SshNet.Common.ExceptionEventArgs>(client_ErrorOccurred);
            
        }
        private void FillTableList()
        {
            treeView1.Nodes.Clear();
            treeView1.Refresh();

            var TopNode = new TreeNode("DB Tables");
            treeView1.Nodes.Add(TopNode);
            var TableNodes = new List<TreeNode>();
            var ColumnNodes = new List<TreeNode>();

            

           // using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            using (client = new SshClient("50.23.169.36", 9042, "esavelle_db", "summ3r10"))
            {               
                listBox1.Items.Add("Retrieving List Of Tables...");
                listBox1.Refresh();
                client.Connect();
                using (SshCommand cmd = client.RunCommand("mysql -psumm3r10 -e\"SELECT table_name FROM information_schema.tables\""))
                {
                    string result = cmd.Result;
                    string error = cmd.Error;

                    if (result != "")
                    {
                        string[] tables = result.Split('\n');
                        int counter = 0;
                        foreach (string table in tables)
                        {
                            if (counter > 0)
                            {
                               
                                TreeNode tmpNode = new TreeNode();
                                tmpNode.Text = table;
                                tmpNode.Name = table;

                                //using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
                                //{
                                    //client.Connect();
                                using (SshCommand cmd2 = client.RunCommand("mysql -psumm3r10 -e\"SELECT column_name FROM information_schema.columns WHERE table_name='" + tmpNode.Text + "'\""))
                                    {
                                        string result2 = cmd2.Result;
                                        string error2 = cmd2.Error;

                                        if (result2 != "")
                                        {
                                            int count = 0;
                                            foreach (string column in result2.Split('\n'))
                                            {
                                                if (count > 0 & column != "")
                                                {
                                                    tmpNode.Nodes.Add(new TreeNode(column));
                                                }
                                                count++;
                                            }
                                        }

                                    }
                                    //client.Disconnect();
                                //}

                                TableNodes.Add(tmpNode);  
                                
                            }
                            counter++;
                        }
                        TopNode.Nodes.AddRange(TableNodes.ToArray());
                        
                        //now our table nodes are good...
                        


                    }
                    if (error != "")
                    {
                        listBox1.Items.Add("Error Pulling Tables: " + error);
                        listBox1.Refresh();
                    }
                }
                client.Disconnect();
                listBox1.Items.Add("Retrieved Tables And Columns...");
                listBox1.Refresh();
            }
        }
        void client_ErrorOccurred(object sender, Renci.SshNet.Common.ExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message + "\r\n\r\n" + e.Exception.Data);
        }

        void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (client.IsConnected == true)
            {
                client.Disconnect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string queryMaster = "mysql -psumm3r10 esavelle_salesorder -e\"";
            string queryMaster = "mysql -psumm3r10 -e\"";
            ClearLVI();
            listBox1.Items.Clear();
            using (client = new SshClient("50.23.169.36", 9042, "db", "summ3r10"))
            {
                listBox1.Items.Add("Connecting...");
                listBox1.Refresh();
                client.Connect();
                listBox1.Items.Add("Connected. Running Command...");
                listBox1.Refresh();
                using (SshCommand cmd = client.RunCommand(queryMaster + textBox2.Text + "\""))
                {
                    string result = cmd.Result;
                    string errors = cmd.Error;

                    if (result != "")
                    {
                        listBox1.Items.Add("Obtained some results!");
                        listBox1.Refresh();
                        //textBox1.Text = result.Replace("\n","\r\n");
                        ThrowDataInLVI(result);
                    }
                    if (errors != "")
                    {
                        listBox1.Items.Add("Command Error: " + errors);
                        listBox1.Refresh();
                    }
                    
                }
                listBox1.Items.Add("Disconnecting from server...");
                listBox1.Refresh();
                client.Disconnect();
                listBox1.Items.Add("Disconnected.");
                listBox1.Refresh();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FillTableList();
        }






    }
}
