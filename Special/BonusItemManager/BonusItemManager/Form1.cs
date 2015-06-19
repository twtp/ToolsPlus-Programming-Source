using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BonusItemManager
{
    public partial class Form1 : Form
    {
        private bool Debug = true;


        public struct PerlSpecial
        {
            public bool IsCommented;
            public bool IsActive;
            public bool AppliesToItems;
            public bool AppliesToLines;
            public string AppliesTo;
            public bool EligibleByQty;
            public bool EligibleByPrice;
            public string EligibleBy;
            public bool BonusFixedQty;
            public bool BonusVariableQty;
            public int BonusQty;
            public string BonusVariation;
            public string BonusItemNumber;
            public string BonusTitle;
            
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Shown += Form1_Shown;
            
        }

        void Form1_Shown(object sender, EventArgs e)
        {
            Common.Variables.DebugFilePath = Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory) + "\\";
            PopulateForm();
        }

        public void PopulateJavascript()
        {
            string results = "";
            if (Debug == false)
            {
                results = Connectivity.HTTPCalls.HTTP_GET("http://p1.hostingprod.com/@tools-plus.com/solidcactus/free-items.js");
            }
            else
            {
                results = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");
            }
            headerTxt.Text = results.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[0] + "TP.FreeItems.specials = [";
            bodyTxt.Text = results.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[1].Split(new string[] { "];" }, StringSplitOptions.None)[0];
            footerTxt.Text = "];" + results.Split(new string[] { "];" }, StringSplitOptions.None)[1];

            List<string> inactive_specials = new List<string>();

            foreach (string str in bodyTxt.Text.Split(new string[] { "//{ active: function ()" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string elem = "//{ active: function ()" + str;
                inactive_specials.Add(elem);
            }
            inactive_specials.RemoveAt(0);

            List<string> specials = new List<string>();

            string tmpBody = bodyTxt.Text;
            tmpBody = tmpBody.Replace("//{ active: function ()", "//{ nonactive: function()");
            foreach (string str in tmpBody.Split(new string[] { "{ active: function ()" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string elem = "{ active: function ()" + str;
                specials.Add(elem);
            }
            specials.RemoveAt(0);



            //foreach (string str in specials)
            //{
            //    dbugTxt.Text += str + "\r\n\r\n----------------------------------------------\r\n\r\n";
            //}
            PopulateListView(inactive_specials, specials);
        }

        public void SetupPerlListView()
        {
            listView2.Items.Clear();
            listView2.Columns.Clear();
            listView2.Clear();
            listView2.View = View.Details;
            listView2.GridLines = true;
            listView2.FullRowSelect = true;

            ColumnHeader colID = new ColumnHeader();
            ColumnHeader colIsCommented = new ColumnHeader();
            ColumnHeader colIsActive = new ColumnHeader();
            ColumnHeader colAppliesToType = new ColumnHeader();
            ColumnHeader colAppliesToData = new ColumnHeader();
            ColumnHeader colEligibleType = new ColumnHeader();
            ColumnHeader colEligibleData = new ColumnHeader();
            ColumnHeader colBonusType = new ColumnHeader();
            ColumnHeader colBonusQty = new ColumnHeader();
            ColumnHeader colBonusItemNumber = new ColumnHeader();
            ColumnHeader colBonusTitle = new ColumnHeader();

            colID.Text = "ID";
            colID.Width = listView2.ClientSize.Width / 25;
            colBonusTitle.Text = "Bonus Title";
            colBonusTitle.Width = listView2.ClientSize.Width / 10;
            colIsCommented.Text = "Commented Out";
            colIsCommented.Width = listView2.ClientSize.Width / 20;
            colIsActive.Text = "Active";
            colIsActive.Width = listView2.ClientSize.Width / 20;
            colAppliesToType.Text = "Applies To Type";
            colAppliesToType.Width = listView2.ClientSize.Width / 10;            
            colAppliesToData.Width = listView2.ClientSize.Width / 6;
            colAppliesToData.Text = "Applies To Data";
            colEligibleType.Text = "Eligible Type";
            colEligibleType.Width = listView2.ClientSize.Width / 10;
            colEligibleData.Text = "Eligible Data";
            colEligibleData.Width = listView2.ClientSize.Width / 6;
            colBonusType.Width = listView2.ClientSize.Width / 10;
            colBonusType.Text = "Bonus Type";
            colBonusQty.Width = listView2.ClientSize.Width / 10;
            colBonusQty.Text = "Bonus Qty";
            colBonusItemNumber.Width = listView2.ClientSize.Width / 8;
            colBonusItemNumber.Text = "Bonus Item";
            
            listView2.Columns.AddRange(new ColumnHeader[] { colID, colBonusTitle, colIsCommented, colIsActive, colAppliesToType, colAppliesToData, colEligibleType, colEligibleData, colBonusType, colBonusQty, colBonusItemNumber });
            
        }
        public void PopulatePerlListView(List<PerlSpecial> Specials)
        {
            SetupPerlListView();
            int x = 0;
            foreach (PerlSpecial ps in Specials)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = x.ToString();
                lvi.SubItems.Add(ps.BonusTitle);
                lvi.SubItems.Add((ps.IsCommented == true ? "Yes" : "No"));
                lvi.SubItems.Add((ps.IsActive == true ? "Yes" : "No"));
                lvi.SubItems.Add((ps.AppliesToItems == true ? "Items" : "Lines"));
                lvi.SubItems.Add(ps.AppliesTo);
                lvi.SubItems.Add((ps.EligibleByPrice == true ? "Price" : "Qty"));
                lvi.SubItems.Add(ps.EligibleBy);
                lvi.SubItems.Add((ps.BonusFixedQty == true ? "Fixed" : "Variation"));
                lvi.SubItems.Add((ps.BonusFixedQty == true ? ps.BonusQty.ToString() : ps.BonusVariation));
                lvi.SubItems.Add(ps.BonusItemNumber);
                listView2.Items.Add(lvi);
                x++;
            }

        }
        public void PopulatePerl()
        {
            string file = "";
            if (Debug == true)
            {
                file = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm");
            }
            else
            {
                MessageBox.Show("If I wasn't in debug mode and didn't code this message the program would literally shit on itself. Put back in debug mode!!!");
                return;
            }
            perlHeaderTxt.Text = file.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[0] + "my @specials = (";
            perlBodyTxt.Text = file.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[1].Split(new string[]{"sub run {"},StringSplitOptions.None)[0].Trim().Trim(';').Trim(')');
            perlFooterTxt.Text = "sub run {" + file.Split(new string[] { "sub run {" }, StringSplitOptions.None)[1];
            List<PerlSpecial> PS = new List<PerlSpecial>();
            foreach (string str in perlBodyTxt.Text.Split(new string[]{"{ active"},StringSplitOptions.None))
            {
                if (str.Length > 5)
                {
                    PerlSpecial ps = new PerlSpecial();
                    ps.BonusTitle = str.Split(new string[] { "#Bonus Name: " }, StringSplitOptions.None)[1].Trim();
                    ps.IsCommented = str.Contains("#},");
                    ps.IsActive = (str.Contains("active  => sub { return 0; }") == true ? false : true);
                    ps.AppliesToItems = str.Contains("return in_list($item") == true;
                    ps.AppliesToLines = (ps.AppliesToItems == true ? false : true);
                    if (ps.AppliesToItems == true)
                    {
                        ps.AppliesTo = str.Split(new string[] { "return in_list($item, " }, StringSplitOptions.None)[1].Split(new string[] { ");" }, StringSplitOptions.None)[0];
                    }
                    else
                    {
                        
                        ps.AppliesTo = str.Split(new string[] { "my $item = lc(shift); return $item =~ " }, StringSplitOptions.None)[1].Split(';')[0];
                    }

                    ps.EligibleByPrice = str.Split(new string[] { "eligible => sub {" }, StringSplitOptions.None)[1].Split('}')[0].Contains("return $item_dollars");
                    ps.EligibleByQty = (ps.EligibleByPrice == true ? false : true);
                    if (ps.EligibleByPrice == true)
                    {
                        ps.EligibleBy = str.Split(new string[] { "eligible => sub {" }, StringSplitOptions.None)[1].Split('}')[0].Split(new string[] { "return $item_dollars " }, StringSplitOptions.None)[1];
                    }
                    else
                    {
                        ps.EligibleBy = str.Split(new string[] { "eligible => sub {" }, StringSplitOptions.None)[1].Split('}')[0].Split(new string[] { "return $item_count " }, StringSplitOptions.None)[1];
                    }
                    string bqty = str.Split(new string[] { "quantity => sub {" }, StringSplitOptions.None)[1].Split('}')[0];
                    if (bqty.Length > 15)
                    {
                        ps.BonusVariableQty = true;
                        ps.BonusFixedQty = false;
                        ps.BonusVariation = bqty;
                    }
                    else
                    {
                        ps.BonusFixedQty = true;
                        ps.BonusVariableQty = false;
                        ps.BonusQty = int.Parse(bqty.Split(new string[] { "return " }, StringSplitOptions.None)[1].Split(';')[0].Trim());
                    }
                    ps.BonusItemNumber = str.Split(new string[] { "item     => '" }, StringSplitOptions.None)[1].Split(new string[] { "'" }, StringSplitOptions.None)[0];
                    PS.Add(ps);
                }
            }
            PopulatePerlListView(PS);
        }
        public void PopulateForm()
        {

            PopulateJavascript();
            PopulatePerl();

        }
        private void PopulateListView(List<string> CommentedOutSpecials, List<string> Specials)
        {
            SetupListView();

            int idCounter = 0;
            foreach (string str in CommentedOutSpecials)
            {
                string IsActive = "";
                string IsCommented = "Yes";
                string AppliesTo = "";
                string Requirement = "";
                string BonusTitle = "";
                string CartMessage = "";

                string tmpActive = str.Split(new string[] { "return " }, StringSplitOptions.None)[1].Split(';')[0].Trim();
                IsActive = (tmpActive.ToLower() == "true" ? "Yes" : "No");
                string tmpApplies = str.Split(new string[] { "applies:" }, StringSplitOptions.None)[1].Split(new string[] { ";" }, StringSplitOptions.None)[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                AppliesTo = tmpApplies;
                string tmpRequirement = str.Split(new string[] { "eligible:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                Requirement = tmpRequirement;
                string tmpBonus = str.Split(new string[] { "message:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                BonusTitle = tmpBonus;
                try
                {
                    string tmpCartMsg = str.Split(new string[] { "teaser:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                    CartMessage = tmpCartMsg;
                }
                catch
                {
                    CartMessage = "undefined";
                }

                ListViewItem lvi = new ListViewItem();
                lvi.Text = idCounter.ToString();
                lvi.SubItems.Add(IsCommented);
                lvi.SubItems.Add(IsActive);
                lvi.SubItems.Add(BonusTitle);                
                lvi.SubItems.Add(Requirement);
                lvi.SubItems.Add(AppliesTo);
                lvi.SubItems.Add(CartMessage);

                listView1.Items.Add(lvi);
                //MessageBox.Show(listView1.Items[0].SubItems[5].Text);
                idCounter++;
            }

            foreach (string str in Specials)
            {
                string IsActive = "";
                string IsCommented = "No";
                string AppliesTo = "";
                string Requirement = "";
                string BonusTitle = "";
                string CartMessage = "";

                string tmpActive = str.Split(new string[] { "return " }, StringSplitOptions.None)[1].Split(';')[0].Trim();
                IsActive = (tmpActive.ToLower() == "true" ? "Yes" : "No");
                string tmpApplies = str.Split(new string[] { "applies:" }, StringSplitOptions.None)[1].Split(new string[] { ";" }, StringSplitOptions.None)[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                AppliesTo = tmpApplies;
                string tmpRequirement = str.Split(new string[] { "eligible:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                Requirement = tmpRequirement;
                string tmpBonus = str.Split(new string[] { "message:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                BonusTitle = tmpBonus;
                try
                {
                    string tmpCartMsg = str.Split(new string[] { "teaser:" }, StringSplitOptions.None)[1].Split(';')[0].Split(new string[] { "return " }, StringSplitOptions.None)[1];
                    CartMessage = tmpCartMsg;
                }
                catch
                {
                    CartMessage = "undefined";
                }
                ListViewItem lvi = new ListViewItem();
                lvi.Text = idCounter.ToString();
                lvi.SubItems.Add(IsCommented);
                lvi.SubItems.Add(IsActive);
                lvi.SubItems.Add(BonusTitle);
                lvi.SubItems.Add(CartMessage);
                lvi.SubItems.Add(Requirement);
                lvi.SubItems.Add(AppliesTo);                
                

                listView1.Items.Add(lvi);

                idCounter++;
            }


        }
        private void SetupListView()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;

            ColumnHeader colID = new ColumnHeader();
            ColumnHeader colIsCommented = new ColumnHeader();
            ColumnHeader colIsActive = new ColumnHeader();
            ColumnHeader colAppliesTo = new ColumnHeader();
            ColumnHeader colRequirement = new ColumnHeader();
            ColumnHeader colBonusTitle = new ColumnHeader();
            ColumnHeader colCartMessage = new ColumnHeader();

            colID.Text = "ID";
            colID.Width = listView1.ClientSize.Width / 25;
            colIsCommented.Text = "Commented Out";
            colIsCommented.Width = listView1.ClientSize.Width / 20;
            colIsActive.Text = "Active";
            colIsActive.Width = listView1.ClientSize.Width / 20;
            colAppliesTo.Text = "Applicable Item(s) / Line(s)";
            colAppliesTo.Width = listView1.ClientSize.Width / 6;
            colRequirement.Text = "Bonus Requirement(s)";
            colRequirement.Width = listView1.ClientSize.Width / 6;
            colBonusTitle.Text = "Bonus Title";
            colBonusTitle.Width = listView1.ClientSize.Width / 4;
            colCartMessage.Text = "Cart Message";
            colCartMessage.Width = (int)((float)listView1.ClientSize.Width / 3.75f);

            listView1.Columns.Add(colID);
            listView1.Columns.Add(colIsCommented);
            listView1.Columns.Add(colIsActive);
            listView1.Columns.Add(colBonusTitle);
            listView1.Columns.Add(colCartMessage);
            listView1.Columns.Add(colRequirement);
            listView1.Columns.Add(colAppliesTo);


        }


        private void bodyTxt_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddBonus ab = new AddBonus();
            ab.debug = true;
            ab.FormClosed += ab_FormClosed;
            ab.Show();
        }

        void ab_FormClosed(object sender, FormClosedEventArgs e)
        {
            PopulateForm();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int selIndex = listView1.SelectedIndices[0];
            string BonusTitle = listView1.Items[selIndex].SubItems[3].Text;
            string theFile = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");


            int index = 0;
            int count = 0;
            foreach (string str in theFile.Split(new string[] {"{ active:"},StringSplitOptions.None))
            {

                if (count > 0)
                {

                    string tmp = str.Split(new string[] { "message: function (item_dollars,item_count) { return " }, StringSplitOptions.None)[1].Split(';')[0];
                    if (tmp == listView1.Items[selIndex].SubItems[3].Text)
                    {
                        index = count;
                        break;
                    }
                }
                count++;
            }
            if (index > 0)
            {
                string toRemove = "{ active:" + theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None)[index].Split(new string[] {"}]},"},StringSplitOptions.None)[0] + "}]},";
                theFile = theFile.Replace(toRemove, "");
            }
            System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "free-items.js", theFile);

            EditBonus eb = new EditBonus();
            eb.debug = true;
            string perlDoc = eb.RemoveOld2(BonusTitle);
            System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm", perlDoc);




            //string toRemove = "{ active:" + theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None)[selIndex];
            //theFile = theFile.Replace(toRemove, "");
            PopulateForm();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedIndices.Count <=0)
            {
                MessageBox.Show("Please select a bonus to edit!");
                return;
            }

            int selIndex = listView1.SelectedIndices[0];


            string ID = listView1.Items[selIndex].Text;
            string Commented = listView1.Items[selIndex].SubItems[1].Text;
            string Active = listView1.Items[selIndex].SubItems[2].Text;
            string BonusTitle = listView1.Items[selIndex].SubItems[3].Text;
            string CartMessage = listView1.Items[selIndex].SubItems[4].Text;
            string Requirement = listView1.Items[selIndex].SubItems[5].Text;
            string AppliesTo = listView1.Items[selIndex].SubItems[6].Text;
            
            //need to link perl to js... by bonus title!
            ListViewItem lvi = new ListViewItem();
            foreach (ListViewItem lvi2 in listView2.Items)
            {
                if (lvi2.SubItems[1].Text == BonusTitle)
                {
                    lvi = lvi2;
                    break;
                }
            }
            string BonusType = "";
            string BonusQty = "";
            string BonusItem = "";

            if (lvi.Text != "")
            {
                BonusType = lvi.SubItems[8].Text;
                BonusQty = lvi.SubItems[9].Text;
                BonusItem = lvi.SubItems[10].Text;

            }
            else
            {
                MessageBox.Show("Error: Cannot find matching rule in Perl module. Tell Tom this shit's broken!");
                return;
            }


            EditBonus eb = new EditBonus();
            eb.FormClosed += eb_FormClosed;
            eb.FillForm(Debug,listView1,listView1.SelectedIndices[0],int.Parse(ID), (Commented == "Yes" ? true : false), (Active == "Yes" ? true : false), BonusTitle, CartMessage, Requirement, AppliesTo,BonusType,BonusQty,BonusItem);

            eb.Show();
            

        }

        void eb_FormClosed(object sender, FormClosedEventArgs e)
        {
            //throw new NotImplementedException();
            PopulateForm();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            bool preliminaryCheck = false;
            if (listView1.SelectedIndices.Count > 0)
            {
                try
                {
                    string BonusTitle = listView1.Items[listView1.SelectedIndices[0]].SubItems[3].Text;
                    string jsDoc = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");
                    string perlDoc = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm");

                    if (jsDoc.Contains(BonusTitle) == true)
                    {
                        if (perlDoc.Contains("#Bonus Name: " + BonusTitle)==true)
                        {
                            preliminaryCheck = true;
                        }
                        else
                        {
                            preliminaryCheck = false;
                        }
                    }
                    else
                    {
                        preliminaryCheck = false;
                    }

                }
                catch
                {
                    MessageBox.Show("There is an issue with this bonus!");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Please select a bonus to verify!");
                return;
            }

            if (preliminaryCheck==true)
            {
                MessageBox.Show("Preliminary Check: OK");
            }
            else
            {
                MessageBox.Show("Preliminary Check: ISSUE");
            }


        }
    }
}
