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
    public partial class EditBonus : Form
    {
        public System.Windows.Forms.ListView LV = new ListView();
        public int TheSelectedIndex = 0;
        public bool debug = false;
        public EditBonus()
        {
            InitializeComponent();
        }

        private void EditBonus_Load(object sender, EventArgs e)
        {

        }


        public string RemoveOld()
        {
            int selIndex = TheSelectedIndex;

            string theFile = "";
            if (debug == true)
            {
                theFile = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");
            }
            else
            {
                MessageBox.Show("Last damn time, turn debug on!");
                return "";
            }


            int index = 0;
            int count = 0;
            foreach (string str in theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None))
            {

                if (count > 0)
                {

                    string tmp = str.Split(new string[] { "message: function (item_dollars,item_count) { return " }, StringSplitOptions.None)[1].Split(';')[0];
                    if (tmp == LV.Items[selIndex].SubItems[3].Text)
                    {
                        index = count;
                        break;
                    }
                }
                count++;
            }
            if (index > 0)
            {
                string toRemove = "{ active:" + theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None)[index].Split(new string[] { "}]}," }, StringSplitOptions.None)[0] + "}]},";
                theFile = theFile.Replace(toRemove, "");
            }
            System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "free-items.js", theFile);
            return theFile;
            //string toRemove = "{ active:" + theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None)[selIndex];
            //theFile = theFile.Replace(toRemove, "");
            
        }
        public string RemoveOld2(string BonusTitle)
        {
            int selIndex = TheSelectedIndex;
            string theFile = "";
            if (debug == true)
            {
                theFile = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm");
            }
            else
            {
                MessageBox.Show("Im losing it now. Why wont you turn on debug? Getting to the point im going to let you screw it all up! (just kidding)");
                return "";
            }


            int index = 0;
            int count = 0;
            bool complete = false;
            foreach (string str in theFile.Split(new string[] { "{ active" }, StringSplitOptions.None))
            {

               if (str.Contains("#Bonus Name: " + BonusTitle) == true)
               {
                   theFile = theFile.Replace("{ active" + str.Split('#')[0], "");
                   theFile = theFile.Replace("#Bonus Name: " + BonusTitle, "");
                   
                   complete = true;
                   break;
               }
            }
            if (complete == true)
            {
                System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm", theFile);
                return theFile;
            }
            else
            {
                MessageBox.Show("Could not locate handler for bonus " + BonusTitle);
                return "";
            }
            //string toRemove = "{ active:" + theFile.Split(new string[] { "{ active:" }, StringSplitOptions.None)[selIndex];
            //theFile = theFile.Replace(toRemove, "");

        }


        public void FillForm(bool isDebug, System.Windows.Forms.ListView lv, int SelectedIndex, int ID, bool IsCommented, bool IsActive, string BonusTitle, string CartMessage, string Requirements, string AppliesTo,string BonusType, string BonusQty, string BonusItem)
        {
            LV = lv;
            TheSelectedIndex = SelectedIndex;
            debug = isDebug;

            commentedChk.Checked = IsCommented;
            activeChk.Checked = IsActive;
            cartTxt.Text = CartMessage;
            bonusTxt.Text = BonusTitle;
            if (Requirements.Contains("item_count") == true)
            {
                qtyRadio.Checked = true;
                qtyTxt.Text = Requirements.Replace("item_count ", "").Replace("item_dollars ", "");
            }
            else
            {
                priceRadio.Checked = true;
                priceTxt.Text = Requirements.Replace("item_count ","").Replace("item_dollars ","");
            }
            if (AppliesTo.Contains("item ==") == true)
            {
                itemsRadio.Checked = true;
                itemsTxt.Text = AppliesTo.Replace("||", ",").Replace("/^", "").Replace("/.test(item)", "").Trim();
            }
            else
            {
                linesRadio.Checked = true;
                linesTxt.Text = AppliesTo.Replace("||",",").Replace("/^","").Replace("/.test(item)","").Trim();
            }
            if (BonusType.ToLower() == "fixed")
            {
                fixedQtyRadio.Checked = true;
                fixedQtyAmountTxt.Text = BonusQty;
            }
            else
            {
                buyOneGetOneRadio.Checked = true;
            }
            bonusItemNumberTxt.Text = BonusItem;

        }

        private void linesRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (linesRadio.Checked == true)
            { linesTxt.Enabled = true; itemsTxt.Enabled = false; }
           
        }

        private void itemsRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (itemsRadio.Checked == true)
            { linesTxt.Enabled = false; itemsTxt.Enabled = true; }
        }

        private void addBtn_Click(object sender, EventArgs e)        
        {
            string doc = "";
            if (debug == true)
            {
                doc = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");
            }
            else
            {
                MessageBox.Show("Debug mode is the only mode coded at the moment. Please set it back!");
                return;
            }
            string finalLine = "";
            finalLine += (commentedChk.Checked == true ? "//" : "");


            string active = "active: function () { return " + (activeChk.Checked == true ? "true;}," : "false;},");

            string applies = "";
            if (itemsRadio.Checked == true)
            {
                applies = "applies: function (item) { return ";
                foreach (string str in itemsTxt.Text.Split(','))
                {
                    applies += "item == " + str.Trim() + " || ";
                }
                applies = applies.Trim().TrimEnd('|');
                applies += "; },";

            }
            else
            {
                applies = "applies: function (item) { return ";
                foreach (string str in linesTxt.Text.Split(','))
                {
                    applies += "/^" + str.Trim() + "/.test(item) || ";
                }
                applies = applies.Trim().TrimEnd('|');
                applies += "; },";
            }


            string eligible = "eligible: function (item_dollars,item_count) { return " + (priceRadio.Checked == true ? "item_dollars " + priceTxt.Text : "item_count " + qtyTxt.Text) + ";},";
            string message = "message: function (item_dollars,item_count) { return " + bonusTxt.Text + ";},";
            string teaser = "teaser: function (item_dollars,item_count) { return " + cartTxt.Text + ";},";

            finalLine = "\r\n" + finalLine + "{ " + active + " " + applies + " bonus: [{" + eligible + " " + message + " " + teaser + " }]},\r\n";
            doc = doc.Replace(finalLine, "");

            //now create the code for AddBogusItems.pm...
            string doc2 = "";
            if (debug ==true)
            {
                doc2 = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm");
            }
            else
            {
                MessageBox.Show("Turn Debug ON for the last damn time!");
                return;
            }
            string finalLine2 = "\r\n" + "{ active  => sub { return " + (activeChk.Checked == true ? "1" : "0") + "; }, applies => sub { my $item = lc(shift); return ";
            string appliesLine = "";
            if (itemsRadio.Checked == true)
            {
                appliesLine = "qw/";
                foreach (string str in itemsTxt.Text.Split(','))
                {
                    appliesLine += str.Trim() + " ";
                }
                appliesLine += "/); },";
            }
            else
            {
                foreach (string str in linesTxt.Text.Split(','))
                {
                    appliesLine = "$item =~ /^" + str.Trim() + "/ || ";
                }
                appliesLine = appliesLine.Trim().Trim('|').Trim() + "; },";
            }
            finalLine2 += appliesLine;
            finalLine2 += " bonus => [{ eligible => sub { my $item_dollars = shift; my $item_count = shift; return ";
            if (priceRadio.Checked == true)
            {
                finalLine2 += "$item_dollars " + priceTxt.Text + "; },";
            }
            else
            {
                finalLine2 += "$item_count " + qtyTxt.Text + "; },";
            }
            if (fixedQtyRadio.Checked == true)
            {
                finalLine2 += " quantity => sub { return " + fixedQtyAmountTxt.Text + ";},";
            }
            else
            {
                finalLine2 += " quantity => sub { my $item_dollars = shift; my $item_count = shift; return $item_count; },";
            }
            finalLine2 += " item     => '" + bonusItemNumberTxt.Text + "',}], comment => 1,},\r\n";
            finalLine2 += " #Bonus Name: " + bonusTxt.Text + "\r\n";


            if (debug == true)
            {
                doc = RemoveOld();
                doc2 = RemoveOld2(bonusTxt.Text);
                
                string header = doc.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[0] + "TP.FreeItems.specials = [";
                string therest = doc.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[1];

                string finalPage = header + finalLine + therest;
                System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "free-items.js", finalPage);

                string header2 = doc2.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[0] + "my @specials = (";
                //string therest2 = "sub run {" + doc2.Split(new string[] { "sub run {" }, StringSplitOptions.None)[1];
                string therest2 = doc2.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[1];
                string finalPage2 = header2 + finalLine2 + therest2;
                System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm", finalPage2);

                MessageBox.Show("Complete. Debug is on, so check your desktop");
                this.Close();
                
            }
        }
    }
}
