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
    public partial class AddBonus : Form
    {
        public bool debug = false;
        public AddBonus()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bonusTxt.Text += "item_count";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            bonusTxt.Text += "item_dollars";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            cartTxt.Text += "item_count";

        }

        private void button4_Click(object sender, EventArgs e)
        {
            cartTxt.Text += "item_dollars";
        }

        private void itemsRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (itemsRadio.Checked == true)
            {
                linesTxt.Enabled = false;
                itemsTxt.Enabled = true;
            }
        }

        private void linesRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (linesRadio.Checked == true)
            {
                linesTxt.Enabled = true;
                itemsTxt.Enabled = false;
            }
        }

        private void priceRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (priceRadio.Checked == true)
            {
                priceTxt.Enabled = true;
                qtyTxt.Enabled = false;
            }
        }

        private void qtyRadio_CheckedChanged(object sender, EventArgs e)
        {
            if (qtyRadio.Checked == true)
            {
                priceTxt.Enabled = false;
                qtyTxt.Enabled = true;
            }
        }

        private void addBtn_Click(object sender, EventArgs e)
        {
            string finalLine = "";            
            finalLine += (commentedChk.Checked == true? "//" :"");
            
            
            string active = "active: function () { return " + (activeChk.Checked == true ? "true;}," : "false;},");
            
            string applies = "";
            if (itemsRadio.Checked==true)
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
            string teaser = "teaser: function (item_dollars,item_count) { return " + cartTxt.Text + ";}";

            finalLine = "\r\n" + finalLine + "{ " + active + " " + applies + " bonus: [{" + eligible + " " + message + " " + teaser + "}]},\r\n";


            //now create the code for AddBogusItems.pm...
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
                string doc = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "free-items.js");
                string header = doc.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[0] + "TP.FreeItems.specials = [";
                string therest = doc.Split(new string[] { "TP.FreeItems.specials = [" }, StringSplitOptions.None)[1];

                string finalPage = header + finalLine + therest;
                System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "free-items.js", finalPage);

                string doc2 = System.IO.File.ReadAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm");
                string header2 = doc2.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[0] + "my @specials = (";
                string therest2 = doc2.Split(new string[] { "my @specials = (" }, StringSplitOptions.None)[1];

                string finalPage2 = header2 + finalLine2 + therest2;
                System.IO.File.WriteAllText(Common.Variables.DebugFilePath + "AddBonusItems.pm", finalPage2);

                MessageBox.Show("Complete. Debug is on, so check your desktop");
                this.Close();
            }
        }


        private void AddNewLine(string Line)
        {
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
