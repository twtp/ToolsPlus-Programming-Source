using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CompetitorLister
{
    public partial class competitorFrm : Form
    {
        public competitorFrm()
        {
            InitializeComponent();
        }

        public void PopulateCompetitors()
        {
            competitorLst.Items.Clear();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT CompetitorName FROM Competitors");
            foreach (string comp in results)
            {
                competitorLst.Items.Add(comp.Split('|')[0]);
            }



        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateCompetitors();
        }

        private void competitorLst_SelectedIndexChanged(object sender, EventArgs e)
        {
            string competitorName = competitorLst.Items[competitorLst.SelectedIndex].ToString();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT Enabled,Blacklisted,Whitelisted FROM Competitors WHERE CompetitorName='" + competitorName + "'");
            if (results.Count > 0)
            {
                if (results.Count > 1)
                {
                    MessageBox.Show("This competitor has been listed twice! Tell Tom to clean out the competitor list.");
                }

                enabledChk.Checked = results[0].Split('|')[0].ToUpper() == "TRUE" ? true : false;
                blacklistChk.Checked = results[0].Split('|')[1].ToUpper() == "TRUE" ? true : false;
                whitelistChk.Checked = results[0].Split('|')[2].ToUpper() == "TRUE" ? true : false;


            }
            else
            {
                MessageBox.Show("Could not find competitor info! Perhaps this competitor was deleted since opening this program?");
            }
        }

        private void applyBtn_Click(object sender, EventArgs e)
        {
            string competitorName = competitorLst.Items[competitorLst.SelectedIndex].ToString();
            string enabledValue = enabledChk.Checked == true ? "1" : "0";
            string blacklistValue = blacklistChk.Checked == true ? "1" : "0";
            string whitelistValue = whitelistChk.Checked == true ? "1" : "0";
            Connectivity.SQLCalls.sqlQuery("UPDATE Competitors SET Enabled=" + enabledValue + ", Blacklisted=" + blacklistValue + ", Whitelisted=" + whitelistValue + " WHERE CompetitorName='" + competitorName + "'");
            updateLbl.Visible = true;
            updateTmr.Enabled = true;
        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
            updateLbl.Visible = false;
            updateTmr.Enabled = false;
        }

        private void competitorTxt_TextChanged(object sender, EventArgs e)
        {

            for (int x = 0; x < competitorLst.Items.Count; x++)
            {
                if (competitorLst.Items[x].ToString().ToUpper().Contains(competitorTxt.Text.ToUpper()) == true)
                {
                    competitorLst.SelectedIndex = x;
                    break;
                }
            }
        }

        private void competitorLst_DoubleClick(object sender, EventArgs e)
        {
            if (blacklistChk.Checked == true & enabledChk.Checked == true)
            {
                blacklistChk.Checked = false;
                enabledChk.Checked = false;
                applyBtn_Click(sender, e);
                return;
            }
            if (blacklistChk.Checked == false & enabledChk.Checked == false)
            {
                blacklistChk.Checked = true;
                enabledChk.Checked = true;
                applyBtn_Click(sender, e);
                return;
            }
        }

    }
}
