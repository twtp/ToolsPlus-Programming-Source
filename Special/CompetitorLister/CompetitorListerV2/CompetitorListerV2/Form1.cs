using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CompetitorListerV2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateLists();
        }
        public void ClearLists()
        {
            enabledLst.Items.Clear();
            disabledLst.Items.Clear();
            enabledLst.Refresh();
            disabledLst.Refresh();
        }
        public void PopulateLists()
        {
            ClearLists();
            List<string> enabledCompetitors = Connectivity.SQLCalls.sqlQuery("SELECT CompetitorName FROM Competitors WHERE Enabled=1 AND Blacklisted=0");
            List<string> disabledCompetitors = Connectivity.SQLCalls.sqlQuery("SELECT CompetitorName FROM Competitors WHERE Blacklisted=1");
            if (enabledCompetitors.Count > 0)
            {
                foreach (string competitor in enabledCompetitors)
                {
                    enabledLst.Items.Add(competitor.Split('|')[0]);
                }
            }

            if (disabledCompetitors.Count > 0)
            {
                foreach (string competitor in disabledCompetitors)
                {
                    disabledLst.Items.Add(competitor.Split('|')[0]);
                }
            }


        }

        private void enabledLst_DoubleClick(object sender, EventArgs e)
        {
            ManageEnabledLst();
        }

        private void disabledLst_DoubleClick(object sender, EventArgs e)
        {
            ManageDisabledLst();
        }

        public void ManageDisabledLst()
        {
            List<string> CompetitorsToAdd = new List<string>();
            foreach (int Index in disabledLst.SelectedIndices)
            {
                string Competitor = disabledLst.Items[Index].ToString();
                enabledLst.Items.Add(Competitor);                             
                EnableCompetitor(Competitor);
                CompetitorsToAdd.Add(Competitor);
                disabledLst.Refresh();
            }
            enabledLst.Sorted = true;
            foreach (string Competitor in CompetitorsToAdd)
            {
                disabledLst.Items.Remove(Competitor);
            }
        }
        public void ManageEnabledLst()
        {
            List<string> CompetitorsToRemove = new List<string>();
            foreach (int Index in enabledLst.SelectedIndices)
            {
                string Competitor = enabledLst.Items[Index].ToString();
                disabledLst.Items.Add(Competitor);                
                DisableCompetitor(Competitor);
                CompetitorsToRemove.Add(Competitor);             
                enabledLst.Refresh();
            }
            disabledLst.Sorted = true;

            foreach (string Competitor in CompetitorsToRemove)
            {
                enabledLst.Items.Remove(Competitor);
            }

                
         

        }


        public void DisableCompetitor(string CompetitorName)
        {
            Connectivity.SQLCalls.sqlQuery("UPDATE Competitors SET Blacklisted=1, Whitelisted=0, Enabled=0 WHERE CompetitorName='" + CompetitorName + "'");
        }
        public void EnableCompetitor(string CompetitorName)
        {
            Connectivity.SQLCalls.sqlQuery("UPDATE Competitors SET Blacklisted=0, Whitelisted=1, Enabled=1 WHERE CompetitorName='" + CompetitorName + "'");
        }

        private void moveToDisabledBtn_Click(object sender, EventArgs e)
        {
            ManageEnabledLst();
        }

        private void moveToEnabledBtn_Click(object sender, EventArgs e)
        {
            ManageDisabledLst();
        }



    }
}
