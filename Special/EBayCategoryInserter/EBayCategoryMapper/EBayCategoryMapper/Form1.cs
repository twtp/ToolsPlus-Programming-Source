using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EBayCategoryMapper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void RefreshCategories()
        {
            biLst.Items.Clear();
            biLst.Refresh();
            hgLst.Items.Clear();
            hgLst.Refresh();
            List<string> BIs = Connectivity.SQLCalls.sqlQuery("SELECT EBayCategory.CategoryName FROM EBayCategory WHERE EBayCategory.CategoryID NOT IN (SELECT BusinessIndustrialID FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID IS NOT NULL)");
            List<string> formattedBIs = new List<string>();
            foreach (string bi in BIs)
            {
                formattedBIs.Add(bi.Split('|')[0]);
            }
            biLst.Items.AddRange(formattedBIs.ToArray());


            List<string> HGs = Connectivity.SQLCalls.sqlQuery("SELECT CategoryName FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID IS NULL");
            List<string> formattedHGs = new List<string>();
            foreach (string hg in HGs)
            {
                formattedHGs.Add(hg.Split('|')[0]);
            }
            hgLst.Items.AddRange(formattedHGs.ToArray());

            biCountLbl.Text = "Total: " + biLst.Items.Count.ToString();
            hgCountLbl.Text = "Total: " + hgLst.Items.Count.ToString();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            RefreshForm(sender, e);
            SetAllToolTips();
            RefreshCategories();
            UpdateRemovalLVI();
            
            this.Resize += new EventHandler(RefreshForm);

        }
        private void SetAllToolTips()
        {
            //statuslbl
            //refreshbtn
            //eBILbl
            //eHGLbl
            //bisearchtxt
            //biupbtn
            //bidownbtn
            //hgsearchtxt
            //hgupbtn
            //hgdownbtn
            //bilst
            //hglst
            //bicount
            //hgcount
            //mapbtn
            //currentmapslbl
            //removaltotlbl
            //hgdeletionbtn
            //removallvi
            //removemappingbtn
            ToolTip newTips = new ToolTip();
            newTips.SetToolTip(statusLbl, "Displays the last status change.");
            newTips.SetToolTip(refreshBtn, "Refreshes both category lists.");
            newTips.SetToolTip(biSearchTxt, "Searches for the input in the B&I categories list.");
            newTips.SetToolTip(hgSearchTxt, "Searches for the input in the H&G categories list.");
            newTips.SetToolTip(biUpBtn, "Previous Find");
            newTips.SetToolTip(biDownBtn, "Next Find");
            newTips.SetToolTip(hgUpBtn, "Previous Find");
            newTips.SetToolTip(hgDownBtn, "Next Find");
            newTips.SetToolTip(mapBtn, "Maps the categories");
            newTips.SetToolTip(hgDeletionBtn, "Removes the selected category(s) from the H&G list.");
            newTips.SetToolTip(removeMappingBtn, "Removes the selected category mapping.");
        }

        private void RefreshForm(object sender, EventArgs e)
        {
            //statuslbl
            //refreshbtn
            //eBILbl
            //eHGLbl
            //bisearchtxt
            //biupbtn
            //bidownbtn
            //hgsearchtxt
            //hgupbtn
            //hgdownbtn
            //bilst
            //hglst
            //bicount
            //hgcount
            //mapbtn
            //currentmapslbl
            //removaltotlbl
            //hgdeletionbtn
            //removallvi
            //removemappingbtn

            statusLbl.Top = 10;
            statusLbl.Left = 5;
            refreshBtn.Top = statusLbl.Bottom + 10;
            refreshBtn.Left = 5;
            eBILbl.Top = refreshBtn.Top - 2;
            eBILbl.Left = refreshBtn.Right + 2;
            biSearchTxt.Top = eBILbl.Top + 2;
            biSearchTxt.Left = eBILbl.Right + 20;
            biUpBtn.Top = biSearchTxt.Top;
            biUpBtn.Left = biSearchTxt.Right + 5;
            biDownBtn.Top = biUpBtn.Top;
            biDownBtn.Left = biUpBtn.Right+5;

            biLst.Width = (this.Width / 2) - 6;
            biLst.Left = 1;
            biLst.Top = biDownBtn.Bottom + 3;
            biLst.Height = (this.Height / 3) + 20;
            hgLst.Width = (this.Width / 2) - 15;
            hgLst.Left = (this.Width / 2) + 1;
            hgLst.Top = biLst.Top;
            hgLst.Height = biLst.Height;

            eHGLbl.Top = eBILbl.Top;
            eHGLbl.Left = hgLst.Left + 2;
            hgSearchTxt.Top = eHGLbl.Top + 2;
            hgSearchTxt.Left = eHGLbl.Right + 20;           
            hgUpBtn.Top = hgSearchTxt.Top - 2;
            hgUpBtn.Left = hgSearchTxt.Right + 5;
            hgDownBtn.Top = hgUpBtn.Top;
            hgDownBtn.Left = hgUpBtn.Right + 5;

            biCountLbl.Top = biLst.Bottom + 3;
            biCountLbl.Left = biLst.Left + ((biLst.Width / 2) - (biCountLbl.Width / 2));
            hgCountLbl.Top = hgLst.Bottom + 3;
            hgCountLbl.Left = hgLst.Left + ((hgLst.Width / 2) - (hgCountLbl.Width / 2));

            mapBtn.Left = (this.Width / 2) - (mapBtn.Width / 2);
            mapBtn.Top = biLst.Bottom + 3;

            hgDeletionBtn.Left = hgLst.Right - hgDeletionBtn.Width;
            hgDeletionBtn.Top = hgLst.Bottom + 3;

            currentMapsLbl.Top = mapBtn.Bottom -10;
            currentMapsLbl.Left = eBILbl.Left;
            removalTotLbl.Top = currentMapsLbl.Top;
            removalTotLbl.Left = currentMapsLbl.Right;

            removalLVI.Top = currentMapsLbl.Bottom + 3;
            removalLVI.Left = biLst.Left;
            removalLVI.Width = this.Width - (biLst.Left * 2);
            removalLVI.Height = (this.Height - removalLVI.Top) - 125;

            removeMappingBtn.Top = removalLVI.Bottom + 2;
            removeMappingBtn.Left = (removalLVI.Width / 2) - (removeMappingBtn.Width / 2);

            this.Refresh();

        }
        private void UpdateRemovalLVI()
        {
            removalLVI.Items.Clear();
            removalLVI.Columns.Clear();
            removalLVI.Clear();
            removalLVI.Refresh();

            ColumnHeader BIName = new ColumnHeader();
            ColumnHeader HGName = new ColumnHeader();
            BIName.Text = "Business & Industrial";
            HGName.Text = "Home & Garden";
            BIName.Width = (removalLVI.Width / 2) - 15;
            HGName.Width = (removalLVI.Width / 2) - 15;
            removalLVI.Columns.Add(BIName);
            removalLVI.Columns.Add(HGName);

            List<string> maps = Connectivity.SQLCalls.sqlQuery("SELECT EBayHomeGardenConversionID.CategoryName, EBayCategory.CategoryName FROM EBayHomeGardenConversionID INNER JOIN EBayCategory ON EBayCategory.CategoryID=EBayHomeGardenConversionID.BusinessIndustrialID");
            if (maps.Count > 0)
            {
                foreach (string map in maps)
                {
                    string bi_str = map.Split('|')[1];
                    string hg_str = map.Split('|')[0];
                    ListViewItem newLVI = new ListViewItem();
                    newLVI.Text = bi_str;
                    newLVI.SubItems.Add(hg_str);
                    removalLVI.Items.Add(newLVI);


                }
            }
            removalTotLbl.Text = "Total: " + removalLVI.Items.Count.ToString();
        }
        private void mapBtn_Click(object sender, EventArgs e)
        {
            MapSelectedCategories();
            /*string SelectedBI = biLst.Items[biLst.SelectedIndex].ToString();
            string SelectedHG = hgLst.Items[hgLst.SelectedIndex].ToString();

            //get the B & I ID, then set it for the H & G's BusinessIndustrialID...
            List<string> BI_id = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayCategory WHERE CategoryName='" + SelectedBI + "'");
            if (BI_id.Count > 0)
            {
                string getID = BI_id[0].Split('|')[0];
                try { int.Parse(getID); }
                catch { MessageBox.Show("The ID returned from SQL cannot be parsed as a number!"); return; }

                List<string> HG_id = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE CategoryName='" + SelectedHG + "'");
                if (HG_id.Count > 0)
                {
                    string getHGID = HG_id[0].Split('|')[0];
                    try { int.Parse(getHGID); }
                    catch { MessageBox.Show("The HG ID return from SQL cannot be parsed as a number!"); return; }

                    Connectivity.SQLCalls.sqlQuery("UPDATE EBayHomeGardenConversionID SET BusinessIndustrialID=" + getID + " WHERE CategoryID=" + getHGID);
                    statusLbl.Text = "Status: Updated H&&G ID#" + getHGID.ToString() + " to link to B&&I ID#" + getID.ToString();
                    statusLbl.Refresh();
                }
            }
            else
            {
                MessageBox.Show("Cannot find the ID for " + SelectedBI);
            }
            RefreshCategories();
            UpdateRemovalLVI();
             * */
        }
        private void MapSelectedCategories()
        {
            List<string> BI_Categories = new List<string>();
            try { foreach (string item in biLst.SelectedItems) { BI_Categories.Add(item); } }
            catch { MessageBox.Show("No Business and Industrial Categories Selected"); return; }

            string HG_Category = "";
            try{HG_Category = hgLst.Items[hgLst.SelectedIndices[0]].ToString();}
            catch {MessageBox.Show("No Home and Garden Category Selected");return;}
            
           
         
            List<string> HG_id = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID, CategoryName FROM EBayHomeGardenConversionID WHERE CategoryName='" + HG_Category + "'");
            if (HG_id.Count > 0)
            {
                string getHGID = HG_id[0].Split('|')[0];
                string getHGName = HG_id[0].Split('|')[1];
                try { int.Parse(getHGID); }
                catch { MessageBox.Show("The HG ID returned from SQL cannot be parsed as a number!"); return; }
                foreach (string biCat in BI_Categories)
                {
                    List<string> biIDs = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayCategory WHERE CategoryName='" + biCat + "'");
                    if (biIDs.Count > 0)
                    {
                        string biID = biIDs[0].Split('|')[0];
                        try { int.Parse(biID); }
                        catch { MessageBox.Show("The B&&I ID is not a number!"); return; }

                        Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayHomeGardenConversionID (CategoryName,CategoryID,BusinessIndustrialID) VALUES('" + getHGName + "'," + getHGID + "," + biID + ")");
                        statusLbl.Text = "Status: Updated H&&G ID#" + getHGID.ToString() + " to link to B&&I ID#" + biCat.ToString();
                        statusLbl.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("Couldn't match ID from B&&I category.");
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Cannot find the ID");
            }
            RefreshCategories();
            UpdateRemovalLVI();
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            RefreshCategories();
            UpdateRemovalLVI();
        }

        private void biSearchTxt_TextChanged(object sender, EventArgs e)
        {
            if (biLst.SelectedIndex == -1)
            {
                biLst.SelectedIndex = 0;
            }
            if (biLst.Items[biLst.SelectedIndex].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
            {
                if (sender == biDownBtn)
                {
                    for (int x = biLst.SelectedIndex+1; x < biLst.Items.Count; x++)
                    {
                        if (biLst.Items[x].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                        {
                            biLst.SelectedIndex = x;
                            return;
                        }
                    }
                    for (int x = 0; x <= biLst.SelectedIndex; x++)
                    {
                        if (biLst.Items[x].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                        {
                            biLst.SelectedIndex = x;
                            return;
                        }
                    }
                }
                else if (sender == biUpBtn)
                {
                    for (int x = biLst.SelectedIndex - 1; x > -1; x--)
                    {
                        if (biLst.Items[x].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                        {
                            biLst.SelectedIndex = x;
                            return;
                        }
                    }
                    for (int x = biLst.Items.Count-1; x >= biLst.SelectedIndex; x--)
                    {
                        if (biLst.Items[x-1].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                        {
                            biLst.SelectedIndex = x;
                            return;
                        }
                    }
                }
                else
                {
                    return;
                }

            }

            for (int x = biLst.SelectedIndex; x < biLst.Items.Count; x++)
            {
                if (biLst.Items[x].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                {
                    biLst.SelectedIndex = x;
                    return;
                }
            }
            for (int x = 0; x <= biLst.SelectedIndex; x++)
            {
                if (biLst.Items[x].ToString().ToUpper().Contains(biSearchTxt.Text.ToUpper()) == true)
                {
                    biLst.SelectedIndex = x;
                    return;
                }
            }
        }

        private void biDownBtn_Click(object sender, EventArgs e)
        {
            biSearchTxt_TextChanged(biDownBtn, e);
        }

        private void biUpBtn_Click(object sender, EventArgs e)
        {
            biSearchTxt_TextChanged(biUpBtn, e);
        }
        

        private void hgSearchTxt_TextChanged(object sender, EventArgs e)
        {
            if (hgLst.SelectedIndex == -1)
            {
                hgLst.SelectedIndex = 0;
            }
            if (hgLst.Items[hgLst.SelectedIndex].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
            {
                if (sender == hgDownBtn)
                {
                    for (int x = hgLst.SelectedIndex + 1; x < hgLst.Items.Count; x++)
                    {
                        if (hgLst.Items[x].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                        {
                            hgLst.SelectedIndex = x;
                            return;
                        }
                    }
                    for (int x = 0; x <= hgLst.SelectedIndex; x++)
                    {
                        if (hgLst.Items[x].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                        {
                            hgLst.SelectedIndex = x;
                            return;
                        }
                    }
                }
                else if (sender == hgUpBtn)
                {
                    for (int x = hgLst.SelectedIndex - 1; x > -1; x--)
                    {
                        if (hgLst.Items[x].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                        {
                            hgLst.SelectedIndex = x;
                            return;
                        }
                    }
                    for (int x = hgLst.Items.Count - 1; x >= hgLst.SelectedIndex; x--)
                    {
                        if (hgLst.Items[x - 1].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                        {
                            hgLst.SelectedIndex = x;
                            return;
                        }
                    }
                }
                else
                {
                    return;
                }

            }

            for (int x = hgLst.SelectedIndex; x < hgLst.Items.Count; x++)
            {
                if (hgLst.Items[x].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                {
                    hgLst.SelectedIndex = x;
                    return;
                }
            }
            for (int x = 0; x <= hgLst.SelectedIndex; x++)
            {
                if (hgLst.Items[x].ToString().ToUpper().Contains(hgSearchTxt.Text.ToUpper()) == true)
                {
                    hgLst.SelectedIndex = x;
                    return;
                }
            }
        }

        private void hgUpBtn_Click(object sender, EventArgs e)
        {
            hgSearchTxt_TextChanged(hgUpBtn, e);
        }

        private void hgDownBtn_Click(object sender, EventArgs e)
        {
            hgSearchTxt_TextChanged(hgDownBtn, e);
        }

        private void removeMappingBtn_Click(object sender, EventArgs e)
        {
            ListViewItem currentItem = removalLVI.SelectedItems[0];
            string BI_txt = currentItem.Text;
            string HG_txt = currentItem.SubItems[1].Text;
            List<string> BI_ID = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayCategory WHERE CategoryName='" + BI_txt + "'");
            List<string> HG_ID = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE CategoryName='" + HG_txt + "'");

            string BIid = "";
            string HGid = "";
            if (HG_ID.Count > 0)
            {
                HGid = HG_ID[0].Split('|')[0];
                try { int.Parse(HGid); }
                catch { MessageBox.Show("The Returned H&&G ID is not a number!"); return; }

            }
            else
            {
                MessageBox.Show("Could not resolve categoryID from the H&&G category of " + HG_txt);
                return;
            }

            if (BI_ID.Count > 0)
            {
                BIid = BI_ID[0].Split('|')[0];
                try { int.Parse(BIid); }
                catch { MessageBox.Show("The returned category id from B&&I's \"" + BI_txt + "\" is not a number!"); return; }
                Connectivity.SQLCalls.sqlQuery("DELETE FROM EBayHomeGardenConversionID WHERE CategoryName='" + HG_txt + "' AND BusinessIndustrialID=" + BIid);
                UpdateRemovalLVI();
                RefreshCategories();
                statusLbl.Text = "Status: Unmapped Category \"" + HG_txt + "\" from \"" + BI_txt + "\"";
                statusLbl.Refresh();
            }
            else
            {
                MessageBox.Show("Could not resolve categoryID from the B&&I category of " + BI_txt);
                return;
            }

           //now check to see if you just made the selected home and garden category extinct....
            List<string> anyLeft = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM EBayHomeGardenConversionID WHERE CategoryID=" + HGid);
            if (anyLeft.Count > 0)
            {
                //good to go!
            }
            else
            {
                Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayHomeGardenConversionID (CategoryName,CategoryID) VALUES('" + HG_txt + "'," + HGid + ")");
                statusLbl.Text = "Status: Extincted Category \"" + HG_txt + "\". Re-added to database.";
                statusLbl.Refresh();
            }


            //Connectivity.SQLCalls.sqlQuery("UPDATE EBayHomeGardenConversionID SET BusinessIndustrialID=NULL WHERE CategoryName='" + HG_txt + "'");
           
        }

        private void hgDeletionBtn_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("Do you wish to remove the selected " + (hgLst.SelectedIndices.Count > 1 ? "categories" : "category") + "?", "Confirm Deletion", MessageBoxButtons.YesNo);
            if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
            {
                string hg_txt = "";
                foreach (object index in hgLst.SelectedIndices)
                {
                    hg_txt = hgLst.Items[int.Parse(index.ToString())].ToString();
                    Connectivity.SQLCalls.sqlQuery("DELETE FROM EBayHomeGardenConversionID WHERE CategoryName='" + hg_txt + "'");
                   
                }
                RefreshCategories();
                UpdateRemovalLVI();
                statusLbl.Text = "Status: Removed Category \"" + hg_txt + "\" From DB.";
                statusLbl.Refresh();
            }
        }




    }
}
