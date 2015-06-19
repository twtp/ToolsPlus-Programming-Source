using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace ShippingMonitor
{
    public static class FunFacts
    {

        public struct FunFactStruct
        {
            public string FastestPack;
            public string FastestAvgBetweenPacks;
            public int TotalPackagesPacked;
            public int MostPackagesPacked;
            
        }
        public static FunFactStruct newFunFacts;
        public static string FunFactSpacer = "                            ";

        public static void LoadFunFacts(ref ListView PackersListView, ref Label funFactLbl, ref Timer funFactTmr)
        {
            if (PackersListView.Items.Count > 0)
            {
                //foreach (ListViewItem lvi in PackersListView.Items)
                for (int topFour = 0; topFour < 4; topFour++)
                {

                    //FASTEST PACKER
                    string wtf = PackersListView.Items[topFour].SubItems[2].Text;
                    if (wtf != "00:00:00")
                    {
                        if (newFunFacts.FastestPack != null)
                        {

                            if (TimeSpan.Parse(wtf) < TimeSpan.Parse(newFunFacts.FastestPack))
                            {
                                newFunFacts.FastestPack = wtf;
                            }
                        }
                        else
                        {
                            newFunFacts.FastestPack = wtf;
                        }

                    }
                    ///////////////////////////
                    //FASTEST BETWEEN
                    string wtf2 = PackersListView.Items[topFour].SubItems[3].Text;
                    if (wtf2 != "00:00:00")
                    {
                        if (newFunFacts.FastestAvgBetweenPacks != null)
                        {

                            if (TimeSpan.Parse(wtf2) < TimeSpan.Parse(newFunFacts.FastestAvgBetweenPacks))
                            {
                                newFunFacts.FastestAvgBetweenPacks = wtf2;
                            }
                        }
                        else
                        {
                            newFunFacts.FastestAvgBetweenPacks = wtf2;
                        }

                    }

                    string wtf3 = PackersListView.Items[topFour].SubItems[1].Text;
                    if (wtf3 != "0")
                    {
                        if (int.Parse(wtf3) > newFunFacts.MostPackagesPacked)
                        {
                            newFunFacts.MostPackagesPacked = int.Parse(wtf3);
                        }
                    }


                }


                for (int topPackers = 0; topPackers < 4; topPackers++)
                {

                    if (PackersListView.Items[topPackers].SubItems[2].Text == newFunFacts.FastestPack)
                    {
                        PackersListView.Items[topPackers].UseItemStyleForSubItems = false;
                        PackersListView.Items[topPackers].SubItems[2].BackColor = Color.Green;
                        PackersListView.Refresh();
                    }
                    if (PackersListView.Items[topPackers].SubItems[3].Text == newFunFacts.FastestAvgBetweenPacks)
                    {
                        PackersListView.Items[topPackers].UseItemStyleForSubItems = false;
                        PackersListView.Items[topPackers].SubItems[3].BackColor = Color.Green;
                        PackersListView.Refresh();
                    }
                    if (int.Parse(PackersListView.Items[topPackers].SubItems[1].Text) == newFunFacts.MostPackagesPacked)
                    {
                        PackersListView.Items[topPackers].UseItemStyleForSubItems = false;
                        PackersListView.Items[topPackers].SubItems[1].BackColor = Color.Green;
                        PackersListView.Refresh();
                    }

                }




                newFunFacts.TotalPackagesPacked = GetTotalPackages(PackersListView);
                funFactLbl.Text = "Fastest Pack Time: " + newFunFacts.FastestPack + FunFactSpacer;
                funFactLbl.Text += "Best Between Packs Avg: " + newFunFacts.FastestAvgBetweenPacks + FunFactSpacer;
                funFactLbl.Text += "Total Packages Packed: " + newFunFacts.TotalPackagesPacked.ToString() + FunFactSpacer;
                funFactLbl.Text += "Most Packages By Single Packer: " + newFunFacts.MostPackagesPacked.ToString() + FunFactSpacer;
                funFactLbl.Text += "The Time Is: " + DateTime.Now.ToShortTimeString() + FunFactSpacer;
                funFactLbl.Refresh();
                funFactTmr.Enabled = true;
            }
        }


        private static int GetTotalPackages(ListView PackersListView)
        {
            int totPackages = 0;
            for (int x = 0; x < 4; x++)
            {
                totPackages += int.Parse(PackersListView.Items[x].SubItems[1].Text);
            }
            return totPackages;
        }


        public static void FunFactTimer_Tick(ref Label funFactHeaderLbl, ref Label funFactLbl, int formWidth)
        {
            funFactHeaderLbl.BringToFront();
            funFactLbl.Left -= 5;
            if (funFactLbl.Right < 0)
            {
                funFactLbl.Left = formWidth + 200;
            }
        }


    }
}
