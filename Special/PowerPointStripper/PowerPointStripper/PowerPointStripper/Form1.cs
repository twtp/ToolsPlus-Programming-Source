using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointStripper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView1.View = System.Windows.Forms.View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.Refresh();
            ColumnHeader pID = new ColumnHeader();
            ColumnHeader pPCE = new ColumnHeader();
            ColumnHeader pType = new ColumnHeader();
            ColumnHeader pItem = new ColumnHeader();
            ColumnHeader pCost = new ColumnHeader();
            ColumnHeader pPromo = new ColumnHeader();
            ColumnHeader pIMAP = new ColumnHeader();
            ColumnHeader pDate1 = new ColumnHeader();
            ColumnHeader pDate2 = new ColumnHeader();
            ColumnHeader pDate3 = new ColumnHeader();
            ColumnHeader pDate4 = new ColumnHeader();

            ColumnHeader tmpPurchase = new ColumnHeader();
            ColumnHeader tmpReceive = new ColumnHeader();
            ColumnHeader tmpFree = new ColumnHeader();

            pID.Text = "ID";
            pPCE.Text = "PCE #";
            pType.Text = "Promo Type";
            pItem.Text = "Item Number";
            pCost.Text = "Item Cost";
            pPromo.Text = "Promo Cost";
            pIMAP.Text = "Item iMAP";
            pDate1.Text = "Pre Order Start Date";
            pDate3.Text = "Pre Order End Date";
            pDate2.Text = "Launch Start Date";
            pDate4.Text = "Launch End Date";

            tmpFree.Text = "FREE CODE:";
            tmpPurchase.Text = "PURCHASE CODE:";
            tmpReceive.Text = "RECEIVE CODE:";

            listView1.Columns.Add(pID);
            listView1.Columns.Add(pPCE);
            listView1.Columns.Add(pType);
            listView1.Columns.Add(pItem);
            listView1.Columns.Add(pCost);
            listView1.Columns.Add(pPromo);
            listView1.Columns.Add(pIMAP);
            listView1.Columns.Add(pDate1);
            listView1.Columns.Add(pDate3);
            listView1.Columns.Add(pDate2);
            listView1.Columns.Add(pDate4);
            listView1.Columns.Add(tmpPurchase);
            listView1.Columns.Add(tmpReceive);
            listView1.Columns.Add(tmpFree);

            listBox1.Items.Clear();
            openFileDialog1.Filter = "*.pptx | *.pptx";
            openFileDialog1.ShowDialog();
            fileTxt.Text = openFileDialog1.FileName;
            if (fileTxt.Text == "")
            {
                return;
            }
            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Presentations ppPresens = ppApp.Presentations;
            Presentation objPres = ppPresens.Open(fileTxt.Text, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            Slides objSlides = objPres.Slides;

            int counter = 0;
            foreach (Slide s in objSlides)
            {
                counter++;
                MilwaukeePromoCapture.MilwaukeePromoCapture.SlideData sd = MilwaukeePromoCapture.MilwaukeePromoCapture.ProcessSlide(s,MilwaukeePromoCapture.MilwaukeePromoCapture.GetSlideLayoutType(s));
                if (sd.PromoItems != null)
                {
                    PromotionTypes.PromotionTypes.NewItemPromo nip = PromotionTypes.PromotionTypes.ConvertSlideDataToNewItemPromo(sd);
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = counter.ToString();
                    if (nip.PromoItemList != null)
                    {
                        if (nip.PromoItemList.Count > 0)
                        {
                            foreach (PromotionTypes.PromotionTypes.PromoItemData pid in nip.PromoItemList)
                            {
                                lvi = new ListViewItem();

                                ListViewItem.ListViewSubItem pce = new ListViewItem.ListViewSubItem();
                                pce.Text = sd.PCENumber;

                                ListViewItem.ListViewSubItem type = new ListViewItem.ListViewSubItem();
                                type.Text = nip.PromoTitle;

                                ListViewItem.ListViewSubItem mpn = new ListViewItem.ListViewSubItem();
                                mpn.Text = pid.MPN;

                                ListViewItem.ListViewSubItem cost = new ListViewItem.ListViewSubItem();
                                cost.Text = pid.HeavyDutyPrice.ToString();

                                ListViewItem.ListViewSubItem promo = new ListViewItem.ListViewSubItem();
                                promo.Text = pid.PromoPrice.ToString();

                                ListViewItem.ListViewSubItem map = new ListViewItem.ListViewSubItem();
                                map.Text = pid.iMAP.ToString();

                                ListViewItem.ListViewSubItem preorderstartdate = new ListViewItem.ListViewSubItem();
                                preorderstartdate.Text = sd.PreOrderStartDate.ToShortDateString();//nip.PreOrderDate.ToString();

                                ListViewItem.ListViewSubItem preorderenddate = new ListViewItem.ListViewSubItem();
                                preorderenddate.Text = sd.PreOrderEndDate.ToShortDateString();


                                ListViewItem.ListViewSubItem launchstartdate = new ListViewItem.ListViewSubItem();
                                launchstartdate.Text = sd.LaunchShipStartDate.ToShortDateString();//nip.LaunchDate.ToString();

                                ListViewItem.ListViewSubItem launchenddate = new ListViewItem.ListViewSubItem();
                                launchenddate.Text = sd.LaunchShipEndDate.ToShortDateString();

                                ListViewItem.ListViewSubItem tmpPurchaseLine = new ListViewItem.ListViewSubItem();
                                tmpPurchaseLine.Text = sd.tmpPurchaseLine;
                                ListViewItem.ListViewSubItem tmpReceiveLine = new ListViewItem.ListViewSubItem();
                                tmpReceiveLine.Text = sd.tmpReceiveLine;
                                ListViewItem.ListViewSubItem tmpFreeLine = new ListViewItem.ListViewSubItem();
                                tmpFreeLine.Text = sd.tmpFreeLine;

                                //lvi.SubItems.Add(counter.ToString());
                                lvi.SubItems.Add(pce);
                                lvi.SubItems.Add(type);
                                lvi.SubItems.Add(mpn);
                                lvi.SubItems.Add(cost);
                                lvi.SubItems.Add(promo);
                                lvi.SubItems.Add(map);
                                lvi.SubItems.Add(preorderstartdate);
                                lvi.SubItems.Add(preorderenddate);
                                lvi.SubItems.Add(launchstartdate);
                                lvi.SubItems.Add(launchenddate);
                                lvi.SubItems.Add(tmpPurchaseLine);
                                lvi.SubItems.Add(tmpReceiveLine);
                                lvi.SubItems.Add(tmpFreeLine);
                                listView1.Items.Add(lvi);
                            }
                        }
                    }
                    //listBox1.Items.Add(s.CustomLayout.Name + " - ID#" + MilwaukeePromoCapture.MilwaukeePromoCapture.GetSlideLayoutType(s).ToString());

                }
                    

            }
            //MessageBox.Show(objSlides[1].CustomLayout.Name);


            //Microsoft.Office.Interop.PowerPoint.SlideShowWindow objSSWs;
            //Microsoft.Office.Interop.PowerPoint.SlideShowSettings objSSS;
            //objSSS = objPres.SlideShowSettings;
            //objSSS.Run();
            
            ppApp.Quit();
        }




        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
    }
}
