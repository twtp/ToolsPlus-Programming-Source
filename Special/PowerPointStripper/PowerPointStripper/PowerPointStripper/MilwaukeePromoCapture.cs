using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
namespace MilwaukeePromoCapture
{
    public static class MilwaukeePromoCapture
    {
        public static SlideData ProcessSlide(Slide slide, int slideType)
        {
            if (slideType == 0)
                return new SlideData();
            if (slideType == 1)
                return new SlideData();
            if (slideType == 2)
                //return SlideType2(slide); //may disable this one as its not exactly a promo but 'kit creator'
                return new SlideData();
            if (slideType == 3)
                return new SlideData();
            if (slideType == 4)
                return SlideType4(slide);

            return new SlideData();
        }
        public static int GetSlideLayoutType(Slide slide)
        {
            if (slide.CustomLayout.Name == "NEW PRODUCT & PROMO TITLE")
                return SlideTypes.NewProductPromoTitle;
            if (slide.CustomLayout.Name == "NP M12 ONE SLIDER")
                return SlideTypes.NewProductM12OneSlide;
            if (slide.CustomLayout.Name == "Blank  Slide")
                return SlideTypes.BlankSlide;
            if (slide.CustomLayout.Name == "2_NLP - M12 Single Item")
                return SlideTypes.NLPM12SingleItem;
            if (slide.CustomLayout.Name == "Blank M12 Promo Slide")
                return SlideTypes.BlankM12PromoSlide;
            if (slide.CustomLayout.Name == "M12 no imap Promo Slide")
                return SlideTypes.M12NoIMAPPromoSlide;
            if (slide.CustomLayout.Name == "NP & NEW PRODUCT M18 TITLE")
                return SlideTypes.NewProductM18Title;
            if (slide.CustomLayout.Name == "NP ONE SLIDER")
                return SlideTypes.NewProductOneSlide;
            if (slide.CustomLayout.Name == "NP M18 ONE SLIDER")
                return SlideTypes.NewProductM18OneSlide;
            if (slide.CustomLayout.Name == "NLP PERMANENT M18")
                return SlideTypes.NLPPermanentM18Slide;
            if (slide.CustomLayout.Name == "Blank M18 Slide")
                return SlideTypes.BlankM18Slide;
            if (slide.CustomLayout.Name == "M18 FUEL no imap Promo Slide")
                return SlideTypes.M18FuelNoIMAPPromoSlide;
            if (slide.CustomLayout.Name == "1_Blank M18 FUEL Promo Slide")
                return SlideTypes.BlankM18FuelPromoSlide;
            if (slide.CustomLayout.Name == "Blank M18 Promo Slide")
                return SlideTypes.BlankM18PromoSlide;
            if (slide.CustomLayout.Name == "1_NLP - M18 FUEL Single Item")
                return SlideTypes.NLPM18FuelSingleItemSlide;
            if (slide.CustomLayout.Name == "NLP -M18 Single Item")
                return SlideTypes.NLPM18SingleItemSlide;
            if (slide.CustomLayout.Name == "Blank Promo Slide")
                return SlideTypes.BlankPromoSlide;
            if (slide.CustomLayout.Name == "Blank no imap Promo Slide")
                return SlideTypes.BlankNoIMAPPromoSlide;
            if (slide.CustomLayout.Name == "ACCY NP & PROMO TITLE")
                return SlideTypes.ACCYNPPromoTitleSlide;
            if (slide.CustomLayout.Name == "2_ACCY NP & PROMO TITLE")
                return SlideTypes.ACCYNPPromoTitleSlide2;
            if (slide.CustomLayout.Name == "NP Blank ONE SLIDER")
                return SlideTypes.NPBlankOneSlide;
            if (slide.CustomLayout.Name == "BUY X GET X Promo Slide")
                return SlideTypes.BuyXGetXPromoSlide;
            if (slide.CustomLayout.Name == "T&M NP & PROMO TITLE")
                return SlideTypes.TMNewProductPromoTitleSlide;
            if (slide.CustomLayout.Name == "Blank")
                return SlideTypes.Blank;
            if (slide.CustomLayout.Name == "1_Blank")
                return SlideTypes.Blank2;
            if (slide.CustomLayout.Name == "2_Title Slide")
                return SlideTypes.TitleSlide;

            return 0;
            
        }
        public static class SlideTypes
        {
            public const int NewProductPromoTitle = 1;
            public const int NewProductM12OneSlide = 2;
            public const int BlankSlide = 3;
            public const int NLPM12SingleItem = 4;
            public const int BlankM12PromoSlide = 5;
            public const int M12NoIMAPPromoSlide = 6;
            public const int NewProductM18Title = 7;
            public const int NewProductOneSlide = 8;
            public const int NewProductM18OneSlide = 9;
            public const int NLPPermanentM18Slide = 10;
            public const int BlankM18Slide = 11;
            public const int M18FuelNoIMAPPromoSlide = 12;
            public const int BlankM18FuelPromoSlide = 13;
            public const int BlankM18PromoSlide = 14;
            public const int NLPM18FuelSingleItemSlide = 15;
            public const int NLPM18SingleItemSlide = 16;
            public const int BlankPromoSlide = 17;
            public const int BlankNoIMAPPromoSlide = 18;
            public const int ACCYNPPromoTitleSlide = 19;
            public const int ACCYNPPromoTitleSlide2 = 20;
            public const int NPBlankOneSlide = 21;
            public const int BuyXGetXPromoSlide = 22;
            public const int TMNewProductPromoTitleSlide = 23;
            public const int Blank = 24;
            public const int Blank2 = 25;
            public const int TitleSlide = 26;

        }

        public struct SlideItemData
        {
            public int SlideType;
            public string PCE;
            public string UPC;
            public decimal SpecialPrice;
            public decimal RecommendedSellingPrice;
            public decimal SuggestedAdvertisedSavings;
            public string ItemNumber;
            public string Description;
            public decimal FiftyPercentPrice;
            public decimal HeavyDutyPrice;
            public string PromotionalPrice;
            public decimal iMAP;
        }
        public struct SlideData
        {
            public int SlideType;
            public List<SlideItemData> PromoItems;
            //public DateTime BuyInStartDate;
            //public DateTime BuyInEndDate;
            public DateTime PreOrderStartDate;
            public DateTime PreOrderEndDate;
            public DateTime LaunchShipStartDate;
            public DateTime LaunchShipEndDate;
            public string PCENumber;
            public string tmpPurchaseLine;
            public string tmpReceiveLine;
            public string tmpFreeLine;
        }

        private static SlideData SlideType2(Slide slide)
        {
            int counter = 0;
            string preorderDate = "";
            string launchshipDate = "";
            List<SlideItemData> SD = new List<SlideItemData>();
            foreach (Shape sh in slide.Shapes)
            {
                counter++;
                if (sh.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    //System.Windows.Forms.MessageBox.Show("OFFSET " + counter.ToString() + sh.TextFrame.TextRange.Text);
                    if (counter == 9)
                    {
                        preorderDate = sh.TextFrame.TextRange.Text;
                    }
                    if (counter == 10)
                    {
                        launchshipDate = sh.TextFrame.TextRange.Text;
                    }
                }
                if (sh.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    string tableheader = "";
                    Row r = sh.Table.Rows[1];
                    int itemNumberOffset = 0, descriptionOffset = 0, upcOffset = 0, fiftyoffOffset = 0, heavydutyOffset = 0 , imapOffset = 0;
                    int cCount = 1;
                        foreach (Cell c in r.Cells)
                        {
                            tableheader += c.Shape.TextFrame.TextRange.Text + " | ";
                            if (c.Shape.TextFrame.TextRange.Text.Contains("ITEM #") == true)
                            {
                                itemNumberOffset = cCount;
                            }
                            if (c.Shape.TextFrame.TextRange.Text.Contains("DESCRIPTION") == true)
                            {
                                descriptionOffset = cCount;
                            }
                            if (c.Shape.TextFrame.TextRange.Text.Contains("UPC") == true)
                            {
                                upcOffset = cCount;
                            }
                            if (c.Shape.TextFrame.TextRange.Text.Contains("50% OFF LIST") == true)
                            {
                                fiftyoffOffset = cCount;
                            }
                            if (c.Shape.TextFrame.TextRange.Text.Contains("HEAVY DUTY") == true)
                            {
                                heavydutyOffset = cCount;
                            }
                            if (c.Shape.TextFrame.TextRange.Text.Contains("iMAP") == true)
                            {
                                imapOffset = cCount;
                            }
                            cCount++;
                        }


                        //System.Windows.Forms.MessageBox.Show(itemNumberOffset.ToString());
                        //System.Windows.Forms.MessageBox.Show(descriptionOffset.ToString());
                        //System.Windows.Forms.MessageBox.Show(upcOffset.ToString());
                        //System.Windows.Forms.MessageBox.Show(fiftyoffOffset.ToString());
                        //System.Windows.Forms.MessageBox.Show(heavydutyOffset.ToString());
                        //System.Windows.Forms.MessageBox.Show(imapOffset.ToString());

                        int rowCount = sh.Table.Rows.Count;
                        for (int rc = 2; rc <= rowCount; rc++)
                        {
                            //System.Windows.Forms.MessageBox.Show("Item Number: " + sh.Table.Rows[rc].Cells[itemNumberOffset].Shape.TextFrame.TextRange.Text);
                            //System.Windows.Forms.MessageBox.Show("Description: " + sh.Table.Rows[rc].Cells[descriptionOffset].Shape.TextFrame.TextRange.Text);
                            //System.Windows.Forms.MessageBox.Show("UPC: " + sh.Table.Rows[rc].Cells[upcOffset].Shape.TextFrame.TextRange.Text);
                            //System.Windows.Forms.MessageBox.Show("50% Off List: " + sh.Table.Rows[rc].Cells[fiftyoffOffset].Shape.TextFrame.TextRange.Text);
                            //System.Windows.Forms.MessageBox.Show("Heavy Duty: " + sh.Table.Rows[rc].Cells[heavydutyOffset].Shape.TextFrame.TextRange.Text);
                            //System.Windows.Forms.MessageBox.Show("IMAP: " + sh.Table.Rows[rc].Cells[imapOffset].Shape.TextFrame.TextRange.Text);
                            
                            SlideItemData tmpSlideItem = new SlideItemData();
                            //tmpSlideItem.PreOrderStartDate = DateTime.Parse(preorderDate);
                            //tmpSlideItem.LaunchShipDate = DateTime.Parse(launchshipDate);
                            tmpSlideItem.ItemNumber = sh.Table.Rows[rc].Cells[itemNumberOffset].Shape.TextFrame.TextRange.Text;
                            tmpSlideItem.Description = sh.Table.Rows[rc].Cells[descriptionOffset].Shape.TextFrame.TextRange.Text;
                            tmpSlideItem.UPC = sh.Table.Rows[rc].Cells[upcOffset].Shape.TextFrame.TextRange.Text;
                            tmpSlideItem.FiftyPercentPrice = decimal.Parse(sh.Table.Rows[rc].Cells[fiftyoffOffset].Shape.TextFrame.TextRange.Text.Replace("$",""));
                            tmpSlideItem.HeavyDutyPrice = decimal.Parse(sh.Table.Rows[rc].Cells[heavydutyOffset].Shape.TextFrame.TextRange.Text.Replace("$",""));
                            tmpSlideItem.iMAP = decimal.Parse(sh.Table.Rows[rc].Cells[imapOffset].Shape.TextFrame.TextRange.Text.Replace("$",""));
                            tmpSlideItem.SlideType = 2;
                            SD.Add(tmpSlideItem);
                        }
                    
                    //sh.Table.Rows[1].Cells[1].Shape.TextFrame.TextRange.Text
                    //System.Windows.Forms.MessageBox.Show("TABLE FOUND!  Rows:" + sh.Table.Rows.Count.ToString() + " Cols:" + sh.Table.Columns.Count.ToString());
                    //System.Windows.Forms.MessageBox.Show(table);
                }
                if (sh.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    System.Windows.Forms.MessageBox.Show("has chart!!");
                }
                //System.Windows.Forms.MessageBox.Show("TEXT " + sh.HasTextFrame.ToString());
                //System.Windows.Forms.MessageBox.Show("TABLE " + sh.HasTable.ToString());
                //System.Windows.Forms.MessageBox.Show("CHART" + sh.HasChart.ToString());
                //System.Windows.Forms.MessageBox.Show(s.Name);
            }

            SlideData tmpData = new SlideData();
            tmpData.SlideType = 2;
            tmpData.PromoItems = SD;
            tmpData.PreOrderStartDate = DateTime.Parse(preorderDate);
            tmpData.LaunchShipStartDate = DateTime.Parse(launchshipDate);

            return tmpData;
        }
        private static SlideData SlideType4(Slide slide)
        {
            int counter = 0;            
            string preorderDate = "";
            string preorderEndDate = "";
            string launchshipDate = "";
            string launchshipEndDate = "";
            string pceNumber = "";
            List<SlideItemData> SD = new List<SlideItemData>();
            bool foundPreOrderDate = false;
            bool foundLaunchOrderDate = false;
            int freebieAmount = 0;
            string PurchaseLine = "";
            string ReceiveLine = "";
            string FreeLine = "";
            foreach (Shape sh in slide.Shapes)
            {
                counter++;
                if (sh.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    //System.Windows.Forms.MessageBox.Show("OFFSET " + counter.ToString() + sh.TextFrame.TextRange.Text);
                    string tmpStr = sh.TextFrame.TextRange.Text.Replace("–", "-");
                    if (tmpStr.Length > 0)
                    {
                        if (tmpStr.Contains("PURCHASE") == true)
                        {
                            PurchaseLine = tmpStr;
                        }
                        if (tmpStr.Contains("RECEIVE") == true)
                        {
                            ReceiveLine = tmpStr;
                        }
                        if (tmpStr.Contains("(") == true & tmpStr.Contains(")") == true & tmpStr.Contains("FREE") ==true)
                        {
                            FreeLine = tmpStr;
                            freebieAmount = int.Parse(tmpStr.Split('(')[1].Split(')')[0]);
                        }
                        if (tmpStr.Substring(0, tmpStr.Length / 2).Contains(".") == true)
                        {
                            if (tmpStr.Substring(tmpStr.Length / 2).Contains(".") == true)
                            {
                                if (foundPreOrderDate == false)
                                {
                                    foundPreOrderDate = true;
                                    try
                                    {
                                        preorderDate = tmpStr.Split(' ')[0].Trim();
                                        preorderEndDate = tmpStr.Split('-')[1].Trim();
                                    }
                                    catch { }
                                }
                                else
                                {
                                    foundLaunchOrderDate = true;
                                    try
                                    {
                                        launchshipDate = tmpStr.Split(' ')[0].Trim();
                                        launchshipEndDate = tmpStr.Split('-')[1].Trim();
                                    }
                                    catch { }
                                }
                            }
                        }
                        long t;
                        if (long.TryParse(tmpStr, out t) == true)
                        {
                            if (tmpStr.Length >= 6)
                            {
                                pceNumber = t.ToString();
                            }
                        }

                    }
                }
                if (sh.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    string tableheader = "";
                    Row r = sh.Table.Rows[1];
                    int itemNumberOffset = 0, descriptionOffset = 0, upcOffset = 0, fiftyoffOffset = 0, heavydutyOffset = 0, imapOffset = 0, promopriceOffset =0;
                    int cCount = 1;
                    foreach (Cell c in r.Cells)
                    {
                        string test = c.Shape.TextFrame.TextRange.Text;

                        tableheader += c.Shape.TextFrame.TextRange.Text + " | ";
                        if (c.Shape.TextFrame.TextRange.Text.Contains("ITEM #") == true)
                        {
                            itemNumberOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("DESCRIPTION") == true)
                        {
                            descriptionOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("UPC") == true)
                        {
                            upcOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("50% OFF LIST") == true)
                        {
                            fiftyoffOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("HEAVY") == true & c.Shape.TextFrame.TextRange.Text.Contains("DUTY") == true)
                        {
                            heavydutyOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("HD PRICE") == true)
                        {
                            heavydutyOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("iMAP") == true)
                        {
                            imapOffset = cCount;
                        }
                        if (c.Shape.TextFrame.TextRange.Text.Contains("PROMO PRICE") == true)
                        {
                            promopriceOffset = cCount;
                        }
                        cCount++;
                    }



                    int rowCount = sh.Table.Rows.Count;
                    for (int rc = 2; rc <= rowCount; rc++)
                    {                       
                        SlideItemData tmpSlideItem = new SlideItemData();

                        if (itemNumberOffset > 0)
                        {
                            tmpSlideItem.ItemNumber = sh.Table.Rows[rc].Cells[itemNumberOffset].Shape.TextFrame.TextRange.Text;
                        }
                        if (descriptionOffset > 0)
                        {
                            tmpSlideItem.Description = sh.Table.Rows[rc].Cells[descriptionOffset].Shape.TextFrame.TextRange.Text;
                        }
                       try
                        {
                            tmpSlideItem.HeavyDutyPrice = decimal.Parse(sh.Table.Rows[rc].Cells[heavydutyOffset].Shape.TextFrame.TextRange.Text.Replace("$", ""));
                        }catch{tmpSlideItem.HeavyDutyPrice = 0;}
                        try{tmpSlideItem.iMAP = decimal.Parse(sh.Table.Rows[rc].Cells[imapOffset].Shape.TextFrame.TextRange.Text.Replace("$", ""));}
                        catch{tmpSlideItem.iMAP = 0;}
                        try
                        {
                            tmpSlideItem.PromotionalPrice = decimal.Parse(sh.Table.Rows[rc].Cells[promopriceOffset].Shape.TextFrame.TextRange.Text.Replace("$","")).ToString();
                        }
                        catch
                        {
                            if (sh.Table.Rows[rc].Cells[promopriceOffset].Shape.TextFrame.TextRange.Text.Contains("FREE") == true)
                            {
                                if (freebieAmount > 0)
                                {
                                    tmpSlideItem.PromotionalPrice = "Get " + freebieAmount.ToString() + " Free w/";
                                }
                            }
                        }

                        tmpSlideItem.SlideType = 4;
                        if (tmpSlideItem.ItemNumber != null)
                        {
                            if (tmpSlideItem.ItemNumber != "")
                            {
                                SD.Add(tmpSlideItem);
                            }
                        }
                    }

                }
                if (sh.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    System.Windows.Forms.MessageBox.Show("has chart!!");
                }

            }

            SlideData tmpData = new SlideData();
            tmpData.SlideType = 4;
            tmpData.PromoItems = SD;
            try { tmpData.PreOrderStartDate = DateTime.Parse(preorderDate); }catch { }
            try { tmpData.PreOrderEndDate = DateTime.Parse(preorderEndDate); }catch { }
            try { tmpData.LaunchShipStartDate = DateTime.Parse(launchshipDate); }catch { }
            try { tmpData.LaunchShipEndDate = DateTime.Parse(launchshipEndDate); }catch { }
            tmpData.PCENumber = pceNumber;

            string[] verbs = PurchaseLine.Split(' ');
            bool containsAllVerbs = true;
            foreach (string verb in verbs)
            {
                //foreach item in this promo....
            }
            if (containsAllVerbs == true)
            {

            }
            tmpData.tmpPurchaseLine = PurchaseLine;
            tmpData.tmpReceiveLine = ReceiveLine;
            tmpData.tmpFreeLine = FreeLine;
            
            
            
            return tmpData;
        }
    }
}
