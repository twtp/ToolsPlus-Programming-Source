using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PromotionTypes
{
    public static class PromotionTypes
    {
        public struct PromoItemData
        {
            public string MPN;
            public string Description;
            public string UPC;
            public decimal FiftyOffPrice;
            public decimal HeavyDutyPrice;
            public string PromoPrice;
            public decimal iMAP;
        }
        public struct NewItemPromo
        {
            //This promotion is for new promo items or new items coming out.
            //Essentially you have a pre-order date and a launch date.
            //Then you have the list of items, descriptions, upc, 50 off, heavy duty (our price to buy) and iMAP
            public string PromoTitle;
            public List<string> PromoDetails;
            public int PromoType;
            public List<PromoItemData> PromoItemList;
            public DateTime PreOrderStartDate;
            public DateTime PreOrderEndDate;
            public DateTime LaunchStartDate;
            public DateTime LaunchEndDate;
            public string tmpPurchaseLine;
            public string tmpReceiveLine;
            public string tmpFreeLine;
        }
        public static NewItemPromo ConvertSlideDataToNewItemPromo(MilwaukeePromoCapture.MilwaukeePromoCapture.SlideData slideData)
        {            
            if (slideData.PromoItems == null | slideData.PromoItems.Count < 1)
            {
                return new NewItemPromo();
            }
            NewItemPromo tmpItm = new NewItemPromo();
            tmpItm.PromoItemList = new List<PromoItemData>();
            tmpItm.PromoTitle = slideData.SlideType.ToString();
            tmpItm.PromoDetails = new List<string>();
            tmpItm.PreOrderStartDate = slideData.PreOrderStartDate;//slideData.PromoItems[0].PreOrderDate;
            tmpItm.PreOrderEndDate = slideData.PreOrderEndDate;//slideData.PromoItems[0].LaunchShipDate;
            tmpItm.LaunchStartDate = slideData.LaunchShipStartDate;
            tmpItm.LaunchEndDate = slideData.LaunchShipEndDate;
            tmpItm.tmpPurchaseLine = slideData.tmpPurchaseLine;
            tmpItm.tmpReceiveLine = slideData.tmpReceiveLine;
            tmpItm.tmpFreeLine = slideData.tmpFreeLine;
            foreach (MilwaukeePromoCapture.MilwaukeePromoCapture.SlideItemData itemData in slideData.PromoItems)
            {
                PromoItemData pid = new PromoItemData();
                pid.Description = itemData.Description;
                pid.FiftyOffPrice = itemData.FiftyPercentPrice;
                pid.HeavyDutyPrice = itemData.HeavyDutyPrice;
                pid.iMAP = itemData.iMAP;
                pid.MPN = itemData.ItemNumber;
                pid.UPC = itemData.UPC;
                pid.PromoPrice = itemData.PromotionalPrice;
                tmpItm.PromoItemList.Add(pid);               
                
            }
            return tmpItm;
        }



















        //Promo Type #1- Buy Amount Of X, Get Promo Price For X
        //Promo Title
        //Promo Description
        //Promo Details
        //Promo Items {Item #, Description, UPC, PromoPrice, iMAP}
        //Promo Minimum Qty For Promo
        //Promo Can Mix-n-Match
        //Promo Order Start Date
        //Promo Order End Date
        //Promo Start Date
        //Promo End Date
        //Promo PCE
        //Promo Suggested Advertised Savings

        public struct PromoItemDetails
        {
            public string MPN;
            public string Description;
            public string UPC;
            public decimal RegularPrice;
            public decimal PromoPrice;
            public decimal iMAP;
        }
        public struct PromoBuyQtyGetPromoPrice
        {
            public string Title;
            public string Description;
            public string Details;
            public List<PromoItemDetails> Items;
            public int MinimumQty;
            public bool CanMixAndMatch;
            public DateTime OrderStartDate;
            public DateTime OrderEndDate;
            public DateTime PromoStartDate;
            public DateTime PromoEndDate;
            public string PromoPCE;
            public decimal SuggestedAdvertisedSavings;
        }






    }
}
