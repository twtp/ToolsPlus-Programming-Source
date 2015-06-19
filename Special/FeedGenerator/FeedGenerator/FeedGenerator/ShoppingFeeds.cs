using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FeedGenerator
{
    public static class ShoppingFeeds
    {
        private static class ToBeRemoved
        {
        private struct ItemRebates
        {
            public string[] ItemNumber;
            public void GetItemRebates()
            {                
                List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM vSpecialsActive WHERE ShowOnRebatePage=1 AND FreeShipping=0")).Rows;
                List<string> parsed = new List<string>();
                foreach (string r in res)
                {
                    parsed.Add(r.Split('|')[0]);
                }
                ItemNumber = parsed.ToArray<string>();
            }
        }
        private struct SubItems
        {
            public string[] ItemNumber;
            public string[] ItemNumberPointer;
            public void GetSubItems()
            {
                List<string> items = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT PartNumberGroupLines.ItemNumber, PartNumberGroupMaster.ItemNumberPointer FROM PartNumberGroupMaster INNER JOIN PartNumberGroupLines ON PartNumberGroupMaster.ID=PartNumberGroupLines.GroupID WHERE SortOrder>=0")).Rows;
                List<string> parsedItemNumbers = new List<string>();
                List<string> parsedItemNumberPointers = new List<string>();
                foreach (string s in items)
                {
                    parsedItemNumbers.Add(s.Split('|')[0]);
                    parsedItemNumberPointers.Add(s.Split('|')[1]);
                }
                ItemNumber = parsedItemNumbers.ToArray<string>();
                ItemNumberPointer = parsedItemNumberPointers.ToArray<string>();

            }


        }
        private struct TemplateIDs
        {
            public string[] ID;
            public string[] ItemNumberPointer;
            public void GetTemplateIDs()
            {
                List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT ID,ItemNumberPointer FROM PartNumberGroupMaster")).Rows;
                List<string> parsedID = new List<string>();
                List<string> parsedPointer = new List<string>();
                foreach (string r in res)
                {
                    parsedID.Add(r.Split('|')[0]);
                    parsedPointer.Add(r.Split('|')[1]);
                }
                ID = parsedID.ToArray<string>();
                ItemNumberPointer = parsedPointer.ToArray<string>();
            }
        }
        private struct MASKitContents
        {
            public string[] SalesKitNo;
            public string[] ComponentItemCode;
            public string[] QuantityPerAssembly;
            public void GetMASKitContents()
            {
                List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT SalesKitNo,ComponentItemCode,QuantityPerAssembly FROM IM_SalesKitDetail")).Rows;
                List<string> parsedKitNo = new List<string>();
                List<string> parsedItemCode = new List<string>();
                List<string> parsedQuantity = new List<string>();
                foreach (string r in res)
                {
                    parsedKitNo.Add(r.Split('|')[0]);
                    parsedItemCode.Add(r.Split('|')[1]);
                    parsedQuantity.Add(r.Split('|')[2]);
                }
                SalesKitNo = parsedKitNo.ToArray<string>();
                ComponentItemCode = parsedItemCode.ToArray<string>();
                QuantityPerAssembly = parsedQuantity.ToArray<string>();
            }

        }
        private struct WebsiteURLs
        {
            public string[] WebsiteID;
            public string[] URL;
            public string[] Description;
            public void GetWebsiteURLs()
            {
                List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT WebsiteID,URL,Description FROM Websites")).Rows;
                List<string> parsedID = new List<string>();
                List<string> parsedURL = new List<string>();
                List<string> parsedDescription = new List<string>();

                foreach (string s in res)
                {
                    parsedID.Add(s.Split('|')[0]);
                    parsedURL.Add(s.Split('|')[1]);
                    parsedDescription.Add(s.Split('|')[2]);
                }
                WebsiteID = parsedID.ToArray<string>();
                URL = parsedURL.ToArray<string>();
                Description = parsedDescription.ToArray<string>();
            }

        }
        
        private struct WebPaths
        {

        }



        public struct ShoppingFeed
        {
            public string FeedName;
            public string FeedData;
            public bool SyncImages;
            public bool SendFeed;
            public bool CreateKeywords;
            public bool DidSyncImages;
            public bool DidSendFeed;
            public bool DidCreateKeywords;
        }
        public static ShoppingFeed CreateFeed(string FeedName, bool SyncImages = false, bool SendFeed = false, bool CreateKeywords = false)
        {
            ShoppingFeed tmpFeed = new ShoppingFeed();
            tmpFeed.SyncImages = SyncImages;
            tmpFeed.SendFeed = SendFeed;
            tmpFeed.CreateKeywords = CreateKeywords;

            if (SyncImages == true)
            {
                throw new Exception("Syncing images is currently not supported!");
                return new ShoppingFeed();
            }
            if (SendFeed == true)
            {
                throw new Exception("Sending feed is currently not supported!");
                return new ShoppingFeed();
            }
            if (CreateKeywords == true)
            {
                throw new Exception("Creating Keywords is currently not supported!");
                return new ShoppingFeed();
            }

            //1.) Sync Feeds


            //2.) Create Keywords



            //3.) Get Item Rebates
            //rebates - _get_item_rebates()
            ItemRebates itemRebates = new ItemRebates();
            itemRebates.GetItemRebates();




            //4.) Get Web Paths
            //path_info_for - _get_web_paths()

            //5.) Get Shopping Engine Keywords
            //onoff - _get_shopeng_kwds(@{$opts{engines}})

            //6.) Get Website URLs
            //websites - _get_website_urls()
            WebsiteURLs websiteURLs = new WebsiteURLs();
            websiteURLs.GetWebsiteURLs();

            //7.) Get MAS Kit Contents
            //g_kits - _get_mas_kit_contents()
            MASKitContents masKitContents = new MASKitContents();
            masKitContents.GetMASKitContents();

            //8.) Get Free Shipping Specials
            //freeship - _get_free_shipping_specials()

            //9.) Get Sale Specials***
            //salesspecials - _get_sale_specials() <===my function/method!!!

            //10.) Get Sale Discount Rebates
            //pct_sales - _get_sales_discount_rebates()

            //11.) Get SubItems
            //is_subitem - _get_subitems()
            SubItems subItems = new SubItems();
            subItems.GetSubItems();


            //12.) Get Template IDs
            //template_ids - _get_template_ids()
            TemplateIDs templateIDs = new TemplateIDs();
            templateIDs.GetTemplateIDs();

            //QUERY FIELDS.... AND A WHOLE LOT MORE..

            return tmpFeed;
        }

        }

        public struct FeedColumns
        {
            
        }

    }
}
