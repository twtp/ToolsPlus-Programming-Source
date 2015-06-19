using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Net;
namespace RealTimePricesConsole
{
    class Program
    {
        internal class EBayConstants
        {
            public const string SortByBestMatch = "sortOrder=BestMatch";
            public const string SortByFewestBidCount = "sortOrder=BidCountFewest";
            public const string SortByMostBidCount = "sortOrder=BidCountMost";
            public const string SortByCountryAscending = "sortOrderCountryAscending";
            public const string SortByCountryDescending = "sortOrderCountryDescending";
            public const string SortByHighestPrice = "sortOrderCurrentPriceHighest";
            public const string SortByDistance = "sortOrderDistanceNearest";
            public const string SortByTimeEnding = "sortOrderEndTimeSoonest";
            public const string SortByPriceAndShippingHighest = "sortOrder=PricePlusShippingHighest";
            public const string SortByPriceAndShippingLowest = "sortOrder=PricePlusShippingLowest";
            public const string SortByStartTime = "sortOrder=StartTimeNewest";

            public const string BaseSearchURL = "http://svcs.ebay.com/services/search/FindingService/v1?";
            public const string AdvancedSearchAPI = "OPERATION-NAME=findItemsAdvanced";
            public const string ServiceVersion = "SERVICE-VERSION=1.7.0";
            public const string SecurityApp = "SECURITY-APPNAME=ToolsPlu-e1bd-4921-ac6b-14497e74488a";
            public const string XMLResponseFormat = "RESPONSE-DATA-FORMAT=XML";
            public const string XMLRESTStatement = "REST-PAYLOAD";
            public const string IncludeSellerInfo = "outputSelector=SellerInfo";
            public const string BuyerPostalCode = "buyerPostalCode=06712";

        }
        private static Dictionary<int, EBaySearchResult> EBaySearchResults = new Dictionary<int, EBaySearchResult>();
        private struct EBaySearchResult
        {
            //public string ItemImageURL;
            public string ItemPageURL;
            public string ItemTitle;
            public string ItemSubtitle;
            //public decimal ItemPrice;
            //public decimal ItemListPrice;
            //public string ItemBuyItNowOrBestOffer;
            public string ItemShippingType;
            public string ItemUnknownField;
            //public string ItemFastNFreeOrSavings;
            public string ItemSellerAuthorizationImageURL;

            public string ItemID;
            public string ItemCategoryID;
            public string ItemCategoryName;
            public string SellerName;
            public string SellerLocation;
            public string SellerPostalCode;
            public string SellerCountry;
            public int SellerFeedbackScore;
            public decimal SellerFeedbackPercentage;
            public string SellerFeedbackRatingStar;
            public bool SellerTopRated;

            public bool SellerExpeditedShipping;
            public bool SellerOneDayShipping;
            public int SellerHandlingTime;
            public string SellerPriceCurrency;
            public decimal SellerPrice;
            public string SellerConvertedCurrency;
            public decimal SellerConvertedPrice;
            public bool SellerAllowsReturns;

            public string ItemSellingState;
            public bool ItemBestOfferEnabled;
            public bool ItemBuyItNowEnabled;
            public DateTime ItemAuctionStartTime;
            public DateTime ItemAuctionEndTime;
            public string ItemListingType;
            public bool ItemIsGift;



            public int ItemConditionID;
            public string ItemConditionName;
            public bool ItemIsMultiVariationListing;
            public bool ItemTopRated;

            public string SellerShippingServiceType;
            public string SellerShippingServiceCurrency;
            public decimal SellerShippingServiceCost;

        }

        static void Main(string[] args)
        {
            string SearchString = args[0];

            string searchURL = "";
            searchURL += EBayConstants.BaseSearchURL;
            searchURL += EBayConstants.AdvancedSearchAPI;
            searchURL += "&" + EBayConstants.ServiceVersion;
            searchURL += "&" + EBayConstants.SecurityApp;
            searchURL += "&" + EBayConstants.XMLResponseFormat;
            searchURL += "&" + EBayConstants.XMLRESTStatement;
            searchURL += "&" + EBayConstants.IncludeSellerInfo;
            searchURL += "&" + EBayConstants.SortByPriceAndShippingLowest;
            searchURL += "&" + EBayConstants.BuyerPostalCode;
            searchURL += "&keywords=" + System.Web.HttpUtility.UrlEncode(SearchString);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(searchURL);
            string results = "";
            using (System.IO.Stream s = request.GetResponse().GetResponseStream())
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                {
                    results = sr.ReadToEnd();
                }
            }

            System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\ebay_api_search.txt", results);
            string htmlOut = "";
            htmlOut = "<html><body>" + MakeEbayResponseTableHTMLFromXML(results) + "</body></html>";

            Console.WriteLine("<!--EBAY-->\r\n" + htmlOut);
        }

        private static void PopulateSearchResults(string XMLResponse)
        {
            EBaySearchResults = new Dictionary<int, EBaySearchResult>();

            int resultsCount = int.Parse(XMLResponse.Split(new string[] { "<searchResult count=\"" }, StringSplitOptions.None)[1].Split('"')[0]);




            for (int x = 1; x <= resultsCount; x++)
            {
                EBaySearchResult esr = new EBaySearchResult();

                string xmlSegment = XMLResponse.Split(new string[] { "<item>" }, StringSplitOptions.None)[x];

                try
                {
                    esr.ItemAuctionEndTime = DateTime.Parse(xmlSegment.Split(new string[] { "<endTime>" }, StringSplitOptions.None)[1].Split('<')[0]);
                }
                catch
                {
                    esr.ItemAuctionEndTime = new DateTime();
                }
                try
                {
                    esr.ItemAuctionStartTime = DateTime.Parse(xmlSegment.Split(new string[] { "<startTime>" }, StringSplitOptions.None)[1].Split('<')[0]);
                }
                catch
                {
                    esr.ItemAuctionStartTime = new DateTime();
                }
                try
                {
                    esr.ItemBestOfferEnabled = (xmlSegment.Split(new string[] { "<bestOfferEnabled>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.ItemBestOfferEnabled = false;
                }
                try
                {
                    esr.ItemBuyItNowEnabled = (xmlSegment.Split(new string[] { "<buyItNowAvailable>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);

                }
                catch
                {
                    esr.ItemBuyItNowEnabled = false;
                }
                try
                {
                    esr.ItemCategoryID = xmlSegment.Split(new string[] { "<categoryId>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemCategoryID = "";
                }
                try
                {
                    esr.ItemCategoryName = xmlSegment.Split(new string[] { "<categoryName>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemCategoryName = "";
                }
                try
                {
                    esr.ItemConditionID = int.Parse(xmlSegment.Split(new string[] { "<conditionId>" }, StringSplitOptions.None)[1].Split('<')[0]);

                }
                catch
                {
                    esr.ItemConditionID = 0;
                }
                try
                {
                    esr.ItemConditionName = xmlSegment.Split(new string[] { "<conditionDisplayName>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemConditionName = "";
                }
                try
                {
                    esr.ItemID = xmlSegment.Split(new string[] { "<itemId>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemID = "";
                }

                try
                {
                    esr.ItemIsGift = (xmlSegment.Split(new string[] { "<gift>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.ItemIsGift = false;
                }
                try
                {
                    esr.ItemIsMultiVariationListing = (xmlSegment.Split(new string[] { "<isMultiVariationListing>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.ItemIsMultiVariationListing = false;
                }
                try
                {
                    esr.ItemListingType = xmlSegment.Split(new string[] { "<listingType>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemListingType = "";
                }
                try
                {
                    esr.ItemPageURL = xmlSegment.Split(new string[] { "<viewItemURL>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemPageURL = "";
                }
                try
                {
                    esr.ItemSellerAuthorizationImageURL = xmlSegment.Split(new string[] { "<galleryURL>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemSellerAuthorizationImageURL = "";
                }
                try
                {
                    esr.ItemSellingState = xmlSegment.Split(new string[] { "<sellingState>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemSellingState = "";
                }
                try
                {
                    esr.ItemShippingType = xmlSegment.Split(new string[] { "<shippingType>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemShippingType = "";
                }
                try
                {
                    esr.ItemSubtitle = xmlSegment.Split(new string[] { "<subtitle>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemSubtitle = "";
                }
                try
                {
                    esr.ItemTitle = xmlSegment.Split(new string[] { "<title>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemTitle = "";
                }
                try
                {
                    esr.ItemTopRated = (xmlSegment.Split(new string[] { "<topRatedListing>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.ItemTopRated = false;
                }
                try
                {
                    esr.SellerAllowsReturns = (xmlSegment.Split(new string[] { "<returnsAccepted>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.SellerAllowsReturns = false;
                }
                try
                {
                    esr.SellerConvertedCurrency = xmlSegment.Split(new string[] { "<convertedCurrentPrice currencyId=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                }
                catch
                {
                    esr.SellerConvertedCurrency = "";
                }
                try
                {
                    esr.SellerConvertedPrice = decimal.Parse(xmlSegment.Split(new string[] { "</convertedCurrentPrice>" }, StringSplitOptions.None)[0].Split('>')[xmlSegment.Split(new string[] { "</convertedCurrentPrice>" }, StringSplitOptions.None)[0].Split('>').GetLength(0) - 1]);
                }
                catch
                {
                    esr.SellerConvertedPrice = 0;
                }
                try
                {
                    esr.SellerCountry = xmlSegment.Split(new string[] { "<country>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerCountry = "";
                }
                try
                {
                    esr.SellerExpeditedShipping = (xmlSegment.Split(new string[] { "<expeditedShipping>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.SellerExpeditedShipping = false;
                }
                try
                {
                    esr.SellerFeedbackPercentage = decimal.Parse(xmlSegment.Split(new string[] { "<positiveFeedbackPercent>" }, StringSplitOptions.None)[1].Split('<')[0]);
                }
                catch
                {
                    esr.SellerFeedbackPercentage = 0;
                }
                try
                {
                    esr.SellerFeedbackRatingStar = xmlSegment.Split(new string[] { "<feedbackRatingStar>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerFeedbackRatingStar = "";
                }
                try
                {
                    esr.SellerFeedbackScore = int.Parse(xmlSegment.Split(new string[] { "<feedbackScore>" }, StringSplitOptions.None)[1].Split('<')[0]);
                }
                catch
                {
                    esr.SellerFeedbackScore = 0;
                }
                try
                {
                    esr.SellerHandlingTime = int.Parse(xmlSegment.Split(new string[] { "<handlingTime>" }, StringSplitOptions.None)[1].Split('<')[0]);
                }
                catch
                {
                    esr.SellerHandlingTime = 0;
                }
                try
                {
                    esr.SellerLocation = xmlSegment.Split(new string[] { "<location>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerLocation = "";
                }
                try
                {
                    esr.SellerName = xmlSegment.Split(new string[] { "<sellerUserName>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerName = "";
                }
                try
                {
                    esr.SellerOneDayShipping = (xmlSegment.Split(new string[] { "<oneDayShippingAvailable>" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);

                }
                catch
                {
                    esr.SellerOneDayShipping = false;
                }
                try
                {
                    esr.SellerPostalCode = xmlSegment.Split(new string[] { "<postalCode>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerPostalCode = "";
                }
                try
                {
                    //MessageBox.Show(xmlSegment.Split(new string[] { "</currentPrice>" }, StringSplitOptions.None)[0].Split('>')[xmlSegment.Split(new string[] { "</currentPrice>" }, StringSplitOptions.None)[0].Split('>').GetLength(0) - 1]);
                    //MessageBox.Show(xmlSegment.Split(new string[] { "</currentPrice>" }, StringSplitOptions.None)[0].Split('>')[1]);
                    esr.SellerPrice = decimal.Parse(xmlSegment.Split(new string[] { "</currentPrice>" }, StringSplitOptions.None)[0].Split('>')[xmlSegment.Split(new string[] { "</currentPrice>" }, StringSplitOptions.None)[0].Split('>').GetLength(0) - 1]);
                }
                catch
                {
                    esr.SellerPrice = 0;
                }
                try
                {
                    esr.SellerPriceCurrency = xmlSegment.Split(new string[] { "<currentPrice currencyId=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                }
                catch
                {
                    esr.SellerPriceCurrency = "";
                }
                try
                {
                    esr.SellerTopRated = (xmlSegment.Split(new string[] { "<topRatedSeller" }, StringSplitOptions.None)[1].Split('<')[0].ToLower() == "false" ? false : true);
                }
                catch
                {
                    esr.SellerTopRated = false;
                }
                try
                {
                    esr.SellerShippingServiceType = xmlSegment.Split(new string[] { "<shippingType>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.SellerShippingServiceType = "";
                }
                try
                {
                    esr.SellerShippingServiceCost = decimal.Parse(xmlSegment.Split(new string[] { "</shippingServiceCost>" }, StringSplitOptions.None)[0].Split('>')[xmlSegment.Split(new string[] { "</shippingServiceCost>" }, StringSplitOptions.None)[0].Split('>').GetLength(0) - 1]);
                }
                catch
                {
                    esr.SellerShippingServiceCost = 0;
                }
                try
                {
                    esr.SellerShippingServiceCurrency = xmlSegment.Split(new string[] { "<shippingServiceCost currencyId=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                }
                catch
                {
                    esr.SellerShippingServiceCurrency = "";
                }
                EBaySearchResults.Add(x, esr);
            }


        }
        private static string MakeEbayResponseTableHTMLFromXML(string XMLResponse)
        {
            string Response = "";

            //check for errors and handle...
            if (XMLResponse.Contains("<ack>Success</ack>") == false)
            {
                Response = "<table width=\"95%\" border=\"1\"><tr><td>Failure To Return EBay Results</td></tr></table>";
                return Response;
            }

            PopulateSearchResults(XMLResponse);

            foreach (EBaySearchResult esr in EBaySearchResults.Values)
            {
                Response += "<table width=\"95%\" border=\"1\"><tr><td rowspan=\"6\" width=\"20%\"><a href=\"" + esr.ItemPageURL + "\"><img src=\"" + esr.ItemSellerAuthorizationImageURL + "\"></a></td>";
                Response += "<td>" + esr.SellerName + "</td></tr>";
                Response += "<tr><td>" + esr.ItemTitle + "</td></tr>";
                Response += "<tr><td>$" + esr.SellerConvertedPrice.ToString() + " " + esr.SellerConvertedCurrency + "</td></tr>";
                Response += "<tr><td>Shipping: $" + esr.SellerShippingServiceCost.ToString() + " " + esr.SellerShippingServiceCurrency + " (" + esr.SellerShippingServiceType + ")</td></tr>";
                Response += "<tr><td>Shipping Types: " + (esr.SellerOneDayShipping == true ? "One Day Shipping; " : "") + (esr.SellerExpeditedShipping == true ? "Expedited; " : "") + (esr.SellerHandlingTime > 0 ? "Handling Time: " + esr.SellerHandlingTime.ToString() : "") + "<td></tr>";
                Response += "</table><br>";
            }

            return Response;
        }
    }
}
