using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace LivePriceComparer
{
    public partial class Form1
    {
        private struct EBaySearchResult
        {
            public string ItemImageURL;
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
        private void ebayBrowser_Navigating(object sender, System.Windows.Forms.WebBrowserNavigatingEventArgs e)
        {
            if (e.Url.ToString().ToLower().Contains("arbitrarymotorola.plex?") == true)
            {

            }
            else
            {
                if (isNavigating == true)
                {
                    return;
                }
                else
                {
                    //cancel the current event
                    e.Cancel = true;
                    //this opens the URL in the user's default browser
                    System.Diagnostics.Process.Start(e.Url.ToString());
                }
            }
        }
        private void ebayBrowser_DocumentCompleted(object sender, System.Windows.Forms.WebBrowserDocumentCompletedEventArgs e)
        {
            if (e.Url.ToString().ToLower().Contains("arbitrarymotorola.plex?") == true)
            {
                //update cache..
                if (e.Url.ToString().ToLower().Contains("blacklisted=1") == true)
                {
                    //was blacklisted
                    string competitorName = e.Url.ToString().Split('\'')[1];
                    Competitor old = CompetitorsData[competitorName.ToLower()];
                    Competitor newCompetitor = new Competitor();

                    newCompetitor.AddTax = old.AddTax;
                    newCompetitor.Blacklisted = true;
                    newCompetitor.CompetitorName = old.CompetitorName;
                    newCompetitor.Enabled = old.Enabled;
                    newCompetitor.Whitelisted = old.Whitelisted;

                    CompetitorsData.Remove(old.CompetitorName);
                    CompetitorsData.Add(newCompetitor.CompetitorName, newCompetitor);


                }
                else if (e.Url.ToString().ToUpper().Contains("INSERT") == true)
                {
                    //default all vars; blacklisted =1...
                    Competitor newOne = new Competitor();
                    newOne.AddTax = false;
                    newOne.Blacklisted = true;
                    newOne.CompetitorName = e.Url.ToString().Split('\'')[1];
                    newOne.Enabled = true;
                    newOne.Whitelisted = false;
                    CompetitorsData.Add(newOne.CompetitorName, newOne);

                }
                else
                {
                    //was un-blacklisted...
                    string competitorName = e.Url.ToString().Split('\'')[1];
                    Competitor old = CompetitorsData[competitorName.ToLower()];
                    Competitor newCompetitor = new Competitor();

                    newCompetitor.AddTax = old.AddTax;
                    newCompetitor.Blacklisted = false;
                    newCompetitor.CompetitorName = old.CompetitorName;
                    newCompetitor.Enabled = old.Enabled;
                    newCompetitor.Whitelisted = old.Whitelisted;

                    CompetitorsData.Remove(old.CompetitorName);
                    CompetitorsData.Add(newCompetitor.CompetitorName, newCompetitor);
                }
                LastSearch = "";
                FindBtn_Click(sender, null);
            }
            else
            {
                if (isNavigating == true)
                {
                    isNavigating = false;
                }
            }
        }
        

        
        private static string MakeEbayResponseTableHTMLFromXML(string XMLResponse)
        {
            //System.IO.File.WriteAllText("C:\\users\\tomwestbrook\\desktop\\ebay_xml.txt",XMLResponse);
            string Response = "";

            string Comment = "<!--";

            //check for errors and handle...
            if (XMLResponse.Contains("<ack>Success</ack>") == false)
            {
                Response = "<table width=\"95%\" border=\"1\" bgcolor=\"#FFFFFF\"><tr><td>Failure To Return EBay Results</td></tr></table>";
                return Response;
            }

            PopulateSearchResults(XMLResponse);

            //try to sort results by putting calculated / N/A shipping at the bottom of the barrel.
            List<EBaySearchResult> SearchResults = new List<EBaySearchResult>();

            foreach (EBaySearchResult esr in EBaySearchResults.Values)
            {
                if (esr.SellerShippingServiceCost > 0)
                {
                    SearchResults.Add(esr);
                }
                if (esr.SellerShippingServiceType.ToUpper() == "FREE")
                {
                    SearchResults.Add(esr);
                }

            }
            foreach (EBaySearchResult esr in EBaySearchResults.Values)
            {
                if (esr.SellerShippingServiceType.ToUpper() != "FREE" & esr.SellerShippingServiceCost == 0)
                {
                    //bottom of barrel
                    SearchResults.Add(esr);
                }
            }
            if (SearchResults.Count != EBaySearchResults.Values.Count)
            {
                MessageBox.Show("Whoa buddy... filtering search results yielded less than input!");
            }



            //foreach (EBaySearchResult esr in EBaySearchResults.Values)
            foreach (EBaySearchResult esr in SearchResults)
            {
                bool writeResult = true;

                if (FilterBlacklistedCompetitors == true)
                {
                    if (CompetitorsData.ContainsKey(esr.SellerName.ToLower()) == true)
                    {
                        if (CompetitorsData[esr.SellerName.ToLower()].Enabled == true)
                        {
                            if (CompetitorsData[esr.SellerName.ToLower()].Blacklisted == true)
                            {
                                writeResult = false;
                            }
                        }
                    }
                }
                if (ShowOnlyNewItems == true)
                {
                    if (esr.ItemConditionID != 1000)
                    {
                        writeResult = false;
                    }
                }

                //Response += "<table width=\"95%\" border=\"1\"><tr><td rowspan=\"6\" width=\"20%\"><a href=\"" + esr.ItemPageURL + "\"><img src=\"" + esr.ItemSellerAuthorizationImageURL + "\"></a></td>";
                //Response += "<td>" + esr.SellerName + "</td></tr>";
                //Response += "<tr><td>" + esr.ItemTitle + "</td></tr>";
                //Response += "<tr><td>$" + esr.SellerConvertedPrice.ToString() + " " + esr.SellerConvertedCurrency + "</td></tr>";
                //Response += "<tr><td>Shipping: $" + esr.SellerShippingServiceCost.ToString() + " " + esr.SellerShippingServiceCurrency + " (" + esr.SellerShippingServiceType + ")</td></tr>";
                //Response += "<tr><td>Shipping Types: " + (esr.SellerOneDayShipping == true ? "One Day Shipping; " : "") + (esr.SellerExpeditedShipping == true ? "Expedited; " : "") + (esr.SellerHandlingTime > 0 ? "Handling Time: " + esr.SellerHandlingTime.ToString() : "") + "<td></tr>";
                //Response += "</table><br>";
                if (writeResult == true)
                {
                    string blacklistingLink = (CompetitorsData.ContainsKey(esr.SellerName.ToLower()) == true ? (CompetitorsData[esr.SellerName.ToLower()].Blacklisted == false ?  "<a href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=1%20WHERE%20CompetitorName=%27" + esr.SellerName.ToLower() + "%27\" class=\"Blacklist\">Blacklist</a>" : "<a href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=0%20WHERE%20CompetitorName=%27" + esr.SellerName.ToLower() + "%27\" class=\"Blacklist\">Un-Blacklist</a>") : "<a href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?INSERT INTO Competitors (CompetitorName,Enabled,Blacklisted,Whitelisted,AddTax) VALUES('" + esr.SellerName.ToLower() + "',1,1,0,0)\" class=\"Blacklist\">Blacklist</a>");
                    string showPicture = "";
                    //showPicture = (ShowPicture == true ? "<td rowspan=\"5\" width=\"20%\"><a href=\"" + esr.ItemPageURL + "\"><img height=\"50px\" src=\"" + esr.ItemImageURL + "\" style=\"border-style: none\"></a></td>" : "");
                    showPicture = (ShowPicture == true ? "<td rowspan=\"5\" style=\"word-wrap:break-word;width=64px;\"><!--{ITEMURL}--><a href=\"" + esr.ItemPageURL + "\"><!--{/ITEMURL}--><img width=\"64px\" src=\"" + esr.ItemImageURL + "\" style=\"border-style: none\"></a></td>" : "");

                    string sellerLine = "<!--{SELLER}-->";
                    sellerLine += (ShowCompetitorName == true ? esr.SellerName + " " + blacklistingLink + " " : "");
                    sellerLine += (ShowCompetitorFeedbackPercent == true ? "" + esr.SellerFeedbackPercentage.ToString() + "% " : "");
                    sellerLine += (ShowCompetitorFeedbackScore == true ? "<br>Feedback Score: " + esr.SellerFeedbackScore : "");

                    if (sellerLine.Length > 0)
                    {
                        sellerLine = "<td><font color=\"#555555\">" + sellerLine + "</font><!--{/SELLER}--></td></tr>";
                    }
                   
                    string productPricing = "";
                    //productPricing = (ShowProductPricing == true ? "<tr><td><font face=\"arial\">$" + esr.SellerConvertedPrice.ToString("0.00") + " " + (esr.SellerShippingServiceCost > 0 ? " + $" + esr.SellerShippingServiceCost.ToString("0.00") + "s/h " + " (<b><font color=\"green\">$" + (esr.SellerShippingServiceCost + esr.SellerConvertedPrice).ToString("0.00") + "</font></b>)" : (esr.ItemShippingType.ToLower().Trim() == "free" ? " (<font color=\"green\"><b>$" + esr.SellerConvertedPrice.ToString("0.00") + "</b></font>)" : "+" + "</font></td></tr>")) : "");

                    if (ShowProductPricing == true)
                    {
                        if (esr.SellerShippingServiceCost > 0)
                        {
                            productPricing = "<tr><td><!--{PRICE}--><b><font face=\"arial\">(<font color=\"green\">$" + (esr.SellerShippingServiceCost + esr.SellerConvertedPrice).ToString("0.00") + "</font>)</b> $" + esr.SellerConvertedPrice.ToString("0.00");                           
                            productPricing += " + $" + esr.SellerShippingServiceCost.ToString("0.00") + "<font style=\"font-size:11px\">s/h</font>";

                            productPricing += "<!--{/PRICE}--></td></tr>";
                        }
                        else if (esr.SellerShippingServiceCost == 0 & esr.SellerShippingServiceType.ToUpper() == "FREE")
                        {
                            productPricing = "<tr><td><!--{PRICE}--><b><font face=\"arial\">(<font color=\"green\">$" + (esr.SellerConvertedPrice).ToString("0.00") + "</font>) ";
                            productPricing += "</b><!--{/PRICE}--></td></tr>";
                        }
                        else
                        {
                            productPricing = "<tr><td><!--{PRICE}--><b><font face=\"arial\">(<a href=\"#\" title=\"Shipping Type: " + esr.SellerShippingServiceType + "\" class=\"Popup\"><font color=\"green\">$" + (esr.SellerConvertedPrice).ToString("0.00") + "</a></font>)* ";
                            productPricing += "</b><!--{/PRICE}--></td></tr>";
                        }

                        

                    }


                    string productTitle = "";
                    productTitle = (ShowProductTitle == true ? "<tr><td><font size=\"2\">" + esr.ItemTitle + "</font></td></tr>" : "");


                    Response += "<table width=\"95%\" bgcolor=\"#FFFFFF\"><tr>" + showPicture;
                    Response += sellerLine;
                    Response += productTitle;
                    Response += productPricing;

                    string shippingDetails = "";
                    shippingDetails += (ShowShippingType == true ? "<font size=\"2\">Shipping: " + esr.ItemShippingType + "</font><br>" : "");
                    shippingDetails += (ShowExpedited == true ? "<font size=\"2\">Expedited: " + (esr.SellerExpeditedShipping == true ? "Yes" : "No") + "</font><br>" : "");
                    shippingDetails += (ShowOneDay == true ? "<font size=\"2\">One Day: " + (esr.SellerOneDayShipping == true ? "Yes" : "No") + "</font><br>" : "");
                    shippingDetails += (ShowShippingServices == true ? "<font size=\"2\">Services: " + esr.SellerShippingServiceType + "<br>" : "");
                    shippingDetails += (ShowHandlingTime == true ? "<font size=\"2\">Handling Time: " + esr.SellerHandlingTime + " days</font>" : "");
                    if (shippingDetails.Length > 0)
                    {
                        shippingDetails = "<tr><td>" + shippingDetails + "</td></tr>";
                    }


                    //Response += "<td>$" + esr.SellerConvertedPrice.ToString() + " " + esr.SellerConvertedCurrency + (esr.SellerShippingServiceCost > 0 ? " + $" + esr.SellerShippingServiceCost + " " + esr.SellerShippingServiceCurrency + " Shipping (<font color=\"green\">$" + (esr.SellerShippingServiceCost + esr.SellerConvertedPrice).ToString() + "</font>)" : (esr.ItemShippingType.ToLower().Trim() == "free" ? " (<font color=\"green\">$" + esr.SellerConvertedPrice + "</font>)" : "+ * - No Shipping Price")) + "</td></tr>";
                    //Response += "<td><font size=\"2\">" + "Shipping: " + esr.ItemShippingType + "<br>Expedited: " + (esr.SellerExpeditedShipping == true ? "Yes" : "No") + "<br>One Day:" + (esr.SellerOneDayShipping == true ? "Yes" : "No") + "<br>Services: " + esr.SellerShippingServiceType + "<br>Handling Time: " + esr.SellerHandlingTime.ToString() + " days" + "</font></td></tr>";
                    Response += shippingDetails;
                    Response += "</table>";
                }
            }

            return Response;
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
                    esr.ItemImageURL = xmlSegment.Split(new string[] { "<galleryURL>" }, StringSplitOptions.None)[1].Split('<')[0];
                }
                catch
                {
                    esr.ItemImageURL = "";
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
        private async Task<string> FindAndProcessEBayResults()
        {


            //build url
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
            searchURL += "&keywords=" + System.Web.HttpUtility.UrlEncode(itemTxt.Text);

            System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(searchURL);
            string results = "";
            using (System.IO.Stream s = request.GetResponse().GetResponseStream())
            {
                using (System.IO.StreamReader sr = new System.IO.StreamReader(s))
                {
                    results = sr.ReadToEnd();
                }
            }


            //File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\ebay_api_search.txt", results);

            //order results if shipping is free but calculated, * it at bottom of list...
          /*  string finalEbayTable = "";
            Dictionary<bool, string> tables = new Dictionary<bool, string>();
            string tmpTables = MakeEbayResponseTableHTMLFromXML(results);
            foreach (string s in tmpTables.Split(new string[]{"<table"},StringSplitOptions.None))
            {
                string t = s.Split(new string[] { "</table>" }, StringSplitOptions.None)[0];
                if (t.Contains(")*") == true)
                {
                    tables.Add(true, "<table" + t + "</table>");
                }
                else
                {
                    tables.Add(false, "<table" + t + "</table>");
                }
            }
            foreach (KeyValuePair<bool,string> kvp in tables)
            {
                if (kvp.Key == false)
                {
                    finalEbayTable += kvp.Value;
                }
            }
            foreach (KeyValuePair<bool,string> kvp2 in tables)
            {
                if (kvp2.Key == true)
                {
                    finalEbayTable += kvp2.Value;
                }
            }*/



            string htmlOut = "";
            string CSSSheet = "<style> a.Blacklist:link{color: #551A8B;} a.Blacklist:visited{color: #551A8B;} a.Blacklist:hover{color: #444400;} a.Popup:link{color: #000000;text-decoration:none;} a.Popup:visited{color: #000000;text-decoration:none;} a.Popup:hover{color: #000000;text-decoration:none;} </style>";
            //string JSIncludes = "<script src=\"https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js\"></script><script src=\"http://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js\"></script>";
            htmlOut = "<html><head>" + CSSSheet + "</head><body bgcolor=\"#FFFFFF\">" + MakeEbayResponseTableHTMLFromXML(results) + "<br><br>* Shipping rate was calculated free to 06712.</body></html>";
            //htmlOut = "<html><head>" + CSSSheet + "</head><body bgcolor=\"#FFFFFF\">" + finalEbayTable + "<br><br>* Shipping rate was calculated free to 06712.</body></html>";


            isNavigating = true;
            ebayBrowser.DocumentText = htmlOut;

            if (ebayStatus.InvokeRequired == true)
            {
                ebayStatus.Invoke(new MethodInvoker(()=>{
                    ebayStatus.Text = "eYes";
                }));
                
            }
            else
            {
                ebayStatus.Text = "eYes";
            }
            //Task t = new Task(new Action(Dummy));
            return "";

            //HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://svcs.ebay.com/services/search/FindingService/v1?OPERATION-NAME=findItemsAdvanced&SERVICE-VERSION=1.7.0&SECURITY-APPNAME=ToolsPlu-e1bd-4921-ac6b-14497e74488a&RESPONSE-DATA-FORMAT=XML&REST-PAYLOAD&outputSelector=SellerInfo&keywords=Adjustable%20Clamp%201410-C");


        }


        private bool rawFlag = false;

        private async Task<string> FindAndProcessEBayResultsRaw()
        {

            string searchURL = "http://www.ebay.com/sch/i.html?_from=R40&_trksid=p2050601.m570.l1313.TR0.TRC0.H0.X{ITEM}.TRS0&_nkw={ITEM}&_sacat=0";

            searchURL = searchURL.Replace("{ITEM}", itemTxt.Text.Replace(" ","+"));
           
                MethodInvoker update = delegate()
                {
                    debugBrowser2.Navigate(searchURL);
                    rawFlag = true;
                };
                debugBrowser2.Invoke(update);
                
                while (rawFlag == true)
                {
                    System.Threading.Thread.Sleep(1000);
                }
                return "";
                debugBrowser2.Navigate("javascript:void(document.getElementById('e1-3').click());");
                System.Threading.Thread.Sleep(1000);
            
                debugBrowser2.Navigate("javascript:void(document.getElementById('e1-13').checked=true);");
                System.Threading.Thread.Sleep(1000);

                debugBrowser2.Navigate("javascript:void(document.getElementById('e1-3').click());");
                System.Threading.Thread.Sleep(1000);
           
          
            MethodInvoker finish = delegate()
            {
                string results_html = debugBrowser2.DocumentText;
                //string results_html2 = debugBrowser2.
                try
                {
                    string header_html = results_html.Split(new string[] { "<div class=\"nojs-msk\">" }, StringSplitOptions.None)[0];
                    //string header_html = results_html.Split(new string[] { "<DIV tabindex" }, StringSplitOptions.None)[0];
                    string footer_html = "</div></div></body></html>";
                    string data_html = "<DIV id=\"Results\">" + results_html.Split(new string[] { "<DIV id=\"Results\">" }, StringSplitOptions.None)[1].Split(new string[] { "<div id=\"t_VRFooterDF\"" }, StringSplitOptions.None)[0];
                    string final_output = header_html + data_html + footer_html;
                   // System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\debug.txt", final_output);
                    isNavigating = true;
                    ebayBrowser.DocumentText = final_output;
                }
                catch
                {
                    ebayBrowser.DocumentText = debugBrowser2.DocumentText;
                    return;
                }
            };
            debugBrowser2.Invoke(finish);
            


            //Task t = new Task(new Action(Dummy));
            return "";

        }



    }
}
