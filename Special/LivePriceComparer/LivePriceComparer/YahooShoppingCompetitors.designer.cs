using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Net;
using mshtml;
using System.Windows.Forms;

namespace LivePriceComparer
{
    public partial class Form1
    {

        private struct YahooShoppingResult
        {
            public string SellerName;
            public string SellerURL;
            public string SellerLogoURL;
            public decimal SellerStarRating;
            public string SellerRatingURL;
            public int SellerReviews;

            public decimal Tax;
            public decimal Shipping;
            public decimal Price;
            public string ItemCondition;
                       
            public string ItemImageURL;
            public string YahooItemID;

        }
        private void yahooBrowser_DocumentCompleted(object sender, System.Windows.Forms.WebBrowserDocumentCompletedEventArgs e)
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

        private static string MakeYahooResponseTableHTMLFromShoppingResults(List<YahooShoppingResult> YSRs)
        {
            string tableHTML = "";

            foreach (YahooShoppingResult ysr in YSRs)
            {
                 bool show = false;
                    if (FilterBlacklistedCompetitors == true)
                    {
                        if (CompetitorsData.ContainsKey(ysr.SellerName.ToLower()) == false)
                        {
                            show = true;
                        }
                        else
                        {
                            if (CompetitorsData[ysr.SellerName.ToLower()].Blacklisted == false)
                            {
                                show = true;
                            }
                        }
                        if (FilterBlacklistedCompetitors == true & show == true)
                        {
                            show = true;
                        }
                        else
                        {
                            show = false;
                        }
                    }
                    else
                    {
                        show = true;
                    }
                    if (show == true)
                    {
                        string blacklistingLink = (CompetitorsData.ContainsKey(ysr.SellerName.ToLower()) == true ? (CompetitorsData[ysr.SellerName.ToLower()].Blacklisted == false ? "[" + "<a class=\"one\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=1%20WHERE%20CompetitorName=%27" + ysr.SellerName + "%27\">Blacklist</a>]" : "[" + "<a class=\"one\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=0%20WHERE%20CompetitorName=%27" + ysr.SellerName.ToLower() + "%27\">Un-Blacklist</a>]") : "[" + "<a class=\"one\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?INSERT INTO Competitors (CompetitorName,Enabled,Blacklisted,Whitelisted,AddTax) VALUES('" + ysr.SellerName.ToLower() + "',1,1,0,0)\">Blacklist</a>]");
                        tableHTML += "<table border=\"1\" width=\"90%\" bgcolor=\"#FFFFFF\">";
                        tableHTML += "<tr>" + (ShowPicture == true ? "<td rowspan=\"5\" width=\"20%\">" + ysr.ItemImageURL + "</td>" : "") + "</tr>";
                        tableHTML += "<tr><td>" + (ShowCompetitorName == true ? "<a class=\"two\" href=\"" + ysr.SellerURL + "\">" + ysr.SellerName + "</a> " + blacklistingLink : "") + " (" + ysr.SellerStarRating.ToString("0.0") + " stars) Reviews: " + ysr.SellerReviews.ToString() + "</td></tr>";
                        tableHTML += "<tr>" + (ShowProductTitle == true ? "<td><font size=\"2\">Description: N/A</font></td>" : "") + "</tr>";
                        tableHTML += "<tr>" + (ShowProductPricing == true ? "<td>$" + ysr.Price.ToString("0.00") + " + $" + (ysr.Tax < 0 ? "N/A" : ysr.Tax.ToString("0.00") + "t") + "+ $" + (ysr.Shipping < 0 ? "N/A" : ysr.Shipping.ToString("0.00") + "s/h") + " (<font color=\"green\">$" + ysr.Price.ToString("0.00") + ")</font></td>" : "") + "</tr>";
                        tableHTML += "<tr><td><font size=\"2\">Item Condition: " + ysr.ItemCondition + "</font></td></tr>";
                        tableHTML += "</table>";
                    }

            }
            return tableHTML;
        }
        private async Task<string> YahooNavigateFirstResultSetOnly(string URL)
        {
            List<YahooShoppingResult> YahooShoppingResults = new List<YahooShoppingResult>();
            if (debugBrowser2.InvokeRequired == true)
            {
                Invoke(new System.Windows.Forms.MethodInvoker(() =>
                {

                    debugBrowser2.Navigate(URL);
                    while (debugBrowser2.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        //System.Threading.Thread.Sleep(1000);
                    }
                    //MessageBox.Show(((IHTMLDocument2)debugBrowser2.Document.DomDocument).readyState);
                    //while (((IHTMLDocument2)debugBrowser2.Document.DomDocument).readyState.ToString() != WebBrowserReadyState.Complete.ToString())
                    // {
                    //     Application.DoEvents();
                    // }
                    string ProductID = "";
                    string ProductImageURL = "";
                    int maxRetries = 5;
                    int retryCount = 0;
                    DateTime start = DateTime.Now;
                    //for (retryCount = 0; retryCount < maxRetries; retryCount++)
                    while (((TimeSpan)(DateTime.Now - start)).TotalSeconds < 10)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        SetStatus("Processing Yahoo Pages...");
                        System.Threading.Thread.Sleep(1000);
                        string htmlCode = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.innerHTML;
                        //string htmlCodeImage = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.outerHTML;
                        //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\yahoo_shopping.txt", htmlCode);
                        //string tmpProductID = (htmlCode.Contains("data-cid") == true ? htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None)[1].Split('"')[0] : "N/A");
                        //return;
                        //string tmpProductImageURL = "";
                        //if (htmlCode.Contains("<img class=\"_ioe\"") == true)
                        //{
                        //    tmpProductImageURL = "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('<')[0];
                        //}
                        //string tmpProductImageURL = (htmlCode.Contains("<img class=\"_ioe\"") == true ? "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('>')[0] + ">" : "");
                        //if (tmpProductImageURL.Length > 0)
                        //{
                        //    ProductImageURL = tmpProductImageURL;
                        //}

                        //if (tmpProductID != "N/A")
                        //{
                        //    ProductID = tmpProductID;
                            //add this product id to db if not exist...
                           // if (settingGoogleSaveGIDsChk.Checked == true)
                           // {
                                //Connectivity.SQLCalls.sqlQuery("INSERT INTO GoogleShoppingIDs (GoogleID,ItemTitle,ItemNumber) VALUES('" + ProductID + "','" + itemTxt.Text + "','" + ItemNumber + "')");
                            //}
                       //     break;
                       // }
                        retryCount++;
                    }
                    

                    //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", output);
                }));
            }
            else
            {
                await Task.Run(() =>
                {
                    debugBrowser2.Navigate(URL);
                });


                while (debugBrowser2.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                {
                    System.Windows.Forms.Application.DoEvents();
                }



                string result = debugBrowser2.Document.Body.InnerHtml + "\r\n\r\n------------------------\r\n\r\n" + debugBrowser2.Document.Body.InnerText + "\r\n\r\n----------------------\r\n\r\n" + debugBrowser2.Document.Body.OuterHtml + "\r\n\r\n-------------------------\r\n\r\n" + debugBrowser2.Document.Body.OuterText;
                //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\yahoo_shopping.txt", result);
            }
            return "";
        }
        private async Task<string> YahooPullInfo(string URL)
        {
            List<YahooShoppingResult> YahooShoppingResults = new List<YahooShoppingResult>();
            List<string> ProductIDs = new List<string>();
            List<string> ProductImageURLs = new List<string>();
            if (debugBrowser2.InvokeRequired == true)
            {
                Invoke(new System.Windows.Forms.MethodInvoker(() =>
                {

                    debugBrowser2.Navigate(URL);
                    while (debugBrowser2.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        //System.Threading.Thread.Sleep(1000);
                    }
                    //MessageBox.Show(((IHTMLDocument2)debugBrowser2.Document.DomDocument).readyState);
                    //while (((IHTMLDocument2)debugBrowser2.Document.DomDocument).readyState.ToString() != WebBrowserReadyState.Complete.ToString())
                    // {
                    //     Application.DoEvents();
                    // }
                    //string ProductID = "";
                    //string ProductImageURL = "";
                    int maxRetries = 5;
                    int retryCount = 0;
                    //for (retryCount = 0; retryCount < maxRetries; retryCount++)
                    DateTime start = DateTime.Now;
                    List<YahooShoppingResult> YSR = new List<YahooShoppingResult>();
                    while (((TimeSpan)(DateTime.Now - start)).TotalSeconds < 10)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        SetStatus("Attempt " + (retryCount + 1).ToString() + " of processing Yahoo results.");
                        System.Threading.Thread.Sleep(1000);
                        string htmlCode = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.innerHTML;
                        //string htmlCodeImage = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.outerHTML;
                        //htmlCode = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.innerHTML;
                        //string htmlCodeImage = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.outerHTML;
                       // System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\yahoo_shopping.txt", htmlCode);



                        //string tmpProductID = (htmlCode.Contains("data-cid") == true ? htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None)[1].Split('"')[0] : "N/A");
                        List<string> tmpProductIDs = new List<string>();
                        List<string> tmpProductImageURLs = new List<string>();

                        string baseURL = "https://shopping.yahoo.com";
                        string pricingURL = "";

                        string ProductImageURL = "";
                        string ProductYahooID = "";



                        if (htmlCode.Contains("<DIV id=predoc>") == true)
                        {
                           if (htmlCode.Contains("<A class=price href=\"")==true)
                           {
                               pricingURL = baseURL + htmlCode.Split(new string[] { "<A class=price href=\"" }, StringSplitOptions.None)[1].Split('"')[0] + "#compare-price";
                               if (pricingURL.ToLower().Contains("comhttp") == true)
                               {
                                   
                                   pricingURL = "http" + pricingURL.Split(new string[] { "comhttp" }, StringSplitOptions.None)[1].Trim();
                                   SetDebug(pricingURL);
                               }
                               debugBrowser2.Navigate(pricingURL);
                               start = DateTime.Now;
                               while (debugBrowser2.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)// & debugBrowser2.ReadyState !=System.Windows.Forms.WebBrowserReadyState.Interactive)
                               {
                                   if (((TimeSpan)(DateTime.Now - start)).TotalSeconds >= 10)
                                   {
                                       return;
                                   }
                                   System.Windows.Forms.Application.DoEvents();
                                   //System.Threading.Thread.Sleep(1000);
                                   
                               }
                               string results = ((IHTMLDocument2)debugBrowser2.Document.DomDocument).body.innerHTML;
                               if (results.Contains("<TABLE id=shcgoffergrid") == false)
                               {
                                  if (debugBrowser2.Document.Body.InnerHtml.Contains("<TABLE id=shcoffergrid") == true)
                                  {
                                      results = debugBrowser2.Document.Body.InnerHtml;
                                  }
                                  else
                                  {
                                      yahooBrowser.Navigate("about:blank");
                                      while (yahooBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                                      {
                                          System.Windows.Forms.Application.DoEvents();
                                          //System.Threading.Thread.Sleep(1000);
                                      }
                                      yahooBrowser.Document.Write("<html><body>No sales grouping was found. Only programmed so far to pull first set of grouped items.");
                                      return; 
                                  }
                               }


                               
                               
                                   if (debugBrowser2.Document.Body.InnerHtml.Contains("<IMG id=shimgproductmain") == true)
                                   {
                                       ProductImageURL = "<center><IMG width=\"75px\" src=\"" + debugBrowser2.Document.Body.InnerHtml.Split(new string[] { "<IMG id=shimgproductmain" }, StringSplitOptions.None)[1].Split(new string[] { "src=\"" }, StringSplitOptions.None)[1].Split('"')[0] + "\"></center>";
                                   }
                                   if (debugBrowser2.Document.Body.InnerHtml.Contains("<TD class=label>Parent Retsku</TD>") == true)
                                   {
                                       ProductYahooID = debugBrowser2.Document.Body.InnerHtml.Split(new string[] { "<TD class=label>Parent Retsku</TD>" }, StringSplitOptions.None)[1].Split(new string[] { "</TD>" }, StringSplitOptions.None)[0].Split('>')[1];
                                   }
                                   results = results.Split(new string[] { "<TABLE id=shcgoffergrid" }, StringSplitOptions.None)[1].Split(new string[] { "</TABLE>" }, StringSplitOptions.None)[0];

                                   List<string> competitorCodeSegment = new List<string>();
                                   foreach (string segment in results.Split(new string[] { "<TR class=" }, StringSplitOptions.None))
                                   {
                                       competitorCodeSegment.Add(segment.Split(new string[] { "</TR>" }, StringSplitOptions.None)[0]);
                                   }
                                   competitorCodeSegment.RemoveAt(0);

                                   foreach (string codeSegment in competitorCodeSegment)
                                   {
                                       YahooShoppingResult ysr = new YahooShoppingResult();
                                       ysr.SellerName = codeSegment.Split(new string[] { "class=merch-name>" }, StringSplitOptions.None)[1].Split('<')[0];
                                       ysr.SellerURL = codeSegment.Split(new string[] { "href=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                                       try { ysr.SellerLogoURL = codeSegment.Split(new string[] { "<IMG src=\"" }, StringSplitOptions.None)[1].Split('"')[0]; }
                                       catch { ysr.SellerLogoURL = ""; }
                                       string starrating = codeSegment.Split(new string[] { "<P class=\"rating sprite4 stars-s-r" }, StringSplitOptions.None)[1].Split('"')[0];
                                       ysr.SellerStarRating = decimal.Parse(starrating.Replace("_", "."));
                                       ysr.SellerRatingURL = baseURL + codeSegment.Split(new string[] { "<P class=\"rating sprite4 stars-s-r" }, StringSplitOptions.None)[1].Split(new string[] { "<A href=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                                       ysr.SellerReviews = int.Parse(codeSegment.Split(new string[] { "<P class=\"rating sprite4 stars-s-r" }, StringSplitOptions.None)[1].Split(new string[] { "<A href=\"" }, StringSplitOptions.None)[1].Split('"')[1].Split('(')[1].Split(')')[0]);
                                       try { ysr.Tax = decimal.Parse(codeSegment.Split(new string[] { "<P class=tax>" }, StringSplitOptions.None)[1].Split(new string[] { "</P>" }, StringSplitOptions.None)[0].Replace("<EM>", "").Replace("</EM>", "").Trim().Replace("$", "")); }
                                       catch { ysr.Tax = -1; }
                                       try { ysr.Shipping = decimal.Parse(codeSegment.Split(new string[] { "<P class=shipping>" }, StringSplitOptions.None)[1].Split(new string[] { "</P>" }, StringSplitOptions.None)[0].Replace("<EM>", "").Replace("</EM>", "").Trim().Replace("$", "")); }
                                       catch { ysr.Shipping = -1; }
                                       try
                                       {
                                           string test = codeSegment.Split(new string[] { "<P class=price>" }, StringSplitOptions.None)[1].Split(new string[] { "</P>" }, StringSplitOptions.None)[0].Replace("<STRONG>", "").Replace("</STRONG>", "").Trim().Replace("$", "");
                                           ysr.Price = decimal.Parse(test);
                                       }
                                       catch { ysr.Price = -1; }
                                       ysr.ItemCondition = codeSegment.Split(new string[] { "<TD class=notes>" }, StringSplitOptions.None)[1].Split(new string[] { "</TD>" }, StringSplitOptions.None)[0].Replace("<EM>", "").Replace("</EM>", "").Trim();
                                       ysr.ItemImageURL = ProductImageURL;
                                       ysr.YahooItemID = ProductYahooID;
                                       YSR.Add(ysr);
                                   }
                               

                           }
                           else
                           {
                               //return;
                           }

                        }
                        else
                        {
                            // return;
                        }
                        retryCount++;
                    }
                    SetStatus("Yahoo Loop Complete");
                    string htmlTableData = MakeYahooResponseTableHTMLFromShoppingResults(YSR);

                    string html = "<html><head><style>a.one:link{color: #FF0000;} a.one:visited{color: #FF0000;} a.one:hover{color: #444400;} a.two:link{color: #000000;} a.two:visited{color: #000000;} a.two:hover{color: #003464;}</style></head><body bgcolor=\"#FFFFFF\">" + htmlTableData + "</body></html>";

                    yahooBrowser.Navigate("about:blank");
                    while (yahooBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        //System.Threading.Thread.Sleep(1000);
                    }
                    yahooBrowser.Document.Write(html);
                    if (yahooStatus.InvokeRequired == true)
                    {
                        yahooStatus.Invoke(new MethodInvoker(() =>
                        {
                            yahooStatus.Text = "yYes";
                        }));
                    }
                    else
                    {
                        yahooStatus.Text = "yYes";
                    }
                    //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", output);
                }));
            }
            else
            {
                await Task.Run(() =>
                {
                    debugBrowser2.Navigate(URL);
                });


                while (debugBrowser2.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                {
                    System.Windows.Forms.Application.DoEvents();
                    //System.Threading.Thread.Sleep(1000);
                }



                string result = debugBrowser2.Document.Body.InnerHtml + "\r\n\r\n------------------------\r\n\r\n" + debugBrowser2.Document.Body.InnerText + "\r\n\r\n----------------------\r\n\r\n" + debugBrowser2.Document.Body.OuterHtml + "\r\n\r\n-------------------------\r\n\r\n" + debugBrowser2.Document.Body.OuterText;
                //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\yahoo_shopping.txt", result);
            }
            return "";
        }
        private async Task<string> FindAndProcessYahooShoppingResults()        
        {
            SetStatus("MADE IT HERE...");

            string searchURL = "https://shopping.yahoo.com/search?fr=yshoppingheader_test2&type=2button&p={ITEM}&did=0";
            searchURL = searchURL.Replace("{ITEM}", itemTxt.Text);
            YahooPullInfo(searchURL);

           // if (YahooJustUseFirstResultSet == true)
           // {
           //     YahooNavigateFirstResultSetOnly(searchURL);
           // }
           // if (YahooUseAllResultSets == true)
           // {
           //     YahooNavigateAllResultSets(searchURL);
           // }
           // if (YahooUseSavedGIDsOnly == true)
           // {

           // }

           // if (YahooUseSavedAndQueriedGIDs == true)
           // {

           // }



            return "";

        }

    }
}
