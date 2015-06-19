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
using System.Windows.Forms;
using mshtml;

namespace LivePriceComparer
{
    public partial class Form1
    {

        private struct GoogleShoppingResult
        {
            public string SellerName;
            public string SellerURL;
            public string SellerRatingGeneral;
            public int SellerRatingNumber;
            public decimal Tax;
            public decimal Shipping;
            public decimal BasePrice;
            public decimal TotalPrice;
            public string SpecialOffers;
            public string ItemImageURL;

        }
        private static string MakeGoogleResponseTableHTMLFromShoppingResults(List<GoogleShoppingResult> GSRs)
        {
            string tableHTML = "";

            foreach (GoogleShoppingResult gsr in GSRs)
            {

                    bool show = false;
                    if (FilterBlacklistedCompetitors == true)
                    {
                        if (CompetitorsData.ContainsKey(gsr.SellerName.ToLower()) == false)
                        {
                            show = true;
                        }
                        else
                        {
                            if (CompetitorsData[gsr.SellerName.ToLower()].Blacklisted == false)
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
                        if (gsr.SellerName == "Tools Plus")
                        {
                            
                        }
                        string blacklistingLink = (CompetitorsData.ContainsKey(gsr.SellerName.ToLower()) == true ? (CompetitorsData[gsr.SellerName.ToLower()].Blacklisted == false ? "<a class=\"two\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=1%20WHERE%20CompetitorName=%27" + gsr.SellerName.ToLower() + "%27\">Blacklist</a>" : "<a class=\"two\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?UPDATE%20Competitors%20SET%20Blacklisted=0%20WHERE%20CompetitorName=%27" + gsr.SellerName + "%27\">Un-Blacklist</a>") : "<a class=\"two\" href=\"http://10.0.50.16/whse/arbitraryMotorola.plex?INSERT INTO Competitors (CompetitorName,Enabled,Blacklisted,Whitelisted,AddTax) VALUES('" + gsr.SellerName.ToLower() + "',1,1,0,0)\">Blacklist</a>");
                        tableHTML += "<table border=\"0\" width=\"95%\" bgcolor=\"#FFFFFF\">";
                        tableHTML += "<tr>" + (ShowPicture == true ? "<td rowspan=\"5\" style=\"word-wrap:break-word;width=64px;\">" + "<!--{ITEMURL}--><a class=\"two\" href=\"" + gsr.SellerURL + "\"><!--{/ITEMURL}-->" + gsr.ItemImageURL + "</a></td>" : "") + "</tr>";
                        tableHTML += "<tr><td><font color=\"#555555\">" + (ShowCompetitorName == true ? "<!--{SELLER}-->" + gsr.SellerName + "<!--{/SELLER}--> " + blacklistingLink : "") + " Rating: " + (gsr.SellerRatingGeneral == "" | gsr.SellerRatingGeneral.Contains("No rating") ? "N/A" : (gsr.SellerRatingGeneral.Contains(" ") ? gsr.SellerRatingGeneral.Split(' ')[0] : gsr.SellerRatingGeneral)) + " (" + gsr.SellerRatingNumber.ToString() + ")</td></tr>";
                        tableHTML += "<tr>" + (ShowProductTitle == true ? "<td><font size=\"2\">Description: N/A</font></td>" : "") + "</tr>";
                        tableHTML += "<tr>" + (ShowProductPricing == true ? "<td><font face=\"Arial\"><b>" + "(<font color=\"green\"><!--{PRICE}-->$" + (gsr.TotalPrice  < 0 ? "N/A" : gsr.TotalPrice.ToString("0.00")) + "<!--{/PRICE}--></font>)</b></font> " + (gsr.BasePrice > 0 ? "$" + gsr.BasePrice.ToString("0.00"):"") + (gsr.Tax > 0 ? " + $" + gsr.Tax.ToString("0.00") + "tax" : "") + (gsr.Shipping > 0 ? " + $" + gsr.Shipping.ToString("0.00") + "s/h":"") +"</td>" : "") + "</tr>";
                        //tableHTML += "<tr><td><font size=\"2\">Shipping Info: N/A<br>Handling Info: N/A</font></td></tr>";
                        tableHTML += "</table>";


                    }
            }
            return tableHTML;
        }        
        private async Task<string> GoogleNavigateFirstResultSetOnly(string URL)
        {

            List<GoogleShoppingResult> GoogleShoppingResults = new List<GoogleShoppingResult>();
            if (debugBrowser.InvokeRequired == true)
            {
                Invoke(new System.Windows.Forms.MethodInvoker(() =>
                {

                    debugBrowser.Navigate(URL);
                    while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    //MessageBox.Show(((IHTMLDocument2)debugBrowser.Document.DomDocument).readyState);
                    //while (((IHTMLDocument2)debugBrowser.Document.DomDocument).readyState.ToString() != WebBrowserReadyState.Complete.ToString())
                    // {
                    //     Application.DoEvents();
                    // }
                    string ProductID = "";
                    string ProductImageURL = "";
                    int maxRetries = 5;
                    int retryCount = 0;
                    DateTime start = DateTime.Now;
                    //for (retryCount = 0; retryCount < maxRetries; retryCount++)
                    while (((TimeSpan)(DateTime.Now-start)).TotalSeconds < 10)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        SetStatus("Attempt " + (retryCount + 1).ToString() + " of processing PID");
                        System.Threading.Thread.Sleep(1000);
                        string htmlCode = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.innerHTML;
                        string htmlCodeImage = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.outerHTML;
                        //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", htmlCode);
                        if  (htmlCode == null)
                        {
                            return;
                        }

                        string tmpProductID = (htmlCode.Contains("data-cid") == true ? htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None)[1].Split('"')[0] : "N/A");
                        //return;
                        string tmpProductImageURL = "";
                        if (htmlCode.Contains("<img class=\"_ioe\"") == true)
                        {
                            tmpProductImageURL = "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('<')[0];
                        }
                        //string tmpProductImageURL = (htmlCode.Contains("<img class=\"_ioe\"") == true ? "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('>')[0] + ">" : "");
                        if (tmpProductImageURL.Length > 0)
                        {
                            ProductImageURL = tmpProductImageURL;
                            ProductImageURL = ProductImageURL.Replace("width: 120px;","width: 64px;").Replace("height:","tmpNode:");
                            ProductImageURL = ProductImageURL.Replace("tmpNode:", "border-style: none;");
                        }

                        if (tmpProductID != "N/A")
                        {
                            ProductID = tmpProductID;
                            //add this product id to db if not exist...
                            if (settingGoogleSaveGIDsChk.Checked==true)
                            {
                                Connectivity.SQLCalls.sqlQuery("INSERT INTO GoogleShoppingIDs (GoogleID,ItemTitle,ItemNumber) VALUES('" + ProductID + "','" + itemTxt.Text + "','" + ItemNumber + "')");
                            }
                            break;
                        }
                        retryCount++;
                    }
                    if (ProductID.Length > 0)
                    {
                        SetStatus("PID=" + ProductID);
                        string injectionURL = "https://www.google.com/shopping/product/{GPN}/online?sclient=psy-ab&biw=1280&bih=904&prds=scoring:tp";
                        injectionURL = injectionURL.Replace("{GPN}", ProductID);
                        debugBrowser.Navigate(injectionURL);
                        while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                        {
                            System.Windows.Forms.Application.DoEvents();
                        }
                        string mytestout = "";
                        string results = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.innerHTML;
                        if (results.Contains("<table class=\"os-main-table\" id=\"os-sellers-table\">") == true)
                        {
                            string htmlTable = results.Split(new string[] { "<table class=\"os-main-table\" id=\"os-sellers-table\">" }, StringSplitOptions.None)[1].Split(new string[] { "</table>" }, StringSplitOptions.None)[0];

                            foreach (string row in htmlTable.Split(new string[] { "<tr class=\"os-row\"" }, StringSplitOptions.None))
                            {
                                if (row.StartsWith(" jsaction=\"mouseover:pdui.misg;mouseout:pdui.mosg\">") == true)
                                {
                                    string DefinedRow = row.Split(new string[] { " jsaction=\"mouseover:pdui.misg;mouseout:pdui.mosg\">" }, StringSplitOptions.None)[1].Split(new string[] { "</tr>" }, StringSplitOptions.None)[0];

                                    GoogleShoppingResult gsr = new GoogleShoppingResult();
                                    gsr.SellerURL = DefinedRow.Split(new string[] { "<a href=\"" }, StringSplitOptions.None)[1].Split(new string[] { "\"" }, StringSplitOptions.None)[0];
                                    gsr.SellerName = DefinedRow.Split(new string[] { "<a href=\"" }, StringSplitOptions.None)[1].Split(new string[] { "\">" }, StringSplitOptions.None)[1].Split('<')[0];
                                    gsr.SellerRatingGeneral = DefinedRow.Split(new string[] { "<td class=\"os-rating-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Split('>')[1].Split('<')[0].Trim();
                                    gsr.SpecialOffers = DefinedRow.Split(new string[] { "<td class=\"os-details-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                    gsr.ItemImageURL = ProductImageURL;

                                    string ratings = DefinedRow.Split(new string[] { "<td class=\"os-rating-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();

                                    string pricings = DefinedRow.Split(new string[] { "<td class=\"os-price-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                    if (pricings.Contains("<span class=\"os-base_price\">") == true)
                                    {
                                        gsr.BasePrice = decimal.Parse(pricings.Split(new string[] { "<span class=\"os-base_price\">" }, StringSplitOptions.None)[1].Split('<')[0].Trim().Replace("$", ""));
                                    }
                                    string taxPhrase = "";
                                    string shippingPhrase = "";

                                    if (pricings.Contains("<div class=\"os-total-description\">") == true)
                                    {
                                        string PhraseLine = pricings.Split(new string[] { "<div class=\"os-total-description\">" }, StringSplitOptions.None)[1].Split('<')[0].Trim();
                                        if (PhraseLine.Contains("tax") == true)
                                        {
                                            taxPhrase = PhraseLine.Split(new string[] { "tax" }, StringSplitOptions.None)[0].Split('$')[1].Trim();
                                        }
                                        if (PhraseLine.Contains("shipping") == true)
                                        {
                                            if (PhraseLine.Contains("and") == true)
                                            {
                                                shippingPhrase = PhraseLine.Split(new string[] { "and" }, StringSplitOptions.None)[1].Split(new string[] { "shipping" }, StringSplitOptions.None)[0].Trim();
                                            }
                                            else
                                            {
                                                shippingPhrase = PhraseLine.Split(new string[] { "shipping" }, StringSplitOptions.None)[0].Split('$')[1].Trim();
                                            }
                                        }


                                    }
                                    if (taxPhrase.Length > 0)
                                    {
                                        gsr.Tax = decimal.Parse(taxPhrase);
                                    }
                                    if (shippingPhrase.Length > 0)
                                    {
                                        shippingPhrase = shippingPhrase.Replace("$", "");
                                        gsr.Shipping = decimal.Parse(shippingPhrase);
                                    }


                                    string totPrice = DefinedRow.Split(new string[] { "<td class=\"os-total-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                    if (totPrice.Contains("$") == false)
                                    {
                                        gsr.TotalPrice = -1;
                                    }
                                    else
                                    {
                                        if (decimal.TryParse(totPrice.Split('$')[1].Trim(), out gsr.TotalPrice) == false)
                                        {
                                            gsr.TotalPrice = -1;
                                        }
                                    }




                                    mytestout += gsr.SellerURL + "\r\n" + gsr.SellerName + "\r\n" + gsr.SellerRatingGeneral + "\r\n" + gsr.SpecialOffers + "\r\n" + pricings + "\r\n" + totPrice + "\r\n\r\n\r\n";


                                    GoogleShoppingResults.Add(gsr);
                                    // break;
                                }
                            }
                            //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", mytestout);
                            string htmlTableData = MakeGoogleResponseTableHTMLFromShoppingResults(GoogleShoppingResults);
                            string html = "<html><head><style>a.one:link{color: #FF0000;} a.one:visited{color: #FF0000;} a.one:hover{color: #444400;} a.two:link{color: #551A8B;} a.two:visited{color: #551A8B;} a.two:hover{color: #444400;}</style></head><body bgcolor=\"#FFFFFF\">" + htmlTableData + "</body></html>";

                            googleBrowser.Navigate("about:blank");
                            while (googleBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                            {
                                System.Windows.Forms.Application.DoEvents();
                            }
                            googleBrowser.Document.Write(html);
                            if (googleStatus.InvokeRequired == true)
                            {
                                googleStatus.Invoke(new MethodInvoker(() =>
                                {
                                    googleStatus.Text = "gYes";
                                }));

                            }
                            else
                            {
                                googleStatus.Text = "gYes";
                            }

                        }
                        else
                        {

                        }



                    }
                    else
                    {
                        //could find product id
                        SetStatus("Could not find Google PID");
                    }

                    //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", output);
                }));
            }
            else
            {
                await Task.Run(() =>
                {
                    debugBrowser.Navigate(URL);
                });


                while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                {
                    System.Windows.Forms.Application.DoEvents();
                }



                string result = debugBrowser.Document.Body.InnerHtml + "\r\n\r\n------------------------\r\n\r\n" + debugBrowser.Document.Body.InnerText + "\r\n\r\n----------------------\r\n\r\n" + debugBrowser.Document.Body.OuterHtml + "\r\n\r\n-------------------------\r\n\r\n" + debugBrowser.Document.Body.OuterText;
                //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", result);
            }
            return "";
        }
        private async Task<string> GoogleNavigateAllResultSets(string URL)
        {
            List<GoogleShoppingResult> GoogleShoppingResults = new List<GoogleShoppingResult>();
            List<string> ProductIDs = new List<string>();
            List<string> ProductImageURLs = new List<string>();
            if (debugBrowser.InvokeRequired == true)
            {
                Invoke(new System.Windows.Forms.MethodInvoker(() =>
                {

                    debugBrowser.Navigate(URL);
                    while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    //MessageBox.Show(((IHTMLDocument2)debugBrowser.Document.DomDocument).readyState);
                    //while (((IHTMLDocument2)debugBrowser.Document.DomDocument).readyState.ToString() != WebBrowserReadyState.Complete.ToString())
                    // {
                    //     Application.DoEvents();
                    // }
                    //string ProductID = "";
                    //string ProductImageURL = "";
                    int maxRetries = 5;
                    int retryCount = 0;
                    //for (retryCount = 0; retryCount < maxRetries; retryCount++)
                    DateTime start = DateTime.Now;
                    while (((TimeSpan)(DateTime.Now - start)).TotalSeconds < 10)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        SetStatus("Attempt " + (retryCount + 1).ToString() + " of processing PID");
                        System.Threading.Thread.Sleep(2500);
                        string htmlCode = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.innerHTML;
                        string htmlCodeImage = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.outerHTML;
                        //htmlCode = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.innerHTML;
                        //string htmlCodeImage = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.outerHTML;
                        //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", htmlCode);

                       
                        
                        //string tmpProductID = (htmlCode.Contains("data-cid") == true ? htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None)[1].Split('"')[0] : "N/A");
                        List<string> tmpProductIDs = new List<string>();
                        List<string> tmpProductImageURLs = new List<string>();
                      
                        if (htmlCode.Contains("data-cid=\"")==true)
                        {
                            for (int x = 0; x < htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None).GetLength(0);x++ )
                            {
                                if (x > 0)
                                {
                                    string tmpPID = htmlCode.Split(new string[] { "data-cid=\"" }, StringSplitOptions.None)[x].Split('"')[0];
                                    if (tmpPID.Length > 0)
                                    {
                                        tmpProductIDs.Add(tmpPID);
                                        SetStatus("Found GID=" + tmpPID);
                                    }
                                    
                                }
                            }

                            

                        
                        
                            //return;
                            //string tmpProductImageURL = "";
                            if (htmlCode.Contains("<img class=\"_ioe\"") == true)
                            {
                                for (int x = 0; x < htmlCode.Split(new string[]{"<img class=\"_ioe\""},StringSplitOptions.None).GetLength(0); x++)                            
                                {
                                    if (x > 0)
                                    {
                                        string imgSrc = "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[x].Split('<')[0];
                                        tmpProductImageURLs.Add(imgSrc);
                                    }
                                }
                                //tmpProductImageURL = "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('<')[0];
                            }
                            //string tmpProductImageURL = (htmlCode.Contains("<img class=\"_ioe\"") == true ? "<img class=\"_ioe\"" + htmlCode.Split(new string[] { "<img class=\"_ioe\"" }, StringSplitOptions.None)[1].Split('>')[0] + ">" : "");
                            if (tmpProductIDs.Count > 0)
                            {
                                ProductIDs = tmpProductIDs;
                                foreach (string PID in ProductIDs)
                                {
                                    Connectivity.SQLCalls.sqlQuery("INSERT INTO GoogleShoppingIDs (GoogleID,ItemTitle,ItemNumber) VALUES('" + PID + "','" + itemTxt.Text + "','" + ItemNumber + "')");
                                }
                                //THIS IS WHERE WE UPDATE DATABASE & ProductIDs IF WE WHERE USING INTERNAL GIDs
                            }
                            if (tmpProductImageURLs.Count > 0)
                            {
                                ProductImageURLs = tmpProductImageURLs;
                            }
                            if (tmpProductIDs.Count > 0)
                            {
                                break;
                            }
                        }
                        else
                        {
                           // return;
                        }
                        retryCount++;
                    }
                   
                    string htmlTableData = "";
                    foreach (string ProductID in ProductIDs)
                    {
                        if (ProductID.Length > 0)
                        {
                        
                            string injectionURL = "https://www.google.com/shopping/product/{GPN}/online?sclient=psy-ab&biw=1280&bih=904&prds=scoring:tp";
                            injectionURL = injectionURL.Replace("{GPN}", ProductID);
                            debugBrowser.Navigate(injectionURL);
                            while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                            {
                                System.Windows.Forms.Application.DoEvents();
                            }
                            string mytestout = "";
                            string results = ((IHTMLDocument2)debugBrowser.Document.DomDocument).body.innerHTML;
                            if (results.Contains("<table class=\"os-main-table\" id=\"os-sellers-table\">") == true)
                            {
                                string htmlTable = results.Split(new string[] { "<table class=\"os-main-table\" id=\"os-sellers-table\">" }, StringSplitOptions.None)[1].Split(new string[] { "</table>" }, StringSplitOptions.None)[0];
                                int prodCount = 0;
                                foreach (string row in htmlTable.Split(new string[] { "<tr class=\"os-row\"" }, StringSplitOptions.None))
                                {
                                    if (row.StartsWith(" jsaction=\"mouseover:pdui.misg;mouseout:pdui.mosg\">") == true)
                                    {
                                        string DefinedRow = row.Split(new string[] { " jsaction=\"mouseover:pdui.misg;mouseout:pdui.mosg\">" }, StringSplitOptions.None)[1].Split(new string[] { "</tr>" }, StringSplitOptions.None)[0];

                                        GoogleShoppingResult gsr = new GoogleShoppingResult();
                                        gsr.SellerURL = DefinedRow.Split(new string[] { "<a href=\"" }, StringSplitOptions.None)[1].Split(new string[] { "\"" }, StringSplitOptions.None)[0];
                                        gsr.SellerName = DefinedRow.Split(new string[] { "<a href=\"" }, StringSplitOptions.None)[1].Split(new string[] { "\">" }, StringSplitOptions.None)[1].Split('<')[0];
                                        //gsr.SellerRatingGeneral = DefinedRow.Split(new string[] { "<td class=\"os-rating-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                        gsr.SpecialOffers = DefinedRow.Split(new string[] { "<td class=\"os-details-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                        try
                                        {
                                            gsr.ItemImageURL = ProductImageURLs[prodCount];
                                        }
                                        catch { gsr.ItemImageURL = ""; }
                                        prodCount++;
                                        string ratings = DefinedRow.Split(new string[] { "<td class=\"os-rating-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();

                                        string pricings = DefinedRow.Split(new string[] { "<td class=\"os-price-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                        if (pricings.Contains("<span class=\"os-base_price\">") == true)
                                        {
                                            gsr.BasePrice = decimal.Parse(pricings.Split(new string[] { "<span class=\"os-base_price\">" }, StringSplitOptions.None)[1].Split('<')[0].Trim().Replace("$", ""));
                                        }
                                        string taxPhrase = "";
                                        string shippingPhrase = "";

                                        if (pricings.Contains("<div class=\"os-total-description\">") == true)
                                        {
                                            string PhraseLine = pricings.Split(new string[] { "<div class=\"os-total-description\">" }, StringSplitOptions.None)[1].Split('<')[0].Trim();
                                            if (PhraseLine.Contains("tax") == true)
                                            {
                                                taxPhrase = PhraseLine.Split(new string[] { "tax" }, StringSplitOptions.None)[0].Split('$')[1].Trim();
                                            }
                                            if (PhraseLine.Contains("shipping") == true)
                                            {
                                                if (PhraseLine.Contains("and") == true)
                                                {
                                                    shippingPhrase = PhraseLine.Split(new string[] { "and" }, StringSplitOptions.None)[1].Split(new string[] { "shipping" }, StringSplitOptions.None)[0].Trim();
                                                }
                                                else
                                                {
                                                    shippingPhrase = PhraseLine.Split(new string[] { "shipping" }, StringSplitOptions.None)[0].Split('$')[1].Trim();
                                                }
                                            }


                                        }
                                        if (taxPhrase.Length > 0)
                                        {
                                            gsr.Tax = decimal.Parse(taxPhrase);
                                        }
                                        if (shippingPhrase.Length > 0)
                                        {
                                            shippingPhrase = shippingPhrase.Replace("$", "");
                                            gsr.Shipping = decimal.Parse(shippingPhrase);
                                        }

                                        bool missingPrice = false;
                                        string totPrice = DefinedRow.Split(new string[] { "<td class=\"os-total-col\">" }, StringSplitOptions.None)[1].Split(new string[] { "</td>" }, StringSplitOptions.None)[0].Trim();
                                        if (totPrice.Length == 0)
                                        {
                                            missingPrice = true;
                                        }
                                        else
                                        {
                                            try
                                            {
                                                gsr.TotalPrice = decimal.Parse(totPrice.Split('$')[1].Trim());
                                            }
                                            catch { missingPrice = true; }
                                        }




                                        mytestout += gsr.SellerURL + "\r\n" + gsr.SellerName + "\r\n" + gsr.SellerRatingGeneral + "\r\n" + gsr.SpecialOffers + "\r\n" + pricings + "\r\n" + totPrice + "\r\n\r\n\r\n";

                                        if (missingPrice == false)
                                        {
                                            GoogleShoppingResults.Add(gsr);
                                        }
                                        // break;
                                    }
                                }
                                //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", mytestout);
                                htmlTableData += MakeGoogleResponseTableHTMLFromShoppingResults(GoogleShoppingResults);
                                

                            }
                            else
                            {

                            }



                        }
                        else
                        {
                            //could find product id
                            SetStatus("Could not find Google PID");
                        }
                       
                    }
                    string html = "<html><head><style> a.Blacklist:link{color: #551A8B;} a.Blacklist:visited{color: #551A8B;} a.Blacklist:hover{color: #444400;} a.Popup:link{color: #000000;text-decoration:none;} a.Popup:visited{color: #000000;text-decoration:none;} a.Popup:hover{color: #000000;text-decoration:none;} </style></head><body bgcolor=\"#FFFFFF\">" + htmlTableData + "</body></html>";

                    googleBrowser.Navigate("about:blank");
                    while (googleBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    googleBrowser.Document.Write(html);

                    if (googleStatus.InvokeRequired == true)
                    {
                        googleStatus.Invoke(new MethodInvoker(() =>
                        {
                            googleStatus.Text = "gYes";
                        }));
                    }
                    else
                    {
                        googleStatus.Text = "gYes";
                    }
                    //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", output);
                }));
            }
            else
            {
                await Task.Run(() =>
                {
                    debugBrowser.Navigate(URL);
                });


                while (debugBrowser.ReadyState != System.Windows.Forms.WebBrowserReadyState.Complete)
                {
                    System.Windows.Forms.Application.DoEvents();
                }



                string result = debugBrowser.Document.Body.InnerHtml + "\r\n\r\n------------------------\r\n\r\n" + debugBrowser.Document.Body.InnerText + "\r\n\r\n----------------------\r\n\r\n" + debugBrowser.Document.Body.OuterHtml + "\r\n\r\n-------------------------\r\n\r\n" + debugBrowser.Document.Body.OuterText;
                //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\google_shopping.txt", result);
            }
            return "";
        }
      


        private async Task<string> FindAndProcessGoogleShoppingResults()
        {


            string searchURL = "https://www.google.com/webhp?hl=en#hl=en&tbm=shop&q={ITEM}";
            searchURL = searchURL.Replace("{ITEM}", itemTxt.Text);
            if (GoogleJustUseFirstResultSet == true)
            {               
                GoogleNavigateFirstResultSetOnly(searchURL);
            }
            if (GoogleUseAllResultSets == true)
            {
                GoogleNavigateAllResultSets(searchURL);
            }
            if (GoogleUseSavedGIDsOnly == true)
            {

            }

            if (GoogleUseSavedAndQueriedGIDs == true)
            {

            }



            return "";

        }


        private void googleBrowser_DocumentCompleted(object sender, System.Windows.Forms.WebBrowserDocumentCompletedEventArgs e)
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


    }
}
