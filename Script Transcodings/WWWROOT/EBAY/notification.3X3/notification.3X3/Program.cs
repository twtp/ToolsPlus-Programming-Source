using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace notification._3X3
{
    class Program
    {
        public const string VERSION = "0.0.0";
        public static bool isScript = false;

        public static void GETVERSION()
        {

            Console.WriteLine("Version " + VERSION);
            Console.WriteLine("  Notification Ebay Order Script");
            Console.WriteLine("   Last Modified 8-27-2014");
            Console.WriteLine("   This script takes in an XML Posted document, parses it, and checks");
            Console.WriteLine("   the order against available inventory. At the end, this emails the");
            Console.WriteLine("   main programmer if this ran successful, had warning(s), or error(s).");


            if (!isScript)
            {
                Console.ReadKey();
            }

        }

        private static string GETQUERY(string[] args)
        {
            string queryString = System.Environment.GetEnvironmentVariable("QUERY_STRING");
            if (queryString == null)
            {
                isScript = false;
                queryString = "";
                foreach (string str in args)
                {
                    queryString += str + " ";
                }
                return queryString;
            }
            else
            {
                isScript = true;
                queryString = queryString.Replace("%20", " ");
            }


            //SET YOUR OUTPUT HEADER IF CALLED AS SCRIPT
            if (isScript)
            {
                Console.WriteLine("HTTP/1.1 200 OK\nContent-type: text/plain\n\n\n");
            }





            if (queryString == null)
            {
                Console.WriteLine("Query String Was Null.");
                if (!isScript)
                {
                    Console.WriteLine("Press Any Key To Continue");
                    Console.ReadKey();
                }
                Environment.Exit(0);
            }

            if (queryString == "")
            {
                Console.WriteLine("Query String Was Empty.");
                if (!isScript)
                {
                    Console.WriteLine("Press Any Key To Continue");
                    Console.ReadKey();
                }
                Environment.Exit(0);
            }

            return queryString;
        }

        private static string GETPOST()
        {
            //lets collect our post data...
            string PostData = "";
            int ContentLength = Convert.ToInt32(System.Environment.GetEnvironmentVariable("CONTENT_LENGTH"));
            //Console.WriteLine("This Post contains " + ContentLength.ToString() + " characters.");
            for (int x = 0; x < ContentLength; x++)
            {
                PostData += Convert.ToChar(Console.Read()).ToString();

            }
            //now we should have our post data stored in PostData

            return PostData;

        }



        private static void PROCESSQUERY(string Query)
        {
            if (Query.ToUpper().Contains("GETVERSION()") == true)
            {
                GETVERSION();
                return;
            }

            



        }

        public struct EbaySalesOrderEntry
        {
            public string TransactionID;
            public string ItemID;
            public string ItemSKU;
            public string BuyerID;
            public string SalesOrderNumber;
            public int Quantity;
            public int EbayAvailableQty;
            public int EbaySoldQty;
            public string EBayStatusID;
            
        }
        public struct WarningErrorMessage
        {
            public string MessageType;
            public string SubjectLine;
            public string BodyMessage;
            public string XML_Code;
            public List<string> Recipients;
        }


        public static List<EbaySalesOrderEntry> EbaySalesOrders = new List<EbaySalesOrderEntry>();
        public static List<WarningErrorMessage> WarningErrorEmails = new List<WarningErrorMessage>();
        static void Main(string[] args)
        {
           
            

            //CHECK #1: Is the current request a 'POST'?   //////////////////////////////
            string RequestMethod = Environment.GetEnvironmentVariable("REQUEST_METHOD");
            if (RequestMethod.ToUpper() == "POST")
            {
                //Request is a 'POST', so move on...
            }
            else
            {
                
               Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nExpecting a POST\n");
               Environment.Exit(0);
            }
            /////////////////////////////////////////////////////////////////////////////

            
            
            //CHECK #2: Is the 'POST' content-length variable available?   //////////////
            int ContentLength = 0;
            try
            {
                ContentLength = int.Parse(Environment.GetEnvironmentVariable("CONTENT_LENGTH"));
                if (ContentLength > 0)
                {
                    //'POST' has content-length and it's greater than zero...

                }
                else
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nMissing/invalid content length\n");
                    Environment.Exit(0);
                }
            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nMissing/invalid content length\n");
                Environment.Exit(0);
            }
            //////////////////////////////////////////////////////////////////////////////


                      
            
            // STEP #1: Get the 'POST' data...
            
            string PostData = GETPOST();

            ///////////////////////////////////



            //CHECK #3: See if 'POST' contains the "<soapenv:Envelope>" tag

            if (PostData.ToUpper().Contains("<SOAPENV:ENVELOPE") == true)
            {
                //'POST' contains the <soapenv:envelope> tag... move on..
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing soap envelope\n");
                Environment.Exit(0);
            }

            ///////////////////////////////////////////////////////////////






            //CHECK #4: See if 'POST' contains the "<soapenv:Body>" tag
            if (PostData.ToUpper().Contains("<SOAPENV:BODY>") == true)
            {
                //Great, the 'POST' has the body tag. Let's move on...
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing soap body\n");
                Environment.Exit(0);
            }

            ////////////////////////////////////////////////////////////




            //CHECK #5: This Check is a bit more confusing in code. We need to
            //check to see if only 1 root element exists. This means only 1
            //pair of <SOAPENV:BODY> and 1 pair of <SOAPENV:ENVELOPE>
            int EnvelopeCount = PostData.ToUpper().Split(new string[] { "SOAPENV:ENVELOPE" },StringSplitOptions.None).GetLength(0);
            int BodyCount = PostData.ToUpper().Split(new string[] { "SOAPENV:BODY" },StringSplitOptions.None).GetLength(0);
            EnvelopeCount -= 2;
            BodyCount -= 2;
            if (EnvelopeCount == 1 & BodyCount == 1)
            {
                //Awesome, the 'POST' contains 1 envelope and 1 body
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, expecting single root element\n");
                Environment.Exit(0);
            }
            ///////////////////////////////////////////////////////////////////




            //CHECK #6: Check to see if NotificationEventName exists ///////////
            if (PostData.ToUpper().Contains("NOTIFICATIONEVENTNAME") == true)
            {
                //The NotificationEventName tag exists, move on...
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing NotificationEventName element");
                Environment.Exit(0);
            }
            ////////////////////////////////////////////////////////////////////



            //CHECK #7: NotificationEventName type must be 'FixedPriceTransaction' and
            //the first tag under <SOAPENV:BODY> must be '<GetItemTransactionsResponse>'
            string NotificationEventName = PostData.ToUpper().Split(new string[] { "NOTIFICATIONEVENTNAME>" }, StringSplitOptions.None)[1];
            NotificationEventName = NotificationEventName.Substring(0, NotificationEventName.Length - 2);
            string TransactionEventName = PostData.ToUpper().Split(new string[] { "<SOAPENV:BODY>" }, StringSplitOptions.None)[1].Split(new string[] { "<TIMESTAMP>" },StringSplitOptions.None)[0];
            
            if (NotificationEventName == "FIXEDPRICETRANSACTION" & TransactionEventName.Contains("GETITEMTRANSACTIONSRESPONSE") == true)
            {
                //Great, its a fixed price transaction and it is a get item transactions response... move on..
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nunsupported notification type");
                Environment.Exit(0);
            }

            ///////////////////////////////////////////////////////////////////////////



            
            //STEP #2: Pre-Processing 'Sanity-Checking'
            //Need to grab item sku/id for lookup, ebay's supposed available qty and sold qty
            //At this point in time it looks as if it'll only be 1 item each time (hopefully)
            /////////////////////////////////////////////////////////////////////////////
            string ItemInfo = "";
            int Quantity = 0;
            try
            {
                ItemInfo = PostData.ToUpper().Split(new string[] { "<ITEM>" }, StringSplitOptions.None)[1].Split(new string[] { "</ITEM>" }, StringSplitOptions.None)[0];
            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item");
                Environment.Exit(0);
            }
            try
            {
                Quantity = int.Parse(ItemInfo.Split(new string[] {"<QUANTITY>"},StringSplitOptions.None)[1].Split(new string[] {"</QUANTITY>"},StringSplitOptions.None)[0]);

            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/Quantity");
                Environment.Exit(0);
            }
            //////////////////////////////////////////////////////////////////////////////


            //STEP #3: Aquire the QuantitySold value... ///////////////////////////////////
            int QuantitySold = 0;
            try
            {
                QuantitySold = int.Parse(ItemInfo.Split(new string[] {"<QUANTITYSOLD>"},StringSplitOptions.None)[1].Split(new string[] {"</QUANTITYSOLD>"},StringSplitOptions.None)[0]);

            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/SellingStatus/QuantitySold");
                Environment.Exit(0);
            }
            ///////////////////////////////////////////////////////////////////////////////





            //CHECK #8: See if the ItemID exists///////////////////////////////////////////
            string ItemID = "";
            try
            {
                ItemID = ItemInfo.Split(new string[] { "<ITEMID>" }, StringSplitOptions.None)[1].Split(new string[] { "</ITEMID>" }, StringSplitOptions.None)[0];
            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/SellingStatus/QuantitySold");
                Environment.Exit(0);
            }
            ////////////////////////////////////////////////////////////////////////////////






            //STEP #4: Use SKU or obtain it to get info from PartNumbers Table in SQLDB.
            string SQLItemData = "";
                
            string SKU = "";
            try
            {
                SKU = ItemInfo.Split(new string[] { "<SKU>" }, StringSplitOptions.None)[1].Split(new string[] { "</SKU>" }, StringSplitOptions.None)[0];
                List<string> itemList = SQLCalls.sqlQuery("SELECT ItemNumber, EBayStatusID, EBayAvailableQty, EBaySoldQty FROM PartNumbers WHERE ItemNumber='" + SKU + "'", 1, false);
                SQLItemData = itemList[0];
            }
            catch
            {
                //No SKU.. try looking up via ebay part id
                try
                {
                    List<string> itemList = SQLCalls.sqlQuery("SELECT ItemNumber, EBayStatusID, EBayAvailableQty, EBaySoldQty FROM PartNumbers WHERE EBayItemID='" + ItemID + "'",1,false);
                    SKU = itemList[0].Split('|')[1];
                    SQLItemData = itemList[0];
                }
                catch
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nunknown itemid $ebay_item_id sku");
                    Environment.Exit(0);
                }

            }

            if (SKU != "" | SQLItemData != "")
            {
                //move on, doing good!
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nunknown itemid $ebay_item_id sku");
                Environment.Exit(0);
            }

            ////////////////////////////////////////////////////////////////////////////////////



            //CHECK #9: Check if TransactionArray exists ////////////////////////////////////////
            string TransactionArray = "";
            try
            {
                TransactionArray = ItemInfo.ToUpper().Split(new string[] { "<TRANSACTIONARRAY>" }, StringSplitOptions.None)[1].Split(new string[] { "</TRANSACTIONARRAY>" }, StringSplitOptions.None)[0];
                
                    
            }
            catch
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing TransactionArray");
                Environment.Exit(0);
            }

            if (TransactionArray != "")
            {
                //transaction exists. carry on...
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing TransactionArray");
                Environment.Exit(0);
            }
            /////////////////////////////////////////////////////////////////////////////////////








            //CHECK #10: Make sure Transaction tag is inside TransactionArray.. /////////////////
            if (TransactionArray.Contains("<TRANSACTION>") == true)
            {
                //alrighty it does in fact exist. move on..
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Transaction");
                Environment.Exit(0);
            }
            /////////////////////////////////////////////////////////////////////////////////////





            //STEP #5: Get Transaction Count and Iterate through them...//////////////////////////
            List<string> Transaction = new List<string>();
            List<string> TransactionSQLData = new List<string>();
            int tCount = TransactionArray.Split(new string[] { "<TRANSACTION>" },StringSplitOptions.None).GetLength(0) - 2;
            for (int x = 0; x < tCount; x++)
            {
                string CurrentTransaction = TransactionArray.Split(new string[] { "<TRANSACTION>" }, StringSplitOptions.None)[x + 1];
                Transaction.Add(CurrentTransaction);

                //Check if QuantityPurchased Exists
                try
                {
                    int tmpQtyPurchased = int.Parse(CurrentTransaction.Split(new string[] { "<QUANTITYPURCHASED>" }, StringSplitOptions.None)[1].Split(new string[] { "</QUANTITYPURCHASED>" }, StringSplitOptions.None)[0]);
                                            
                }
                catch
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/TransactionArray/Transaction/QuantityPurchased");
                    Environment.Exit(0);
                }

                //Check if TransactionID Exists
                try
                {
                    string tmpTXID = CurrentTransaction.Split(new string[] { "<TRANSACTIONID>" }, StringSplitOptions.None)[1].Split(new string[] { "</TRANSACTIONID>" }, StringSplitOptions.None)[0];

                }
                catch
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/TransactionArray/Transaction/TransactionID");
                    Environment.Exit(0);
                }
                //Check if ShippingDetails\SellingManagerSalesRecordNumber Exists
                try
                {
                    string tmpShippingDetails = CurrentTransaction.Split(new string[] { "<SELLINGMANAGERSALESRECORDNUMBER>" }, StringSplitOptions.None)[1].Split(new string[] { "</SELLINGMANAGERSALESRECORDNUMBER>" }, StringSplitOptions.None)[0];

                }
                catch
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/TransactionArray/Transaction/ShippingDetails/SellingManagerSalesRecordNumber");
                    Environment.Exit(0);
                }

                //Check if Buyer's UserID Exists
                try
                {
                    string tmpUserID = CurrentTransaction.Split(new string[] { "<USERID>" }, StringSplitOptions.None)[1].Split(new string[] { "</USERID>" }, StringSplitOptions.None)[0];

                }
                catch
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, missing Item/TransactionArray/Transaction/Buyer/UserID");
                    Environment.Exit(0);
                }

                //Get Transaction Data from SQL Table 'EBayTransactionLog' and load into our ebay var
                EbaySalesOrderEntry tmpEntry = new EbaySalesOrderEntry();
                
                List<string> tmpTSQLData = SQLCalls.sqlQuery("SELECT COUNT(*) FROM EBayTransactionLog WHERE ItemID='" + ItemID + "' AND TransactionID='" + CurrentTransaction.Split(new string[] {"<TRANSACTIONID>"},StringSplitOptions.None)[1].Split(new string[] {"</TRANSACTIONID>"},StringSplitOptions.None)[0] + "'",1,false);
                TransactionSQLData.Add(tmpTSQLData[0]);
                //so what if there is more than one result?? Can there be?
                tmpEntry.BuyerID = CurrentTransaction.Split(new string[] { "<USERID>" }, StringSplitOptions.None)[1].Split(new string[] { "</USERID>" }, StringSplitOptions.None)[0];
                tmpEntry.ItemID = ItemID;
                tmpEntry.TransactionID = CurrentTransaction.Split(new string[] { "<TRANSACTIONID>" }, StringSplitOptions.None)[1].Split(new string[] { "</TRANSACTIONID>" }, StringSplitOptions.None)[0];
                tmpEntry.SalesOrderNumber = CurrentTransaction.Split(new string[] { "<ORDERID>" }, StringSplitOptions.None)[1].Split(new string[] { "</ORDERID>" }, StringSplitOptions.None)[0];
                tmpEntry.Quantity = int.Parse(CurrentTransaction.Split(new string[] { "<QUANTITYPURCHASED>" }, StringSplitOptions.None)[1].Split(new string[] { "</QUANTITYPURCHASED>" }, StringSplitOptions.None)[0]);
                tmpEntry.EbayAvailableQty = int.Parse(PostData.ToUpper().Split(new string[] { "<QUANTITY>" }, StringSplitOptions.None)[1].Split(new string[] { "</QUANTITY>" }, StringSplitOptions.None)[0]);
                tmpEntry.EbaySoldQty = int.Parse(PostData.ToUpper().Split(new string[] { "<QUANTITYSOLD>" }, StringSplitOptions.None)[1].Split(new string[] { "</QUANTITYSOLD>" }, StringSplitOptions.None)[0]);
                tmpEntry.ItemSKU = SKU;

                EbaySalesOrders.Add(tmpEntry);

                //Most Excellent... We have done most of the work so far, now to do final error
                //checks and actually post/update sql data (outside of loop)


            }
            //////////////////////////////////////////////////////////////////////////////////////








            //CHECK #11: Check all quantities for negatives. error if so... if zero, warning///////
            foreach (EbaySalesOrderEntry eBe in EbaySalesOrders)
            {
                if (eBe.EbayAvailableQty >= 0 & eBe.EbaySoldQty >= 0 & eBe.Quantity >= 0)
                {
                    //ok we are good... continue...
                }
                else
                {
                    Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid xml, expected positive quantity over all transaction elements");
                    Environment.Exit(0);
                }
                if (eBe.Quantity == 0)
                {
                    //WARNING EMAIL TIME...
                    WarningErrorMessage wem = new WarningErrorMessage();
                    wem.SubjectLine = "ebay notification api warning";
                    wem.MessageType = "WARNING";
                    wem.BodyMessage = "transaction has already been handled?\r\n\r\n";
                    wem.XML_Code = PostData;
                    List<string> rec = new List<string>();
                    rec.Add("tom@tools-plus.com");
                    wem.Recipients = rec;

                    WarningErrorEmails.Add(wem);

                }
            }
            ///////////////////////////////////////////////////////////////////////////////////////


            //STEP #6: See if the quantity available in the sql table matches that of this order
            string EBayStatusID = "";
            foreach (EbaySalesOrderEntry eBe in EbaySalesOrders)
            {
                List<string> itmQty = SQLCalls.sqlQuery("SELECT ItemNumber, EBayStatusID, EBayAvailableQty, EBaySoldQty FROM PartNumbers WHERE ItemNumber='" + eBe.ItemSKU +"'", 1, false);
                EBayStatusID = itmQty[0].Split('|')[2];
                if (int.Parse(itmQty[0].Split('|')[3]) == eBe.EbayAvailableQty)
                {
                    //good so far...
                }
                else
                {
                    //throw a warning that quantities are off...
                    WarningErrorMessage wem = new WarningErrorMessage();
                    wem.SubjectLine = "ebay notification api warning";
                    wem.MessageType = "WARNING";
                    wem.BodyMessage = eBe.ItemSKU + " available quantity of " + int.Parse(itmQty[0].Split('|')[3]).ToString() + " does not match qty from ebay " + eBe.EbayAvailableQty.ToString();
                    wem.XML_Code = PostData;
                    List<string> rec = new List<string>();
                    rec.Add("tom@tools-plus.com");
                    wem.Recipients = rec;
                    WarningErrorEmails.Add(wem);
                }
                //update the sql database...
                SQLCalls.sqlQuery("UPDATE PartNumbers SET EBayAvailableQty=" + eBe.EbayAvailableQty.ToString() + " WHERE ItemNumber='" + eBe.ItemSKU + "'", 1, false);



            }
            ////////////////////////////////////////////////////////////////////////////////////
            
            
            //STEP #7: Check if database is in sync...////////////////////////////////////////////
            foreach (EbaySalesOrderEntry eBe in EbaySalesOrders)
            {
                if (eBe.EbaySoldQty + eBe.Quantity == QuantitySold)
                {
                    //great! keep movin...
                }
                else
                {
                    //warning time!
                    WarningErrorMessage wem = new WarningErrorMessage();
                    wem.SubjectLine = "ebay notification api warning";
                    wem.MessageType = "WARNING";
                    wem.BodyMessage = "item " + eBe.ItemSKU + " total sold qty of " + (eBe.EbaySoldQty + eBe.Quantity).ToString() + " does not match qty from ebay " + eBe.EbaySoldQty.ToString();
                    wem.XML_Code = PostData;
                    List<string> rec = new List<string>();
                    rec.Add("tom@tools-plus.com");
                    wem.Recipients = rec;
                    WarningErrorEmails.Add(wem);
                }
            }

            //////////////////////////////////////////////////////////////////////////////////////

            
            
            
            //STEP #8: Update the sql database table 'PartNumbers' using item number and sold
            //quantity.
            foreach (EbaySalesOrderEntry eBe in EbaySalesOrders)
            {
                SQLCalls.sqlQuery("UPDATE PartNumbers SET EBaySoldQty=" + eBe.EbaySoldQty.ToString() + " WHERE ItemNumber='" + eBe.ItemSKU + "'",1,false);

                //now, if item's available qty is equal to sold qty, notify us via warning message...

                if (eBe.EbayAvailableQty == eBe.EbaySoldQty)
                {
                    if (EBayStatusID != "0")
                    {
                        //warning time!
                        WarningErrorMessage wem = new WarningErrorMessage();
                        wem.SubjectLine = "ebay notification api warning";
                        wem.MessageType = "WARNING";
                        wem.BodyMessage = "item " + eBe.ItemSKU + " is now out of stock on EBay.";
                        wem.XML_Code = PostData;
                        List<string> rec = new List<string>();
                        rec.Add("tom@tools-plus.com");
                        wem.Recipients = rec;
                        WarningErrorEmails.Add(wem);
                    }
                    else
                    {
                        WarningErrorMessage wem = new WarningErrorMessage();
                        wem.SubjectLine = "ebay notification api warning";
                        wem.MessageType = "WARNING";
                        wem.BodyMessage = "item " + eBe.ItemSKU + " is now out of stock on EBay, but we already knew that?";
                        wem.XML_Code = PostData;
                        List<string> rec = new List<string>();
                        rec.Add("tom@tools-plus.com");
                        wem.Recipients = rec;
                        WarningErrorEmails.Add(wem);
                    }
                    
                }
            }
            //////////////////////////////////////////////////////////////////////////////////


            //Step #9: WE FINALLY MADE IT!!!! 200 OK AND NOW TO FINALIZE SAID TRANSACTION(s)///////////
            Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\nSuccess\n");
            foreach (EbaySalesOrderEntry eBe in EbaySalesOrders)
            {
                SQLCalls.sqlQuery("INSERT INTO EBayTransactionLog (ItemID, TransactionID, Quantity, SalesOrderNo, BuyerID) VALUES ('" + eBe.ItemID + "','" + eBe.TransactionID + "'," + eBe.Quantity.ToString() + ",'" + eBe.SalesOrderNumber + "','" + eBe.BuyerID + "')", 1, false);

            }
            ///////////////////////////////////////////////////////////////////////////////////


            
            
            
            
            
            
            //Step #10: Send any emails pending.../////////////////////////////////////////////
            foreach (WarningErrorMessage wem in WarningErrorEmails)
            {
                Emailer.SendEmail(wem.SubjectLine, wem.BodyMessage + "\r\n\r\n" + wem.XML_Code, false, wem.Recipients);
            }
            ///////////////////////////////////////////////////////////////////////////////////


            //THIS SCRIPT HAS YET TO BE TESTED!


            File.WriteAllText("C:\\Users\\tomwestbrook\\desktop\\POST_RESULTS.txt", PostData);
            string QueryString = GETQUERY(args);
            PROCESSQUERY(QueryString);

            
            
 

        }
    }
}
