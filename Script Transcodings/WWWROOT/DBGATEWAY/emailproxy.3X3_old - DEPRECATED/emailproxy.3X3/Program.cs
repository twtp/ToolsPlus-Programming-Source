using System;
using System.Collections.Generic;

using System.Text;
using System.Web;
using System.Net;
using System.Net.Sockets;
using System.IO;




namespace emailproxy._3X3
{



    class Program
    {


        public const string VERSION = "0.0.0";
        public static bool isScript = false;
        public const decimal TaxRate = 0.0635M;

        public static List<OrderDetails> orderDetailsList = new List<OrderDetails>();

        public static void GETVERSION()
        {

            Console.WriteLine("Version " + VERSION);
            Console.WriteLine("  Email Proxy Script");
            Console.WriteLine("   Last Modified 8-27-2014");
            Console.WriteLine("   This Script checks to see if the remote IP is an acceptable one.");
            Console.WriteLine("   If so, it will take the 'POST' data in, aquire its variables, then");
            Console.WriteLine("   send a confirmation email to the buyer of this particular order.");


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

        private static void EnumerateAcceptedRemoteIPs()
        {
            AcceptedRemoteIPs.Add("64.22.97.20");
            AcceptedRemoteIPs.Add("50.23.169.36");
            AcceptedRemoteIPs.Add("50.23.169.67");

        }

        public static List<string> AcceptedRemoteIPs = new List<string>();

        static void Main(string[] args)
        {
            //System.Threading.Thread.Sleep(1000);
            string PostData = GETPOST();
            //PostData = args[0];
            orderDetailsList = YAMLtoOrderDetails(PostData);
            string HTMLOUT = "";
            foreach (OrderDetails od in orderDetailsList)
            {
                HTMLOUT += Get_EmailBody(od);

            }
            List<string>EmailList = new List<string>();
            EmailList.Add("tom@tools-plus.com");

            
            

            //Emailer.SendEmail(Get_FromName(orderDetailsList[0].OrderNumber), HTMLOUT, false, EmailList);
           // ImpersonationDemo.SendEmailImpersonated(Get_FromName(orderDetailsList[0].OrderNumber),HTMLOUT,false,EmailList);
          //  Environment.Exit(0);

            Console.WriteLine("Status: 200 OK\nContent-type: text/html\n\n");
            //Console.WriteLine(HTMLOUT);

            //ImpersonationDemo.SendEmailImpersonated(Get_FromName(orderDetailsList[0].OrderNumber), HTMLOUT, false, EmailList);
            Emailer.SendEmailv2(Get_FromName();

            Environment.Exit(0);
            DebugOutput();

            Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n" + PostData);
            Environment.Exit(0);

            string RemoteIP = "";
            try
            {
                RemoteIP = Environment.GetEnvironmentVariable("REMOTE_ADDR");
            }
            catch 
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nRemote Address Not Found.");
                Environment.Exit(0);
            }

            bool isAccepted = false;
            foreach (string aIP in AcceptedRemoteIPs)
            {
                if (RemoteIP == aIP)
                {
                    isAccepted = true;
                    break;
                }
            }
            if (isAccepted == true)
            {
                //Great Move on...
            }
            else
            {
                Console.WriteLine("Status: 403 Forbidden\nContent-type: text/plain\n\nYou are not authorized for this page.\n");
                Environment.Exit(0);
            }


            if (PostData.Length > 0)
            {
                ProcessPostData(PostData);
            }
            else
            {
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\nPOST Data Missing.");
                Environment.Exit(0);
            }


        }

        public List<OrderDetails> OrderInfo = new List<OrderDetails>();
        private static void ProcessPostData(string POST_DATA)
        {
            OrderDetails OD = new OrderDetails();
            ItemDetails ID = new ItemDetails();
            

            //string InData = YAMLtoQuery(POST_DATA);

            Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n" + POST_DATA);
            Environment.Exit(0);



        }


        public struct ItemDetails
        {
            public string ItemNumber;
            public decimal UnitPrice;
            public int Quantity;
            public string Description;
            public decimal TotalItemCost;
        }
    
        public struct OrderDetails
        {
            public string OrderNumber;
            public string OrderDate;
            public string OrderStatus;
            public string PaymentMethod;

            public string PONumber;

            public string BillingName;
            public string BillingAddress1;
            public string BillingAddress2;
            public string BillingCity;
            public string BillingState;
            public string BillingZipcode;
            public string BillingCountry;
            public string BillingPhone;
            public string ShippingName;
            public string ShippingAddress1;
            public string ShippingAddress2;
            public string ShippingCity;
            public string ShippingState;
            public string ShippingZipcode;
            public string ShippingCountry;
            public string ShippingPhone;
            public string EmailAddress;            
            public string ShippingMethod;
            public bool IsTaxable;
            public decimal TaxAmount;
            public decimal SubTotal;
            public decimal Total;

            
            public int ItemCount;

            public List<TransactionDetails> TransactionList;
            
            public List<ItemDetails> ItemsList;            



        }

        public struct TransactionDetails
        {
            public decimal Cost;
            public bool IsTaxable;
            public string Text;
            
        }

        public struct EmailVariables
        {
            public string From_Name;
            public string From_Address;
            public string Subject;
            public string TextEmail;
            public string HTMLEmail;

        }

        private static string Get_FromName(string OrderNumber)
        {
            if (OrderNumber.Substring(0,1) == "P")
            {
                return "Tools Plus Order Confirmation";
            }
            else
            {

                return "ToolPartsStore.com Order Confirmation";
            }
        }
        private static string Get_EmailBody(OrderDetails OD)
        {
            string BodyHTML = "";
            if (OD.OrderNumber.Substring(0, 1) == "P")
            {
                
                BodyHTML += "<body style=\"font-family: sans-serif;\">\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n<td>\r\n";
                BodyHTML += "<p>Hi and welcome to ToolPartsStore.com<br />We wish to thank you for your order.</p>\r\n";
                BodyHTML += "<p><b>** Please save this email, this has your order number **</b></p>\r\n";
                BodyHTML += "<p>Our crew will now go out of their way to make sure that you receive excellent service.\r\n";
                BodyHTML += "We want your shopping experience to be fun, easy, and simple.</p>\r\n";
                BodyHTML += "<p>Here is a copy of the order that you placed. Please double check this information\r\n";
                BodyHTML += "for accuracy. If there are any changes please email our order processing department\r\n";
                BodyHTML += "(<a href=\"mailto:sales@toolpartsstore.com\">sales@toolpartsstore.com</a>) or call us\r\n";
                BodyHTML += "(1-203-573-0750) as soon as possible. You will receive an email with details if there\r\n";
                BodyHTML += "are any problems in processing your order or if any items are unavailable for shipping\r\n";
                BodyHTML += "at this time.</p>\r\n";
                BodyHTML += "</td>\r\n</tr>\r\n</table>\r\n<br />\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td width=\"150\"><img src=\"http://lib.store.yahoo.net/lib/yhst-130072382124186/invoice-logo.png\" alt=\"ToolPartsStore.com Logo\" /></td>\r\n";
                BodyHTML += "<td width=\"450\">\r\n";
                BodyHTML += "<h2>ORDER CONFIRMATION</h2>\r\n";
                BodyHTML += "<table width=\"450\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td valign=\"top\" width=\"225\">\r\n";
                BodyHTML += "<b>Order Number:</b> " + OD.OrderNumber + "<br />\r\n";



               
            }
            else
            {
                //BodyHTML = "";

                //TOP SECTION
                BodyHTML += "<body style=\"font-family: sans-serif;\">\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n<td>\r\n";
                BodyHTML += "<p>Hi and welcome to Tools-Plus.com<br />We wish to thank you for your order.</p>";
                BodyHTML += "<p><b>** Please save this email, this has your order number **</b></p>\r\n";
                BodyHTML += "<p>Our crew will now go out of their way to make sure that you receive excellent service.\r\n";
                BodyHTML += "We want your shopping experience to be fun, easy, and simple.</p>\r\n";
                BodyHTML += "<p>Here is a copy of the order that you placed. Please double check this information\r\n";
                BodyHTML += "for accuracy. If there are any changes please email our order processing department\r\n";
                BodyHTML += "(<a href=\"mailto:orderprocessing@tools-plus.com\">orderprocessing@tools-plus.com)</a>\r\n";
                BodyHTML += "or call us (1-800-222-6133) as soon as possible. You will receive an email with\r\n";
                BodyHTML += "details if there are any problems in processing your order or if any items are\r\n";
                BodyHTML += "unavailable for shipping at this time.</p>\r\n";
                BodyHTML += "</td>\r\n</tr>\r\n</table>\r\n<br />\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td width=\"150\"><img src=\"http://lib.store.yahoo.net/lib/toolsplus/invoice-logo.gif\" alt=\"Tools Plus Logo\" /></td>\r\n";
                BodyHTML += "<td width=\"450\">\r\n";
                BodyHTML += "<h2>ORDER CONFIRMATION</h2>\r\n";
                BodyHTML += "<table width=\"450\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td valign=\"top\" width=\"225\">\r\n";
                BodyHTML += "<b>Order Number:</b> " + OD.OrderNumber + "<br />\r\n";

                //MID SECTION
                if (OD.PONumber != "")
                {
                    BodyHTML += "<b>PO Number:</b> " + OD.PONumber + "<br />\r\n";
                }

                BodyHTML += "<b>Date:</b> " + OD.OrderDate + "<br />\r\n";
                BodyHTML += "<b>Order Status:</b> " + OD.OrderStatus + "<br />\r\n";
                BodyHTML += "<b>Delivery Method:</b> " + OD.ShippingMethod + "<br />\r\n";
                BodyHTML += "<b>Payment Method:</b> " + OD.PaymentMethod + "\r\n";
                BodyHTML += "</td>\r\n";
                BodyHTML += "<td valign=\"top\" width=\"225\" style=\"text-align:right;\">\r\n";
                BodyHTML += "<b>Tools Plus.com</b><br />\r\n";
                BodyHTML += "60 Scott Road<br />\r\n";
                BodyHTML += "Prospect, CT 06712<br />\r\n";
                BodyHTML += "Toll-Free: (800) 222-6133<br />\r\n";
                BodyHTML += "Fax: (203) 753-9042<br />\r\n";
                BodyHTML += "E-Mail: <a href=\"mailto:orderprocessing@tools-plus.com\">orderprocessing@tools-plus.com</a><br />\r\n";
                BodyHTML += "Visit our Retail Store:<br />\r\n";
                BodyHTML += "153 Meadow Street<br />\r\n";
                BodyHTML += "Waterbury, CT 06702\r\n";
                BodyHTML += "</td>\r\n</tr>\r\n</table>\r\n</td>\r\n</tr>\r\n</table>\r\n";
                BodyHTML += "<hr width=\"600\" align=\"left\" />\r\n";
                BodyHTML += "<br />\r\n";
                
                
                
                //BILLING INFO SECTION
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td width=\"295\" valign=\"top\">\r\n";
                BodyHTML += "<b>Billing Address</b>\r\n";
                BodyHTML += "<hr />\r\n";
                BodyHTML += "<table width=\"295\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr><td><b>Name:</b></td><td>" + OD.BillingName + "</td></tr>\r\n";
                
                //NOT SURE HOW TO WRITE ADDRESSLINE1 + ADDRESSLINE2... SO I DID THIS..
                BodyHTML += "<tr><td><b>Address:</b></td><td>" + OD.BillingAddress1 + "<br>" + OD.BillingAddress2 + "</td></tr>\r\n";
                
                BodyHTML += "<tr><td><b>City:</b></td><td>" + OD.BillingCity + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>State:</b></td><td>" + OD.BillingState + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Zip/Postal Code:</b></td><td>" + OD.BillingZipcode + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Country:</b></td><td>" + OD.BillingCountry +"</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Phone:</b></td><td>" + OD.BillingPhone + "</td></tr>\r\n";
                BodyHTML += "</table>\r\n";
                BodyHTML += "</td>\r\n";
                BodyHTML += "<td width=\"10\"></td>\r\n";
                BodyHTML += "<td width=\"295\" valign=\"top\">\r\n";
                BodyHTML += "<b>Shipping Address</b>\r\n";
                BodyHTML += "<hr />\r\n";
                BodyHTML += "<table width=\"295\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr><td><b>Name:</b></td><td>" + OD.ShippingName + "</td></tr>\r\n";

                //AGAIN, NOT SURE HOW TO HANDLE ADDRESS FIELDS, SO I DID THIS...
                BodyHTML += "<tr><td><b>Address:</b></td><td>" + OD.ShippingAddress1 + "<br>" + OD.ShippingAddress2 + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>City:</b></td><td>" + OD.ShippingCity + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>State:</b></td><td>" + OD.ShippingState + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Zip/Postal Code:</b></td><td>" + OD.ShippingZipcode + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Country:</b></td><td>" + OD.ShippingCountry + "</td></tr>\r\n";
                BodyHTML += "<tr><td><b>Phone:</b></td><td>" + OD.ShippingPhone + "</td></tr>\r\n";
                BodyHTML += "</table>\r\n";
                BodyHTML += "</td>\r\n";
                BodyHTML += "</tr>\r\n";
                BodyHTML += "</table>\r\n";
                BodyHTML += "<br />\r\n";
                BodyHTML += "<h3 width=\"600\" style=\"font-size:16px;\">Products Ordered</h3>\r\n";
                BodyHTML += "<table width=\"600\" border=\"1\" cellspacing=\"0\" cellpadding=\"1\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<th width=\"125\">Item Number</th>\r\n";
                BodyHTML += "<th width=\"223\">Description</th>\r\n";
                BodyHTML += "<th width=\"75\">Unit Price</th>\r\n";
                BodyHTML += "<th width=\"75\">Quantity</th>\r\n";
                BodyHTML += "<th width=\"80\">Total</th>\r\n";
                BodyHTML += "</tr>\r\n";

                
                foreach (ItemDetails ID in OD.ItemsList)
                {
                    BodyHTML += "<tr>\r\n";
                    BodyHTML += "<td>" + ID.ItemNumber + "</td>\r\n";
                    BodyHTML += "<td>" + ID.Description + "</td>\r\n";
                    BodyHTML += "<td align=\"center\">" + ID.UnitPrice + "</td>\r\n";
                    BodyHTML += "<td align=\"center\">" + ID.Quantity +"</td>\r\n";
                    //total item * qty cost...
                    BodyHTML += "<td align=\"center\">" + ID.TotalItemCost.ToString() + "</td>\r\n";
                    
                    
                }
                BodyHTML += "</tr>\r\n</table>\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\">\r\n";
                BodyHTML += "<tr>\r\n<td>\r\n";
                BodyHTML += "<table width=\"180\" align=\"right\" border=\"1\" cellspacing=\"0\" cellpadding=\"1\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr><td width=\"100\" align=\"right\">Subtotal:</td><td width=\"80\" align=\"center\">" + OD.SubTotal.ToString("0.00") +"</td></tr>\r\n";
                BodyHTML += "<tr><td width=\"100\" align=\"right\">Tax:</td><td width=\"80\" align=\"center\">" + OD.TaxAmount.ToString("0.00") + "</td></tr>\r\n";
                BodyHTML += "<tr><td width=\"100\" align=\"right\">Total:</td><td width=\"80\" align=\"center\">" + OD.Total.ToString("0.00") +"</td></tr>\r\n";
                BodyHTML += "</table>\r\n</td>\r\n</tr>\r\n</table>";    
                
                
                
                //SKIPPED ITEMS INFO SECTION FOR NOW UNTIL MORE INFORMATION IS GIVEN OR FOUND...

                BodyHTML += "<br />\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size 12px;\">\r\n";
                BodyHTML += "<tr>\r\n<td>\r\n";
                BodyHTML += "<p>If you requested a freight quote, we should email you within 24 hours for you approval.</p>\r\n";
                BodyHTML += "<p>You will also receive an automatic email with tracking information once your order has\r\n";
                BodyHTML += " been processed and shipped. All Shipments can be tracked online at the Tools Plus website.</p>\r\n";
                BodyHTML += "<p>We're hoping you can give us an excellent rating, but if you feel your experience was any less, \r\n";
                BodyHTML += "we want to know immediately! We are here to make your shopping experience easier, so please email \r\n";
                BodyHTML += "us for anything. Please include your order number to speed this process.</p>\r\n";
                BodyHTML += "<p>Thanks,<br />\r\n";
                BodyHTML += "Tools Plus Order Processing Department<br />\r\n";
                BodyHTML += "<a href=\"mailto:orderprocessing@tools-plus.com\">orderprocessing@tools-plus.com</a><br />\r\n";
                BodyHTML += "60 Scott Road, Prospect, CT 06712<br />\r\n";
                BodyHTML += "Telephone: (203) 573-0750 or 1-800-222-6133<br />\r\n";
                BodyHTML += "Visit our Retail Store:<br />\r\n";
                BodyHTML += "153 Meadow Street, Waterbury, CT 06702</font></p>\r\n";
                BodyHTML += "</td>\r\n</tr>\r\n</table>\r\n";
                
                
                
                

                //ADD THE FOOTER IN HERE...
                BodyHTML += "<hr width=\"600\" align=\"left\" />\r\n";
                BodyHTML += "<img src=\"http://qr.kaywa.com/img.php?s=8&d=http%3A%2F%2Fwww.facebook.com%2FToolsPlusCT\" alt=\"Connect with Us on Facebook\"/ width=\"140\" height=\"140\">\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td width=\"200\"><a href=\"http://twitter.com/#!/toolsplus\"><img src=\"http://tools-plus.com.p1.hostingprod.com/images/emails/footer-twitter-200.gif\" width=\"200\" height=\"31\" border=\"0\" alt=\"\"></a></td>\r\n";
                BodyHTML += "<td width=\"200\"><a href=\"http://www.facebook.com/pages/Waterbury-CT/Tools-Plus/126025707444288\"><img src=\"http://tools-plus.com.p1.hostingprod.com/images/emails/footer-facebook-200.gif\" width=\"200\" height=\"31\" border=\"0\" alt=\"\"></a></td>\r\n";
                BodyHTML += "<td width=\"200\"><a href=\"http://www.tool-talker.com/\"><img src=\"http://tools-plus.com.p1.hostingprod.com/images/emails/footer-tool-talker-200.gif\" width=\"200\" height=\"31\" border=\"0\" alt=\"\"></a></td>\r\n";
                BodyHTML += "</tr>\r\n";
                BodyHTML += "</table>\r\n";
                BodyHTML += "<table width=\"600\" cellspacing=\"0\" cellpadding=\"0\" style=\"font-size: 12px;\">\r\n";
                BodyHTML += "<tr>\r\n";
                BodyHTML += "<td width=\"302\"><a href=\"http://www.tools-plus.com/newsletter.html\"><img src=\"http://tools-plus.com.p1.hostingprod.com/images/emails/confirm-links-tools-plus-exclusive.jpg\" width=\"302\" height=\"66\" border=\"0\" alt=\"\"></a></td>\r\n";
                BodyHTML += "<td width=\"298\"><a href=\"http://myaccount.tools-plus.com/mod_myRewards/rewardsLanding.php\"><img src=\"http://tools-plus.com.p1.hostingprod.com/images/emails/confirm-links-tools-plus-rewards.jpg\" width=\"298\" height=\"66\" border=\"0\" alt=\"\"></a></td>\r\n";
                BodyHTML += "</tr>\r\n";
                BodyHTML += "</table>\r\n";








            }

            return BodyHTML;

        }


        private static string YAMLtoQuery(string YAMLinput)
        {

            //there must be a better way to do this...
            //URL DECODER THAT WORKS
            //string temp = HttpUtility.HtmlDecode(YAMLinput);
            //return temp;
            
            string returnData = YAMLinput;
            returnData = returnData.Replace("%0A", "\r\n");
            returnData = returnData.Replace("%3A", ":");
            returnData = returnData.Replace("%40", "@");
            returnData = returnData.Replace("+", " ");
            returnData = returnData.Replace("%2F", "/");

            return returnData;
        }

        private static List<OrderDetails> YAMLtoOrderDetails(string YAMLinput)
        {
            List<OrderDetails> tmpDetails = new List<OrderDetails>();
            


            string Query = YAMLtoQuery(YAMLinput);

            int YAMLTransactionCount = Query.ToUpper().Split(new string[] { "BILL_ADDRESS_1" }, StringSplitOptions.None).GetLength(0) - 1;
            for (int counter = 0; counter < YAMLTransactionCount; counter++)
            {
                
                //QueryQ is each individual YAML Transaction, not to be confused with sales transactions...
                string QueryQ = Query.Split(new string[] {"shipping_method: "},StringSplitOptions.None)[counter] + "shipping_method: " + Query.Split(new string[] {"shipping_method: "},StringSplitOptions.None)[counter+1].Split(new string[] {"\r\n"},StringSplitOptions.None)[0];

                OrderDetails od = new OrderDetails();
                
                //not sure if this is always true or if it's determined by the MISC_TX YAML code block.
                od.IsTaxable = true;



                od.BillingAddress1 = QueryQ.Split(new string[] { "bill_address_1: " }, StringSplitOptions.None)[counter+1].Split(new string[] {"\r\n"},StringSplitOptions.None)[0].Trim();
                od.BillingAddress2 = QueryQ.Split(new string[] { "bill_address_2: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                if (od.BillingAddress2 == "''")
                {
                    od.BillingAddress2 = "";
                }
                od.BillingCity = QueryQ.Split(new string[] { "bill_city: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.BillingCountry = QueryQ.Split(new string[] { "bill_country: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.BillingName = QueryQ.Split(new string[] { "bill_name: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.BillingPhone = QueryQ.Split(new string[] { "bill_phone: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.BillingState = QueryQ.Split(new string[] { "bill_state: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.BillingZipcode = QueryQ.Split(new string[] { "bill_zip: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.EmailAddress = QueryQ.Split(new string[] { "email: " }, StringSplitOptions.None)[counter + 1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();

                
                //now for the fun part. we must iterate through all items and transactions

                int YAMLTransactionLinesCount = Query.Split(new string[] {"shipping_method: "},StringSplitOptions.None)[counter].Split(new string[] { " - \r\n" }, StringSplitOptions.None).GetLength(0) - 1;
                List<ItemDetails> IDs = new List<ItemDetails>();
                List<TransactionDetails> TDs = new List<TransactionDetails>();
                
                for (int lineC = 0; lineC < YAMLTransactionLinesCount; lineC++)
                {
                    


                    string Line = QueryQ.Split(new string[] { " - \r\n" }, StringSplitOptions.None)[lineC+1];

                    
                    //get the type of line...
                    string LineType = Line.Split(new string[] { "type:" }, StringSplitOptions.None)[1].Split(new string[] {"\r\n"},StringSplitOptions.None)[0];

                    if (LineType.ToUpper().Trim() == "ITEM")
                    {
                        
                        ItemDetails ID = new ItemDetails();
                        ID.Description = Line.Split(new string[] { "description: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                        ID.ItemNumber = Line.Split(new string[] { "code: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                        ID.Quantity = int.Parse(Line.Split(new string[] { "quantity: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim());

                        //Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n" + Decimal.Parse(Line.Split(new string[] { "cost: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim().Split('\'')[1]).ToString());
                        //Environment.Exit(0);
                        ID.UnitPrice = Decimal.Parse(Line.Split(new string[] { "cost: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim().Split('\'')[1]);
                        ID.TotalItemCost = (decimal)(ID.UnitPrice * ID.Quantity);
                        IDs.Add(ID);
                        
                        
                    }
                    else if (LineType.ToUpper().Trim() == "MISC_TX")
                    {
                        TransactionDetails TX = new TransactionDetails();

                        //Console.WriteLine("Status: 200 OK\n Content-type: text/plain\n\n" + Decimal.Parse(Line.Split(new string[] { "cost: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim().Split('\'')[1]).ToString());
                        //Environment.Exit(0);
                        TX.Cost = Decimal.Parse(Line.Split(new string[] { "cost: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim().Split('\'')[1]);
                        string Taxable = Line.Split(new string[] { "taxable: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                        if (Taxable.Trim() == "1")
                        {
                            TX.IsTaxable = true;
                        }
                        else
                        {
                            TX.IsTaxable = false;
                        }
                        TX.Text = Line.Split(new string[] { "text: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                        TDs.Add(TX);


                    }
                    else
                    {
                        //eh, we weren't expecting this "type" of line code... let's email me a message letting me know what
                        //the hell just happened...
                        string ErrorMessage = "So, this script was built to detect two different types of \"Line\" types ";
                        ErrorMessage += "in a YAML Transaction Post to this script. We accept 'ITEM' and 'MISC_TX'. This ";
                        ErrorMessage += "line type was '" + LineType + "'. Maybe we need to add it?\r\n\r\nThanks,\r\n  ";
                        ErrorMessage += "Tom's Automated Response System\r\n\r\n\r\nYAML MESSAGE:\r\n" + YAMLinput + "\r\n";
                        ErrorMessage += "\r\nDECODED YAML MESSAGE:\r\n" + Query + "\r\n";

                        List<string> EmailAddresses = new List<string>();
                        EmailAddresses.Add("tom@tools-plus.com");

                        //bool sent = Emailer.SendEmail("Email Proxy Script Error",ErrorMessage,false,EmailAddresses);

                        Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\nType = " + LineType);
                        Environment.Exit(0);

                        
                    }


                }
               
                od.ItemsList = IDs;
                od.TransactionList = TDs;
                od.ItemCount = od.ItemsList.Count;
                //now process the order details themselves...
                od.OrderDate = QueryQ.Split(new string[] { "order_date: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.OrderNumber = QueryQ.Split(new string[] { "order_number: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                if (od.OrderNumber == "''")
                {
                    od.OrderNumber = "";
                }
                od.OrderStatus = QueryQ.Split(new string[] { "order_status: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.PaymentMethod = QueryQ.Split(new string[] { "payment_method: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.PONumber = QueryQ.Split(new string[] { "po_number: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                if (od.PONumber == "''")
                {
                    od.PONumber = "";
                }

                

                //now process the shipping info...
                od.ShippingAddress1 = QueryQ.Split(new string[] { "ship_address_1: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.ShippingAddress2 = QueryQ.Split(new string[] { "ship_address_2: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                if (od.ShippingAddress2 == "''")
                {
                    od.ShippingAddress2 = "";
                }
                od.ShippingCity = QueryQ.Split(new string[] { "ship_city: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.ShippingCountry = QueryQ.Split(new string[] { "ship_country: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.ShippingName = QueryQ.Split(new string[] { "ship_name: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
                od.ShippingPhone = QueryQ.Split(new string[] { "ship_phone: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();               
                od.ShippingState = QueryQ.Split(new string[] { "ship_state: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();              
                od.ShippingZipcode = QueryQ.Split(new string[] { "ship_zip: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();              
                od.ShippingMethod = QueryQ.Split(new string[] { "shipping_method: " }, StringSplitOptions.None)[1].Split(new string[] { "\r\n" }, StringSplitOptions.None)[0].Trim();
             

                
                //now subtotal, get tax amount, and total amounts...
                decimal SubTotal = 0;
                foreach (ItemDetails id in od.ItemsList)
                {
                    SubTotal += id.TotalItemCost;
                }
                decimal Tax = SubTotal * TaxRate;

                decimal GrandTotal = SubTotal + Tax;


                od.SubTotal = SubTotal;
                od.Total = GrandTotal;
                od.TaxAmount = Tax;

                
                //now add all that info to the main YAML Transaction List....
                tmpDetails.Add(od);


            }

            return tmpDetails;

        }


        public static void DebugOutput()
        {
            //orderDetailsList;
            Console.WriteLine("Status: 200 OK\nContent-type: text/html\n\n");
            Console.WriteLine("<html><body bgcolor=\"#000000\"><font color=\"#ffffff\">");
            Console.WriteLine("<b><h2>Total YAML Transactions: " + orderDetailsList.Count.ToString() + "</h2></b><br>");
            
            for (int x = 0; x < orderDetailsList.Count; x++)
            {
               Console.WriteLine("<table width=\"500px\" bgcolor=\"#FFFFFF\"><tr><td><b><center>YAML</center></b></td><td><b>Transaction #" + (x+1).ToString() + "</b></td></tr>");
               Console.WriteLine("<tr><td>Billing Info</td></tr>");
               Console.WriteLine("<tr><td>Billing Address:</td><td>" + orderDetailsList[x].BillingAddress1 + "</td></tr>");
               Console.WriteLine("<tr><td></td><td>" + orderDetailsList[x].BillingAddress2 + "</td></tr>");
               Console.WriteLine("<tr><td>Billing City:</td><td>" + orderDetailsList[x].BillingCity + "</td></tr>");
               Console.WriteLine("<tr><td>Billing Zipcode:</td><td>" + orderDetailsList[x].BillingZipcode + "</td></tr>");
               Console.WriteLine("<tr><td>Billing State:</td><td>" + orderDetailsList[x].BillingState + "</td></tr>");
               Console.WriteLine("<tr><td>Billing Country:</td><td>" + orderDetailsList[x].BillingCountry + "</td></tr>");
               Console.WriteLine("<tr><td>Billing Name:</td><td>" + orderDetailsList[x].BillingName + "</td></tr>");
               Console.WriteLine("<tr><td>Billing Phone:</td><td>" + orderDetailsList[x].BillingPhone + "</td></tr>");             
               Console.WriteLine("<tr><td>Email Address:</td><td>" + orderDetailsList[x].EmailAddress + "</td></tr>");

               Console.WriteLine("<tr><td>&nbsp</td><td>&nbsp</td></tr>");
               Console.WriteLine("<tr><td>Items</td></tr>");
               foreach (ItemDetails id in orderDetailsList[x].ItemsList)
               {
                   Console.WriteLine("<tr><td><hr></td><td><hr></td></tr>");
                   Console.WriteLine("<tr><td>Name:</td><td><font color=\"#0000DE\">" + id.ItemNumber + "</font></td></tr>");
                   Console.WriteLine("<tr><td>Description:</td><td>" + id.Description + "</td></tr>");
                   Console.WriteLine("<tr><td>Quantity:</td><td>" + id.Quantity + "</td></tr>");
                   Console.WriteLine("<tr><td>Unit Price:</td><td>" + "$" + id.UnitPrice + "</td></tr>");
                   
               }
               Console.WriteLine("<tr><td><hr></td><td><hr></td></tr>");
               Console.WriteLine("<tr><td>&nbsp</td><td>&nbsp</td></tr>");
               Console.WriteLine("<tr><td>Transactions</td></tr>");
               foreach (TransactionDetails tx in orderDetailsList[x].TransactionList)
               {
                   Console.WriteLine("<tr><td><hr></td><td><hr></td></tr>");
                   if (tx.IsTaxable == true)
                   {
                       Console.WriteLine("<tr><td>Is Taxable:</td><td>Yes.</td></tr>");
                   }
                   else
                   {
                       Console.WriteLine("<tr><td>Is Taxable:</td><td>No.</td></tr>");
                   }
                   Console.WriteLine("<tr><td>Trans-Type:</td><td>" + tx.Text + "</td></tr>");
                   Console.WriteLine("<tr><td>Cost:</td><td>" + "$" + tx.Cost.ToString("0.00") + "</td></tr>");
                   
               }
               Console.WriteLine("<tr><td><hr></td><td><hr></td></tr>");
               Console.WriteLine("<tr><td>&nbsp</td><td>&nbsp</td></tr>");
               Console.WriteLine("<tr><td>Shipping Info</td></tr>");
               Console.WriteLine("<tr><td>Shipping Address:</td><td>" + orderDetailsList[x].ShippingAddress1 + "</td></tr>");
               Console.WriteLine("<tr><td></td><td>" + orderDetailsList[x].ShippingAddress2 + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping City:</td><td>" + orderDetailsList[x].ShippingCity + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping Zipcode:</td><td>" + orderDetailsList[x].ShippingZipcode + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping State:</td><td>" + orderDetailsList[x].ShippingState + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping Country:</td><td>" + orderDetailsList[x].ShippingCountry + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping Name:</td><td>" + orderDetailsList[x].ShippingName + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping Phone:</td><td>" + orderDetailsList[x].ShippingPhone + "</td></tr>");
               Console.WriteLine("<tr><td>Shipping Method:</td><td>" + orderDetailsList[x].ShippingMethod + "</td></tr>");

               Console.WriteLine("<tr><td>&nbsp</td><td>&nbsp</td></tr>");
                Console.WriteLine("<tr><td>Total Items:</td><td>" + orderDetailsList[x].ItemCount.ToString() + "</td></tr>");
                Console.WriteLine("<tr><td>PO Number:</td><td><font color=\"#DE0000\">" + orderDetailsList[x].PONumber + "</font></td></tr>");
                Console.WriteLine("<tr><td>Order Number:</td><td><font color=\"#00DE00\">" + orderDetailsList[x].OrderNumber + "</font></td></tr>");
                Console.WriteLine("<tr><td>Order Date:</td><td>" + orderDetailsList[x].OrderDate + "</td></tr>");
                Console.WriteLine("<tr><td>Order Status:</td><td>" + orderDetailsList[x].OrderStatus + "</td></tr>");
                Console.WriteLine("<tr><td>Payment Method:</td><td>" + orderDetailsList[x].PaymentMethod + "</td></tr>");
                

                    
               Console.WriteLine("</table></body></html>");


            }
            Environment.Exit(0);
        }



        

    }
}
