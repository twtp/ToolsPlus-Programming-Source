using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;

namespace freightquote
{
    class Program
    {
        private const string Version = "1.0.0";


        //COMMON CALLS//////////////////////////////////
        public static string zipCode = "";
        public static string state = "";
        public static string country = "";
        public static string residential = "";
        public static string ups_username = "bdonorfio";
        public static string ups_password = "7539042";
        public static string ups_xml_key = "7BCADEDB93046728";
        public static string ups_account_number = "XX3034";
        public static string usps_userid = "683TOOLS0983";
        public static string usps_password = "279IT37QH898";

        public static string ship_from_zipcode = "06712";
        public static string ship_from_state = "CT";
        public static string ship_from_country = "US";
        public static string ups_domain = "onlinetools.ups.com";
        public static string ups_port = "443";
        public static string ups_path = "/ups.app/xml/Rate";
        public static string ups_contentType = "application/x-www-form-urlencoded";
        public static string ups_userAgent = "ToolsPlus UPS Rates/3.0";

        /////////////////////////////////////////////////

        //ITEM CALL/////////////////////////////////////
        public static List<string> itemsList = new List<string>();
        public static List<string> qtyList = new List<string>();

        ////////////////////////////////////////////////

        //BOX CALL///////////////////////////////////////
        public static List<string> boxweightList = new List<string>();
        public static List<string> boxlengthList = new List<string>();
        public static List<string> boxwidthList = new List<string>();
        public static List<string> boxheightList = new List<string>();
        /////////////////////////////////////////////////

        //ENGINE CALL////////////////////////////////////
        public static List<string> carrierList = new List<string>();
        /////////////////////////////////////////////////


        //PACKAGING//////////////////////////////////////
        public static List<PackageVolumeInfo> PackageInfo = new List<PackageVolumeInfo>();
        public static PackageVolumeInfo pvi = new PackageVolumeInfo();
        public static int pNum = 1;
        /////////////////////////////////////////////////




        public struct ComponentVolumeInfo
        {
            public int PackageNumber;
            public string ComponentName;
            public decimal Length;
            public decimal Width;
            public decimal Height;
            public decimal Weight;
            public decimal Volume;
            public int Quantity;
            public decimal LongestSide;
        }

        public struct PackageVolumeInfo
        {
            public int PackageNumber;
            public decimal Length;
            public decimal Width;
            public decimal Height;
            public decimal Volume;
            public decimal Weight;

        }


        
        //////////////////////////////////////////////////////////////
        // GET VERSION
        //      Gets the version number of this program
        //
        //////////////////////////////////////////////////////////////
      
        static bool GETVERSION()
        {
            if (System.Environment.GetEnvironmentVariable("QUERY_STRING").ToUpper().Contains("GETVERSION()") == true)
            {
                Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n");
                Console.WriteLine("Freight Quote Version: " + Version);
                return true;
            }
            else
            {
                return false;
            }

        }








        //////////////////////////////////////////////////////////////
        // Throw Bad Request
        //     This function returns a 400 error if called
        //
        //////////////////////////////////////////////////////////////

        static void ThrowBadRequest()
        {
            Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\n\n");
        }









        //////////////////////////////////////////////////////////////
        // Main Entry Point
        //    This method collects all arguments thrown at it.. should
        //  be made to accept either GET/POST or ARGS[] without recoding.
        //  Anyways, there are two primary methods, If requested 'type'
        //  is 'item' then the item processing section that leads to the
        //  item processing methods will be utilized. If requested 'type'
        //  was 'box' then the box processing section that leads to the
        //  box processing methods will be utilized. After the process
        //  completes in the called methods, this method returns the
        //  final result, whatever it may be.
        //
        //////////////////////////////////////////////////////////////


        static void Main(string[] args)
        {
            if (GETVERSION() == true)
            {
                return;
            }

            string queryString = System.Environment.GetEnvironmentVariable("QUERY_STRING");


            List<string> queryItems = new List<string>();

            foreach (string item in queryString.Split('&'))
            {
                queryItems.Add(item);
            }





            zipCode = queryString.Substring(queryString.ToUpper().IndexOf("ZIPCODE=") + 8).Split('&')[0];
            state = queryString.Substring(queryString.ToUpper().IndexOf("STATE=") + 6).Split('&')[0];
            country = queryString.Substring(queryString.ToUpper().IndexOf("COUNTRY=") + 8).Split('&')[0];
            residential = queryString.Substring(queryString.ToUpper().IndexOf("RESIDENTIAL=") + 12).Split('&')[0];






            int itemCount = 0;
            int qtyCount = 0;
            int carrierCount = 0;
            string currentType = "";

            bool hasZipCode = false;
            bool hasCountry = false;
            bool hasState = false;
            bool hasResidential = false;
            bool hasCarrier0 = false;

            string requestType = "";

            foreach (string isItem in queryItems)
            {
                if (isItem.ToUpper().StartsWith("ITEM") == true)
                {
                    itemCount++;
                }
                if (isItem.ToUpper().StartsWith("QTY") == true)
                {
                    qtyCount++;
                }
                if (isItem.ToUpper().StartsWith("CARRIER") == true)
                {
                    carrierCount++;
                }
                if (isItem.ToUpper().StartsWith("TYPE") == true)
                {
                    currentType = isItem.ToUpper().Split('=')[1];
                }
                if (isItem.ToUpper().StartsWith("ZIPCODE") == true)
                {
                    hasZipCode = true;
                }
                if (isItem.ToUpper().StartsWith("COUNTRY") == true)
                {
                    hasCountry = true;
                }
                if (isItem.ToUpper().StartsWith("STATE") == true)
                {
                    hasState = true;
                }
                if (isItem.ToUpper().StartsWith("RESIDENTIAL") == true)
                {
                    hasResidential = true;
                }
                if (isItem.ToUpper().StartsWith("CARRIER0") == true)
                {
                    hasCarrier0 = true;
                }

            }

            if (hasCarrier0 = true & hasCountry == true & hasResidential == true & hasState == true & hasZipCode == true)
            {
                //GOOD TO GO...
            }
            else
            {
                ThrowBadRequest();
                return;
            }


           
            for (int iCount = 0; iCount < itemCount; iCount++)
            {
                itemsList.Add(queryString.Substring(queryString.ToUpper().IndexOf("ITEM" + iCount.ToString()) + 5 + iCount.ToString().Length).Split('&')[0]);
                
            }
         
            for (int qCount = 0; qCount < qtyCount; qCount++)
            {
                qtyList.Add(queryString.Substring(queryString.ToUpper().IndexOf("QTY" + qCount.ToString()) + 4 + qCount.ToString().Length).Split('&')[0]);
            }
            
            for (int cCount = 0; cCount < carrierCount; cCount++)
            {
                carrierList.Add(queryString.Substring(queryString.ToUpper().IndexOf("CARRIER" + cCount.ToString()) + 8 + cCount.ToString().Length).Split('&')[0]);
            }

            if (currentType == "ITEM" | currentType == "BOX")
            {
                if (currentType == "ITEM")
                {
                    if (queryString.ToUpper().Contains("ITEM0") == true & queryString.ToUpper().Contains("QTY0") == true)
                    {
                        //Good To Go
                        requestType = "ITEM";
                    }
                    else
                    {
                        ThrowBadRequest();
                        return;
                    }



                }
                else if (currentType == "BOX")
                {
                    if (queryString.ToUpper().Contains("BOXWEIGHT0") == true &
                        queryString.ToUpper().Contains("BOXLENGTH0") == true &
                        queryString.ToUpper().Contains("BOXWIDTH0") == true &
                        queryString.ToUpper().Contains("BOXHEIGHT0") == true &
                        queryString.ToUpper().Contains("ZIPCODE") == true &
                        queryString.ToUpper().Contains("STATE") == true &
                        queryString.ToUpper().Contains("COUNTRY") ==true)
                    {
                        //GOOD TO GO...
                        requestType = "BOX";
                    }
                    else
                    {
                        ThrowBadRequest();
                        return;
                    }

                }


            }
            else
            {
                ThrowBadRequest();
                return;
            }


        if (requestType == "ITEM")
        {

            //ProcessDataIntoJSON();

            //this contains our package data
            string PackageInfo = CreatePackages();
            string JSON_Encoded = "";
            //this line is the JSON!

            foreach (string carrier in carrierList)
            {
                if (carrier.ToUpper() == "UPS")
                {
                    JSON_Encoded = CreateUPSQuery(PackageInfo);
                }
                if (carrier.ToUpper() == "USPS")
                {
                    JSON_Encoded += CreateUSPSQuery(PackageInfo);
                }

            }

            JSON_Encoded += "}";

            if (JSON_Encoded.Substring(JSON_Encoded.Length - 3, 3) != "}}}")
            {
                JSON_Encoded = JSON_Encoded.Substring(0, JSON_Encoded.Length - 2) + "}}";
            }


            //Write the final data to the output and call it a day!
            Console.WriteLine(JSON_Encoded);
            return;
        }


        if (requestType == "BOX")
        {
            //Now to process the 'BOX' in an easier, simpler, more efficient manner.
            //Also, this is used not nearly as often as 'ITEM'


            bool boxExists = true;
            int Box_x = 0;
            while (boxExists == true)
            {
                try
                {
                    boxweightList.Add(queryString.Split(new string[] { "boxweight" + Box_x.ToString() + "=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    boxheightList.Add(queryString.Split(new string[] { "boxheight" + Box_x.ToString() + "=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    boxlengthList.Add(queryString.Split(new string[] { "boxlength" + Box_x.ToString()+ "=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    boxwidthList.Add(queryString.Split(new string[] { "boxwidth" + Box_x.ToString() + "=" }, StringSplitOptions.None)[1].Split('&')[0]);

                   


                }
                catch
                {
                    boxExists = false;
                }
                Box_x++;
            }
            string JSON_BOXCODE = "";


            JSON_BOXCODE += CreateBoxesJSON();


            foreach (string carrier in carrierList)
            {
                if (carrier.ToUpper() == "UPS")
                {
                    //GET OUR UPS RESPONSE!!
                    JSON_BOXCODE += UPS_processBoxQuery();
                }
                if (carrier.ToUpper() == "USPS")
                {
                    JSON_BOXCODE += USPS_processBoxQuery();
                }

            }

            if (JSON_BOXCODE.Substring(JSON_BOXCODE.Length - 2, 2) == "},")
            {
                JSON_BOXCODE = JSON_BOXCODE.Substring(0,JSON_BOXCODE.Length - 2) + "}}}";
            }

            Console.WriteLine("HTTP/1.1 200 OK\nContent-type: text/plain\n\n\n" + JSON_BOXCODE);
            return;
        }

        ThrowBadRequest();
        return;

        }








        //////////////////////////////////////////////////////////////
        // Create Packages Method
        //     This method takes our itemList global that was made in
        //   the main method. We use the itemnumbers to call up sql and
        //   get a list of data about said items. Now, after taking all
        //   data in and placing it in the appropriate variables, we
        //   JSON encode it because Brian was "special". Really, JSON
        //   should never be utilized again. Who wants a one line code
        //   that is a million miles long and takes 20 minutes to read.
        //   Anyways, take the JSON encoded work and send it back to the
        //   calling method.
        //////////////////////////////////////////////////////////////

        public static string CreatePackages()
        {
            string ReturnData = "HTTP/1.1 200 OK\nContent-type: text/plain\n\n\n";
            ReturnData += "{\"packages\":[";

            int runningQtyCount = 0;
            decimal TotWeight = 0;
            List<string> PackageData = new List<string>();
            for (int x = 0; x < itemsList.Count; x++)
            {
                //if (x == 0)
                //{
                //    ReturnData += "{\"cube\":null,\"width\":null,\"dimensional_weight\":null,\"contents\":{";
                //}

                //1.) Get Item Data - 
                List<string> ItemData = SQLCalls.sqlQuery("SELECT * FROM InventoryComponents,InventoryComponentMap WHERE ID=ComponentID AND ItemNumber='" + itemsList[x] + "'");

                //2.) Parse Out Item Name
                string ItemName = ItemData[0].Split('|')[1];

                //3.) Parse Out Weight
                decimal Weight = decimal.Parse(ItemData[0].Split('|')[5]);

                //4.) Parse Out Qty
                int Quantity = int.Parse(ItemData[0].Split('|')[18]);

                //5.) Parse Out Width
                decimal Width = decimal.Parse(ItemData[0].Split('|')[3]);

                //6.) Parse Out Height
                decimal Height = decimal.Parse(ItemData[0].Split('|')[4]);

                //7.) Parse Out Length
                decimal Length = decimal.Parse(ItemData[0].Split('|')[2]);

                //8.) Parse Out ComponentID
                string ComponentID = ItemData[0].Split('|')[0];

                //9.) Parse Out HazMat
                string isHazMat = "";
                if (ItemData[0].Split('|')[9] == "1")
                {
                    isHazMat = "true";
                }
                else
                {
                    isHazMat = "false";
                }

                //10.) Parse Out YahooID (just use itemnumber for now)
                string YahooID = ItemName;
                YahooID = YahooID.Substring(0, 3) + "-" + YahooID.Substring(3);
                if (Quantity > 1)
                {
                    YahooID += "-" + Quantity.ToString();
                }
                //Console.WriteLine("Debug #3: " + YahooID);

                //11.) Parse Out UniqueID
                string UniqueID = ItemName;
                if (Quantity > 1)
                {
                    UniqueID += "-" + Quantity.ToString();
                    ItemName += "-" + Quantity.ToString();
                }
                UniqueID += "_" + ComponentID;

                //Console.WriteLine("Debug #1,4: " + UniqueID);
                //Console.WriteLine("Debug #2: " + ItemName);

                int packageStart = 0;
                bool justCreatedPackage = false;
                int xCount = 0;
                for (xCount = 0; xCount < int.Parse(qtyList[x]); xCount++)
                {
                    justCreatedPackage = false;
                    runningQtyCount++;
                    TotWeight += (Weight * Quantity);
                    if (TotWeight > 60)
                    {
                        ReturnData += "\"" + UniqueID + "\":{\"component\":{";
                        ReturnData += "\"width\":\"" + Width.ToString() + "\",";
                        ReturnData += "\"quantity\":\"" + Quantity.ToString() + "\",";
                        ReturnData += "\"component_id\":\"" + ComponentID + "\",";
                        ReturnData += "\"height\":\"" + Height.ToString() + "\",";
                        ReturnData += "\"length\":\"" + Length.ToString() + "\",";
                        ReturnData += "\"hazmat\":\"" + isHazMat + "\",";
                        ReturnData += "\"item_number\":\"" + ItemName + "\",";
                        ReturnData += "\"weight\":\"" + Weight.ToString() + "\",";
                        ReturnData += "\"yahoo_id\":\"" + YahooID + "\",";
                        ReturnData += "\"unique_id\":\"" + UniqueID + "\"},";
                        TotWeight = Math.Ceiling(TotWeight) + 2;
                        ReturnData += "\"quantity\":" + (int.Parse(ItemData[0].Split('|')[18]) * runningQtyCount).ToString() + "}},\"height\":null,\"length\":null,\"girth\":null,\"weight\":" + ((int)TotWeight).ToString() + "},";
                        //Console.WriteLine("Debug #6: " + ReturnData);
                        //int.Parse(qtyList[x])).ToString()

                        runningQtyCount = 0;

                        if (x != itemsList.Count - 1 & xCount != Quantity - 1)
                        {
                            ReturnData += "{\"cube\":null,\"width\":null,\"dimensional_weight\":null,\"contents\":{";
                        }
                        packageStart = xCount;
                        TotWeight = 0;
                        justCreatedPackage = true;
                        //PROCESS AND COMPLETE PACKAGE.. THEN START ANOTHER PACKAGE IF THERES MORE
                        //PackageData.Add(ItemName + "|" + xCount.ToString() + (xCount - packagedQty).ToString());
                    }
                }

                if (justCreatedPackage == false)
                {
                    ReturnData += "{\"cube\":null,\"width\":null,\"dimensional_weight\":null,\"contents\":{";
                    ReturnData += "\"" + UniqueID + "\":{\"component\":{";
                    ReturnData += "\"width\":\"" + Width.ToString() + "\",";
                    ReturnData += "\"quantity\":\"" + Quantity.ToString() + "\",";
                    ReturnData += "\"component_id\":\"" + ComponentID + "\",";
                    ReturnData += "\"height\":\"" + Height.ToString() + "\",";
                    ReturnData += "\"length\":\"" + Length.ToString() + "\",";
                    ReturnData += "\"hazmat\":\"" + isHazMat + "\",";
                    ReturnData += "\"item_number\":\"" + ItemName + "\",";
                    ReturnData += "\"weight\":\"" + Weight.ToString() + "\",";
                    ReturnData += "\"yahoo_id\":\"" + YahooID + "\",";
                    ReturnData += "\"unique_id\":\"" + UniqueID + "\"},";
                    ReturnData += "\"quantity\":" + (int.Parse(ItemData[0].Split('|')[18]) * runningQtyCount).ToString() + "},";  //int.Parse(qtyList[x])).ToString() + "},";
                    runningQtyCount = 0;

                    //Console.WriteLine("Debug #5: " + ReturnData);

                    if (x == itemsList.Count - 1)
                    {
                        ReturnData = ReturnData.Substring(0, ReturnData.Length - 1);
                        ReturnData += "},";
                        //ReturnData += "\"quantity\":" + runningQtyCount.ToString() + "},";
                        ReturnData += "\"height\":null,\"length\":null,\"girth\":null,\"weight\":" + ((int)Math.Ceiling(TotWeight) + 2).ToString() + "},";

                    }
                    //TotWeight = 0;

                }





            }

            ReturnData = ReturnData.Substring(0, ReturnData.Length - 1);
            ReturnData += "],";
            return ReturnData;

        }



        //////////////////////////////////////////////////////////////
        // Create UPS Query
        //     As the name implies, all this function does is prepares
        //   the XML query to send to ups, takes in the JSON made in
        //   the CreatePackages() method, and then produces a spiffy
        //   XML Sheet (also utilizing PackageSizeInfo() method). This
        //   Gets pumped out to SendUPSQuery() to actually get the data
        //   over to UPS, and catch it's response. 



        public static string CreateUPSQuery(string PackageData)
        {


            //string Ship_From_State="CT";
            //string Ship_From_Zipcode="06712";
            //string Ship_From_Country="US";
            string Ship_To_State = state;
            string Ship_To_Zipcode = zipCode;
            string Ship_To_Country = country;




            string QueryDefault = "<?xml version=\"1.0\"?>\n";
            QueryDefault += "<AccessRequest xml:lang=\"en-US\">\n";
            QueryDefault += " <AccessLicenseNumber>" + ups_xml_key + "</AccessLicenseNumber>\n";
            QueryDefault += " <UserId>" + ups_username + "</UserId>\n";
            QueryDefault += " <Password>" + ups_password + "</Password>\n";
            QueryDefault += "</AccessRequest>\n";
            QueryDefault += "<?xml version=\"1.0\"?>\n";
            QueryDefault += "<RatingServiceSelectionRequest xml:lang=\"en-US\">\n";
            QueryDefault += " <Request>\n";
            QueryDefault += "  <TransactionReference>\n";
            QueryDefault += "   <CustomerContext>Rating and Service</CustomerContext>\n";
            QueryDefault += "   <XpciVersion>1.0001</XpciVersion>\n";
            QueryDefault += "  </TransactionReference>\n";
            QueryDefault += "  <RequestAction>Rate</RequestAction>\n";
            QueryDefault += "  <RequestOption>Shop</RequestOption>\n";
            QueryDefault += " </Request>\n";
            QueryDefault += " <PickupType>\n";
            QueryDefault += "  <Code>01</Code>\n";
            QueryDefault += " </PickupType>\n";
            QueryDefault += " <CustomerClassification>\n";
            QueryDefault += "  <Code>01</Code>\n";
            QueryDefault += " </CustomerClassification>\n";
            QueryDefault += " <Shipment>\n";
            QueryDefault += "  <Shipper>\n";
            QueryDefault += "   <ShipperNumber>" + ups_account_number + "</ShipperNumber>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + ship_from_state + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + ship_from_zipcode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + ship_from_country + "</CountryCode>\n";
            QueryDefault += "   </Address>\n";
            QueryDefault += "  </Shipper>\n";
            QueryDefault += "  <RateInformation>\n";
            QueryDefault += "   <NegotiatedRatesIndicator />\n";
            QueryDefault += "  </RateInformation>\n";
            QueryDefault += "  <ShipFrom>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + ship_from_state + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + ship_from_zipcode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + ship_from_country + "</CountryCode>\n";
            QueryDefault += "   </Address>\n";
            QueryDefault += "  </ShipFrom>\n";
            QueryDefault += "  <ShipTo>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + Ship_To_State + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + Ship_To_Zipcode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + Ship_To_Country + "</CountryCode>\n";

            if (residential == "1")
            {
                QueryDefault += "    <ResidentialAddressIndicator />\n";
            }

            QueryDefault += "   </Address>\n";
            QueryDefault += "  </ShipTo>\n";

            int totalPackages = 0;
            int i = 0;
            while ((i = PackageData.IndexOf("\"cube\"", i)) != -1)
            {
                i += "\"cube\"".Length;
                totalPackages++;
            }



           

            List<PackageVolumeInfo> pList = PackageSizeInfo(PackageData);

            
            for (int x = 0; x < pList.Count; x++)
            {
                QueryDefault += "  <Package>\n";
                QueryDefault += "   <PackagingType>\n";
                QueryDefault += "    <Code>02</Code>\n";
                QueryDefault += "   </PackagingType>\n";
                QueryDefault += "   <Dimensions>\n";
                QueryDefault += "    <Length>" + Math.Ceiling(pList[x].Length).ToString() + "</Length>\n";
                QueryDefault += "    <Width>" + Math.Ceiling(pList[x].Width).ToString() + "</Width>\n";
                QueryDefault += "    <Height>" + Math.Ceiling(pList[x].Height).ToString() + "</Height>\n";
                QueryDefault += "   </Dimensions>\n";
                QueryDefault += "   <PackageWeight>\n";
                QueryDefault += "    <UnitOfMeasurement>\n";
                QueryDefault += "     <Code>LBS</Code>\n";
                QueryDefault += "    </UnitOfMeasurement>\n";
                QueryDefault += "    <Weight>" + Math.Ceiling(pList[x].Weight).ToString() + "</Weight>\n";
                QueryDefault += "   </PackageWeight>\n";
                QueryDefault += "  </Package>\n";



            }


            QueryDefault += " </Shipment>\n";
            QueryDefault += "</RatingServiceSelectionRequest>\n";

            //Console.WriteLine("Wrote UPS XML Query. Now to send it to them for results.");

            string UPSResults = SendUPSQuery(QueryDefault);
            string UPSParsedResults = ParseUPSResults(UPSResults);


            string returnData = PackageData + "\"carriers\":{\"UPS\":{\"xml_sent\":\"" + QueryDefault.Replace("\n", "\\n") + "\",";
            returnData += UPSParsedResults;

            //Console.WriteLine("Returning JSON Encoding including XML Query...");
            return returnData;



        }


        public static List<PackageVolumeInfo> PackageSizeInfo(string packageData)
        {

            //Console.WriteLine("HTTP/1.1 200 OK\nContent-type: text/plain\n\n\n");
            
            //So it looks like we programmatically create the package size
            //1.)Let's Pull Out Individual Packages!
            List<string> Packages = new List<string>();
            List<int> Offsets = new List<int>();
            int totalPackages = 0;
            int i = 0;
            while ((i = packageData.IndexOf("\"cube\"", i)) != -1)
            {
                Offsets.Add(i);
                i += "\"cube\"".Length;
                totalPackages++;
            }

            for (int pCount = 0; pCount < Offsets.Count; pCount++)
            {
                try
                {

                    Packages.Add(packageData.Substring(Offsets[pCount] - 2, Offsets[pCount + 1] - (Offsets[pCount] - 2)));


                }
                catch
                {
                    //last one doesnt have another "cube", but has a "]"...

                    int endOffset = packageData.IndexOf(']');
                   


                    Packages.Add(packageData.Substring(Offsets[pCount] - 2));

                }

            }

            if (totalPackages == 2)
            {
                //Console.WriteLine("3 packages...\r\n");
                foreach (string tmp in Packages)
                {
                    //Console.WriteLine(tmp);
                }
            }


            List<ComponentVolumeInfo> componentData = new List<ComponentVolumeInfo>();
            List<PackageVolumeInfo> returnData = new List<PackageVolumeInfo>();
            int packageNumber = 0;
            foreach (string Package in Packages)
            {
                packageNumber++;


                string[] ComponentsInPackage = Package.Split(new string[] { "\"width\":\"" }, StringSplitOptions.None);




                List<string> Components = new List<string>();
                foreach (string str in ComponentsInPackage)
                {
                    

                    if (str.Contains("\"cube\":") == false)
                    {
                        Components.Add("\"" + str);
                    }





                }
               
                //Console.WriteLine("Debugging Time: \r\n");
                foreach (string str in Components)
                {
                    ComponentVolumeInfo cvi = new ComponentVolumeInfo();



                    cvi.ComponentName = str.Split(new string[] { "\"item_number\":" }, StringSplitOptions.None)[1].Split('\"')[1];
                    cvi.Width = decimal.Parse(str.Split('\"')[1]);
                    cvi.Quantity = int.Parse(str.Split(new string[] { "\"quantity\":" }, StringSplitOptions.None)[2].Split('}')[0]);
                    cvi.Height = decimal.Parse(str.Split(new string[] { "\"height\":" }, StringSplitOptions.None)[1].Split('\"')[1]);
                    cvi.Length = decimal.Parse(str.Split(new string[] { "\"length\":" }, StringSplitOptions.None)[1].Split('\"')[1]);
                    cvi.Weight = decimal.Parse(str.Split(new string[] { "\"weight\":" }, StringSplitOptions.None)[1].Split('\"')[1]);
                    cvi.Volume = cvi.Height * cvi.Length * cvi.Width;
                    cvi.PackageNumber = packageNumber;
                    cvi.LongestSide = Math.Max(Math.Max(cvi.Height, cvi.Length), cvi.Width);
                    componentData.Add(cvi);
                    




                }


                //for each component in specific package, create a "box" for them


                pvi.PackageNumber = packageNumber;
                decimal totVolume = 0;
                decimal Width_Height = 0;
                decimal LargestSide = 0;
                foreach (ComponentVolumeInfo cvi in componentData)
                {
                    //Console.WriteLine("------------------------------------------------------------");
                    //Console.WriteLine("So its package " + cvi.PackageNumber.ToString() + "... cvi is package " + cvi.PackageNumber.ToString());
                    if (cvi.PackageNumber > packageNumber)
                    {
                        //Console.WriteLine("Time to finalize this package as this item is for another package...");
                        break;
                    }
                    if (cvi.PackageNumber == packageNumber)
                    {
                        totVolume += cvi.Volume * cvi.Quantity;
                        pvi.Weight += cvi.Weight * cvi.Quantity;
                        LargestSide = Math.Max(Math.Max(cvi.Length, cvi.Width), cvi.Height);
                        //Console.WriteLine("The total volume so far is " + (totVolume).ToString());
                        //Console.WriteLine("The total weight so far is " + (pvi.Weight).ToString());
                        //Console.WriteLine("The largest side so far is " + (LargestSide).ToString());
                        //Console.WriteLine("---------------------------------------------------");
                    }
                    else
                    {
                        //Console.WriteLine("Ah, we already took care of this package!");
                    }

                }
                //Console.WriteLine("Finalization Begins!");
                Width_Height = (decimal)Math.Sqrt((double)(totVolume / LargestSide));
                pvi.Width = Math.Ceiling(Width_Height);
                pvi.Height = Math.Ceiling(Width_Height);
                pvi.Length = Math.Ceiling(LargestSide);
                pvi.Volume = pvi.Width * pvi.Height * pvi.Length;
                PackageInfo.Add(pvi);
                pvi.Weight = 0;




            }

            List<PackageVolumeInfo> tmpPackageInfo = new List<PackageVolumeInfo>();
            foreach (PackageVolumeInfo pvi2 in PackageInfo)
            {
                
                tmpPackageInfo.Add(FindBoxToFit(pvi2));
              


            }
            PackageInfo = tmpPackageInfo;

           
            return PackageInfo;


        }


        public static PackageVolumeInfo FindBoxToFit(PackageVolumeInfo package)
        {
            //in case no box fits, this will make sure we return the original data.
            PackageVolumeInfo tmpPVI = package;

            string boxFindQuery = "SELECT * FROM BoxDimensions WHERE Deprecated=0 AND Length > " +
                Math.Ceiling(package.Length).ToString() + " AND Width >= " +
                Math.Ceiling(package.Width).ToString() + " AND Height >= " +
                Math.Ceiling(package.Height).ToString() + " AND (MaximumWeight >= " + package.Weight.ToString() +
                " OR MaximumWeight IS NULL) ORDER BY Weight";

            //Console.WriteLine("Querying Database for Boxes");
            List<string> Boxes = SQLCalls.sqlQuery(boxFindQuery);
            if (Boxes.Count > 0)
            {
                //Console.WriteLine("Found " + Boxes.Count.ToString() + " Boxes. Using the smallest available..");
                tmpPVI.Height = decimal.Parse(Boxes[0].Split('|')[3]);
                tmpPVI.Width = decimal.Parse(Boxes[0].Split('|')[2]);
                tmpPVI.Length = decimal.Parse(Boxes[0].Split('|')[1]);
                tmpPVI.Weight = package.Weight + decimal.Parse(Boxes[0].Split('|')[4]);
                tmpPVI.PackageNumber = package.PackageNumber;
                tmpPVI.Volume = tmpPVI.Width * tmpPVI.Height * tmpPVI.Length;
              


            }
            else
            {
                //IDK man! no boxes to fit this... damn what should i do?!?!
                //Console.WriteLine("Unfortunately no box will fit this package. That sucks!");
            }

            return tmpPVI;

        }

        public static string SendUPSQuery(string query)
        {
            string URL = "https://" + ups_domain + ups_path;
            //Console.WriteLine("Created URL: " + URL);


            return PostHTTPData(URL, query);



        }


        private static string PostHTTPData(string Url, string postData)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "POST";

            request.ContentType = ups_contentType;
            request.UserAgent = ups_userAgent;

            byte[] bytes = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = bytes.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(bytes, 0, bytes.Length);

            WebResponse response = request.GetResponse();
            Stream stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);

            string result = reader.ReadToEnd();
            //Console.WriteLine("Holy shit... we got data!");
            stream.Dispose();
            reader.Dispose();

            //Console.WriteLine("UPS SENT US THIS IN RETURN:\r\n");
            //Console.WriteLine(result);
            //Console.WriteLine("===============================================================");
            return result;

        }


        public static string ParseUPSResults(string UPSResults)
        {
            UPSResults = UPSResults.Replace("\n", "\\n");
            string parsedResults = "";
            string resultsState = UPSResults.Split(new string[] { "<ResponseStatusCode>" }, StringSplitOptions.None)[1].Split(new string[] { "</ResponseStatusCode>" }, StringSplitOptions.None)[0];
           

            parsedResults += "\"state_desc\":\"" + UPSResults.Split(new string[] { "<ResponseStatusDescription>" },
                StringSplitOptions.None)[1].Split(new string[] { "</ResponseStatusDescription>" },
                StringSplitOptions.None)[0].ToLower() + "\",\"results\":{";

            //filter the shipment choices out of the response

            string[] TmpRatedShipments = UPSResults.Split(new string[] { "<RatedShipment>" }, StringSplitOptions.None);
            List<string> RatedShipments = new List<string>();
            //Console.WriteLine("Shipment Choices: \r\n");
            foreach (string rs in TmpRatedShipments)
            {
                if (rs.Substring(0, 9) == "<Service>")
                {
                    RatedShipments.Add(rs);
                    //Console.WriteLine(rs);
                }
                else
                {
                    //Console.WriteLine(rs.Substring(0, 9) + " did not contain <Service> tag. Skipping.");
                }
            }
            //Console.WriteLine("\r\n=======================================================");

            foreach (string result in RatedShipments)
            {
                //now we need to parse all results, and gather "code", "tp_rate", "negotiated_rate", "list_rate", "service"


                string Code = result.Split(new string[] { "<Code>" }, StringSplitOptions.None)[1].Split(new string[] { "</Code>" }, StringSplitOptions.None)[0];
                //string TP_Rate = result.Split(new string[] {"<"};
                foreach (string mv in result.Split(new string[] { "<MonetaryValue>" }, StringSplitOptions.None))
                {
                    //Console.WriteLine("MonetaryValue Value: " + mv);
                }

                string Negotiated_Rate = result.Split(new string[] { "<GrandTotal>" }, StringSplitOptions.None)[1].Split(new string[] { "</MonetaryValue>" }, StringSplitOptions.None)[0];
                Negotiated_Rate = Negotiated_Rate.Split(new string[] { "<MonetaryValue>" }, StringSplitOptions.None)[1];


                string List_Rate = result.Split(new string[] { "<TotalCharges>" }, StringSplitOptions.None)[1].Split(new string[] { "</MonetaryValue>" }, StringSplitOptions.None)[0]; //acts like total charges
                List_Rate = List_Rate.Split(new string[] { "<MonetaryValue>" }, StringSplitOptions.None)[1];


                string TP_Rate = ((decimal)(float.Parse(Negotiated_Rate) / 0.83f)).ToString("0.##");

                string Service = result.Split(new string[] { "<Code>" }, StringSplitOptions.None)[1].Split(new string[] { "</Code>" }, StringSplitOptions.None)[0];
                string ServiceCode = Service;
                if (ServiceCode.Length == 1)
                {
                    ServiceCode = "0" + ServiceCode;
                }
                Service = GetUPSServiceName(int.Parse(Service));
                //Console.WriteLine("Service: " + Service);
                //Console.WriteLine(" Negotiated Rate: " + Negotiated_Rate); //acts like grand total
                //Console.WriteLine(" List Rate: " + List_Rate);
                //Console.WriteLine(" TP Rate: " + TP_Rate);
                //Console.WriteLine("=====================================");

                parsedResults += "\"" + ServiceCode + "\":{\"tp_rate\":" + TP_Rate + ",\"negotiated_rate\":\"" + Negotiated_Rate + "\",\"list_rate\":\"" + List_Rate + "\",\"service\":\"" + Service + "\"},";


            }

            parsedResults = parsedResults.Substring(0, parsedResults.Length - 1);
            parsedResults += "},\"xml_received\":\"" + UPSResults + "\",\"state\":" + resultsState + "},";

            return parsedResults;
        }

        public static string GetUPSServiceName(int ServiceCode)
        {
            switch (ServiceCode)
            {
                case 1:
                    return "Next Day Air";
                case 2:
                    return "2nd Day Air";
                case 3:
                    return "Ground";
                case 7:
                    return "Worldwide Express";
                case 8:
                    return "Worldwide Expedited";
                case 11:
                    return "Standard";
                case 12:
                    return "3 Day Select";
                case 13:
                    return "Next Day Air Saver";
                case 14:
                    return "Next Day Air Early AM";
                case 54:
                    return "Worldwide Express Plus";
                case 59:
                    return "2nd Day Air AM";
                case 65:
                    return "World Wide Saver";
            }
            return "Unknown Shipping Method";
        }


        public static string CreateUSPSQuery(string PackageData)
        {
            string returnDataHeader = "\"USPS\":{\"xml_sent\":\"";
            string returnData = "<RateV4Request USERID=\"" + usps_userid + "\" PASSWORD=\"" + usps_password + "\"> ";
            returnData += " <Revision>2</Revision>";

            //UPS Module did all the work for us in creating the packages, sizes, weight, etc... so lets use it
            //instead of being redundant!

            int packageCounter = 0;
            foreach (PackageVolumeInfo pvi in PackageInfo)
            {
                returnData += " <Package ID=\"" + packageCounter.ToString() + "\">";
                returnData += "  <Service>All</Service>";
                returnData += "  <ZipOrigination>" + ship_from_zipcode + "</ZipOrigination>";
                returnData += "  <ZipDestination>" + zipCode + "</ZipDestination>";

                decimal pounds = (decimal)Math.Floor(pvi.Weight);
                decimal fractions = pvi.Weight - pounds;
                decimal ounces = 0;

                if (fractions > 0)
                {
                    ounces = fractions * 16;
                }





                returnData += "  <Pounds>" + pounds.ToString() + "</Pounds>";
                returnData += "  <Ounces>" + ounces.ToString() + "</Ounces>";
                returnData += "  <Container>RECTANGULAR</Container>";
                returnData += "  <Size>LARGE</Size>";
                returnData += "  <Width>" + pvi.Width + "</Width>";
                returnData += "  <Length>" + pvi.Length + "</Length>";
                returnData += "  <Height>" + pvi.Height + "</Height>";


                //now to figure out girth: take 2 longest sides, multiply each figure by two, and add them together.
                decimal girth = 0;
                decimal largestSide = 0;
                decimal secondLargestSide = 0;

                largestSide = Math.Max(Math.Max(pvi.Length, pvi.Width), pvi.Height);

                if (pvi.Length == largestSide)
                {
                    secondLargestSide = Math.Max(pvi.Height, pvi.Width);
                }
                if (pvi.Width == largestSide)
                {
                    secondLargestSide = Math.Max(pvi.Height, pvi.Length);
                }
                if (pvi.Height == largestSide)
                {
                    secondLargestSide = Math.Max(pvi.Length, pvi.Width);
                }

                girth = (largestSide * 2) + (secondLargestSide * 2);


                returnData += "  <Girth>" + girth.ToString() + "</Girth>";
                returnData += "  <Machinable>True</Machinable>";
                returnData += " </Package>";


                packageCounter++;
            }


            returnData += "</RateV4Request>";

            string ResponseHeader = "\",\"state_desc\":\"";
            //Console.WriteLine("USPS DATA:");
            //Console.WriteLine(returnDataHeader + returnData);
            //return returnDataHeader + returnData;
            string USPSQuery = SendUSPSQuery(returnData);
            //Console.WriteLine("USPS QUERY: " + USPSQuery);

            //string responseCode = USPSQuery.Split(new string[] { "<ResponseStatusCode>" },StringSplitOptions.None)[1].Split(new string[] {"</ResponseStatusCode>"},StringSplitOptions.None)[0];
            //string responseDetail = USPSQuery.Split(new string[] { "<ResponseStatusDescription>" }, StringSplitOptions.None)[1].Split(new string[] { "</ResponseStatusDescription>" }, StringSplitOptions.None)[0];

            string responseDetail = "";
            if (USPSQuery.Contains("<RateV4Response>") == true)
            {
                responseDetail = "success";

            }
            else
            {
                responseDetail = "failure";
            }

            USPSQuery = USPSQuery.Replace("\n", "\\n");

            ResponseHeader += responseDetail + "\",\"results\":{";

            //now to parse the results! fun is...
            List<string> USPS_Parsings = new List<string>();
            string[] uspsParsings = USPSQuery.Split(new string[] { "<Package " }, StringSplitOptions.None);

            //Console.WriteLine("USPS Results: ");
            foreach (string str in uspsParsings)
            {
                if (str.Substring(0, 2) == "ID")
                {
                    USPS_Parsings.Add(str);
                    //Console.WriteLine(str);
                }


            }
            //Console.WriteLine("===========================================");
            string parsedData = "";

   
            // NOTE: First look through all packages and find all common shipping methods.
            //       Next, take all pricings for that shipping method and add em up.
            //       Finally, compute the TP_rate and the other rate (either list or negotiated; i forget)
            // "3-Priority Mail Express 2-Day":{"tp_rate":0.00,"negotiated_rate":0.00,"list_rate":0.00,"service":"Priority Mail Express 2-Day"},
            // Also seems as tho list and negotiated share the same value

            //USPS_Parsings

            List<string> USPS_ShipInfo = new List<string>();
            List<string> USPS_AllShippingInfo = new List<string>();
            int TotPackages = USPS_Parsings.Count;


            foreach (string pkg in USPS_Parsings)
            {
                //string PostageData = pkg.Split('\"')[1] + "|" + pkg.Split(new string[] { "<Postage CLASSID=" }, StringSplitOptions.None)[1].Split(new string[] { "</Postage>" },StringSplitOptions.None)[0];

                foreach (string postInfo in pkg.Split(new string[] { "<Postage CLASSID=" }, StringSplitOptions.None))
                {
                    if (postInfo[0] == '\"')
                    {
                        //Console.WriteLine("Individual Postage: Package #" + pkg.Split('\"')[1] + " - " + postInfo);
                        USPS_AllShippingInfo.Add(pkg.Split('\"')[1] + "|" + postInfo);
                    }
                }
            }

            //JSON Time... :(
            string JSON_Addition = "";

            foreach (string pkg in USPS_AllShippingInfo)
            {
                JSON_Addition += "\"" + pkg.Split('\"')[1] + "-";
                JSON_Addition += pkg.Split(new string[] { "<MailService>" }, StringSplitOptions.None)[1].Split(new string[] { "</MailService>" }, StringSplitOptions.None)[0];
                JSON_Addition += "\":{\"tp_rate\":" + ((decimal.Parse(pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0]) * (decimal)1.2)).ToString("0.##");
                JSON_Addition += ",\"negotiated_rate\":" + pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0];
                JSON_Addition += ",\"list_rate\":" + pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0];
                JSON_Addition += ",\"service\":\"" + pkg.Split(new string[] { "<MailService>" }, StringSplitOptions.None)[1].Split(new string[] { "</MailService>" }, StringSplitOptions.None)[0];
                JSON_Addition += "\"},";

            }

            JSON_Addition = JSON_Addition.Substring(0, JSON_Addition.Length - 1);

            JSON_Addition += "},";

            //NOW TO ADD THE XML RAW DATA WE GOT FROM USPS...

            JSON_Addition += "\"xml_received\":\"";
            JSON_Addition += USPSQuery;
            JSON_Addition += "\",";
            JSON_Addition += "\"state\":" + "1"; //WHERE IS THIS 'STATE' Value coming from? ASSUMING 0 or 1 depending on error.
            JSON_Addition += "}}";
            //LAST LINE OF JSON ENCODING!!!! YAY!!!




            return returnDataHeader + returnData + ResponseHeader + parsedData + USPSQuery + JSON_Addition;



        }


        public static string SendUSPSQuery(string query)
        {
            //firstly usps has two urls, one for domestic and one for international.
            //zipcode can be used to determine this, so lets zip map!

            //also, domestics include: USA,US,AS,FM,PR,VI,UM

            List<string> Countries = SQLCalls.sqlQuery("SELECT Country FROM LookupCityZipCode WHERE ZipCode='" + zipCode + "'");

            bool isDomestic = false;

            if (Countries.Count < 1)
            {
                //Console.WriteLine("Damn, no zipcodes here!");
                return "No ZipCodes.";
            }

            string country = Countries[0].Split('|')[0].ToUpper();

            if (country == "USA" | country == "US" | country == "AS" | country == "FM" | country == "PR" | country == "VI" | country == "UM")
            {
                //Console.WriteLine("SHIPMENT IS DOMESTIC");
                isDomestic = true;
            }
            else
            {
                //Console.WriteLine("SHIPMENT IS INTERNATIONAL");
            }

            string URL_Domestic = "http://production.shippingapis.com/ShippingAPI.dll?API=RateV4&XML=" + query;
            string URL_International = "http://production.shippingapis.com/ShippingAPI.dll?API=IntlRateV2&XML=" + query;
            string URL_ToUse = "";
            if (isDomestic == true)
            {
                URL_ToUse = URL_Domestic;
            }
            else
            {
                URL_ToUse = URL_International;
            }

            string USPS_Response = GetHTTPData(URL_ToUse);
            return USPS_Response;
        }


        private static string GetHTTPData(string Url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "GET";


            WebResponse response = request.GetResponse();
            Stream stream = response.GetResponseStream();
            StreamReader reader = new StreamReader(stream);

            string result = reader.ReadToEnd();
            //Console.WriteLine("Holy shit... we got data!");
            stream.Dispose();
            reader.Dispose();

            //Console.WriteLine("USPS SENT US THIS IN RETURN:\r\n");
            //Console.WriteLine(result);
            //Console.WriteLine("===============================================================");
            return result;

        }


        public static string UPS_processBoxQuery()
        {
            //USES HTTP POST
            string QueryDefault = "<?xml version=\"1.0\"?>\n";
            QueryDefault += "<AccessRequest xml:lang=\"en-US\">\n";
            QueryDefault += " <AccessLicenseNumber>" + ups_xml_key + "</AccessLicenseNumber>\n";
            QueryDefault += " <UserId>" + ups_username + "</UserId>\n";
            QueryDefault += " <Password>" + ups_password + "</Password>\n";
            QueryDefault += "</AccessRequest>\n";
            QueryDefault += "<?xml version=\"1.0\"?>\n";
            QueryDefault += "<RatingServiceSelectionRequest xml:lang=\"en-US\">\n";
            QueryDefault += " <Request>\n";
            QueryDefault += "  <TransactionReference>\n";
            QueryDefault += "   <CustomerContext>Rating and Service</CustomerContext>\n";
            QueryDefault += "   <XpciVersion>1.0001</XpciVersion>\n";
            QueryDefault += "  </TransactionReference>\n";
            QueryDefault += "  <RequestAction>Rate</RequestAction>\n";
            QueryDefault += "  <RequestOption>Shop</RequestOption>\n";
            QueryDefault += " </Request>\n";
            QueryDefault += " <PickupType>\n";
            QueryDefault += "  <Code>01</Code>\n";
            QueryDefault += " </PickupType>\n";
            QueryDefault += " <CustomerClassification>\n";
            QueryDefault += "  <Code>01</Code>\n";
            QueryDefault += " </CustomerClassification>\n";
            QueryDefault += " <Shipment>\n";
            QueryDefault += "  <Shipper>\n";
            QueryDefault += "   <ShipperNumber>" + ups_account_number + "</ShipperNumber>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + ship_from_state + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + ship_from_zipcode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + ship_from_country + "</CountryCode>\n";
            QueryDefault += "   </Address>\n";
            QueryDefault += "  </Shipper>\n";
            QueryDefault += "  <RateInformation>\n";
            QueryDefault += "   <NegotiatedRatesIndicator />\n";
            QueryDefault += "  </RateInformation>\n";
            QueryDefault += "  <ShipFrom>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + ship_from_state + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + ship_from_zipcode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + ship_from_country + "</CountryCode>\n";
            QueryDefault += "   </Address>\n";
            QueryDefault += "  </ShipFrom>\n";
            QueryDefault += "  <ShipTo>\n";
            QueryDefault += "   <Address>\n";
            QueryDefault += "    <StateProvinceCode>" + state + "</StateProvinceCode>\n";
            QueryDefault += "    <PostalCode>" + zipCode + "</PostalCode>\n";
            QueryDefault += "    <CountryCode>" + country + "</CountryCode>\n";

            //if (residential == "1")
            //{
                QueryDefault += "    <ResidentialAddressIndicator />\n";
            //}

            QueryDefault += "   </Address>\n";
            QueryDefault += "  </ShipTo>\n";

            for (int bCount = 0; bCount < boxheightList.Count; bCount++)
            {
                QueryDefault += "  <Package>\n";
                QueryDefault += "   <PackagingType>\n";
                QueryDefault += "    <Code>02</Code>\n";
                QueryDefault += "   </PackagingType>\n";
                QueryDefault += "   <Dimensions>\n";
                QueryDefault += "    <Length>" + boxlengthList[bCount] + "</Length>\n";
                QueryDefault += "    <Width>" + boxwidthList[bCount] + "</Width>\n";
                QueryDefault += "    <Height>" + boxheightList[bCount] + "</Height>\n";
                QueryDefault += "   </Dimensions>\n";
                QueryDefault += "   <PackageWeight>\n";
                QueryDefault += "    <UnitOfMeasurement>\n";
                QueryDefault += "     <Code>LBS</Code>\n";
                QueryDefault += "    </UnitOfMeasurement>\n";
                QueryDefault += "    <Weight>" + boxweightList[bCount] + "</Weight>\n";
                QueryDefault += "   </PackageWeight>\n";
                QueryDefault += "  </Package>\n";

            }

            QueryDefault += " </Shipment>\n";
            QueryDefault += "</RatingServiceSelectionRequest>\n";



            string UPS_Response = SendUPSQuery(QueryDefault);
            UPS_Response = UPS_Response.Replace("\n", "\\n");
            QueryDefault = QueryDefault.Replace("\n", "\\n");
                       
            string Parsed_Response = ParseUPSResults(UPS_Response);
           
             // Parsed_Response;
            return "\"carriers\":{\"UPS\":{\"xml_sent\":\"" + QueryDefault + "\"," + Parsed_Response;

            //return "";

        }

        public static string USPS_processBoxQuery()
        {
            //USES HTTP GET
            string returnDataHeader = "\"USPS\":{\"xml_sent\":\"";
            string returnData = "<RateV4Request USERID=\"" + usps_userid + "\" PASSWORD=\"" + usps_password + "\"> ";
            returnData += " <Revision>2</Revision>";

            for (int i = 0; i < boxweightList.Count; i++)
            {
                decimal smallestSide = 0;
                decimal secondSmallestSide = 0;

                smallestSide = Math.Min(Math.Min(decimal.Parse(boxheightList[i]),decimal.Parse(boxlengthList[i])),decimal.Parse(boxwidthList[i]));

                if (decimal.Parse(boxwidthList[i]) == smallestSide)
                {
                    secondSmallestSide = Math.Min(decimal.Parse(boxlengthList[i]),decimal.Parse(boxheightList[i]));
                }
                if (decimal.Parse(boxheightList[i]) == smallestSide)
                {
                    secondSmallestSide = Math.Min(decimal.Parse(boxwidthList[i]),decimal.Parse(boxlengthList[i]));
                }
                if (decimal.Parse(boxlengthList[i]) == smallestSide)
                {
                    secondSmallestSide = Math.Min(decimal.Parse(boxheightList[i]),decimal.Parse(boxwidthList[i]));
                }


                decimal currentWeight = Math.Floor(decimal.Parse(boxweightList[i]));
                decimal weight_decimal = (decimal.Parse(boxweightList[i]) - currentWeight) * 16;
                returnData += " <Package ID=\"" + i + "\">";
                returnData += "  <Service>All</Service>";
                returnData += "  <ZipOrigination>" + ship_from_zipcode + "</ZipOrigination>";
                returnData += "  <ZipDestination>" + zipCode + "</ZipDestination>";
                returnData += "  <Pounds>" + currentWeight.ToString() + "</Pounds>";
                returnData += "  <Ounces>" + weight_decimal.ToString("0.##") + "</Ounces>";
                returnData += "  <Container>RECTANGULAR</Container>";
                returnData += "  <Size>LARGE</Size>";
                returnData += "  <Width>" + boxwidthList[i] + "</Width>";
                returnData += "  <Length>" + boxlengthList[i] + "</Length>";
                returnData += "  <Height>" + boxheightList[i] + "</Height>";
                returnData += "  <Girth>" + ((secondSmallestSide * 2) + (smallestSide * 2)).ToString("0.##") + "</Girth>";
                returnData += "  <Machinable>True</Machinable>";
                returnData += " </Package>";



            }

            returnData += "</RateV4Request>";

            string USPS_Response = SendUSPSQuery(returnData);



            string ResponseHeader = "\",\"state_desc\":\"";
            string responseDetail = "";
            if (USPS_Response.Contains("<RateV4Response>") == true)
            {
                responseDetail = "success";

            }
            else
            {
                responseDetail = "failure";
            }
            string JSON_Addition = "";



            
            USPS_Response = USPS_Response.Replace("\n", "\\n");

            ResponseHeader += responseDetail + "\",\"results\":{";

            //now to parse the results! fun is...
            List<string> USPS_Parsings = new List<string>();
            string[] uspsParsings = USPS_Response.Split(new string[] { "<Package " }, StringSplitOptions.None);

            //Console.WriteLine("USPS Results: ");
            foreach (string str in uspsParsings)
            {
                if (str.Substring(0, 2) == "ID")
                {
                    USPS_Parsings.Add(str);
                    //Console.WriteLine(str);
                }


            }
            //Console.WriteLine("===========================================");
            string parsedData = "";

   
            // NOTE: First look through all packages and find all common shipping methods.
            //       Next, take all pricings for that shipping method and add em up.
            //       Finally, compute the TP_rate and the other rate (either list or negotiated; i forget)
            // "3-Priority Mail Express 2-Day":{"tp_rate":0.00,"negotiated_rate":0.00,"list_rate":0.00,"service":"Priority Mail Express 2-Day"},
            // Also seems as tho list and negotiated share the same value

            //USPS_Parsings

            List<string> USPS_ShipInfo = new List<string>();
            List<string> USPS_AllShippingInfo = new List<string>();
            int TotPackages = USPS_Parsings.Count;


            foreach (string pkg in USPS_Parsings)
            {
                //string PostageData = pkg.Split('\"')[1] + "|" + pkg.Split(new string[] { "<Postage CLASSID=" }, StringSplitOptions.None)[1].Split(new string[] { "</Postage>" },StringSplitOptions.None)[0];

                foreach (string postInfo in pkg.Split(new string[] { "<Postage CLASSID=" }, StringSplitOptions.None))
                {
                    if (postInfo[0] == '\"')
                    {
                        //Console.WriteLine("Individual Postage: Package #" + pkg.Split('\"')[1] + " - " + postInfo);
                        USPS_AllShippingInfo.Add(pkg.Split('\"')[1] + "|" + postInfo);
                    }
                }
            }

            //JSON Time... :(


            foreach (string pkg in USPS_AllShippingInfo)
            {
                JSON_Addition += "\"" + pkg.Split('\"')[1] + "-";
                string tmpStr = pkg.Split(new string[] { "<MailService>" }, StringSplitOptions.None)[1].Split(new string[] { "</MailService>" }, StringSplitOptions.None)[0];
                if (tmpStr.Contains("&") == true)
                {
                    tmpStr=tmpStr.Split('&')[0];
                }

                JSON_Addition += tmpStr;
                JSON_Addition += "\":{\"tp_rate\":" + ((decimal.Parse(pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0]) * (decimal)1.2)).ToString("0.##");
                JSON_Addition += ",\"negotiated_rate\":" + pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0];
                JSON_Addition += ",\"list_rate\":" + pkg.Split(new string[] { "<Rate>" }, StringSplitOptions.None)[1].Split(new string[] { "</Rate>" }, StringSplitOptions.None)[0];
                JSON_Addition += ",\"service\":\"" + tmpStr;
                JSON_Addition += "\"},";

            }

            JSON_Addition = JSON_Addition.Substring(0, JSON_Addition.Length - 1);

            JSON_Addition += "},";

            //NOW TO ADD THE XML RAW DATA WE GOT FROM USPS...

            JSON_Addition += "\"xml_received\":\"";
            JSON_Addition += USPS_Response;
            JSON_Addition += "\",";
            JSON_Addition += "\"state\":" + "1"; //WHERE IS THIS 'STATE' Value coming from? ASSUMING 0 or 1 depending on error.
            JSON_Addition += "}}}";











            return returnDataHeader + returnData + ResponseHeader + JSON_Addition;
        }

        public static string CreateBoxesJSON()
        {
            string JSON_Boxes = "{\"packages\":[";

            for (int i = 0; i < boxheightList.Count; i++)
            {
                JSON_Boxes += "{\"cube\":" + (decimal.Parse(boxheightList[i]) * decimal.Parse(boxlengthList[i]) * decimal.Parse(boxwidthList[i])).ToString();
                JSON_Boxes += ",\"width\":" + boxwidthList[i] + ",";
                JSON_Boxes += "\"dimensional_weight\":" + boxweightList[i] + ",";
                JSON_Boxes += "\"contents\":{},\"height\":" + boxheightList[i] + ",";
                JSON_Boxes += "\"length\":" + boxlengthList[i] + ",";
                

                decimal SmallestSide = 0;
                decimal SecondSmallestSide = 0;

                SmallestSide = Math.Min(Math.Min(decimal.Parse(boxheightList[i]), decimal.Parse(boxlengthList[i])), decimal.Parse(boxwidthList[i]));

                if (decimal.Parse(boxheightList[i]) == SmallestSide)
                {
                    SecondSmallestSide = Math.Min(decimal.Parse(boxlengthList[i]),decimal.Parse(boxwidthList[i]));
                }
                if (decimal.Parse(boxlengthList[i]) == SmallestSide)
                {
                    SecondSmallestSide = Math.Min(decimal.Parse(boxheightList[i]),decimal.Parse(boxwidthList[i]));

                }
                if (decimal.Parse(boxwidthList[i]) == SmallestSide)
                {
                    SecondSmallestSide = Math.Min(decimal.Parse(boxheightList[i]),decimal.Parse(boxlengthList[i]));
                }

                SmallestSide = SmallestSide * 2;
                SecondSmallestSide = SecondSmallestSide * 2;

                string Girth = (SmallestSide + SecondSmallestSide).ToString("0.##");
                JSON_Boxes += "\"girth\":" + Girth + ",";
                JSON_Boxes += "\"weight\":" + boxweightList[i] + "},";

            }
            
            JSON_Boxes = JSON_Boxes.Substring(0, JSON_Boxes.Length - 1) + "],";

            return JSON_Boxes;

        }

    }
}
