using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace PhoneValidation
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.GetLength(0) > 0)
            {
                string phoneNumber = args[0];
                Console.WriteLine("Validating Phone Number: " + phoneNumber +  ". Please Wait...");

                //string xmlOut = "<?xml version=\"1.0\" encoding=\"ISO8859-1\"?>\r\n";
                //xmlOut += " <number>" + phoneNumber + "</number>\r\n";

                //var postData = new List<KeyValuePair<string, string>>();
                //postData.Add(new KeyValuePair<string, string>("PhoneNumber", phoneNumber));
                //postData.Add(new KeyValuePair<string,string>("APIKey","pv-cda1617baaa4f7b8e0d4c85197ccc191");
                
                string pn = "PhoneNumber=" + phoneNumber;
                string api = "APIKey=" + "pv-cda1617baaa4f7b8e0d4c85197ccc191";

                //List<string> content = new List<string>();
                //content.Add(pn);
                //content.Add(api);
                string Getout = pn + "&" + api;


                string results = Connectivity.HTTPCalls.HTTP_POST("http://api.phone-validator.net/api/v2/verify",Getout);
                string Status, LineType, Geolocation, RegionCode, FormatNational, FormatInternational = "";

                if (results.ToUpper().Contains("FAILED") == false)
                {

                    Status = results.Split(new string[]{"status\":\""},StringSplitOptions.None)[1].Split('"')[0];
                    LineType = results.Split(new string[] { "linetype\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    Geolocation = results.Split(new string[] { "geolocation\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    RegionCode = results.Split(new string[] { "regioncode\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    FormatNational = results.Split(new string[] { "formatnational\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    FormatInternational = results.Split(new string[] { "formatinternational\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    Console.WriteLine("Status: " + Status);
                    Console.WriteLine("Line Type: " + LineType);
                    Console.WriteLine("Geolocation: " + Geolocation);
                    Console.WriteLine("Region Code: " + RegionCode);
                    Console.WriteLine("Format National: " + FormatNational);
                    Console.WriteLine("Format International: " + FormatInternational);
                    Console.ReadKey();

                }

                return 0;
            }
            return 0;
        }
    }
}
