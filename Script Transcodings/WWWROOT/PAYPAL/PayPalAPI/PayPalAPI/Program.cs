using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace PayPalAPI
{
    class Program
    {
        public const string VERSION = "0.0.0";
        public const string LAST_BUILD_DATE = "9/02/2014 - 2:10pm";
        public static bool isScript = false;


        //all of these must be changed as last argument may not have a '&'... figure it out...lol
        public struct OperationSearchAll
        {
            //DateTime gets received in bits...
            public string sy,ey;
            public string sm,em;
            public string sd,ed;
            public string sh,eh;
            public string sn,en;
            public string ss,es;


            public void QueryToVar(string queryString)
            {
                try
                {
                    sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                    ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                    ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                    em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                    ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                    eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                    en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                    es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
                }
                catch
                {
                    //invalid inputs... laterz
                    Environment.Exit(0);
                }
                

            }
            public string buildQuery()
            {
                string Query = "type=search&subtype=all&sy=" + sy + "&sm=" + sm + "&sd=" + sd + "&sh=" + sh + "&sn=" + sn + "&ss=" + ss;
                return Query;
            }

            public DateTime returnStartDateTime()
            {
                try
                {
                    string dTime = "";
                    DateTime newDateTime = new DateTime();
                    dTime += sm + "/" + sd + "/" + sy + " " + sh + ":" + sn + ":" + ss;
                    newDateTime = DateTime.Parse(dTime);
                    return newDateTime;
                }
                catch
                {
                    //command is corrupt, exit after an error...
                    Environment.Exit(0);
                }
                //this line never gets hit...
                return new DateTime();
            }
            public DateTime returnEndDateTime()
            {
                try
                {
                    string dTime = "";
                    DateTime newDateTime = new DateTime();
                    dTime += em + "/" + ed + "/" + ey + " " + eh + ":" + en + ":" + es;
                    newDateTime = DateTime.Parse(dTime);
                    return newDateTime;
                }
                catch
                {
                    //command is corrupt, exit after an error...
                    Environment.Exit(0);
                }
                //this line never gets hit...
                return new DateTime();
            }

        }
        public struct OperationSearchCaptures
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;
            public List<string> KVPargs;

            public void QueryToVar(string queryString)
            {
                try
                {
                    sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                    sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                    ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                    ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                    em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                    ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                    eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                    en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                    es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];

                    string kvp_section = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1];
                    while (kvp_section.StartsWith("&") == false)
                    {
                        kvp_section = kvp_section.Substring(1, kvp_section.Length);
                    }
                    kvp_section = kvp_section.Substring(1, kvp_section.Length);

                    string[] KVPs = kvp_section.Split('&');
                    foreach (string kvps in KVPs)
                    {
                        KVPargs.Add(kvps);
                    }

                }
                catch
                {
                    //invalid inputs... laterz
                    Environment.Exit(0);
                }
            }


        }
        public struct OperationSearchRefunds
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;

            public void QueryToVar(string queryString)
            {
                sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
            }
        }
        public struct OperationSearchPayments
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;

            public void QueryToVar(string queryString)
            {
                sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
            }
        }
        public struct OperationSearchHolds
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;

            public void QueryToVar(string queryString)
            {
                sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
            }
        }
        public struct OperationCapture
        {
            public string SalesOrderNo;
            public string AuthorizationID;
            public string CaptureAmount;

            public void QueryToVar(string queryString)
            {
                SalesOrderNo = queryString.Split(new string[] { "salesorderno=" }, StringSplitOptions.None)[1].Split('&')[0];
                AuthorizationID = queryString.Split(new string[] { "authid=" }, StringSplitOptions.None)[1].Split('&')[0];
                CaptureAmount = queryString.Split(new string[] { "amount=" }, StringSplitOptions.None)[1].Split('&')[0];
            }

        }
        public struct OperationVoid
        {
            public string AuthorizationID;

            public void QueryToVar(string queryString)
            {
                AuthorizationID = queryString.Split(new string[] {"authid="},StringSplitOptions.None)[1];
            }

        }
        public struct OperationDetail
        {
            public string TransactionID;

            public void QueryToVar(string queryString)
            {
                TransactionID = queryString.Split(new string[] { "transid=" }, StringSplitOptions.None)[1];
            }
        }
        public struct OperationRefund
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;

            public void QueryToVar(string queryString)
            {
                sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
            }
        }
        public struct OperationSearchBankDeposits
        {
            public string sy, ey;
            public string sm, em;
            public string sd, ed;
            public string sh, eh;
            public string sn, en;
            public string ss, es;

            public void QueryToVar(string queryString)
            {
                sy = queryString.Split(new string[] { "sy=" }, StringSplitOptions.None)[1].Split('&')[0];
                sm = queryString.Split(new string[] { "sm=" }, StringSplitOptions.None)[1].Split('&')[0];
                sd = queryString.Split(new string[] { "sd=" }, StringSplitOptions.None)[1].Split('&')[0];
                sh = queryString.Split(new string[] { "sh=" }, StringSplitOptions.None)[1].Split('&')[0];
                sn = queryString.Split(new string[] { "sn=" }, StringSplitOptions.None)[1].Split('&')[0];
                ss = queryString.Split(new string[] { "ss=" }, StringSplitOptions.None)[1].Split('&')[0];

                try
                {
                    ey = queryString.Split(new string[] { "ey=" }, StringSplitOptions.None)[1].Split('&')[0];
                    em = queryString.Split(new string[] { "em=" }, StringSplitOptions.None)[1].Split('&')[0];
                    ed = queryString.Split(new string[] { "ed=" }, StringSplitOptions.None)[1].Split('&')[0];
                    eh = queryString.Split(new string[] { "eh=" }, StringSplitOptions.None)[1].Split('&')[0];
                    en = queryString.Split(new string[] { "en=" }, StringSplitOptions.None)[1].Split('&')[0];
                    es = queryString.Split(new string[] { "es=" }, StringSplitOptions.None)[1].Split('&')[0];
                }
                catch
                {
                }

            }
        }


        public static void GETVERSION()
        {

            string VersionData = "Version " + VERSION + "\r\n" +
                "PayPal API Script\r\n" +
                "   Last Modified " + LAST_BUILD_DATE + "\r\n" +
                "";


            if (!isScript)
            {
                Console.WriteLine(VersionData);
                Console.WriteLine("Press any key to exit or press 'S' to save to 'c:\\emailproxy.txt'");
                ConsoleKeyInfo cki = Console.ReadKey();
                if (cki.Key == ConsoleKey.S)
                {
                    try
                    {
                        File.WriteAllText("c:\\emailproxy.txt", VersionData);
                    }
                    catch
                    {
                        //doesnt seem to catch... hmm..
                        Console.WriteLine("Failed to write to c:\\emailproxy.txt. You may not\r\nhave sufficient priviledges. Try Again. Or Not.\r\n\r\n");
                    }
                }

            }
            else
            {
                Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n");
                Console.WriteLine(VersionData);
            }
            Environment.Exit(0);
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

        private static void ThrowHTTPException(string ExceptionDesc)
        {
            Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\n\n" + ExceptionDesc);
        }

        private static void PROCESSQUERY(string Query)
        {
            if (Query.ToUpper().Contains("GETVERSION()") == true)
            {
                GETVERSION();
                return;
            }

            if (Query.ToUpper().Contains("TYPE=") == false)
            {
                ThrowHTTPException("Bad Request: No 'Type' descriptor found.");
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH") == true)
            {
                try
                {
                    int StartYears = int.Parse(Query.ToUpper().Split(new string[] { "SY="},StringSplitOptions.None)[1].Split('&')[0]);
                    int StartMonths = int.Parse(Query.ToUpper().Split(new string[] { "SM=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    int StartDays = int.Parse(Query.ToUpper().Split(new string[] { "SD=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    int StartHours = int.Parse(Query.ToUpper().Split(new string[] { "SH=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    int StartMinutes = int.Parse(Query.ToUpper().Split(new string[] { "SN=" }, StringSplitOptions.None)[1].Split('&')[0]);
                    int StartSeconds = int.Parse(Query.ToUpper().Split(new string[] { "SS=" }, StringSplitOptions.None)[1].Split('&')[0]);

                    if (StartYears.ToString().Length != 4)
                    {
                        ThrowHTTPException("Bad Request: Starting Years must contain 4 digits.");
                    }
                    if (StartMonths > 12 & StartMonths < 1)
                    {
                        ThrowHTTPException("Bad Request: Starting Months must be a number between 1 and 12.");
                    }
                    if (StartDays > 31 | StartDays < 1)
                    {
                        ThrowHTTPException("Bad Request: Starting Days must be a number between 1 and 31.");
                    }
                    if (StartHours > 24 | StartHours < 0)
                    {
                        ThrowHTTPException("Bad Request: Starting Hours must be a number between 0 and 23.");
                    }
                    if (StartMinutes > 59 | StartMinutes < 0)
                    {
                        ThrowHTTPException("Bad Request: Starting Minutes must be a number between 0 and 59.");
                    }
                    if (StartSeconds > 59 | StartSeconds < 0)
                    {
                        ThrowHTTPException("Bad Request: Starting Seconds must be a number between 0 and 59.");
                    }

                }
                catch (Exception wtf)
                {
                    ThrowHTTPException("Bad Request: Error: " + wtf.Message);
                }
            }

            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=ALL") == true)
            {

                ProcessSearchAll(Query);
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=CAPTURES") == true)
            {
                ProcessSearchCaptures(Query);
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=REFUNDS") == true)
            {
                ProcessSearchRefunds(Query);
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=PAYMENTS") == true)
            {
                ProcessSearchPayments(Query);
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=HOLDS") == true)
            {
                ProcessSearchHolds(Query);
            }
            if (Query.ToUpper().Contains("TYPE=CAPTURE&") == true)
            {
                ProcessCapture(Query);
            }
            if (Query.ToUpper().Contains("TYPE=VOID") == true)
            {
                ProcessVoid(Query);
            }
            if (Query.ToUpper().Contains("TYPE=REFUND") == true)
            {
                ProcessRefund(Query);

            }
            if (Query.ToUpper().Contains("TYPE=DETAIL") == true)
            {
                ProcessDetail(Query);
            }
            if (Query.ToUpper().Contains("TYPE=SEARCH&SUBTYPE=BANKDEPOSITS") == true)
            {
                ProcessBankDeposits(Query);
            }



        }

        static void Main(string[] args)
        {
            string QueryString = Environment.GetEnvironmentVariable("QUERY_STRING");
            PROCESSQUERY(QueryString);



        }

        static void ProcessSearchAll(string Query)
        {
            OperationSearchAll osa = new OperationSearchAll();
            osa.QueryToVar(Query);
            
        }
        static void ProcessSearchCaptures(string Query)
        {
            OperationSearchCaptures osc = new OperationSearchCaptures();
            osc.QueryToVar(Query);

        }
        static void ProcessSearchRefunds(string Query)
        {
            OperationSearchRefunds osr = new OperationSearchRefunds();
            osr.QueryToVar(Query);

        }
        static void ProcessSearchPayments(string Query)
        {
            OperationSearchPayments osp = new OperationSearchPayments();
            osp.QueryToVar(Query);
        }
        static void ProcessSearchHolds(string Query)
        {
            OperationSearchHolds osh = new OperationSearchHolds();
            osh.QueryToVar(Query);
        }
        static void ProcessCapture(string Query)
        {
            OperationCapture oc = new OperationCapture();
            oc.QueryToVar(Query);
        }
        static void ProcessVoid(string Query)
        {
            OperationVoid ov = new OperationVoid();
            ov.QueryToVar(Query);
        }
        static void ProcessRefund(string Query)
        {
            OperationRefund or = new OperationRefund();
            or.QueryToVar(Query);
        }
        static void ProcessDetail(string Query)
        {
            OperationDetail od = new OperationDetail();
            od.QueryToVar(Query);
        }
        static void ProcessBankDeposits(string Query)
        {
            OperationSearchBankDeposits osbd = new OperationSearchBankDeposits();
            osbd.QueryToVar(Query);
        }



        
    }
}
