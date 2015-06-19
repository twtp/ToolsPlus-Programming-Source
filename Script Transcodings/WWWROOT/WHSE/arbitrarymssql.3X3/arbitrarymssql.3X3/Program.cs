using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace arbitrarymssql._3X3
{
    class Program
    {
        public const string VERSION = "0.0.1";
        public static bool isScript = false;

        public static void GETVERSION()
        {

            Console.WriteLine("Version " + VERSION);
            Console.WriteLine("  Arbitrary MS-SQL Script");
            Console.WriteLine("   Last Modified 8-27-2014");
            Console.WriteLine("   This script takes in an SQL Request for the MSSQL Database. If the");
            Console.WriteLine("   request is a SELECT statement, then allow it to proceed. Simple.");
           

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


        private static void PROCESSQUERY(string Query)
        {
            if (Query.ToUpper().Contains("GETVERSION()") == true)
            {
                GETVERSION();
                return;
            }

            if (Query.ToUpper().Contains("SELECT") == true & Query.ToUpper().Contains("FROM") == true)
            {
                //let's rock and roll... yea with JSON too...
                List<string> JSON_OUTPUT = SQLCalls.sqlQuery(Query,1,true);
                string JSON = JSON_OUTPUT[0];
                Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n" + JSON);
            }
            else
            {
                //nope no updates or inserts inside this script!
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\ninvalid sql\n");
            }





        }
        
        
        static void Main(string[] args)
        {
            string QueryString = GETQUERY(args);
            PROCESSQUERY(QueryString);

        }
    }
}
