using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace arbitrarySQL._3X3
{
    class Program
    {
        public const string VERSION = "0.0.1";
        public static bool isScript = false;

        public static void GETVERSION()
        {
            
            Console.WriteLine("Version " + VERSION);
            Console.WriteLine("  Arbitrary SQL Script");
            Console.WriteLine("   Last Modified 8-26-2014");
            Console.WriteLine("   This script takes in an SQL query aimed for MAS, sends it out");
            Console.WriteLine("   to MAS and then gets the return result. Although not-so-arbitrary");
            Console.WriteLine("   anymore, this script gets the results we need in a timely fashion.");

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
                //Houston, we have an SQL query... let's process it!
                
               // Console.WriteLine(Query);

                List<string> Results = SQLCalls.sqlQuery(Query, 2,true);
                
                foreach (string result in Results)
                {
                    Console.WriteLine(result);
                }
            }



        }

        static void Main(string[] args)
        {
            string queryString = GETQUERY(args);
            PROCESSQUERY(queryString);
            if (!isScript)
            {
                Console.ReadKey();
            }
        }
    }
}
