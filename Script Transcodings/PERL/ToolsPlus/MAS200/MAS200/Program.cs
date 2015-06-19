using System;
using System.Collections.Generic;
using System.Text;

namespace MAS200
{
    class Program
    {
        public const bool useCommandArgs = true;
        public const bool useQueryArgs = true;

        static void Main(string[] args)
        {
            Console.WriteLine("Testing MAS200 Connection.. Please Wait..");
            List<string> Arguments = CommonCalls.GetRuntimeArguments(useCommandArgs, useQueryArgs, args, System.Environment.GetEnvironmentVariable("QUERY_STRING"));
            List<string> QueryResults = SQLCalls.sqlQuery("SELECT * FROM IM_ItemWarehouse WHERE ItemCode='AMA51534'", 2);

            Console.WriteLine("Results: " + QueryResults.Count.ToString());
            //string results = "";
            foreach (string res in QueryResults)
            {
                Console.WriteLine("  " + res);
                //results += res + "\r\n";

            }


            //System.IO.File.WriteAllText("C:\\users\\tomwestbrook\\desktop\\columns.txt", results);

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();




        }




    }
}
