using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace test
{
    class Program
    {
        public const string Version = "1.0.0";


        static bool GETVERSION()
        {
            if (System.Environment.GetEnvironmentVariable("QUERY_STRING").ToUpper().Contains("GETVERSION()") == true)
            {
                Console.WriteLine("Content-Type: text/html\n\n");
                Console.WriteLine("Test Version: " + Version);
                return true;
            }
            else
            {
                return false;
            }
        }

        static void Main(string[] args)
        {
            if (GETVERSION() == true)
            {
                return;
            }
            Console.WriteLine("Content-type: text/html\n\n");
            Console.WriteLine("Hello World.\n");
        }
    }
}
