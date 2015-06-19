using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;

namespace barcodedb
{
    class Program
    {

        private const string Version = "1.0.0";

        static bool GETVERSION()
        {
            if (System.Environment.GetEnvironmentVariable("QUERY_STRING").ToUpper().Contains("GETVERSION()") == true)
            {
                Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n");
                Console.WriteLine("BarcodeDB Version: " + Version);
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

            string queryString = System.Environment.GetEnvironmentVariable("QUERY_STRING");

            if (queryString.ToUpper().Contains("ACTION=") == true)
            {

                if (queryString.ToUpper().Contains("ACTION=LOOKUP") == true)
                {
                    if (queryString.ToUpper().Contains("PARAM="))
                    {



                        //if action=lookup&param= is in the query
                        string param = queryString.Substring(queryString.ToUpper().IndexOf("PARAM=") + 6);
                        if (param.Contains("&") == true)
                        {
                            param = param.Split('&')[0];
                        }



                        List<string> queryResult = GetSQLQuery("SELECT ItemNumber, Barcode FROM vBarcodesComponents WHERE Quantity=1 AND ItemNumber IS NOT NULL AND Barcode='" + param + "'");

                        Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n");
                        //Console.WriteLine("Content-Type: text/plain\n\n");
                        Console.WriteLine(queryResult[0]);
                        return;
                    }
                    else
                    {
                        Throw400Error();
                    }
                }


                if (queryString.ToUpper().Contains("ACTION=FULL") == true)
                {
                    //if action=full in query
                    List<string> queryResult = GetSQLQuery("SELECT ItemNumber, Barcode FROM vBarcodesComponents WHERE Quantity=1 AND ItemNumber IS NOT NULL");
                    Console.WriteLine("Status: 200 OK\nContent-type: text/plain\n\n");
                    //Console.WriteLine("Content-type: text/plain\n\n");
                    foreach (string result in queryResult)
                    {
                        Console.WriteLine(result);
                    }
                    return;
                }

                Throw400Error();


            }
            else
            {
                Throw400Error();
                return;
            }


        }


        static List<string> GetSQLQuery(string sqlQuery)
        {
            //Console.WriteLine("HTTP/1.1 200 OK\nContent-Type: text/plain\n\n\n" + sqlQuery);
            //return "";
            List<string> results = new List<string>();

            string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";
            //string results = "";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            results.Add(reader.GetValue(0).ToString() + "," + reader.GetValue(1).ToString());
                        }
                        reader.Close();
                    }
                }
                connection.Close();
            }

            return results;
        }

        static void Throw400Error()
        {
            Console.WriteLine("Status: 400 Bad Request\nContent-Type: text/plain\n\n\n");
        }

    }
}
