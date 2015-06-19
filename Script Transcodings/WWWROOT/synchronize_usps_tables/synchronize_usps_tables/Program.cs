using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.IO;


namespace synchronize_usps_tables
{
    class Program
    {
        public const string Version = "1.0.0";
        public static string JSON_Output = "";


        static bool CheckForVersion()
        {
            if (System.Environment.GetEnvironmentVariable("QUERY_STRING").ToUpper().Contains("GETVERSION()"))
            {
                Console.WriteLine("Status: 200 OK\nContent-Type: text/html\n\n");
                Console.WriteLine("Synchronize_USPS_Tables Version: " + Version);
                return true;
            }
            else
            {
                return false;
            }
        }
        
        
        
        static void Main(string[] args)
        {
            List<string> USPSZoneChart_ZoneNo = new List<string>();
            List<string> USPSZoneChart_ZipCodePrefix = new List<string>();
            List<string> USPSBoxes_Length = new List<string>();
            List<string> USPSBoxes_ID = new List<string>();
            List<string> USPSBoxes_HandledByApi = new List<string>();
            List<string> USPSBoxes_Height = new List<string>();
            List<string> USPSBoxes_MaximumWeight = new List<string>();
            List<string> USPSBoxes_Width = new List<string>();
            List<string> USPSBoxes_Name = new List<string>();
            List<string> USPSZoneRates_ZoneID = new List<string>();
            List<string> USPSZoneRates_BoxID = new List<string>();
            List<string> USPSZoneRates_Cost = new List<string>();

            if (CheckForVersion() == true)
            {
                return;
            }


            string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";
            string queryString = "SELECT * FROM USPSZoneChart";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(queryString,connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            USPSZoneChart_ZipCodePrefix.Add(reader.GetValue(0).ToString());
                            USPSZoneChart_ZoneNo.Add(reader.GetValue(1).ToString());
                            
                        }
                        reader.Close();
                    }
                }
                connection.Close();
            }



            queryString = "SELECT * FROM USPSBoxes";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(queryString,connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            USPSBoxes_ID.Add(reader.GetValue(0).ToString());
                            USPSBoxes_Name.Add(reader.GetValue(1).ToString());
                            USPSBoxes_Length.Add(reader.GetValue(2).ToString());
                            USPSBoxes_Width.Add(reader.GetValue(3).ToString());
                            USPSBoxes_Height.Add(reader.GetValue(4).ToString());
                            USPSBoxes_MaximumWeight.Add(reader.GetValue(5).ToString());
                            USPSBoxes_HandledByApi.Add(reader.GetValue(6).ToString());

                        }
                        reader.Close();
                    }
                }
                connection.Close();
            }


            queryString = "SELECT * FROM USPSZoneRates";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(queryString,connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            USPSZoneRates_BoxID.Add(reader.GetValue(0).ToString());
                            USPSZoneRates_ZoneID.Add(reader.GetValue(1).ToString());
                            USPSZoneRates_Cost.Add(reader.GetValue(2).ToString());
                        }
                        reader.Close();
                    }
                }
                connection.Close();
            }


            //start uspszonechart sector
            JSON_Output = "{\"USPSZoneChart\":[";
            for (int offset = 0; offset < USPSZoneChart_ZoneNo.Count; offset++)
            {
                JSON_Output += "{\"ZoneNo\":\"" + USPSZoneChart_ZoneNo[offset] + "\",\"ZipCodePrefix\":\"" + USPSZoneChart_ZipCodePrefix[offset] + "\"}";
                if (offset < USPSZoneChart_ZoneNo.Count - 1)
                {
                    JSON_Output += ",";
                }
            }
            JSON_Output += "],";


            //now lets add the uspsboxes sector
            JSON_Output += "\"USPSBoxes\":[";
            for (int offset = 0; offset < USPSBoxes_Length.Count; offset++)
            {
                JSON_Output += "{\"Length\":\"" + USPSBoxes_Length[offset] + "\",";
                JSON_Output += "\"ID\":\"" + USPSBoxes_ID[offset] + "\",";
                if (USPSBoxes_HandledByApi[offset] == "True")
                {
                    JSON_Output += "\"HandledByAPI\":\"" + "1" + "\",";
                }
                else
                {
                    JSON_Output += "\"HandledByAPI\":\"" + "0" + "\",";
                }
                
                
                JSON_Output += "\"Height\":\"" + USPSBoxes_Height[offset] + "\",";
                JSON_Output += "\"MaximumWeight\":\"" + USPSBoxes_MaximumWeight[offset] + "\",";
                JSON_Output += "\"Width\":\"" + USPSBoxes_Width[offset] + "\",";
                JSON_Output += "\"Name\":\"" + USPSBoxes_Name[offset] + "\"}";
                if (offset < USPSBoxes_Length.Count - 1)
                {
                    JSON_Output += ",";
                }

            }
            JSON_Output += "],";


            //finally, lets add in the uspszonerates
            JSON_Output += "\"USPSZoneRates\":[";
            for (int offset = 0; offset < USPSZoneRates_ZoneID.Count; offset++)
            {
                JSON_Output += "{\"ZoneID\":\"" + USPSZoneRates_ZoneID[offset] + "\",";
                JSON_Output += "\"BoxID\":\"" + USPSZoneRates_BoxID[offset] + "\",";
                JSON_Output += "\"Cost\":\"" + USPSZoneRates_Cost[offset] + "\"}";
                if (offset < USPSZoneRates_ZoneID.Count - 1)
                {
                    JSON_Output += ",";
                }
            }
            JSON_Output += "]}";


            //File.WriteAllText("C:\\Users\\tomwestbrook\\Documents\\JSON-output.txt",JSON_Output);
            //Console.ReadKey();

            Console.WriteLine("Status: 200 OK\nContent-type: application/json\n");
            Console.WriteLine(JSON_Output);


        }
    }
}
