using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.Data.Odbc;


namespace arbitrarymssql._3X3
{
    public static class SQLCalls
    {
        
        //last updated 8-18-2014 10:02am     
        
        public const string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";
        public const string masConnectionsString = "Driver={MAS 90 4.0 ODBC Driver};Directory=\\\\toolsplus04\\databases\\mas45server\\MAS90;Prefix=\\\\toolsplus04\\databases\\mas45server\\MAS90\\SY,\\\\toolsplus04\\databases\\mas45server\\MAS90\\==;ViewDLL=\\\\toolsplus04\\databases\\mas45server\\MAS90\\Home;LogFile=\\PVXODBC.LOG;CacheSize=4;DirtyReads=1;BurstMode=1;StripTrailingSpaces=1;SERVER=NotTheServer;UID=bd|TOO;PWD=brian";

        /// <summary>
        /// ConnectionType 1 = ToolsPlus DB / ConnectionType 2 = MAS200 DB
        /// </summary>
        /// <param name="queryString"></param>
        /// <param name="ConnectionType"></param>
        /// <returns></returns>
        public static List<string> sqlQuery(string queryString,int ConnectionType, bool JSON_ENCODE)
        {
            List<string> QueryResults = new List<string>();
            QueryResults.Capacity = 1000;
            if (ConnectionType == 1)
            {

                try
                {

                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        using (SqlCommand command = new SqlCommand(queryString, connection))
                        {
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    if (reader.GetValue(0) == null | reader.GetValue(0) == "")
                                    {
                                        reader.Close();
                                        connection.Close();
                                        //MessageBox.Show("Should never return null data from SQL Query...");
                                        return new List<string>();
                                    }
                                    string tmpResults = "";

                                    for (int x = 0; x < reader.FieldCount; x++)
                                    {
                                        tmpResults += reader.GetValue(x).ToString() + "|";
                                    }
                                    QueryResults.Add(tmpResults);
                                }


                                List<string> JSON_Line = new List<string>();
                                if (JSON_ENCODE == true)
                                {
                                    DataTable dt = reader.GetSchemaTable();



                                    List<string> ColumnNames = new List<string>();
                                    List<string> ColumnDataType = new List<string>();
                                    List<string> ColumnSize = new List<string>();




                                    foreach (DataRow row in dt.Rows)
                                    {
                                        ColumnNames.Add(row[0].ToString());
                                        ColumnDataType.Add(getDataTypeCode(row[5].ToString()).ToString());

                                        ColumnSize.Add(row[3].ToString());
                                    }

                                    /*foreach (DataColumn dc in dt.Columns)
                                    {
                                        masReader.GetFieldType
                                        ColumnNames.Add(dc.Caption);
                                        ColumnDataType.Add(dc.DataType.ToString());
                                        ColumnSize.Add(dc.MaxLength.ToString());
                                    }*/

                                    JSON_Line.Add(EncodeToJSON(ColumnNames, ColumnDataType, ColumnSize, QueryResults));

                                }



                                reader.Close();

                            }
                        }
                        connection.Close();
                    }
                }
                catch (Exception A)
                {
                    List<string> errorList = new List<string>();
                    errorList.Add(A.Message);
                    //MessageBox.Show(A.Message);
                    return errorList;
                }
                //MessageBox.Show("IM LEAVING HERE");
                return QueryResults;
            }

            if (ConnectionType == 2)
            {
                //dont know why i have to do this but...
               
                    OdbcConnection masConnection = new OdbcConnection(masConnectionsString);
                    OdbcCommand masCommand = new OdbcCommand(queryString, masConnection);
                
             
                try
                {
                    masConnection.Open();
                }
                catch (Exception B)
                {
                    Console.WriteLine("Caught Error ON OPEN: " + B.Message);
                }
                OdbcDataReader masReader = masCommand.ExecuteReader();


                while (masReader.Read())
                {
                    if (masReader.GetValue(0) == null | masReader.GetValue(0) == "")
                    {
                        masReader.Close();
                        masConnection.Close();
                       
                        //MessageBox.Show("Should never return null data from SQL Query...");
                        return new List<string>();
                    }
                    string tmpResults = "";

                    for (int x = 0; x < masReader.FieldCount; x++)
                    {
                        tmpResults += masReader.GetValue(x).ToString() + "|";
                    }
                    

                    QueryResults.Add(tmpResults);
                }

                List<string> JSON_Line = new List<string>();
                if (JSON_ENCODE == true)
                {
                    DataTable dt = masReader.GetSchemaTable();

                    

                    List<string> ColumnNames = new List<string>();
                    List<string> ColumnDataType = new List<string>();
                    List<string> ColumnSize = new List<string>();
                    

                    
                   
                    foreach (DataRow row in dt.Rows)
                    {
                        ColumnNames.Add(row[0].ToString());
                        ColumnDataType.Add(getDataTypeCode(row[5].ToString()).ToString());
                        
                        ColumnSize.Add(row[3].ToString());
                    }
                    
                    /*foreach (DataColumn dc in dt.Columns)
                    {
                        masReader.GetFieldType
                        ColumnNames.Add(dc.Caption);
                        ColumnDataType.Add(dc.DataType.ToString());
                        ColumnSize.Add(dc.MaxLength.ToString());
                    }*/

                    JSON_Line.Add(EncodeToJSON(ColumnNames,ColumnDataType,ColumnSize,QueryResults));
                    
                }
                

              //  foreach(DataRow dc in dt.Rows)
              //  {
              //      foreach (object c in dc.ItemArray)
              //      {
              //          QueryResults.Add(c.ToString());
              //      }
             //   }

                masConnection.Close();


                
                if (JSON_ENCODE == true)
                {
                    return JSON_Line;
                }
                else
                {
                    return QueryResults;
                }

            }
            return new List<string>();

        }



        private static string EncodeToJSON(List<string> ColumnNames, List<string> ColumnTypes, List<string> ColumneSizes, List<string> Rows)
        {
            string JSON_OUT = "{\"cols\":[";

            for (int x = 0; x < ColumnNames.Count; x++)
            {
                JSON_OUT += "{\"name\":\"" + ColumnNames[x] + "\",\"typeID\":" + ColumnTypes[x] + ",\"size\":" + ColumneSizes[x] + "},";

            }
            JSON_OUT = JSON_OUT.Substring(0,JSON_OUT.Length-1);
            JSON_OUT += "],\"rows\":[";

            foreach (string row in Rows)
            {
                string sRow = row.Substring(0, row.Length - 1);
                JSON_OUT += "[";
                foreach (string cell in sRow.Split('|'))
                {
                  
                    
                        JSON_OUT += "\"" + cell + "\",";
                    
                }
                JSON_OUT = JSON_OUT.Substring(0, JSON_OUT.Length - 1);
                JSON_OUT += "],";
                    
            }
            JSON_OUT = JSON_OUT.Substring(0, JSON_OUT.Length - 1);
            JSON_OUT += "]}";

            return JSON_OUT;
        }

        private static int getDataTypeCode(string DataTypeStr)
        {
            //This function should be as easy as two lines of code, but it's not. This function has to
            //convert typenumbers to an older typecode that is most certainly compatible with what it's
            //interfacing to. In the future we should use the below two lines of code to test it's
            //workings. Returns a -1 if it can't find a code.

            //TypeCode tp = Type.GetTypeCode(Type.GetType(DataTypeStr));
            //return (int)tp;
            
            if (DataTypeStr == "System.Decimal")
            {
                return 3;
            }
            if (DataTypeStr == "System.String")
            {
                return 12;
            }
            if (DataTypeStr == "System.DateTime")
            {
                return 91;
            }
            //Add more...
            if (DataTypeStr == "System.Boolean")
            {
            }
            if (DataTypeStr == "System.Double")
            {
            }
            if (DataTypeStr == "System.Integer" | DataTypeStr == "System.Int")
            {
                return 4;
            }


            return -1;
            

        }


    }
}
