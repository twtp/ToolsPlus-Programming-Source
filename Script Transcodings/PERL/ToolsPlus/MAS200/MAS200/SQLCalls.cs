using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net;
using System.Data.Odbc;


namespace MAS200
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
        public static List<string> sqlQuery(string queryString,int ConnectionType)
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

                masConnection.Open();

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

              //  DataTable dt = masReader.GetSchemaTable();

              //  foreach(DataRow dc in dt.Rows)
              //  {
              //      foreach (object c in dc.ItemArray)
              //      {
              //          QueryResults.Add(c.ToString());
              //      }
             //   }

                masConnection.Close();




                return QueryResults;

            }
            return new List<string>();

        }

    }
}
