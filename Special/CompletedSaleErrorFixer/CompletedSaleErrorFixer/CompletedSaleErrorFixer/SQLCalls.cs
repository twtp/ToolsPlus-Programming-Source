using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;



namespace CompletedSaleErrorFixer
{
    public static class SQLCalls
    {
        
      

        
        
        public const string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";

        public static List<string> sqlQuery(string queryString)
        {
            List<string> QueryResults = new List<string>();
            QueryResults.Capacity = 1000;
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
                //MessageBox.Show(A.Message);
                return new List<string>();
            }
            //MessageBox.Show("IM LEAVING HERE");
            return QueryResults;
        }

    }
}
