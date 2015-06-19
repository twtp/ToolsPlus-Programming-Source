using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;


namespace PriceSnagger
{
    public class SQLCalls
    {
        
       
        
        
        
        public const string connectionString = "Data Source=10.0.50.17;Initial Catalog=toolsplus;User Id=sa;Password=admin1";

        public List<string> sqlQuery(string queryString)
        {
            //if (queryString.StartsWith("SELECT") == false)
            //{
            //    MessageBox.Show(queryString);
            //}
            List<string> QueryResults = new List<string>();
            QueryResults.Capacity = 1000;
            try
            {
                //MessageBox.Show("IN SQL 2");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    //MessageBox.Show("IN SQL 3");
                    using (SqlCommand command = new SqlCommand(queryString, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            //MessageBox.Show("IN SQL 4");
                            while (reader.Read())
                            {
                                if (reader.GetValue(0) == null | reader.GetValue(0) == "")
                                {
                                    reader.Close();
                                    MessageBox.Show("Should never return null data from SQL Query...");
                                    return new List<string>();
                                }
                                string tmpResults = "";

                                for (int x = 0; x < reader.FieldCount; x++)
                                {
                                    tmpResults += reader.GetValue(x).ToString() + "|";
                                }
                                QueryResults.Add(tmpResults);
                            }
                            //MessageBox.Show("IN SQL 5");
                            reader.Close();

                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception A)
            {
                MessageBox.Show(A.Message);
                return new List<string>();
            }
            //MessageBox.Show("IM LEAVING HERE");
            return QueryResults;
        }

        

        private void SQLCalls_Load(object sender, EventArgs e)
        {

        }

    }
}
