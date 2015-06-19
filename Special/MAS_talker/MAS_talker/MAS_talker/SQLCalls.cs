using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;
using System.Data.Odbc;


namespace SQLCalls
{
    public static class MASCalls
    {
        public static class MAS200
        {            
            private static string connstr = "DSN=SOTAMAS90;UID=bd|TOO;PWD=brian";
            //private static string connstr = "DSN=SOTAMAS90;UID=bd|TST;PWD=brian";
            private delegate object queryRunner();

            private static bool? forceJSON = null;

            private static object runQuery(queryRunner qr)
            {
                int retryCount = 0;
                while (retryCount < 5)
                {
                    try
                    {
                        object retval = qr();
                        return retval;
                    }
                    catch (OdbcException oex)
                    {
                        if (oex.Message.Contains("IM002"))
                        { // data source name not found...
                            //throw new MasNotInstalledException(oex);
                        }
                        else
                        {
                            throw;
                        }
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                throw new Exception("MAS200 connection retry count exceeded");
            }

            // MAS doesn't support C#-style parameterized queries, so interpret the params here.
            // Any weird types will need to be added here, but this should be sufficient for 99%.
            //
            // Also, unrelated to parameters, but also important: MAS inner joining is messed up
            // when also using an orderby, so either don't do orderings in the sql, or join using
            // commas and set it up in the where clause instead.
            //
            // Finally, the way MAS handles integers is also messed up, you can't just cast to an
            // int, it needs a Convert.ToInt32() call.
          /*  private static string deParameterize(string sql, params QueryParameter[] parameters)
            {
                foreach (QueryParameter qp in parameters)
                {
                    switch (qp.Type)
                    {
                        case SqlDbType.Int:
                            int num = (int)qp.Value;
                            sql = sql.Replace(qp.Name, num.ToString());
                            break;
                        case SqlDbType.VarChar:
                            sql = sql.Replace(qp.Name, "'" + qp.Value.ToString().Replace("'", "''") + "'");
                            break;
                        case SqlDbType.Date:
                            string dateStr = string.Format("{0:yyyy-MM-dd}", (DateTime)qp.Value);
                            sql = sql.Replace(qp.Name, "{ d '" + dateStr + "' }");
                            break;
                        case SqlDbType.Decimal:
                            decimal dec = (decimal)qp.Value;
                            sql = sql.Replace(qp.Name, dec.ToString());
                            break;
                        default:
                            throw new Exception("SqlDbType " + qp.GetType().ToString() + " not supported!");
                    }
                }
                return sql;
            }*/

            public static string Retrieve(string sql)
            {
                //sql = deParameterize(sql, parameters);

                if (forceJSON == null)
                {
                    forceJSON = false;
                    //TP.Base.Logger.Log("mas200 force json queries state is: " + forceJSON.ToString());
                }

               // if (forceJSON == true)
               // {
               return JSONInterface.Retrieve(sql);
               // }

                try
                {
                    object retval = runQuery(delegate()
                    {
                        OdbcConnection conn = new OdbcConnection(connstr);
                        conn.Open();
                        OdbcCommand sth = new OdbcCommand(sql, conn);

                        OdbcDataAdapter da = new OdbcDataAdapter(sth);
                        DataSet ds = new DataSet("t");
                        da.Fill(ds, "t");
                        conn.Close();
                        return ds.Tables["t"];
                    });
                    return (string)retval;
                }
                catch { return sql; }
               // catch (MasNotInstalledException)
               // {
                    //TP.Base.Logger.Log("mas is not installed, forcing json alternative query");
              //      forceJSON = true;
              //      return sql;
              //  }

            }






            private static class JSONInterface
            {

                //public static System.Data.DataTable Retrieve(string sql)
                public static string Retrieve(string sql)
                {
                    string ro = query(sql);
                    if (ro == null)
                    {
                        return null;
                    }
                    else
                    {
                        return ro;
                    }

                   /* System.Data.DataTable dt = new System.Data.DataTable();
                    foreach (ReturnObjectColumnInfo roci in ro.cols)
                    {
                        dt.Columns.Add(roci.name, roci.columnType);
                    }

                    foreach (List<object> r in ro.rows)
                    {
                        System.Data.DataRow row = dt.NewRow();
                        for (int i = 0; i < r.Count; i++)
                        {
                            row[i] = r[i] == null ? System.DBNull.Value : r[i];
                        }
                        dt.Rows.Add(row);
                    }

                    return dt;
                    * */
                }

                private static string query(string sql)
                {
                    try
                    {
                        string content = null;
                        System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create("http://toolsplus04/whse/arbitrary_sql2.plex?" + System.Web.HttpUtility.UrlEncode(sql));
                        req.Timeout = 600000;
                        using (System.Net.WebResponse resp = (System.Net.HttpWebResponse)req.GetResponse())
                        {
                            using (System.IO.Stream s = resp.GetResponseStream())
                            {
                                using (System.IO.StreamReader r = new System.IO.StreamReader(s))
                                {
                                    content = r.ReadToEnd();
                                }
                            }
                        }
                        return content;
                    }
                    catch (Exception ex)
                    {
                        //TP.Base.Logger.Log(" --> JSONInterface HTTP (probably?) error: " + ex.Message);
                        return null;
                    }
                }

                public class ReturnObject
                {
                    private List<ReturnObjectColumnInfo> _cols;
                    public List<ReturnObjectColumnInfo> cols { get { return this._cols; } set { this._cols = value; } }
                    private List<List<object>> _rows;
                    public List<List<object>> rows { get { return this._rows; } set { this._rows = value; } }
                }

                public class ReturnObjectColumnInfo
                {
                    private string _name;
                    public string name { get { return this._name; } set { this._name = value; } }
                    private int _typeID;
                    public int typeID { get { return this._typeID; } set { this._typeID = value; } }
                    private int _size;
                    public int size { get { return this._size; } set { this._size = value; } }
                    public System.Type columnType
                    {
                        get
                        {
                            switch (this._typeID)
                            {
                                case -7: // System.Data.SqlDbType.Bit
                                    return typeof(bool);
                                case -1: // System.Data.SqlDbType.Text
                                case 1:  // System.Data.SqlDbType.Char
                                case 12: // System.Data.SqlDbType.VarChar
                                    return typeof(string);
                                case 3:  // System.Data.SqlDbType.Decimal
                                    return typeof(decimal);
                                case 9:  // System.Data.SqlDbType.Date
                                case 91: // System.Data.SqlDbType.DateTime
                                    return typeof(DateTime);
                                default:
                                    return typeof(string); // default stringify?
                            }
                        }
                    }
                }

            }


        }
    


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
}
