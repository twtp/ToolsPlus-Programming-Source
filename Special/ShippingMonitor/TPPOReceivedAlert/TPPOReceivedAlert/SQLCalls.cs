using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;



namespace POConnectivity
{
    public static class HTTPCalls
    {
        public static string HTTP_GET(string Url)
        {
            System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(Url);
            req.Credentials = System.Net.CredentialCache.DefaultCredentials;
            req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko";
            System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse();
            //System.Net.WebRequest webReq = System.Net.WebRequest.Create(Url);
            //System.Net.WebResponse webResp = webReq.GetResponse();
            System.IO.Stream str = resp.GetResponseStream();
            System.IO.StreamReader read = new System.IO.StreamReader(str);
            string results = read.ReadToEnd();
            return results;
        }

        public static string HTTP_POST(string Url, string Data)
        {
            string Out = String.Empty;
            System.Net.WebRequest req = System.Net.WebRequest.Create(Url);

            try
            {
                req.Method = "POST";
                req.Timeout = 100000;
                req.ContentType = "application/x-www-form-urlencoded";
                byte[] sentData = Encoding.UTF8.GetBytes(Data);
                req.ContentLength = sentData.Length;
                using (System.IO.Stream sendStream = req.GetRequestStream())
                {
                    sendStream.Write(sentData, 0, sentData.Length);
                    sendStream.Close();
                }
                System.Net.WebResponse res = req.GetResponse();
                System.IO.Stream ReceiveStream = res.GetResponseStream();
                string str = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8))
                {
                    Char[] read = new Char[256];
                    int count = sr.Read(read, 0, 256);

                    while (count > 0)
                    {
                        str = new String(read, 0, count);
                        Out += str;
                        count = sr.Read(read, 0, 256);
                    }
                }
                //MessageBox.Show("The Server Responded With: " + Out);

                //webBrowser1.DocumentText = Out;

            }
            catch (ArgumentException ex)
            {
                Out = string.Format("HTTP_ERROR :: The second HttpWebRequest object has raised an Argument Exception as 'Connection' Property is set to 'Close' :: {0}", ex.Message);
            }
            catch (Exception ex)
            {
                Out = string.Format("HTTP_ERROR :: WebException raised! :: {0}", ex.Message);
            }
            catch
            {
                Out = string.Format("DAMNIT ERROR");
            }

            return Out;
        }
    }

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
