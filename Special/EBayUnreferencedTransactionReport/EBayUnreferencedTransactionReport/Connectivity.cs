using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Data.Odbc;
//using Renci.SshNet;
using System.Windows.Forms;
using System.Net.Mail;



//NEWEST MODEL: Version 1.1


namespace Connectivity
{
    public static class FTPCalls
    {
        public static string FTP_UPLOAD(string Server, string LocalFilePath)
        {
            return FTP_UPLOAD(Server, LocalFilePath, "anonymous", "anonymous");
        }
        public static string FTP_UPLOAD(string Server, string LocalFilePath, string Username, string Password)
        {
            return FTP_UPLOAD(Server, "", LocalFilePath, Username, Password, 21);
        }
        public static string FTP_UPLOAD(string Server, string DirectoryPath, string LocalFilePath, string Username, string Password, int Port)
        {
            // Get the object used to communicate with the server.
            string ftpPath = "";
            if (DirectoryPath.StartsWith("/") == true)
            {
                ftpPath = DirectoryPath;
            }
            else
            {
                ftpPath = "/" + DirectoryPath;
            }
            //if (DirectoryPath.EndsWith("/") == false)
            //{
            //    ftpPath = ftpPath + "/";
           // }
            string ftpServerPath = Server + ":" + Port.ToString() + ftpPath;
           

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftpServerPath);
            request.Method = WebRequestMethods.Ftp.UploadFile;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(Username, Password);

            StreamReader sourceStream = new StreamReader(LocalFilePath);
            byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            request.ContentLength = fileContents.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            return response.StatusCode + ": " + response.StatusDescription;
        }
        public static string FTP_GET(string Server)
        {
            return FTP_GET(Server, "anonymous", "anonymous");
        }
        public static string FTP_GET(string Server,string Username, string Password)
        {
            return FTP_GET(Server, Username, Password, 21);
        }
        public static string FTP_GET(string Server, string Username, string Password, int Port)
        {
            return FTP_GET(Server, "", Username, Password, Port);
        }
        public static string FTP_GET(string Server, string DirectoryPath, string Username, string Password, int PortNumber)
        {
            // Get the object used to communicate with the server.
            string ftpPath = "";
            if (DirectoryPath.StartsWith("/") == true)
            {
                ftpPath = DirectoryPath;
            }
            else
            {
                ftpPath = "/" + DirectoryPath;
            }

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Server + ":" + PortNumber.ToString() + ftpPath);
            request.Method = WebRequestMethods.Ftp.DownloadFile;

            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(Username, Password);

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string results = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return results;
        }
        public static string FTP_LIST(string Server)
        {
            return FTP_LIST(Server, "", "anonymous", "anonymous",21);
        }
        public static string FTP_LIST(string Server, string Username, string Password)
        {
            return FTP_LIST(Server, "", Username, Password, 21);
        }
        public static string FTP_LIST(string Server, string DirectoryPath, string Username, string Password)
        {
            return FTP_LIST(Server, DirectoryPath, Username, Password, 21);
        }
        public static string FTP_LIST(string Server, string DirectoryPath, string Username, string Password, int PortNumber)
        {
            string ftpPath = "";
            if (DirectoryPath.StartsWith("/") == true)
            {
                ftpPath = DirectoryPath;
            }
            else
            {
                ftpPath = "/" + DirectoryPath;
            }
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(Server + ":" + PortNumber + ftpPath);
            request.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            
            // This example assumes the FTP site uses anonymous logon.
            request.Credentials = new NetworkCredential(Username, Password);

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);
            string results = reader.ReadToEnd();
            reader.Close();
            response.Close();
            return results;
        }
    }
    public static class HTTPCalls
    {
        public static string HTTP_GET(string Url)
        {
            System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(Url);
            req.Credentials = System.Net.CredentialCache.DefaultCredentials;
            req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko";
            System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse();    
            System.IO.Stream str = resp.GetResponseStream();
            System.IO.StreamReader read = new System.IO.StreamReader(str);
            string results = read.ReadToEnd();
            return results;
        }
        public static string HTTP_POST(string URL, string Data)
        {
            return HTTP_POST(URL, Data, new List<string>());
        }
        public static string HTTP_POST(string Url, string Data, List<string> ContentHeaders)
        {
            string Out = String.Empty;
            System.Net.WebRequest req = System.Net.HttpWebRequest.Create(Url);
            
            foreach(string header in ContentHeaders)
            {
                if (header.ToUpper().StartsWith("CONTENT-TYPE") == true)
                {
                    req.ContentType = header;
                }
                else if (header.ToUpper().StartsWith("USER-AGENT") == true)
                {
                    ((HttpWebRequest)req).UserAgent = header;
                }
                else
                {
                    req.Headers.Add(header);
                }
            }

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
                return new List<string>();
            }
            
            return QueryResults;
        }
    }
    public static class MASCalls
        {
            private static string connstr = "DSN=SOTAMAS90;UID=bd|TOO;PWD=brian";
            //private static string connstr = "DSN=SOTAMAS90;UID=bd|TST;PWD=brian";
            private delegate object queryRunner();
            private static bool? forceJSON = null;
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
                                //s.Close(); //shouldnt these lines be here?
                            }
                            //resp.Close(); //shouldnt these lines be here?
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
            
            public static bool preloadSpeedBoost()
            {
                string res = Retrieve("SELECT TOP 1 SalesOrderNo FROM SO_SalesOrderHistoryHeader");
                try
                {
                    if (res.Length > 0)
                    {
                        if (res.Contains("rows") == true)
                        {

                            //return true in the future     
                            //MessageBox.Show("Golden!");
                            return true;
                        }
                    }
                }
                catch { }
                //return false in the future
                //MessageBox.Show("There is a problem trying to establish a connection to the ERP System.");
                return false;
            }

        }
    /*public static class VPSCalls
    {
        public struct SSHResult
        {
            public string Results;
            public string Errors;
        }
        public static SSHResult SendCommand(string Command)
        {
            SSHResult sshResult = new SSHResult();
            using (SshClient client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {
                client.Connect();
                using (SshCommand command = client.RunCommand(Command))
                {
                    string result = command.Result;
                    string error = command.Error;
                    if (result != "") { sshResult.Results = result; } else { sshResult.Results = ""; }
                    if (error != "") { sshResult.Errors = error; } else { sshResult.Results = ""; }
                }
                client.Disconnect();
            }
            return sshResult;
        }

        public static Dictionary<string,string> SQLRetrieve(string queryString)
        {

            string queryMaster = "mysql -psumm3r10 -e\"";
            Dictionary<string, string> tmpDict = new Dictionary<string, string>();
            SshClient client;
            using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {

                client.Connect();

                using (SshCommand cmd = client.RunCommand(queryMaster + queryString + "\""))
                {
                    string result = cmd.Result;
                    string errors = cmd.Error;

                   
                    tmpDict.Add("RESULTS", result);
                    tmpDict.Add("ERRORS", errors);


                }

                client.Disconnect();

            }
            return tmpDict;
        }
        public static System.Windows.Forms.TreeView GetVPSTableStructure()
        {
            System.Windows.Forms.TreeView FinalTree = new System.Windows.Forms.TreeView();


            FinalTree.Nodes.Clear();
            FinalTree.Refresh();

            var TopNode = new System.Windows.Forms.TreeNode("DB Tables");
            FinalTree.Nodes.Add(TopNode);
            var TableNodes = new List<System.Windows.Forms.TreeNode>();
            var ColumnNodes = new List<System.Windows.Forms.TreeNode>();


            SshClient client;
            using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {
                
                client.Connect();
                using (SshCommand cmd = client.RunCommand("mysql -psumm3r10 -e\"SELECT table_name FROM information_schema.tables\""))
                {
                    string result = cmd.Result;
                    string error = cmd.Error;

                    if (result != "")
                    {
                        string[] tables = result.Split('\n');
                        int counter = 0;
                        foreach (string table in tables)
                        {
                            if (counter > 0)
                            {

                                System.Windows.Forms.TreeNode tmpNode = new System.Windows.Forms.TreeNode();
                                tmpNode.Text = table;
                                tmpNode.Name = table;

                                //using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
                                //{
                                //client.Connect();
                                using (SshCommand cmd2 = client.RunCommand("mysql -psumm3r10 -e\"SELECT column_name FROM information_schema.columns WHERE table_name='" + tmpNode.Text + "'\""))
                                {
                                    string result2 = cmd2.Result;
                                    string error2 = cmd2.Error;

                                    if (result2 != "")
                                    {
                                        int count = 0;
                                        foreach (string column in result2.Split('\n'))
                                        {
                                            if (count > 0 & column != "")
                                            {
                                                tmpNode.Nodes.Add(new System.Windows.Forms.TreeNode(column));
                                            }
                                            count++;
                                        }
                                    }

                                }
                                //client.Disconnect();
                                //}

                                TableNodes.Add(tmpNode);

                            }
                            counter++;
                        }
                        TopNode.Nodes.AddRange(TableNodes.ToArray());

                        //now our table nodes are good...



                    }
                    if (error != "")
                    {
                        return new System.Windows.Forms.TreeView();
                    }
                }
                client.Disconnect();
                FinalTree.Nodes.Add(TopNode);
                return FinalTree;
            }
        }

    }
    */
    public static class EmailCalls
    {
        //public string SMTPServer = new TcpClient("toolsplus06.toolbox.local", 25);
        public static string[] EmailAddresses = { "tom@tools-plus.com" };
        public static string Subject = "Default Error";
        public static bool SendEmail(string Message, bool isDebugging)
        {
            Socket sending = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

            IPHostEntry ipHostInfo = Dns.GetHostEntry("toolsplus06.toolbox.local");
            IPAddress ipAddress = ipHostInfo.AddressList[0];
            IPEndPoint remoteEP = new IPEndPoint(ipAddress, 25);

            try
            {
                sending.Connect(remoteEP);
                //MessageBox.Show("Connected To: " + ipAddress.ToString());
                byte[] SayHello = Encoding.ASCII.GetBytes("HELO\r\n");
                byte[] SetEmailFrom = Encoding.ASCII.GetBytes("MAIL FROM: no-reply@tools-plus.com\r\n");

                byte[] EmailBodyHeader = Encoding.ASCII.GetBytes("DATA\r\nSubject: " + Subject + "\r\nMime-Version: 1.0;\r\nContent-Type: text/html; charset=\"ISO-8859-1\";\r\nContent-Transfer-Encoding: 7bit;\r\n\r\n");


                sending.Send(SayHello);
                byte[] MessageBack0 = new byte[1024];
                int response0_size = sending.Receive(MessageBack0);
                //MessageBox.Show("Received hello response from server...");
                string ServerResponse0 = Encoding.ASCII.GetString(MessageBack0, 0, response0_size);
                //MessageBox.Show(ServerResponse0);
                //MessageBox.Show("Server Code: " + ServerResponse0[0] + ServerResponse0[1] + ServerResponse0[2]);

                if (ServerResponse0[0].ToString() == "2")
                {
                    //    MessageBox.Show("Server said \"Hi\"!");
                }
                else
                {
                    MessageBox.Show("Email Error: SMTP didn't say hello.");
                    return false;
                }







                sending.Send(SetEmailFrom);

                // MessageBox.Show("Talking To Server...");
                //test
                byte[] MessageBack1 = new byte[1024];
                int response1_size = sending.Receive(MessageBack1);
                // MessageBox.Show("Recieved result from server.");
                string ServerResponse1 = Encoding.ASCII.GetString(MessageBack1, 0, response1_size);
                //  MessageBox.Show(ServerResponse1);
                // MessageBox.Show("Server Code: " + ServerResponse1[0] + ServerResponse1[1] + ServerResponse1[2]);
                if (ServerResponse1[0].ToString() == "2")
                {
                    //     MessageBox.Show("Detected OK So Far...");

                    if (isDebugging == false)
                    {
                        foreach (string emailList in EmailAddresses)
                        {

                            byte[] SetEmailTo = Encoding.ASCII.GetBytes("RCPT TO: " + emailList + "\r\n");
                            sending.Send(SetEmailTo);
                            byte[] MessageBackEmail = new byte[1024];
                            int responseEmail_size = sending.Receive(MessageBackEmail);
                            string EmailListResponse = ASCIIEncoding.ASCII.GetString(MessageBackEmail, 0, responseEmail_size);
                            // MessageBox.Show("Email Response: " + EmailListResponse);



                        }
                    }
                    else
                    {
                        byte[] SetEmailTo = Encoding.ASCII.GetBytes("RCPT TO: tom@tools-plus.com\r\n");
                        sending.Send(SetEmailTo);
                        byte[] MessageBackEmail = new byte[1024];
                        int responseEmail_size = sending.Receive(MessageBackEmail);
                        string EmailListResponse = ASCIIEncoding.ASCII.GetString(MessageBackEmail, 0, responseEmail_size);
                        // MessageBox.Show("Email Response: " + EmailListResponse);

                    }

                    byte[] MessageBack2 = new byte[1024];
                    int response2_size = sending.Receive(MessageBack2);
                    //    MessageBox.Show("Received result #2 from server...");
                    string ServerResponse2 = Encoding.ASCII.GetString(MessageBack2, 0, response2_size);
                    //     MessageBox.Show(ServerResponse2);
                    //     MessageBox.Show("Server Code: " + ServerResponse2[0] + ServerResponse2[1] + ServerResponse2[2]);




                    if (ServerResponse2[0].ToString() == "2")
                    {


                        sending.Send(EmailBodyHeader);

                        byte[] MessageBody = new byte[1024];
                        int responsebody_size = sending.Receive(MessageBody);

                        string ServerBodyResponse = Encoding.ASCII.GetString(MessageBody, 0, responsebody_size);
                        //  MessageBox.Show(ServerBodyResponse);

                        string ServerBodyCode = ServerBodyResponse[0].ToString() + ServerBodyResponse[1].ToString() + ServerBodyResponse[2].ToString();

                        if (ServerBodyCode == "354" | ServerBodyCode == "250")
                        {









                            Message = Message + "\r\n.\r\n";
                            //     MessageBox.Show("Sending the message...");
                            sending.Send(Encoding.ASCII.GetBytes(Message));

                            byte[] MessageBack3 = new byte[1024];
                            int response3_size = sending.Receive(MessageBack3);
                            //     MessageBox.Show("Recieved result #3 from server...");
                            string ServerResponse3 = Encoding.ASCII.GetString(MessageBack3, 0, response3_size);
                            //      MessageBox.Show(ServerResponse3);
                            //      MessageBox.Show("Server Code: " + ServerResponse3[0] + ServerResponse3[1] + ServerResponse3[2]);


                            if (ServerResponse3[0].ToString() == "2" | ServerResponse3[0].ToString() == "3")
                            {
                                //           MessageBox.Show("Email Sent!");
                            }
                            else
                            {
                                MessageBox.Show("Email Error: Server failed to receive message");
                                return false;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Email Error: Server Did Not Want Data! Code: " + ServerBodyCode);
                            return false;
                        }


                    }
                    else
                    {
                        MessageBox.Show("Email Error: Server recieved incorrect \"Send To\" data!");
                        return false;
                    }



                }
                else
                {

                    MessageBox.Show("Email Error: Server recieved incorrect \"Send From\" data!");
                    return false;
                }

                return true;
            }
            catch
            {

                MessageBox.Show("Email Error: Error sending email (general fault)");
                return false;
            }
        }
        public static void SendEmail2(string Subject, string Body, bool isHTML)
        {
            SmtpClient sc = new SmtpClient("toolsplus06.toolbox.local");
            MailMessage msg = null;
            string Emails = "";
            foreach (string e in EmailAddresses)
            {
                Emails += e + "; ";
            }
            Emails = Emails.Substring(0, Emails.Length - 2);
            try
            {
                msg = new MailMessage("no-reply@tools-plus.com", Emails, Subject, Body);
                msg.IsBodyHtml = isHTML;
                sc.Send(msg);
            }
            catch (Exception x)
            {
                throw x;
            }
            finally
            {
                if (msg != null)
                {
                    msg.Dispose();
                }
                
            }
        }
    }
    
}
