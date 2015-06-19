using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Net;
using Renci.SshNet;
using System.Windows.Forms;

namespace Connectivity
{

    public static class VPSCalls
    {
        public struct VPSResult
        {
            public string Results;
            public string Errors;
        }
        public static TreeView GetDatabaseTree(TreeView treeViewTemplate)
        {
            TreeView treeView1 = treeViewTemplate;
            treeView1.Nodes.Clear();
            treeView1.Refresh();

            var TopNode = new TreeNode("DB Tables");
            treeView1.Nodes.Add(TopNode);
            var TableNodes = new List<TreeNode>();
            var ColumnNodes = new List<TreeNode>();



            using (SshClient client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {
                //listBox1.Items.Add("Retrieving List Of Tables...");
                //listBox1.Refresh();
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

                                TreeNode tmpNode = new TreeNode();
                                tmpNode.Text = table;
                                tmpNode.Name = table;

                                //using (client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
                                //{
                                //client.Connect();
                                using (SshCommand cmd2 = client.RunCommand("mysql -psumm3r10 -e\"SELECT column_name,data_type,character_maximum_length,is_nullable,column_default FROM information_schema.columns WHERE table_name='" + tmpNode.Text + "'\""))
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
                                                string[] columnFields = column.Split('\t');                                                
                                                tmpNode.Nodes.Add(new TreeNode(columnFields[0] + " (" + columnFields[1] + "(" + columnFields[2] + ")," + (columnFields[3] == "YES" ? "null":"not null") + ", Default(" + columnFields[4] + "))"));
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
                        //listBox1.Items.Add("Error Pulling Tables: " + error);
                        //listBox1.Refresh();
                    }
                }
                client.Disconnect();
                //listBox1.Items.Add("Retrieved Tables And Columns...");
                //listBox1.Refresh();
            }
            return treeView1;
        }
        public static VPSResult QueryDB(string sqlQuery)
        {
            VPSResult tmpResult = new VPSResult();
            using (SshClient client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {
                client.Connect();
                using (SshCommand command = client.RunCommand("mysql -psumm3r10 esavelle_salesorder -e\"" + sqlQuery + "\""))
                {
                    string result = command.Result;
                    string error = command.Error;
                    if (result != ""){tmpResult.Results = result;}else{tmpResult.Results = "";}
                    if (error != ""){tmpResult.Errors = error;}else{tmpResult.Errors = "";}
                }
                client.Disconnect();
            }
            return tmpResult;
        }
        public static VPSResult SendCommand(string Command)
        {
            VPSResult vpsResult = new VPSResult();
            using (SshClient client = new SshClient("50.23.169.36", 9042, "esavelle", "summ3r10"))
            {
                client.Connect();
                using (SshCommand command = client.RunCommand(Command))
                {
                    string result = command.Result;
                    string error = command.Error;
                    if (result != "") { vpsResult.Results = result; } else { vpsResult.Results = ""; }
                    if (error != "") { vpsResult.Errors = error; } else { vpsResult.Results = ""; }
                }
                client.Disconnect();
            }
            return vpsResult;
        }
    }









    public static class SSHCalls
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
                    if (result != ""){sshResult.Results = result;}else{sshResult.Results = "";}
                    if (error != ""){sshResult.Errors = error;}else{sshResult.Results = "";}
                }
                client.Disconnect();
            }
            return sshResult;
        }
    }










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
}
