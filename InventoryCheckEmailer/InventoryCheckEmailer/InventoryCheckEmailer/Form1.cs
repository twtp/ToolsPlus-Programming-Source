using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net.Sockets;
using System.Net;


namespace InventoryCheckEmailer
{
    public partial class Form1 : Form
    {

        public static string[] EmailAddresses = { "tom@tools-plus.com", "john@tools-plus.com", "eric@tools-plus.com" };
        public static string[] VarianceOnlyEmails = { "jeanne@tools-plus.com", "tom@tools-plus.com" };


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            string InventoryInfo = "";
            string ColumnData = "Problem Type,Username,Version,Time Scanned,Warehouse/Store,Location Type,Location,Item Name,Item Description,Quantity,Scanned QTY\r\n";
            try
            {
                 InventoryInfo = File.ReadAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\data.dat");
            }
            catch
            {
                if (File.Exists("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\data.dat") == true)
                {
                    File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + " -  The Data.dat File cannot be open. Someone probably has it opened and more than likely it's Tom's fault!\r\n\r\n");
                }
                else
                {
                    File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + " -  The Data.dat File doesn't Exist. Either no inventorying was done or the data.dat file cannot be found!\r\n\r\n");
               
                }
                Application.Exit();
                return;
            }
            InventoryInfo = InventoryInfo.Replace("\r\n", "\n");

            string FinalCSVFile = ColumnData + InventoryInfo;

            List<string> Variances = new List<string>();
            foreach (string Row in InventoryInfo.Split('\n'))
            {
                Variances.Add(Row);
            }


            
            InventoryInfo= ColumnData + SortRowsByLocationType(Variances);
            





            File.WriteAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\backup_inventory_issues.csv", FinalCSVFile);

            string HTMLTable = "<html><body><table border=2><tr><td><b>Problem Type</b></td><td><b>Username</b></td><td><b>Version</b></td><td><b>Time Scanned</b></td><td><b>Warehouse/Store</b></td><td><b>Location Type</b></td><td><b>Location</b></td><td><b>Item Name</b></td><td><b>Item Description</b></td><td><b>Quantity<b></td><td><b>Actual QTY</b></td></tr>";

            foreach(string invItems in InventoryInfo.Split('\n'))
            {
                HTMLTable += "\r\n<tr>";
                foreach (string cell in invItems.Split(','))
                {
                    HTMLTable += "<td>" + cell + "</td>";
                }
                HTMLTable += "</tr>";
                
            }

            HTMLTable += "</table><br><br>You can also get the Excel Spreadsheet of the above table <a href=\"\\\\toolsplus04\\wwwroot\\whse\\inventory_logs\\backup_inventory_issues.csv\">here</a></body></html>";

            textBox1.Text = HTMLTable;


            string EmailBodyMessage = "Subject: Variance / Manual Entry Report\r\n";
            EmailBodyMessage += "MIME-Version: 1.0\r\nContent-Type: text/html\r\n";
           // EmailBodyMessage += "\r\nContent-Type: text/plain; charset=us-ascii; name=variance_report.csv\r\nContent-Transfer-Encoding: 7bit\r\nContent-Disposition: attachment; filename=\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\backup_inventory_issues.csv\n\n";
           // EmailBodyMessage += FinalCSVFile;
            EmailBodyMessage += HTMLTable;

            bool test = SendEmail(EmailBodyMessage, false,false);

            //MessageBox.Show(test.ToString());


            if (test == true)
            {
                bool test2 = SendVarianceEmail(InventoryInfo);
                if (test2 == true)
                {
                    File.Delete("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\data.dat");
                    Application.Exit();
                }
                else
                {
                    File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + " -  The Variance Report failed. This is probably either due to nothing to process or someone has the data.dat file open or the backup_inventory_issues.csv file open\r\n\r\n");

                }
            }
            else
            {
                File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + " -  The Variance / Manual Entry Report failed. This is probably either due to nothing to process or someone has the data.dat file open or the backup_inventory_variances.csv file open.\r\n\r\n");


            }

            
            
        }










        public string SortRowsByLocationType(List<string> Variances)
        {
            //11 total columns (0-10) location type = 5
           

            List<string> SortedDataKey = new List<string>();

            int xcount = 0;
            foreach (string row in Variances)
            {
                if (row.Length > 5)
                {
                    SortedDataKey.Add(row.Split(',')[5] + "|" + xcount.ToString());
                }
                xcount++;
            }

            SortedDataKey.Sort();


            List<string> SortedData = new List<string>();
            foreach (string row in SortedDataKey)
            {
                int indexer = int.Parse(row.Split('|')[1]);

                SortedData.Add(Variances[indexer]);


            }
            string finalResult = "";
            foreach (string row in SortedData)
            {
                finalResult += row + "\n";
            }

            return finalResult;


        }













        public bool SendVarianceEmail(string DataIn)
        {
            string VarianceOnlyReport = "";

            foreach (string row in DataIn.Split('\n'))
            {
                
                    if (row.ToUpper().Contains("VARIANCE") == true)
                    {
                        //add this row to the report
                        VarianceOnlyReport += row + "\n";
                    }
                
            }

            File.WriteAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\backup_inventory_variances.csv", VarianceOnlyReport);
            
            string EmailBodyMessage = "Subject: Variance Report\r\n";
            EmailBodyMessage += "MIME-Version: 1.0\r\nContent-Type: text/html\r\n";


            string HTMLTable = "<html><body><table border=2><tr><td><b>Problem Type</b></td><td><b>Username</b></td><td><b>Version</b></td><td><b>Time Scanned</b></td><td><b>Warehouse/Store</b></td><td><b>Location Type</b></td><td><b>Location</b></td><td><b>Item Name</b></td><td><b>Item Description</b></td><td><b>Quantity<b></td><td><b>Actual QTY</b></td></tr>";

            foreach (string invItems in VarianceOnlyReport.Split('\n'))
            {
                HTMLTable += "\r\n<tr>";
                foreach (string cell in invItems.Split(','))
                {
                    HTMLTable += "<td>" + cell + "</td>";
                }
                HTMLTable += "</tr>";

            }

            HTMLTable += "</table><br><br>You can also get the Excel Spreadsheet of the above table <a href=\"\\\\toolsplus04\\wwwroot\\whse\\inventory_logs\\backup_inventory_variances.csv\">here</a></body></html>";
            EmailBodyMessage += HTMLTable;




            
            bool sentEmail = SendEmail(EmailBodyMessage, false,true);

            if (sentEmail == true)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public static bool SendEmail(string Message, bool isDebugging, bool varianceOnly)
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

                byte[] EmailBodyHeader = Encoding.ASCII.GetBytes("DATA\r\n");


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
                    File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: SMTP didn't say hello.");
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
                        if (varianceOnly == false)
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
                            foreach (string emailList in VarianceOnlyEmails)
                            {

                                byte[] SetEmailTo = Encoding.ASCII.GetBytes("RCPT TO: " + emailList + "\r\n");
                                sending.Send(SetEmailTo);
                                byte[] MessageBackEmail = new byte[1024];
                                int responseEmail_size = sending.Receive(MessageBackEmail);
                                string EmailListResponse = ASCIIEncoding.ASCII.GetString(MessageBackEmail, 0, responseEmail_size);
                                // MessageBox.Show("Email Response: " + EmailListResponse);



                            }
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
                                File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: Server failed to receive message");
                                return false;
                            }
                        }
                        else
                        {
                            File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: Server Did Not Want Data! Code: " + ServerBodyCode);
                            return false;
                        }


                    }
                    else
                    {
                        File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: Server recieved incorrect \"Send To\" data!");
                        return false;
                    }



                }
                else
                {

                    File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: Server recieved incorrect \"Send From\" data!");
                    return false;
                }

                return true;
            }
            catch
            {

                File.AppendAllText("\\\\TOOLSPLUS04\\wwwroot\\whse\\inventory_logs\\error.log", DateTime.Now.ToString() + "Email Error: Error sending email (general fault)");
                return false;
            }
        }
    }
}