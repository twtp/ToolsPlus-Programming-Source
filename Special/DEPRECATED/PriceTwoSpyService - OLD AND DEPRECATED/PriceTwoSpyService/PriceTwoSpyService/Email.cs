using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;

namespace Email
{
    public static class Email
    {

        //public string SMTPServer = new TcpClient("toolsplus06.toolbox.local", 25);


        public static string[] EmailAddresses = {"tom@tools-plus.com"};



        public static bool SendEmail(string Message,bool isDebugging)
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
                
                byte[] EmailBodyHeader = Encoding.ASCII.GetBytes("DATA\r\nSubject: Price2Spy Errors\r\nMime-Version: 1.0;\r\nContent-Type: text/html; charset=\"ISO-8859-1\";\r\nContent-Transfer-Encoding: 7bit;\r\n\r\n");


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




    }
}
