using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Reflection;
using System.Net;
using System.Net.Sockets;



namespace carousel
{
    class Program
    {
        private static bool isRunning = true; 
        static void Main(string[] args)
        {
            SynchronousSocketListener newListener = new SynchronousSocketListener();

            newListener.ServerMain(args);


            //AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            //System.Reflection.Assembly.LoadFile("C:\\inetpub\\wwwroot\\handheld\\TP.Base.dll");
           // System.Reflection.Assembly.LoadFile("C:\\inetpub\\wwwroot\\handheld\\TP.Warehousing.White.dll");

          //  InitializeCarousels();

          //  TP.Warehousing.White.Pod.Controller.CarouselStatusUpdate += new TP.Warehousing.White.Pod.CarouselStatusUpdateEventHandler(Controller_CarouselStatusUpdate);
          //  TP.Warehousing.White.Pod.Controller.CarouselOff += new TP.Warehousing.White.Pod.CarouselOffEventHandler(Controller_CarouselOff);
          //  TP.Warehousing.White.Pod.Controller.EmergencyStop += new TP.Warehousing.White.Pod.EmergencyStopEventHandler(Controller_EmergencyStop);
          //  TP.Warehousing.White.Pod.Controller.KeypadModeFault += new TP.Warehousing.White.Pod.KeypadModeFaultEventHandler(Controller_KeypadModeFault);
          //  TP.Warehousing.White.Pod.Controller.LightActivation += new TP.Warehousing.White.Pod.LightActivationEventHandler(Controller_LightActivation);
          //  TP.Warehousing.White.Pod.Controller.LightDeactivation += new TP.Warehousing.White.Pod.LightDeactivationEventHandler(Controller_LightDeactivation);
          //  TP.Warehousing.White.Pod.Controller.LightreeButton += new TP.Warehousing.White.Pod.LightreeButtonEventHandler(Controller_LightreeButton);
          //  TP.Warehousing.White.Pod.Controller.MotorFault += new TP.Warehousing.White.Pod.MotorFaultEventHandler(Controller_MotorFault);
          //  TP.Warehousing.White.Pod.Controller.NoFaults += new TP.Warehousing.White.Pod.NoFaultsEventHandler(Controller_NoFaults);
          //  TP.Warehousing.White.Pod.Controller.OtherTextMessage += new TP.Warehousing.White.Pod.OtherTextMessageEventHandler(Controller_OtherTextMessage);
          //  TP.Warehousing.White.Pod.Controller.PhotoEyeFault += new TP.Warehousing.White.Pod.PhotoEyeFaultEventHandler(Controller_PhotoEyeFault);

            //HTTP_RESPONSE("Program Loaded And Running.\r\nWaiting For Status Updates...\r\nEXIT to terminate\r\nHELP for help\r\n");
            //HTTP_RESPONSE("=======================================================================\r\n\r\n");

            //string input = "";
        /*    try
            {
                if (System.Environment.GetEnvironmentVariable("QUERY_STRING").Length > 0)
                {
                    ProcessUserCommands(Console.ReadLine());

                    System.Threading.Thread.Sleep(3000);
                }
            }
            catch
            {
                HTTP_RESPONSE("Something bad happened!");
            }
*/
        }

        static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            

            System.IO.File.WriteAllText("c:\\inetpub\\wwwroot\\handheld\\debug.txt", args.Name);
            //HTTP_RESPONSE(args.Name);
            
            System.Reflection.Assembly ass = System.Reflection.Assembly.LoadFile("C:\\inetpub\\wwwroot\\handheld\\" + args.Name);

            return ass;
            //System.Reflection.Assembly.LoadFile("C:\\inetpub\\wwwroor\\handheld\\TP.Warehousing.White.dll");
            
        }

        static void Controller_PhotoEyeFault(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " has a photo eye fault.");
           
        }

        static void Controller_OtherTextMessage(object sender, TP.Warehousing.White.Pod.TextMessageEventArgs e)
        {
            HTTP_RESPONSE("Light Tree #" + e.LightreeAddress + " has other text message.");
         

        }

        static void Controller_NoFaults(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " has reported no faults.");
           
        }

        static void Controller_MotorFault(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " has had a motor fault.");
          
            
        }

        static void Controller_LightreeButton(object sender, TP.Warehousing.White.Pod.TextMessageEventArgs e)
        {
            HTTP_RESPONSE("Light Tree Number " + e.LightreeAddress + "'s Button Pressed.");
          
            
        }

        static void Controller_LightDeactivation(object sender, TP.Warehousing.White.Pod.LightDeactivationEventArgs e)
        {
            string results = "Carousel #" + e.AssociatedCarouselAddr + " lights have been deactivated.\r\n";
            results += "   Light Number: " + e.LightNumber.ToString() + "\r\n";
            results += "   Light Tree Address: " + e.LightTreeAddress.ToString() + "\r\n";
            HTTP_RESPONSE(results);
            
            

        }

        static void Controller_LightActivation(object sender, TP.Warehousing.White.Pod.LightActivationEventArgs e)
        {
            string results = "Carousel #" + e.AssociatedCarouselAddr + " lights have been activated.\r\n";
            results += "   Light Number: " + e.LightNumber.ToString() + "\r\n";
            results += "   Light Tree Address: " + e.LightTreeAddress.ToString() + "\r\n";
            results += "   Light Text: " + e.Text + "\r\n";

            HTTP_RESPONSE(results);
           
            
        }

        static void Controller_KeypadModeFault(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " has had a keypad mode fault.");
            
        }

        static void Controller_EmergencyStop(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " emergency stopped.");
          
        }

        static void Controller_CarouselOff(object sender, TP.Warehousing.White.Pod.CarouselFaultEventArgs e)
        {
            HTTP_RESPONSE("Carousel #" + e.CarouselAddress + " is now off.");
           
        }

        static void InitializeCarousels()
        {
            //MessageBox.Show("BEGIN");
            try
            {
                bool gotCfg = TP.Base.Config.Read(Application.ExecutablePath);
                //MessageBox.Show("HERE1");
                if (!gotCfg)
                {
                    HTTP_RESPONSE("Missing configuration file!");
                    
                    Environment.Exit(0);
                }
            }
            catch (Exception ex)
            {
                // probably a problem in the file format, what to do?
                HTTP_RESPONSE(ex.Message);
               
                //MessageBox.Show(ex.Message);
                Environment.Exit(0);
                //throw;
            }
            //MessageBox.Show("HERE");
            try
            {
                if (TP.Base.Config.Pod.IsConfigured)
                {
                    //Console.WriteLine("Initializing Carousels...");
                    if (!TP.Warehousing.White.Pod.Controller.PodActive)
                    {
                        HTTP_RESPONSE("Carousel Initialization Failure");
                       
                        //MessageBox.Show("FAILED");
                    }
                    else
                    {
                        //Console.WriteLine("Carousel Initialized!");
                       // MessageBox.Show("Init");
                        
                    }
                }
                else
                {
                    HTTP_RESPONSE("Carousel Not Configured Here!");
                  
                }
            }
            catch (Exception ey)
            {
                HTTP_RESPONSE("Exception: " + ey.Message);
               
                //MessageBox.Show("Exception: " + ey.Message);
                
            }
        }

        static void Controller_CarouselStatusUpdate(object sender, TP.Warehousing.White.Pod.CarouselStatusUpdateEventArgs e)
        {
            HTTP_RESPONSE("Carousel Address: " + e.CarouselAddress + "\r\nCarousel Mode: " + e.CarouselMode.ToString() + "\r\nCarousel Position: " + e.CurrentPosition.ToString() + "\r\nCarousel Status: " + e.CurrentStatus.ToString());
           
            //Console.WriteLine("========================================================================\r\n\r\n");

        }

        static void ProcessUserCommands(string theCommand)
        {            
            
            //if (theCommand.ToUpper() == "HELP" | theCommand == "?")
            //{
            //    Console.WriteLine("HELP command:\r\n");
            //    Console.WriteLine("   ASKFORSTATUS - Ask the carousel for status");
            //    Console.WriteLine("   ISAVAILABLE - See if a carousel is available");
            //    Console.WriteLine("   MOVETO=<int> - Moves Carousel To Specific Location");
            //    Console.WriteLine("   GETFAULTS - Checks for any faults");
            //    Console.WriteLine("   LIGHTBARMESSAGE=<string> - Display message on Lightbar 1 (N/A)");
            //    Console.WriteLine("   EXIT - Quit the Carousel Server Test");
            //    Console.WriteLine("========================================================================\r\n\r\n");
                    
            //}
            //if (theCommand.ToUpper() == "EXIT")
            //{
            //    isRunning = false;
            //}
            if (theCommand.ToUpper() == "ISAVAILABLE")
            {
               // TP.Base.ConfigParam.CarouselList test = TP.Base.Config.Carousel;
               // string address = test[0].Address.ToString();
               // Console.WriteLine("CAROUSEL ADDRESS: " + address);
               // return;
                


                if (TP.Warehousing.White.Pod.Controller.IsDeviceAvailable(TP.Warehousing.White.DeviceType.Carousel, "1") == true)
                {
                    HTTP_RESPONSE("THIS CAROUSEL IS AVAILABLE!");
                }
                else
                {
                    HTTP_RESPONSE("THIS CAROUSEL IS NOT AVAILABLE.");
                }
            }
            if (theCommand.ToUpper() == "ASKFORSTATUS")
            {
                TP.Warehousing.White.Pod.Controller.AskForStatus("1");
            }
            if (theCommand.ToUpper().Contains("MOVETO=") == true)
            {
                string newLocation = theCommand.Split('=')[1];
                TP.Warehousing.White.Pod.Controller.MoveToBin("1", int.Parse(newLocation));
            }
            if (theCommand.ToUpper() == "GETFAULTS")
            {
                bool faults = TP.Warehousing.White.Pod.Controller.AskForFaults("1");
                if (faults == true)
                {
                    HTTP_RESPONSE("There have been detected faults.");
                }
                else
                {
                    HTTP_RESPONSE("There are no faults detected.");
                }

            }
            if (theCommand.ToUpper().Contains("LIGHTBARMESSAGE=") == true)
            {
                TP.Warehousing.White.Pod.Controller.LightTreeMessage("1", 1, theCommand.Split('=')[1]);
            }
        }


        private static void HTTP_RESPONSE(string bodyText)
        {
            string response = "HTTP/1.1 200 OK\nContent-type: text/plain\n\n\n" + bodyText;

            Console.WriteLine(response);
            isRunning = false;
        }
    }

    public class SynchronousSocketListener
    {

        // Incoming data from the client.
        public static string data = null;

        public static void StartListening()
        {
            // Data buffer for incoming data.
            byte[] bytes = new Byte[1024];

            // Establish the local endpoint for the socket.
            // Dns.GetHostName returns the name of the 
            // host running the application.
            IPHostEntry ipHostInfo = Dns.Resolve(Dns.GetHostName());
            IPAddress ipAddress = ipHostInfo.AddressList[0];
            IPEndPoint localEndPoint = new IPEndPoint(ipAddress, 11000);

            // Create a TCP/IP socket.
            Socket listener = new Socket(AddressFamily.InterNetwork,
                SocketType.Stream, ProtocolType.Tcp);

            // Bind the socket to the local endpoint and 
            // listen for incoming connections.
            try
            {
                listener.Bind(localEndPoint);
                listener.Listen(10);

                // Start listening for connections.
                while (true)
                {
                    Console.WriteLine("Waiting for a connection...");
                    // Program is suspended while waiting for an incoming connection.
                    Socket handler = listener.Accept();
                    data = null;

                    // An incoming connection needs to be processed.
                    while (true)
                    {
                        bytes = new byte[1024];
                        int bytesRec = handler.Receive(bytes);
                        data += Encoding.ASCII.GetString(bytes, 0, bytesRec);
                        if (data.IndexOf("<EOF>") > -1)
                        {
                            break;
                        }
                    }

                    // Show the data on the console.
                    Console.WriteLine("HTTP/1.1 200 OK\nContent-type=text/plain\n\n\nText received : {0}", data);

                    // Echo the data back to the client.
                    byte[] msg = Encoding.ASCII.GetBytes(data);

                    handler.Send(msg);
                    handler.Shutdown(SocketShutdown.Both);
                    handler.Close();
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("\nPress ENTER to continue...");
            Console.Read();

        }

        public int ServerMain(String[] args)
        {
            StartListening();
            return 0;
        }
    }

}
