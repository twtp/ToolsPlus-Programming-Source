using System;
using System.Collections.Generic;

using System.Text;
using System.Printing;
using System.Windows.Forms;
using System.Windows;
using System.Drawing;
using System.Drawing.Printing;
using System.Net;
using System.Net.Sockets;


namespace barcode_printing_server
{
    class Program
    {
        public const int ServerPort = 7677;
        public static Image imageToPrint;

        static void Main(string[] args)
        {
            while (true)
            {
                StartListening();
                
            }
            //SetupPrinter("123456789");
        }

        static void StartListening()
        {
            IPAddress serverIP = IPAddress.Parse("10.0.50.16");
            //IPAddress serverIP = IPAddress.Parse("10.0.50.140");
            TcpListener serverListener = new TcpListener(serverIP,ServerPort);
            serverListener.Start();
            
            Console.WriteLine("Local EP = " + serverListener.LocalEndpoint);
            Console.WriteLine("Server Listening For Connection On Port #" + ServerPort.ToString() + "...");

            Socket serverSck = serverListener.AcceptSocket();
            Console.WriteLine("Connection From: " + serverSck.RemoteEndPoint);

            string FromClient = "";
            byte[] inData = new byte[256];
            while (FromClient.Contains("<EOR>") == false)
            {
                int inCount = serverSck.Receive(inData);
                for (int count = 0; count < inCount; count++)
                {
                    FromClient += Convert.ToChar(inData[count]);
                }
            }

            Console.WriteLine("END OF REQUEST <EOR> Tag Found. Writing Request To Text...");


            Console.WriteLine("Received Request: " + FromClient);

            bool isOK = ProcessRequest(FromClient);

            serverSck.Close();
            serverListener.Stop();

            


        }
        static bool ProcessRequest(string DataIn)
        {
            Console.WriteLine("Preparing To Process Request Information...");
            try
            {
                Console.Write("Does The Request Contain 'BARCODE='? ");
                if (DataIn.ToUpper().Contains("BARCODE=") == false)
                {
                    Console.Write("No. Closing Connection.\r\n");
                    return false;
                }
                else
                {
                    Console.Write("Yes.\r\n");
                }
                if (DataIn.ToUpper().Contains("&QTY=") == false)
                {
                    Console.WriteLine("No Quantity Inserted. Not Printing.");
                    return false;
                }

                Console.WriteLine("Putting Data Where It Should Go...");
                string BarcodeToPrint = "";
                int QtyToPrint = 0;
                string ItemNumber = "";
                if (DataIn.ToUpper().Split(new string[] { "BARCODE=" }, StringSplitOptions.None)[1].Contains("&") == true)
                {
                    BarcodeToPrint = DataIn.ToUpper().Split(new string[] { "BARCODE=" }, StringSplitOptions.None)[1].Split('&')[0];

                }
                else
                {
                    BarcodeToPrint = DataIn.ToUpper().Split(new string[] { "BARCODE=" }, StringSplitOptions.None)[1].Trim();
                }
                ItemNumber = DataIn.ToUpper().Split(new string[] {"ITEM="},StringSplitOptions.None)[1].Split('&')[0];
                QtyToPrint = int.Parse(DataIn.Split(new string[] { "&QTY=" }, StringSplitOptions.None)[1].Split('<')[0]);
                Console.WriteLine("Ready To Print...");
                //SetupPrinter(BarcodeToPrint,QtyToPrint);
                SetupEPL2(BarcodeToPrint,ItemNumber,QtyToPrint);



                return true;
            }
            catch
            {
                return false;
            }
        }


        static void SetupEPL2(string barcodeNumber, string itemNumber, int QtyToPrint)
        {
            TP.Base.Printing.EPL2.EPL2BarcodeDocument epl2 = new TP.Base.Printing.EPL2.EPL2BarcodeDocument(barcodeNumber,itemNumber);
            epl2.Print("Barcode Printer Receiving", QtyToPrint);
            Console.WriteLine("Printing Completed. Closing Connection And Restarting.\r\n\r\n\r\n\r\n\r\n");
           
        }


        static void SetupPrinter(string barcodeNumber,int QtyToPrint)
        {
            Console.Write("Aquiring Barcode Image For #" + barcodeNumber + "...");
            AquireBarcodeImage(barcodeNumber);
            Console.Write("OK.\r\n");
            Console.Write("Setting Up Printer...");
            PrintServer TP_PrintServer = new PrintServer("\\\\10.0.50.16");
            System.Drawing.Printing.PrintDocument barcodeDocument = new System.Drawing.Printing.PrintDocument();
            barcodeDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(barcodeDocument_PrintPage);
            barcodeDocument.PrinterSettings.PrinterName = "Barcode Printer Receiving";
            Console.Write("OK.\r\n");
            Console.WriteLine("Printing...");
            for (int PrintCount = 0; PrintCount < QtyToPrint; PrintCount++)
            {
                barcodeDocument.Print();
            }
            Console.WriteLine("Complete. Closing Connections Then Restarting Listener.\r\n\r\n\r\n\r\n");
        }
        static void AquireBarcodeImage(string barcodeNumber)
        {
            
            var webURL = WebRequest.Create("http://10.0.50.16/barcode.pl?barcode=" + barcodeNumber + "&height=70&notext=0");

            using (var response = webURL.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                {
                    imageToPrint = Bitmap.FromStream(stream);
                }
            }

            //MessageBox.Show("Alright, the size of the barcode is: " + imageToPrint.Size.Width.ToString() + "px x " + imageToPrint.Size.Height.ToString() + "px (" + imageToPrint.Size.ToString() + "px)");


        }

        static void barcodeDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Graphics graphics = e.Graphics;

            graphics.DrawImage(imageToPrint, 0, 0);


        }
    }
}
