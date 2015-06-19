using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Media;
using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.IO;


namespace Code39_Barcode
{
    class Program
    {
        private const string Version = "1.0.0";

        private static Dictionary<string, byte[]> chars = new Dictionary<string, byte[]>();
        private static int _WideNarrowRatio = 3;
        private static int _TextHeightOffset = 20;
        private static int _Width = 191;
        public int WideNarrowRatio



        {
            get { return _WideNarrowRatio; }
            set { _WideNarrowRatio = value; }
        }

        static void Main(string[] args)
        {

            string allargs = "";
            allargs = System.Environment.GetEnvironmentVariable("QUERY_STRING");
         

            string[] eachArgument = allargs.Split('&');
            string barcode = "";
            int height = 100;
            bool showText = false;


            foreach(string eachArg in eachArgument)
            {
                if (eachArg.Split('=')[0].ToUpper() == "BARCODE")
                {
                    try{
                    barcode = eachArg.Split('=')[1];
                    }catch{}
                }
                if (eachArg.Split('=')[0].ToUpper() == "HEIGHT")
                {
                    try{
                        height = int.Parse(eachArg.Split('=')[1]);
                    }catch{}

                }
                if (eachArg.Split('=')[0].ToUpper() == "NOTEXT")
                {
                    try{
                        if (eachArg.Split('=')[1].ToUpper() == "TRUE" | eachArg.Split('=')[1].ToUpper() == "1" | eachArg.Split('=')[1].ToUpper() == "T")
                        {
                            showText = false;
                        }
                        else{
                            showText = true;
                        }
                    }catch{}
                }
                if (eachArg.Split('=')[0].ToUpper() == "WIDTH")
                {
                    try
                    {
                        _Width = int.Parse(eachArg.Split('=')[1]);
                    }
                    catch{ }
                }
                if (eachArg.Split('=')[0].ToUpper() == "GETVERSION()")
                {
                    Console.WriteLine("HTTP/1.1 200 OK\nContent-Type: text/html\n\n");
                    Console.WriteLine("Barcode Version: " + Version);
                    return;
                }

            }

            if (barcode == "")
            {
                
                //Console.WriteLine("HTTP/1.1 400 Bad Request\n");
                Console.WriteLine("Status: 400 Bad Request\nContent-type: text/plain\n\n\n");
                //Console.WriteLine("Content-Type: text/html\n\n");
                //Console.WriteLine("<html><body><h2>400</h2></body></html>\n\n");
                return;
            }

               /* try
                {
                    int index = allargs.ToUpper().IndexOf("BARCODE");
                    barcode = allargs.Substring(index).Split('&')[0].Split('=')[1];
                }
                catch { }

                try
                {
                    int index = allargs.ToUpper().IndexOf("HEIGHT");
                    height = int.Parse(allargs.Substring(index).Split('&')[0].Split('=')[1]);

                }
                catch { }
                try
                {
                    int index = allargs.IndexOf("SHOWTEXT");
                    string showTextTmp = allargs.Substring(index).Split('&')[0].Split('=')[1];
                    if (showTextTmp.ToUpper() == "TRUE" | showTextTmp.ToUpper() == "1" | showTextTmp.ToUpper() == "T")
                    {
                        showText = true;
                    }
                    else
                    {
                        showText = false;
                    }

                    Console.WriteLine(showTextTmp);
                    Console.ReadKey();

                }
                catch { }
            */
                EnumerateDictionary();
                Bitmap BarcodeImage = DrawBarcode(barcode, height, showText);

                ReturnImage(BarcodeImage);

            }


        private static byte[] JoinByteArrays(byte[] ByteArray1, byte[] ByteArray2)
        {
            byte[] newByteArray = new byte[ByteArray1.Length+ByteArray2.Length];
            ByteArray1.CopyTo(newByteArray, 0);
            ByteArray2.CopyTo(newByteArray, ByteArray1.Length - 1);
            return newByteArray;

        }

        private static void ReturnImage(Bitmap BarcodeImage)
        {

            StringBuilder s = new StringBuilder();
            s.Append("HTTP/1.1 200 OK\r\n");
            s.Append("Date: " + DateTime.Now.DayOfWeek.ToString() + ", " + DateTime.Now.Day.ToString() + " " + DateTime.Now.Month.ToString() + " " + DateTime.Now.Year.ToString() + DateTime.Now.TimeOfDay.ToString() + "GMT\r\n"); // Tue, 17 Aug 2010 11:40:00 GMT\r\n");
            s.Append("Vary: *\r\n");
            s.Append("Server: Custommade\r\n");
            s.Append("Content-Type: image/jpeg\r\n");

            MemoryStream memStream = new MemoryStream();
            BarcodeImage.Save(memStream, ImageFormat.Png);
            byte[] bitmapData = memStream.ToArray();
            memStream.Close();
            s.Append("Content-Length: " + bitmapData.Length.ToString() + "\n\n\n");
            byte[] headers = Encoding.ASCII.GetBytes(s.ToString());

            byte[] finalizedData = JoinByteArrays(headers, bitmapData);
            StreamWriter sw = new StreamWriter(Console.OpenStandardOutput());
            sw.BaseStream.Write(finalizedData,0,finalizedData.Length);
            Console.Write(Encoding.ASCII.GetChars(finalizedData));
           // foreach (byte i in finalizedData)
           // {
           //     finalizedData.
           // }
            //Console.ReadKey();
        }

        

        private static void EnumerateDictionary()
        {
            chars.Add("1", new byte[] { 1, 0, 0, 1, 0, 0, 0, 0, 1 });
            chars.Add("2", new byte[] { 0, 0, 1, 1, 0, 0, 0, 0, 1 });
            chars.Add("3", new byte[] { 1, 0, 1, 1, 0, 0, 0, 0, 0 });
            chars.Add("4", new byte[] { 0, 0, 0, 1, 1, 0, 0, 0, 1 });
            chars.Add("5", new byte[] { 1, 0, 0, 1, 1, 0, 0, 0, 0 });
            chars.Add("6", new byte[] { 0, 0, 1, 1, 1, 0, 0, 0, 0 });
            chars.Add("7", new byte[] { 0, 0, 0, 1, 0, 0, 1, 0, 1 });
            chars.Add("8", new byte[] { 1, 0, 0, 1, 0, 0, 1, 0, 0 });
            chars.Add("9", new byte[] { 0, 0, 1, 1, 0, 0, 1, 0, 0 });
            chars.Add("0", new byte[] { 0, 0, 0, 1, 1, 0, 1, 0, 0 });
            chars.Add("A", new byte[] { 1, 0, 0, 0, 0, 1, 0, 0, 1 });
            chars.Add("B", new byte[] { 0, 0, 1, 0, 0, 1, 0, 0, 1 });
            chars.Add("C", new byte[] { 1, 0, 1, 0, 0, 1, 0, 0, 0 });
            chars.Add("D", new byte[] { 0, 0, 0, 0, 1, 1, 0, 0, 1 });
            chars.Add("E", new byte[] { 1, 0, 0, 0, 1, 1, 0, 0, 0 });
            chars.Add("F", new byte[] { 0, 0, 1, 0, 1, 1, 0, 0, 0 });
            chars.Add("G", new byte[] { 0, 0, 0, 0, 0, 1, 1, 0, 1 });
            chars.Add("H", new byte[] { 1, 0, 0, 0, 0, 1, 1, 0, 0 });
            chars.Add("I", new byte[] { 0, 0, 1, 0, 0, 1, 1, 0, 0 });
            chars.Add("J", new byte[] { 0, 0, 0, 0, 1, 1, 1, 0, 0 });
            chars.Add("K", new byte[] { 1, 0, 0, 0, 0, 0, 0, 1, 1 });
            chars.Add("L", new byte[] { 0, 0, 1, 0, 0, 0, 0, 1, 1 });
            chars.Add("M", new byte[] { 1, 0, 1, 0, 0, 0, 0, 1, 0 });
            chars.Add("N", new byte[] { 0, 0, 0, 0, 1, 0, 0, 1, 1 });
            chars.Add("O", new byte[] { 1, 0, 0, 0, 1, 0, 0, 1, 0 });
            chars.Add("P", new byte[] { 0, 0, 1, 0, 1, 0, 0, 1, 0 });
            chars.Add("Q", new byte[] { 0, 0, 0, 0, 0, 0, 1, 1, 1 });
            chars.Add("R", new byte[] { 1, 0, 0, 0, 0, 0, 1, 1, 0 });
            chars.Add("S", new byte[] { 0, 0, 1, 0, 0, 0, 1, 1, 0 });
            chars.Add("T", new byte[] { 0, 0, 0, 0, 1, 0, 1, 1, 0 });
            chars.Add("U", new byte[] { 1, 1, 0, 0, 0, 0, 0, 0, 1 });
            chars.Add("V", new byte[] { 0, 1, 1, 0, 0, 0, 0, 0, 1 });
            chars.Add("W", new byte[] { 1, 1, 1, 0, 0, 0, 0, 0, 0 });
            chars.Add("X", new byte[] { 0, 1, 0, 0, 1, 0, 0, 0, 1 });
            chars.Add("Y", new byte[] { 1, 1, 0, 0, 1, 0, 0, 0, 0 });
            chars.Add("Z", new byte[] { 0, 1, 1, 0, 1, 0, 0, 0, 0 });
            chars.Add("-", new byte[] { 0, 1, 0, 0, 0, 0, 1, 0, 1 });
            chars.Add(".", new byte[] { 1, 1, 0, 0, 0, 0, 1, 0, 0 });
            chars.Add(" ", new byte[] { 0, 1, 1, 0, 0, 0, 1, 0, 0 });
            chars.Add("*", new byte[] { 0, 1, 0, 0, 1, 0, 1, 0, 0 });
            //chars.Add("*", new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 });
            chars.Add("$", new byte[] { 0, 1, 0, 1, 0, 1, 0, 0, 0 });
            chars.Add("/", new byte[] { 0, 1, 0, 1, 0, 0, 0, 1, 0 });
            chars.Add("+", new byte[] { 0, 1, 0, 0, 0, 1, 0, 1, 0 });
            chars.Add("%", new byte[] { 0, 0, 0, 1, 0, 1, 0, 1, 0 });
        }


        private static Bitmap DrawBarcode(string barcode,int height,bool showText)
        {

            
            
            //Bitmap tmpBitmap = new Bitmap(height * _WideNarrowRatio, height);
            Bitmap tmpBitmap = new Bitmap(_Width, height);
            Graphics insGraphics = Graphics.FromImage((Image)tmpBitmap);
            
            string text = barcode.ToUpper();
            if (!text.StartsWith("*"))
            {
                text = "*" + text;
            }
            if (!text.EndsWith("*"))
            {
                text = text + "*";
            }
            insGraphics.FillRectangle(new SolidBrush(Color.White), 0, 0, _Width, height);
            int narrowCount = 0;
            int wideCount = 0;
            int blankCount = 0;
            for (int i = 0; i < text.Length; i++)
            {
                byte[] insByteArray = chars[text.Substring(i, 1)];
                foreach (byte insByte in insByteArray)
                {
                    if (insByte == 1)
                    {
                        wideCount++;
                    }
                    else
                    {
                        narrowCount++;
                    }
                }
                if (i + 1 != text.Length)
                {
                    blankCount++;
                }
            }
            decimal widthPerUnit = Math.Round(Convert.ToDecimal(_Width) / Convert.ToDecimal(((wideCount * _WideNarrowRatio) + blankCount + narrowCount)), 2);
            //decimal widthPerUnit = Math.Round(Convert.ToDecimal(tmpBitmap.Width) / Convert.ToDecimal(_Width + blankCount + narrowCount),2);
            decimal currentLeft = 0;
            for (int i = 0; i < text.Length; i++)
            {
                byte[] insByteArray = chars[text.Substring(i, 1)];
                int index = 0;
                foreach (byte insByte in insByteArray)
                {
                    if (index % 2 == 0)
                    {
                        insGraphics.FillRectangle(new SolidBrush(Color.Black), (float)currentLeft, 0, (float)(insByte == 1 ? widthPerUnit * _WideNarrowRatio : widthPerUnit), height);

                    }
                    currentLeft += (insByte == 1 ? widthPerUnit * _WideNarrowRatio : widthPerUnit);

                    index++;
                }
                currentLeft += widthPerUnit;
            }

            if (showText == true)
            {
                StringFormat format = new StringFormat();
                format.LineAlignment = StringAlignment.Far;
                format.Alignment = StringAlignment.Center;
                //insGraphics.FillRectangle(Brushes.White,new Rectangle(0,height-_TextHeightOffset,height * _WideNarrowRatio,_TextHeightOffset));
                insGraphics.FillRectangle(Brushes.White, new Rectangle(0, height - _TextHeightOffset, _Width, _TextHeightOffset));
                //insGraphics.DrawString(text, new Font(FontFamily.GenericSansSerif, 12.0f), Brushes.Black, new RectangleF(0,0, height * _WideNarrowRatio, height), format);
                insGraphics.DrawString(text, new Font(FontFamily.GenericSansSerif, 12.0f), Brushes.Black, new RectangleF(0, 0, _Width, height), format);
                
                
            }

            //tmpBitmap.Save("c:\\users\\mpgmofo\\desktop\\barcode.bmp");

            return tmpBitmap;

        }

    }
}
