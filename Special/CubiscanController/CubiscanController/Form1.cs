using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;

namespace CubiscanController
{
    public partial class Form1 : Form
    {

        SerialPort ComPort = new SerialPort();
        delegate void SetTextCallback(string text);
        string InputData = String.Empty;
        string ResponseData = "";
        bool ScaleIsStarted = false;
        int readingsCount = 0;

        public struct CubiscanScaleData
        {
            public string Length;
            public string Width;
            public string Height;
            public string Weight;
            public string DeemedWeight;
            public string Volume;

            public bool SetData(string CubiscanOutput)
            {
                if (CubiscanOutput.Contains("MAH") == false)
                {
                    return false;
                }
                if (CubiscanOutput.Contains("\r\n") == false)
                {
                    return false;
                }
                if (CubiscanOutput.Substring(2, 1) == "N")
                {
                    MessageBox.Show("The Cubiscan replied with a NACK, meaning it did NOT acknowledge the command.");
                    return false;
                }
                if (CubiscanOutput.Substring(32, 1) == "M" | CubiscanOutput.Substring(50,1) == "M")
                {
                    MessageBox.Show("Please set the scale in English dimensioning mode by pressing the button on the scale itself. Inches and Pounds, not Centimeters and Kilos.");
                    return false;
                }

               
               

                    //Length: 0 based offsets of 12 - 16
                    Length = CubiscanOutput.Substring(12, 5).Trim();
                    if (Length == "-----" | Length == "_____")
                    {
                        Length = "N/A";
                    }
                    //Width: 0 based offsets of 19 - 23
                    Width = CubiscanOutput.Substring(19, 5).Trim();
                    if (Width == "-----" | Width == "_____")
                    {
                        Width = "N/A";
                    }
                    //Height: 0 based offsets of 26 - 30
                    Height = CubiscanOutput.Substring(26, 5).Trim();
                    if (Height == "-----" | Height =="_____")
                    {
                        Height = "N/A";
                    }
                    //Weight: 0 based offsets of 35 - 40
                    Weight = CubiscanOutput.Substring(35, 6).Trim();
                    if (Weight == "------" | Weight == "______")
                    {
                        Weight = "N/A";
                    }
                    //Dimensional Weight: 0 based offsets of 43 - 48
                    DeemedWeight = CubiscanOutput.Substring(43, 6).Trim();
                    if (DeemedWeight == "------" | DeemedWeight == "______")
                    {
                        DeemedWeight = "N/A";
                    }
                    try
                    {
                        decimal length = decimal.Parse(Length);
                        decimal width = decimal.Parse(Width);
                        decimal height = decimal.Parse(Height);
                        decimal volume = length * width * height;
                        Volume = volume.ToString("#.00");

                    }
                    catch
                    {

                    }
                    
                    return true;
               
                
                }
            }



        CubiscanScaleData CurrentScaleData = new CubiscanScaleData();

        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            ComPort.DataReceived += new SerialDataReceivedEventHandler(ComPort_DataReceived);
            ComPort.ErrorReceived += new SerialErrorReceivedEventHandler(ComPort_ErrorReceived);

            List<string> ComPorts = GetOpenComPorts();
            comboBox1.Items.Clear();
            foreach (string str in ComPorts)
            {
                comboBox1.Items.Add(str);
            }
            List<string> Bauds = EnumerateBaudList();
            comboBox2.Items.Clear();
            foreach (string str in Bauds)
            {
                comboBox2.Items.Add(str);
            }
            List<string> DataBits = EnumerateDataBitsList();
            comboBox3.Items.Clear();
            foreach (string str in DataBits)
            {
                comboBox3.Items.Add(str);
            }
            List<string> StopBits = EnumerateStopBitsList();
            comboBox4.Items.Clear();
            foreach (string str in StopBits)
            {
                comboBox4.Items.Add(str);
            }
            List<string> ParityBits = EnumerateParitiesList();
            comboBox5.Items.Clear();
            foreach (string str in ParityBits)
            {
                comboBox5.Items.Add(str);
            }
            List<string> HandshakeBits = EnumerateHandshakingList();
            comboBox6.Items.Clear();
            foreach (string str in HandshakeBits)
            {
                comboBox6.Items.Add(str);
            }


            comboBox1.SelectedIndex = 1;
            comboBox2.SelectedIndex = 4;
            comboBox3.SelectedIndex = 3;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;


        }

        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ScaleIsStarted == true & ComPort.IsOpen == true)
            {
                button3_Click(sender, (EventArgs)e);
            }
        }

        void ComPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            MessageBox.Show("Well, we received an error. yes... we.... ***RECEIVED*** an error. good sign!");
        }

        void ComPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("HOLY SHIT! WE HAVE DATA!!!");
            readingsCount++;
            readingsCountLbl.Text = "Reads: " + readingsCount.ToString();
            
            InputData = ComPort.ReadExisting();
            if (InputData != String.Empty)
            {
                this.BeginInvoke(new SetTextCallback(SetText), new object[] { InputData });
            }


        }

        private void SetText(string text)
        {
            ResponseData += text;
            //textBox1.Text = ResponseData;
            if (ResponseData.EndsWith("\r\n") == true)
            {
                textBox1.Text = ResponseData;
                //time to process!
                CurrentScaleData.SetData(ResponseData);
                lengthLbl.Text = "Length: " + CurrentScaleData.Length + "in";
                widthLbl.Text = "Width: " + CurrentScaleData.Width + "in";
                heightLbl.Text = "Height: " + CurrentScaleData.Height + "in";
                weightLbl.Text = "Weight: " + CurrentScaleData.Weight + "lbs";
                dimensionalWeightLbl.Text = "D-Weight: " + CurrentScaleData.DeemedWeight + "lbs";
                volumeLbl.Text = "Volume: " + CurrentScaleData.Volume + "in^2";

                
                textBox1.Text = "Length: " + CurrentScaleData.Length + "in\r\n";
                textBox1.Text += "Width: " + CurrentScaleData.Width + "in\r\n";
                textBox1.Text += "Height: " + CurrentScaleData.Height + "in\r\n";
                textBox1.Text += "Weight: " + CurrentScaleData.Weight + "lbs\r\n";
                textBox1.Text += "D-Weight: " + CurrentScaleData.DeemedWeight + "lbs\r\n";
                textBox1.Text += "Volume: " + CurrentScaleData.Volume + "in^2\r\n";

                ResponseData = "";
                
            }
        }

        


        private List<string> GetOpenComPorts()
        {
            List<string> AvailablePorts = new List<string>();
            string[] openPorts = SerialPort.GetPortNames();
            foreach (string str in openPorts)
            {
                AvailablePorts.Add(str);
            }
            return AvailablePorts;
        }

        private List<string> EnumerateBaudList()
        {
            List<string> BaudList = new List<string>();
            BaudList.Add("300");
            BaudList.Add("600");
            BaudList.Add("1200");
            BaudList.Add("2400");
            BaudList.Add("9600");
            BaudList.Add("14400");
            BaudList.Add("19200");
            BaudList.Add("38400");
            BaudList.Add("57600");
            BaudList.Add("115200");

            return BaudList;

        }

        private List<string> EnumerateDataBitsList()
        {
            List<string> DataBits = new List<string>();
            int MostBits = 9;
            for (int x = 5; x < MostBits; x++)
            {
                DataBits.Add(x.ToString());
            }
            return DataBits;
        }
        private List<string> EnumerateStopBitsList()
        {
            List<string> StopBits = new List<string>();
            StopBits.Add("1");
            StopBits.Add("2");
            return StopBits;
        }
        private List<string> EnumerateParitiesList()
        {
            List<string> ParitiesList = new List<string>();
            ParitiesList.Add("None");
            ParitiesList.Add("Even");
            ParitiesList.Add("Odd");
            ParitiesList.Add("Mark");
            ParitiesList.Add("Space");
            return ParitiesList;
        }
        private List<string> EnumerateHandshakingList()
        {
            List<string> HandshakeList = new List<string>();
            HandshakeList.Add("None");
            HandshakeList.Add("XOn/XOff");
            HandshakeList.Add("RequestToSend");
            HandshakeList.Add("RequestToSendXOnXOff");
            return HandshakeList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Turn On")
            {
                //button1.Enabled = false;
                try
                {
                    ComPort.PortName = comboBox1.Text;
                    ComPort.BaudRate = int.Parse(comboBox2.Text);
                    ComPort.DataBits = int.Parse(comboBox3.Text);
                    ComPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), comboBox4.Text);
                    ComPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), comboBox5.Text);
                    ComPort.Parity = (Parity)Enum.Parse(typeof(Parity), comboBox6.Text);
                }
                catch (Exception eh)
                {
                    MessageBox.Show("Must have missed a setting field: " + eh.Message);
                }
                try
                {
                    ComPort.Open();
                }
                catch (UnauthorizedAccessException wtf)
                {
                    MessageBox.Show("Now you did it. The Matrix has been destroyed.");
                    MessageBox.Show("Actually, the problem is " + wtf.Message);

                }
                while (ComPort.IsOpen == false)
                {
                    System.Threading.Thread.Sleep(200);
                }
                //MessageBox.Show("Port is now open.");

                if (ScaleIsStarted == true)
                {
                    ScaleIsStarted = false;
                    button3_Click(sender, e);
                }

                pictureBox1.Left += pictureBox1.Width / 2;
                //button2.Enabled = true;
                button1.Text = "Turn Off";
            }
            else
            {
                if (ScaleIsStarted == true)
                {
                    ScaleIsStarted = false;
                    button3_Click(sender, e);
                }

                ComPort.Close();
                while (ComPort.IsOpen == true)
                {
                    System.Threading.Thread.Sleep(200);
                }
                //MessageBox.Show("Port is now closed.");
                pictureBox1.Left -= pictureBox1.Width / 2;
                //button2.Enabled = false;
               // button1.Enabled = true;
                button1.Text = "Turn On";
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ScaleIsStarted == true)
            {
                button3_Click(sender, e);
            }
            
            ComPort.Close();
            while (ComPort.IsOpen == true)
            {
                System.Threading.Thread.Sleep(200);
            }
            //MessageBox.Show("Port is now closed.");
            pictureBox1.Left -= pictureBox1.Width / 2;
            button2.Enabled = false;
            button1.Enabled = true;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (sender != button5)
            {
                if (button3.Text == "Start Scale")
                {
                    button3.Text = "Stop Scale";
                }
                else
                {
                    button3.Text = "Start Scale";
                }
            }
            char StartOfCommand = (char)2;
            char TheCommand = (char)67;
            char EndOfCommand = (char)3;
            char CarriageReturn = (char)13;
            char LineFeed = (char)10;
            char[] CommandString = new char[] {StartOfCommand,TheCommand,EndOfCommand,CarriageReturn,LineFeed};


            byte[] STX = System.Text.Encoding.ASCII.GetBytes(CommandString);
           

            
            //char CarriageReturn = (char)13;
            //char LineFeed = (char)10;
            //MessageBox.Show("Sending: " + STX.Length.ToString() + "bytes.");
            
            //textBox1.Text = Command;
            if (ComPort.IsOpen == true)
            {
                ComPort.RtsEnable = true;
                ComPort.Write(STX,0,STX.Length);
                ComPort.RtsEnable = false;


            }
            else
            {
                MessageBox.Show("The Com Port is Closed. If restarting the software AND cycling the power on the scale doesn't fix this, call Tom.");
            }

            if (ScaleIsStarted == true)
            {
               // ScaleIsStarted = false;
            }
            else
            {
               // ScaleIsStarted = true;
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (SettingsPnl.Visible == false)
            {
                SettingsPnl.Visible = true;
                SettingsPnl.Top = 0;
                SettingsPnl.Left = 0;
                SettingsPnl.BringToFront();
            }
            else
            {
                SettingsPnl.Visible = false;

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button3_Click(sender, e);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (ComPort.IsOpen == true)
            {
                if (ComPort.DtrEnable == true)
                {
                    dtrLbl.Text = "DTR: ON";
                }
                else
                {
                    dtrLbl.Text = "DTR: OFF";
                }

                if (ComPort.DsrHolding == true)
                {
                    dsrLbl.Text = "DSR: ON";

                }
                else
                {
                    dsrLbl.Text = "DSR: OFF";
                }
                if (ComPort.CtsHolding == true)
                {
                    ctsLbl.Text = "CTS: ON";
                }
                else
                {
                    ctsLbl.Text = "CTS: OFF";
                }
                if (ComPort.RtsEnable == true)
                {
                    rtsLbl.Text = "RTS: ON";
                }
                else
                {
                    rtsLbl.Text = "RTS: OFF";
                }
            }
        }


    }
}
