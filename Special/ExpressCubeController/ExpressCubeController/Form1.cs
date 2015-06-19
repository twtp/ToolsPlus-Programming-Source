using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;

namespace ExpressCubeController
{
    public partial class Form1 : Form
    {
        SerialPort ComPort = new SerialPort();
        delegate void SetTextCallback(string text);
        string InputData = String.Empty;
        string ResponseData = "";

        public const string ExpressCube_InitializerCode = "EC,SET,4,1,5,0,6,0,13,194.0,\r\n";
        
        public const string ExpressCube_GetDimWeight = "EC,GET,40,44,45,46,47,48,50,41,42,39,49,\r\n";


        public struct ExpressCubeData
        {
            public string Length;
            public string Width;
            public string Height;
            public string Weight;


            public bool SetData(string datafeed)
            {
                //MessageBox.Show(datafeed.Length.ToString());
                if (datafeed.Length > 38)
                {
                    Height = datafeed.Split(',')[7];
                    Length = datafeed.Split(',')[9];
                    Width = datafeed.Split(',')[11];
                    Weight = datafeed.Split(',')[13];
                    return true;
                }
                else
                {
                    //was the intializer
                    return false;
                }
            }
        }

        
        //the expresscube can now be placed on ANY com port we want. SizeIt2 limited us to com1. Bastards!!!


        public ExpressCubeData ExpressCubeLiveData = new ExpressCubeData();


        public Form1()
        {
            InitializeComponent();

        }


        private void SetText(string text)
        {
            ResponseData += text;
            //textBox1.Text = ResponseData;
            if (ResponseData.EndsWith("\r\n") == true)
            {
                

                bool test = ExpressCubeLiveData.SetData(ResponseData);
                if (test == false)
                {
                    //scale has been initialized, now lets pump the data out by enabling it's timer...
                    //MessageBox.Show("Timer getting activated!");
                    scaleTmr.Enabled = true;
                }
                if (test == true)
                {
                    lengthLbl.Text = "Length: " + ExpressCubeLiveData.Length;
                    widthLbl.Text = "Width: " + ExpressCubeLiveData.Width;
                    heightLbl.Text = "Height: " + ExpressCubeLiveData.Height;
                    weightLbl.Text = "Weight: " + ExpressCubeLiveData.Weight;
                }


                /*textBox1.Text = ResponseData;
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
                */
                //MessageBox.Show("Response: " + ResponseData);
                ResponseData = "";
            }
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


        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //don't think this scale 'shuts down' like the cubiscan. dont do anything here if thats the case.
        }
        void ComPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //
            InputData = ComPort.ReadExisting();
            if (InputData != String.Empty)
            {
                this.BeginInvoke(new SetTextCallback(SetText), new object[] { InputData });
            }
        }
        void ComPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            //
            MessageBox.Show("Error Received!!!");
        }

        private void button1_Click(object sender, EventArgs e)
        {
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

            MessageBox.Show("Port Open!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ComPort.Write(textBox1.Text + "\r\n");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ComPort.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "Start")
            {
                button4.Text = "Stop";
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

                if (ComPort.IsOpen == true)
                {
                    //now we send our initializer code out...                
                    //MessageBox.Show("was here too");
                    ComPort.Write(ExpressCube_InitializerCode);
                }
            }
            else
            {
                ComPort.Close();
                button4.Text = "Start";
            }
            
        }

        private void scaleTmr_Tick(object sender, EventArgs e)
        {
            ComPort.Write(ExpressCube_GetDimWeight);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (SettingsPnl.Visible == false)
            {
                SettingsPnl.Top = 0;
                SettingsPnl.Left = 0;
                SettingsPnl.Visible = true;
            }
            else
            {
                SettingsPnl.Visible = false;
            }
        }
    }
}
