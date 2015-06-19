using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PriceSnagger
{
    public partial class Form2 : Form
    {
        public struct ConfigParameters
        {
            public DateTime TimeToRun;
            public bool Monday_Run;
            public bool Tuesday_Run;
            public bool Wednesday_Run;
            public bool Thursday_Run;
            public bool Friday_Run;
            public bool Saturday_Run;
            public bool Sunday_Run;
        }

        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string dataFile = "<starttime>" + dateTimePicker1.Text + "</starttime>\r\n";

            if (checkBox1.Checked == true)
            {
                dataFile += "<monday>true</monday>\r\n";
            }
            else
            {
                dataFile += "<monday>false</monday>\r\n";
            }
            if (checkBox2.Checked == true)
            {
                dataFile += "<tuesday>true</tuesday>\r\n";
            }
            else
            {
                dataFile += "<tuesday>false</tuesday>\r\n";
            }
            if (checkBox3.Checked == true)
            {
                dataFile += "<wednesday>true</wednesday>\r\n";
            }
            else
            {
                dataFile += "<wednesday>false</wednesday>\r\n";
            }
            if (checkBox4.Checked == true)
            {
                dataFile += "<thursday>true</thursday>\r\n";
            }
            else
            {
                dataFile += "<thursday>false</thursday>\r\n";
            }
            if (checkBox5.Checked == true)
            {
                dataFile += "<friday>true</friday>\r\n";
            }
            else
            {
                dataFile += "<friday>false</friday>\r\n";
            }
            if (checkBox6.Checked == true)
            {
                dataFile += "<saturday>true</saturday>\r\n";
            }
            else
            {
                dataFile += "<saturday>false</saturday>\r\n";
            }
            if (checkBox7.Checked == true)
            {
                dataFile += "<sunday>true</sunday>\r\n";
            }
            else
            {
                dataFile += "<sunday>false</sunday>\r\n";
            }
            File.WriteAllText(Application.StartupPath + "\\config.dat", dataFile);
            MessageBox.Show("Saved OK.");
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            GetConfig();
        }

        public ConfigParameters GetConfig()
        {
            ConfigParameters CP = new ConfigParameters();
            CP.Monday_Run = false;
            CP.Tuesday_Run = false;
            CP.Wednesday_Run = false;
            CP.Thursday_Run = false;
            CP.Friday_Run = false;
            CP.Saturday_Run = false;
            CP.Sunday_Run = false;
            CP.TimeToRun = new DateTime();
            if (File.Exists(Application.StartupPath + "\\config.dat") == true)
            {
               

                string ConfigFile = File.ReadAllText(Application.StartupPath + "\\config.dat");

                if (ConfigFile.ToUpper().Contains("<STARTTIME>") == true)
                {
                    CP.TimeToRun = DateTime.Parse(ConfigFile.Split(new string[] { "<starttime>" }, StringSplitOptions.None)[1].Split(new string[] { "</starttime>" }, StringSplitOptions.None)[0]);
                }
                if (ConfigFile.ToUpper().Contains("<MONDAY>TRUE") == true)
                {
                    CP.Monday_Run = true;
                }
                if (ConfigFile.ToUpper().Contains("<TUESDAY>TRUE") == true)
                {
                    CP.Tuesday_Run= true;
                }
                if (ConfigFile.ToUpper().Contains("<WEDNESDAY>TRUE") == true)
                {
                    CP.Wednesday_Run = true;
                }
                if (ConfigFile.ToUpper().Contains("<THURSDAY>TRUE") == true)
                {
                    CP.Thursday_Run = true;
                }
                if (ConfigFile.ToUpper().Contains("<FRIDAY>TRUE") == true)
                {
                    CP.Friday_Run = true;
                }
                if (ConfigFile.ToUpper().Contains("<SATURDAY>TRUE") == true)
                {
                    CP.Saturday_Run = true;
                }
                if (ConfigFile.ToUpper().Contains("<SUNDAY>TRUE") == true)
                {
                    CP.Sunday_Run = true;
                }

                dateTimePicker1.Value = CP.TimeToRun;
                checkBox1.Checked = CP.Monday_Run;
                checkBox2.Checked = CP.Tuesday_Run;
                checkBox3.Checked = CP.Wednesday_Run;
                checkBox4.Checked = CP.Thursday_Run;
                checkBox5.Checked = CP.Friday_Run;
                checkBox6.Checked = CP.Saturday_Run;
                checkBox7.Checked = CP.Sunday_Run;
                return CP;

            }
            else
            {                
                //defaults...
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                checkBox4.Checked = false;
                checkBox5.Checked = true;
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                dateTimePicker1.Value = DateTime.Parse("8:00:00 PM");
                CP.TimeToRun = dateTimePicker1.Value;
                return CP;
            }
        }
    }
}
