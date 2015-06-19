using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MotorolaOffloader
{
    public partial class Form1 : Form
    {
        public Dictionary<uint, string> PortableDevices = new Dictionary<uint, string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (System.IO.Directory.Exists(@"\\?\activesyncwpdenumerator#umb#2&306b293b&0&activesyncwpddevice-088edf83-7e88-b2e5-080c-a0ced441e979-#{6ac27878-a6fa-4155-ba85-f98f491d4f33}\\") == true)
            {
                MessageBox.Show("Woot!");
            }
            test();
        }

        private void test()
        {
            PortableDevices = new Dictionary<uint, string>();
            listBox1.Items.Clear();
            PortableDeviceApiLib.PortableDeviceManagerClass WindowsPortableDevices = new PortableDeviceApiLib.PortableDeviceManagerClass();
            string device = "";            

            for (int x = 0; x < 10; x++)
            {
                uint uintX = (uint)x;
                WindowsPortableDevices.GetDevices(ref device, ref uintX);
                if (device.Length > 0 & uintX == (uint)x)
                {
                    PortableDevices.Add(uintX, device);
                }
                device = "";
            }

            foreach (KeyValuePair<uint,string> kvp in PortableDevices)
            {
                uint cFriendlyName = 0;
                ushort usFriendlyName = 0;
                string strFriendlyName = String.Empty;

                WindowsPortableDevices.GetDeviceFriendlyName(kvp.Value, 0, ref cFriendlyName);
                //usFriendlyName = new ushort[cFriendlyName];

                    WindowsPortableDevices.GetDeviceFriendlyName(kvp.Value, ref usFriendlyName, ref cFriendlyName);


                //WindowsPortableDevices.GetDeviceFriendlyName(kvp.Value,ref tmp,ref tmp2);
                
               // listBox1.Items.Add(tmp);

            }
        
        }
    }
}
