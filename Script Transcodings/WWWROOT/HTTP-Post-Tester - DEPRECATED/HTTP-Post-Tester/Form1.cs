using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Collections.Specialized;
using System.Data.Linq;

namespace HTTP_Post_Tester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string theURL = URLTxt.Text;

            List<string> postElements = new List<string>();
            List<string> postData = new List<string>();

            string tmpPOST = POSTTxt.Text;
            tmpPOST = tmpPOST.Replace("\r\n","\n");

            foreach(string element in tmpPOST.Split('\n'))
            {
                postElements.Add(element.Split('=')[0]);
                postData.Add(element.Split('=')[1]);
            }
            
            

            NameValueCollection thePOST = new NameValueCollection();

            for (int xcount = 0; xcount < postElements.Count; xcount++)
            {
                thePOST[postElements[xcount]] = postData[xcount];
            }

            
            using (WebClient wb = new WebClient())
            {
                byte[] response = wb.UploadValues(URLTxt.Text, "POST", thePOST);

                //MessageBox.Show(ASCIIEncoding.ASCII.GetString(response));


                
               // resultHTML.Write(ASCIIEncoding.ASCII.GetString(response));
                webBrowser1.DocumentText = ASCIIEncoding.ASCII.GetString(response);

            }




        }
    }
}