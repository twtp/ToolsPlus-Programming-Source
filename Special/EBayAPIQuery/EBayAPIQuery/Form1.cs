using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EBayAPIQuery
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string GetAuth()
        {
            string Auth = "<RequesterCredentials>\r\n";
            Auth += " <eBayAuthToken>AgAAAA**AQAAAA**aAAAAA**4l7QUg**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6AFkoCgDJOKogudj6x9nY+seQ**+4cBAA**AAMAAA**+9vd97XY1RzNQaC3Ny/35At4dmh29/7mJimJm7iynlxb9R5IPHYGl1/ekf0mmYQ/XivwI8VBQlkPWVFVi8w6AIXNpkZsCBUnsFYSmShD2oeRZO1UpMQnjKCv3VmYXMrb+LO9NZAHh0NwgdC97zVbS3fsJIMUll1O1ct4IcibXDQ+LunjRpiHW2j7kEILWR7j7e2nxLRNG7lSTNTI5WnvsXcs+CgilfGnli358i+KaWNN6lRDPOP8PrHaX1yjhJSQZ3ZC0mjf6oEyYEOl0O9Igmf0U4mfP980iRYWR9kInH8l2wEuKtyrGOx6zX4Pgrw//kWvOgxKng9rzE3X3+UtP4lTGNLGbSfQ0AmL1a+fQ0wmHBYSsi3r1a/PiifWEHZyoySEKw2mUgB9R3dUMQk+zTsRsB4HRDM3y6Ygapa1MAOYjqPTJCYhns+W2MPuXYtDcdKG46exHx2h1XjHOkbWby2QPlrDcOhXMdSPeE6c5y7KB0RPNp5VtHgkxZVsseWUXMZtowJ05Glv+cfhVMveJPgHkbNU7Ql0aW5hgthWTjD1xFTeF3v19BbnwVhvq5aC1uJuYEKc06YEb9qJ/AZ/Cj8X/4arrVzIBvY7ri0famvgvQAXuOo2DLA8ElVyL/c5XdRJrVfBJ58gvY0SCm/pQZdRmsPWnlGRQzJLxfSLwdXdlbjef8DxTu7+8FMB43ScCZGcI1lkgbuhDP4qC1ywZ/SYpO+N0/r81VSQEj8sYCLd1W+S3KQEybOtkKd2XKJa</eBayAuthToken>\r\n";
            Auth += "</RequesterCredentials>\r\n";
            return Auth;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string[] reqText = xmlRequestTxt.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string reqs = "";
            int count = 0;
            foreach (string req in reqText)
            {
                if (count == 2)
                {
                    reqs += GetAuth();
                }
                reqs += req + "\r\n";
                count++;
            }

            responseTxt.Text = "";
            responseTxt.Refresh();
            List<string> ContentHeaders = new List<string>();
            foreach (string cHeader in contentHeaderTxt.Text.Split(new string[]{"\r\n"},StringSplitOptions.RemoveEmptyEntries))
            {
                ContentHeaders.Add(cHeader);
            }
            ContentHeaders.Add("X-EBAY-API-CALL-NAME: " + ebayCallNameTxt.Text);
            //responseTxt.Text = reqs;
            //return;
            string results = Connectivity.HTTPCalls.HTTP_POST(ebayEndPointTxt.Text,reqs,ContentHeaders);
            responseTxt.Text = results;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
