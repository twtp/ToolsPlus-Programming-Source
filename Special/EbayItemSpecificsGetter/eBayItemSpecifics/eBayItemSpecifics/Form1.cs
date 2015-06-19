using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace eBayItemSpecifics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<String> headers = new List<string>();
            headers.Add("content-type: application/x-www-form-urlencoded");
            headers.Add("user-agent: ToolsPlus EBay Interface v0.1");
            headers.Add("X-EBAY-API-COMPATIBILITY-LEVEL: 791");
            headers.Add("X-EBAY-API-DEV-NAME: 18b4e21d-d682-4afe-b8d4-bb14c5cf62ad");
            headers.Add("X-EBAY-API-APP-NAME: ToolsPlu-e1bd-4921-ac6b-14497e74488a");
            headers.Add("X-EBAY-API-CERT-NAME: e54d8580-281d-4a70-89ea-884c77e3704c");
            headers.Add("X-EBAY-API-SITEID: 0");
            headers.Add("X-EBAY-API-CALL-NAME: GetCategorySpecifics");

            string results = Connectivity.HTTPCalls.HTTP_POST("https://api.ebay.com/ws/api.dll",damnit(),headers);
            textBox1.Text = results;
        }
        private string damnit()
        {
            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<RequesterCredentials>\r\n <eBayAuthToken>AgAAAA**AQAAAA**aAAAAA**4l7QUg**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6AFkoCgDJOKogudj6x9nY+seQ**+4cBAA**AAMAAA**+9vd97XY1RzNQaC3Ny/35At4dmh29/7mJimJm7iynlxb9R5IPHYGl1/ekf0mmYQ/XivwI8VBQlkPWVFVi8w6AIXNpkZsCBUnsFYSmShD2oeRZO1UpMQnjKCv3VmYXMrb+LO9NZAHh0NwgdC97zVbS3fsJIMUll1O1ct4IcibXDQ+LunjRpiHW2j7kEILWR7j7e2nxLRNG7lSTNTI5WnvsXcs+CgilfGnli358i+KaWNN6lRDPOP8PrHaX1yjhJSQZ3ZC0mjf6oEyYEOl0O9Igmf0U4mfP980iRYWR9kInH8l2wEuKtyrGOx6zX4Pgrw//kWvOgxKng9rzE3X3+UtP4lTGNLGbSfQ0AmL1a+fQ0wmHBYSsi3r1a/PiifWEHZyoySEKw2mUgB9R3dUMQk+zTsRsB4HRDM3y6Ygapa1MAOYjqPTJCYhns+W2MPuXYtDcdKG46exHx2h1XjHOkbWby2QPlrDcOhXMdSPeE6c5y7KB0RPNp5VtHgkxZVsseWUXMZtowJ05Glv+cfhVMveJPgHkbNU7Ql0aW5hgthWTjD1xFTeF3v19BbnwVhvq5aC1uJuYEKc06YEb9qJ/AZ/Cj8X/4arrVzIBvY7ri0famvgvQAXuOo2DLA8ElVyL/c5XdRJrVfBJ58gvY0SCm/pQZdRmsPWnlGRQzJLxfSLwdXdlbjef8DxTu7+8FMB43ScCZGcI1lkgbuhDP4qC1ywZ/SYpO+N0/r81VSQEj8sYCLd1W+S3KQEybOtkKd2XKJa</eBayAuthToken>\r\n</RequesterCredentials>\r\n<GetCategorySpecificsRequest xmlns=\"urn:ebay:apis:eBLBaseComponents\">\r\n<CategoryID>1504</CategoryID>\r\n</GetCategorySpecificsRequest>\r\n";

        }
    }
}
