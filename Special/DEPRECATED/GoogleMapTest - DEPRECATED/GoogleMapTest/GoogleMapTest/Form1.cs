using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GoogleMapTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string gAPIKey = "AIzaSyDFZwtjvDola6BpCYe1p6MMiITg7QUY2qw";
            //string gMap = "<html><body><iframe width=\"300\" height=\"300\" frameborder=\"0\" style=\"border:0\" src=\"https://www.google.com/maps/embed/v1/place?key=" + gAPIKey + "&q=Space+Needle,Seattle+WA\"></iframe></body></html>";
            string gMap = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n html, body, #map-canvas {height : 100%; margin: 0; padding 0;}";
            gMap += "\r\n</style><script type=\"text/javascript\" src=\"https://maps.googleapis.com/maps/api/js?key=" + gAPIKey + "\"></script>";
            gMap += "\r\n<script type=\"text/javascript\">\r\nfunction initialize() {\r\nvar mapOptions = {\r\ncenter: { lat: 41.49, lng: -72.97},\r\n zoom: 8 \r\n};"; //41.499444, -72.975278
            gMap += "\r\nvar map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);\r\n}\r\n";
            gMap += "\r\ngoogle.maps.event.addDomListener(window,'load',initialize);\r\n</script>\r\n</head>\r\n<body>\r\n<div id=\"map-canvas\"></div>\r\n</body>\r\n</html>";
            
            textBox1.Text = gMap;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = webBrowser1.DocumentText;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            webBrowser1.DocumentText = textBox1.Text;
        }
    }
}
