using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportGenerator
{
    public static class Report_ItemSalesGeo
    {



        public struct ItemSalesGeoOptions
        {
            public string ItemNumber;
            public DateTime StartDate;
            public DateTime EndDate;

        }
        public static string LoadMapBase()
        {
            /*string gAPIKey = "AIzaSyDFZwtjvDola6BpCYe1p6MMiITg7QUY2qw";
            //string gMap = "<html><body><iframe width=\"300\" height=\"300\" frameborder=\"0\" style=\"border:0\" src=\"https://www.google.com/maps/embed/v1/place?key=" + gAPIKey + "&q=Space+Needle,Seattle+WA\"></iframe></body></html>";
            string gMap = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n html, body, #map-canvas {height : 100%; margin: 0; padding 0;}";
            gMap += "\r\n</style><script type=\"text/javascript\" src=\"https://maps.googleapis.com/maps/api/js?key=" + gAPIKey + "\"></script>";
            gMap += "\r\n<script type=\"text/javascript\">\r\n";
            return gMap;*/
            string mapBase = "<!DOCTYPE html>\r\n";
            mapBase += "<html>\r\n";
            mapBase += " <head>\r\n";
            mapBase += "  <meta name=\"viewport\" content=\"initial-scale=1.0, user-scalable=no\">\r\n";
            mapBase += "  <meta charset=\"utf-8\">\r\n";
            mapBase += "  <title>Test</title>\r\n";
            mapBase += "  <style>\r\n";
            mapBase += "   html, body, #map-canvas {\r\n";
            mapBase += "    height: 100%;\r\n";
            mapBase += "    margin: 0px;\r\n";
            mapBase += "    padding: 0px\r\n";
            mapBase += "  }\r\n";
            mapBase += " </style>\r\n";
            mapBase += " <script src=\"https://maps.googleapis.com/maps/api/js?v=3.exp&signed_in=true\"></script>\r\n";
            mapBase += " <script>\r\n";
            return mapBase;
        }
        public static string CreateJavaScriptDataArray(ItemSalesGeoOptions options)
        {
            string StartDate = options.StartDate.Year.ToString("0000") + "-" + options.StartDate.Month.ToString("00") + "-" + options.StartDate.Day.ToString("00");
            string EndDate = options.EndDate.Year.ToString("0000") + "-" + options.EndDate.Month.ToString("00") + "-" + options.EndDate.Day.ToString("00");


            ListView dataLVI = ListViewTools.MAS_to_LVI(Connectivity.MASCalls.Retrieve("SELECT TOP 10 ShipToZipCode,QuantityShipped FROM {OJ SO_SalesOrderHistoryHeader SO_SalesOrderHistoryHeader INNER JOIN SO_SalesOrderHistoryDetail SO_SalesOrderHistoryDetail ON SO_SalesOrderHistoryDetail.SalesOrderNo = SO_SalesOrderHistoryHeader.SalesOrderNo} WHERE OrderDate > { d '" + StartDate + "'} AND OrderDate < { d '" + EndDate + "'} AND ItemCode='" + options.ItemNumber + "'"));

            Dictionary<string, int> ZipSales = new Dictionary<string, int>();

            foreach (ListViewItem lvi in dataLVI.Items)
            {
                if (ZipSales.ContainsKey(lvi.SubItems[0].Text))
                {
                    ZipSales[lvi.SubItems[0].Text] += (int)decimal.Parse(lvi.SubItems[1].Text);
                }
                else
                {
                    ZipSales.Add(lvi.SubItems[0].Text, (int)decimal.Parse(lvi.SubItems[1].Text));
                }
            }

            string output = "var citymap = {};\r\n";
            foreach (KeyValuePair<string,int> entry in ZipSales)
            {
                List<string> location = Connectivity.SQLCalls.sqlQuery("SELECT Latitude,Longitude FROM GeoLocation WHERE ZipCode='" + entry.Key + "'");
                if (location.Count > 0)
                {
                    string Latitude = location[0].Split('|')[0];
                    string Longitude = location[0].Split('|')[1];
                    output += "citymap['" + entry.Key + "'] = {\r\n";
                    output += " center: new google.maps.LatLng(" + Latitude + "," + Longitude + "),\r\n";
                    output += " qtysold: " + entry.Value.ToString() + "\r\n";
                    output += "};\r\n";
                }
                else{
                    //couldnt find.. skip.
                }
            }
            return output;

        }
        public static string InitMap()
        {
            string output = "var cityCircle;\r\n";
            output += "function initialize() {\r\n";
            output += " var mapOptions = {\r\n";
            output += " zoom: 4,\r\n";
            output += " center: new google.maps.LatLng(39.50, -98.35),\r\n";
            output += " mapTypeId: google.maps.MapTypeId.TERRAIN\r\n";
            output += "};\r\n";
            output += "var map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);\r\n";
            return output;
        }



        public static string PopulateMap()
        {
            string output = "for (var city  in citymap) {\r\n";
            output += "var salesOptions = {\r\n";
            output += " strokeColor: '#FF0000',\r\n";
            output += " strokeOpacity: 0.8,\r\n";
            output += " strokeWeight: 2,\r\n";
            output += " fillColor: '#FF0000',\r\n";
            output += " fillOpacity: 0.35,\r\n";
            output += " map: map,\r\n";
            output += " center: citymap[city].center,\r\n";
            output += " radius: Math.sqrt(citymap[city].qtysold) * 10000\r\n";
            output += "};\r\n";
            output += "cityCircle = new google.maps.Circle(salesOptions);\r\n";
            return output;

        }
        public static string BelowCode()
        {
            string belowcode = "\r\n}\r\n}\r\n";
            belowcode += "\r\ngoogle.maps.event.addDomListener(window,'load',initialize);\r\n</script>\r\n</head>\r\n<body>\r\n<div id=\"map-canvas\"></div>\r\n</body>\r\n</html>";
            return belowcode;
        }
        public static string getPageSource(ItemSalesGeoOptions options)
        {
            string finalResult = LoadMapBase();
            finalResult += CreateJavaScriptDataArray(options);
            finalResult += InitMap();
            finalResult += PopulateMap();
            finalResult += BelowCode();
            return finalResult;
        }
    }
}
