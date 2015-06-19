using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReportGenerator
{
    public static class Report_StarshipShippedPackages
    {
        public struct StarshipShippedPackagesReportSettings
        {
            public DateTime StartDate;
            public DateTime EndDate;
            public int Days;
            public bool UseDateRange;

            public bool showProcessDateColumn;
            public bool showServiceName;
            public bool showDeliveryCharge;
            public bool showBilledWeight;
            public bool showPostalCode;
            public bool showPackageDimensions;
            public bool showDestinationAddress;
            public bool showPackingUser;
            public bool showTrackingNumber;
            public bool showPackageType;
        }

        public struct StarshipShippedPackagesReportLine
        {
            public DateTime ProcessDateTime;
            public String ServiceName;
            public decimal DeliveryCharge;
            public decimal BilledWeight;
            public string PostalCode;
            public decimal PackageLength;
            public decimal PackageWidth;
            public decimal PackageHeight;
            public string DestAddress1;
            public string DestAddress2;
            public string DestAddress3;
            public string DestCity;
            public string DestState;
            public string PackingUser;
            public string TrackingNumber;
            public string PackingType;
        }
        public struct StarShipShippedPackages
        {
            public List<StarshipShippedPackagesReportLine> StarshipShippedPackagesReport;
            public List<string> Columns;
        }

        //private static List<StarshipShippedPackagesReportLine> StarshipShippedPackagesReport = new List<StarshipShippedPackagesReportLine>();

        public static StarShipShippedPackages CreateReport(StarshipShippedPackagesReportSettings settings)
        {
            StarShipShippedPackages tmpReport = new StarShipShippedPackages();
            tmpReport.StarshipShippedPackagesReport = new List<StarshipShippedPackagesReportLine>();
            tmpReport.Columns = new List<string>();

            string SelectColumns = "";
            if (settings.showProcessDateColumn == true)
            {
                SelectColumns += "ProcessDateTime,";
                tmpReport.Columns.Add("Processed Time");
            }
            if (settings.showServiceName == true)
            {
                SelectColumns += "ServiceName,";
                tmpReport.Columns.Add("Service");
            }
            if (settings.showDeliveryCharge == true)
            {
                SelectColumns += "DeliveryCharge,";
                tmpReport.Columns.Add("Delivery Charge");
            }
            if (settings.showBilledWeight == true)
            {
                SelectColumns += "Shipment.BilledWeight,";
                tmpReport.Columns.Add("Billed Weight");
            }
            if (settings.showPostalCode == true)
            {
                SelectColumns += "PostalCode,";
                tmpReport.Columns.Add("Postal Code");
            }
            if (settings.showPackageDimensions == true)
            {
                SelectColumns += "Pack.Length,Pack.Width,Pack.Height,";
                tmpReport.Columns.Add("Length");
                tmpReport.Columns.Add("Width");
                tmpReport.Columns.Add("Height");
            }
            if (settings.showPackageType == true)
            {
                SelectColumns += "PackName,";
                tmpReport.Columns.Add("Package Type");
            }
            if (settings.showDestinationAddress == true)
            {
                SelectColumns += "Address.Address1,Address.Address2,Address.Address3,Address.City,Address.StateProvinceCode,";
                tmpReport.Columns.Add("Address 1");
                tmpReport.Columns.Add("Address 2");
                tmpReport.Columns.Add("Address 3");
                tmpReport.Columns.Add("City");
                tmpReport.Columns.Add("State");
            }
            if (settings.showPackingUser == true)
            {
                SelectColumns += "StarShipUser,";
                tmpReport.Columns.Add("Packing User");
            }
            if (settings.showTrackingNumber == true)
            {
                SelectColumns += "MasterTrackingID,";
                tmpReport.Columns.Add("Tracking Number");
            }

            SelectColumns = SelectColumns.TrimStart(',');
            SelectColumns = SelectColumns.TrimEnd(',');


            string query = "";
            if (settings.UseDateRange == true)
            {
                query = "SELECT " + SelectColumns + " FROM Shipment INNER JOIN ShipCarrier ON ShipCarrier.InternalID=Shipment.ShipCarrierID INNER JOIN Address ON Address.InternalID=Shipment.RecipientAddrID INNER JOIN Pack ON Pack.ShipmentID=Shipment.InternalID WHERE ProcessDateTime > " + settings.EndDate.ToShortDateString() + " AND ProcessDateTime < " + settings.StartDate.ToShortDateString();                
            }
            else
            {
                query = "SELECT " + SelectColumns + " FROM Shipment INNER JOIN ShipCarrier ON ShipCarrier.InternalID=Shipment.ShipCarrierID INNER JOIN Address ON Address.InternalID=Shipment.RecipientAddrID INNER JOIN Pack ON Pack.ShipmentID=Shipment.InternalID WHERE ProcessDateTime > DATEADD(DD,-" + settings.Days + ",GETDATE())";           
            }



            List<string> QueryResults = Connectivity.StarShipCalls.ssQuery(query);
            
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.GridLines = true;
            tmpView.FullRowSelect = true;
            ColumnHeader dateCol = new ColumnHeader();
            ColumnHeader serviceCol = new ColumnHeader();
            ColumnHeader deliveryCol = new ColumnHeader();
            ColumnHeader weightCol = new ColumnHeader();
            ColumnHeader postalCol = new ColumnHeader();
            dateCol.Text = "Date Processed:";
            serviceCol.Text = "Service Name:";
            deliveryCol.Text = "Delivery Charge:";
            weightCol.Text = "Package Weight:";
            postalCol.Text = "Postal Code:";            

            foreach (string result in QueryResults)
            {
                /*ListViewItem tmpItm = new ListViewItem();
                tmpItm.Text = result.Split('|')[0];
                tmpItm.SubItems.Add(result.Split('|')[1]);
                tmpItm.SubItems.Add(result.Split('|')[2]);
                tmpItm.SubItems.Add(result.Split('|')[3]);
                tmpItm.SubItems.Add(result.Split('|')[4]);
                tmpView.Items.Add(tmpItm);*/

                StarshipShippedPackagesReportLine tmp = new StarshipShippedPackagesReportLine();

                int colCount = 0;
                if (settings.showProcessDateColumn == true)
                {
                    tmp.ProcessDateTime = DateTime.Parse(result.Split('|')[colCount]);
                    colCount++;
                }
                if (settings.showServiceName == true)
                {
                    tmp.ServiceName = result.Split('|')[colCount];
                    colCount++;
                }
                if (settings.showDeliveryCharge == true)
                {
                    tmp.DeliveryCharge = decimal.Parse(result.Split('|')[colCount]);
                    colCount++;
                }
                if (settings.showBilledWeight == true)
                {
                    tmp.BilledWeight = decimal.Parse(result.Split('|')[colCount]);
                    colCount++;
                }
                if (settings.showPostalCode==true)
                {
                    tmp.PostalCode = result.Split('|')[colCount];
                    colCount++;
                }
                if (settings.showPackageDimensions == true)
                {
                    tmp.PackageLength = decimal.Parse(result.Split('|')[colCount]);
                    colCount++;
                    tmp.PackageWidth = decimal.Parse(result.Split('|')[colCount]);
                    colCount++;
                    tmp.PackageHeight = decimal.Parse(result.Split('|')[colCount]);
                    colCount++;
                }
                if (settings.showDestinationAddress == true)
                {
                    tmp.DestAddress1 = result.Split('|')[colCount];
                    colCount++;
                    tmp.DestAddress2 = result.Split('|')[colCount];
                    colCount++;
                    tmp.DestAddress3 = result.Split('|')[colCount];
                    colCount++;
                    tmp.DestCity = result.Split('|')[colCount];
                    colCount++;
                    tmp.DestState = result.Split('|')[colCount];
                    colCount++;
                }
                if (settings.showPackingUser == true)
                {
                    tmp.PackingUser = result.Split('|')[colCount];
                    colCount++;                   
                }
                if (settings.showTrackingNumber == true)
                {
                    tmp.TrackingNumber = result.Split('|')[colCount];
                    colCount++;
                }
                if (settings.showPackageType == true)
                {
                    tmp.PackingType = result.Split('|')[colCount];
                }

                tmpReport.StarshipShippedPackagesReport.Add(tmp);
                //StarshipShippedPackagesReport.Add(tmp);
            }
            return tmpReport;

        }


        private static void LVItoCSVFile(string FilePath, ListView ListViewControl)
        {
            string tmpFile = "";
            int Cols = ListViewControl.Columns.Count;

            foreach (ColumnHeader ch in ListViewControl.Columns)
            {
                if (ch.Text.Contains("\"") == true | ch.Text.Contains(",") == true)
                {
                    tmpFile += "\"" + ch.Text + "\",";
                }
                else
                {
                    tmpFile += ch.Text + ",";
                }
            }
            tmpFile += "\r\n";


            foreach (ListViewItem lvi in ListViewControl.Items)
            {
                for (int x = 0; x < Cols; x++)
                {
                    if (lvi.SubItems[x].Text.Contains("\"") == true | lvi.SubItems[x].Text.Contains(",") == true)
                    {
                        tmpFile += "\"" + lvi.SubItems[x].Text + "\",";
                    }
                    else
                    {
                        tmpFile += lvi.SubItems[x].Text + ",";
                    }
                }
                tmpFile += "\r\n";
            }

            System.IO.File.WriteAllText(FilePath, tmpFile);
            //statusLbl.Text = "File saved to \"" + FilePath + "\"!";
        }

    }
}
