using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator
{
    public static class Report_P2LinecodeSalesInfo
    {
        public struct P2LinecodeSalesInfoSettings
        {
            public List<string> Linecodes;           
            public bool showCustomerNumber;
            public bool showCustomerName;
            public bool showCustomerAddress;
            public bool showCustomerPhone;
            public bool showCustomerEmail;
            public bool showCustomerURL;
            public bool showSalesPersonNo;
            public bool showDateAccountCreated;
            public bool showDateLastPurchased;
            //public bool showItemNumber;
        }

        public struct P2LinecodeSalesInfoResults
        {
            public List<string> Columns;
            public List<string> Rows;
        }


        public static P2LinecodeSalesInfoResults CreateReport(P2LinecodeSalesInfoSettings settings)
        {
            P2LinecodeSalesInfoResults p2Results = new P2LinecodeSalesInfoResults();
            p2Results.Columns = new List<string>();
            p2Results.Rows = new List<string>();

            string LineCodeString = "";
            foreach (string l in settings.Linecodes)
            {
                LineCodeString += " OR ItemCode LIKE '" + l + "%'";
            }
        
            LineCodeString = LineCodeString.TrimStart(' ').TrimStart('O').TrimStart('R');
            LineCodeString = "(" + LineCodeString.Trim() + ")";
            

            string results = Connectivity.MASCalls.Retrieve("SELECT DISTINCT CustomerNo FROM P2_SalesHistory WHERE CustomerNo IS NOT NULL AND CustomerNo<>'' AND CustomerNo<>' 0 00' AND " + LineCodeString);
            if (results.Length > 0)
            {
                if (results.Contains("\"rows\":[]") == false)
                {
                    //get the darn customer numbers and put them neatly in an sql list ==> '0002344','0000233','2349943'   <==...etc...
                    string customers = results.Split(new string[] { "\"rows\":[" }, StringSplitOptions.None)[1].Split(new string[] { "]}" }, StringSplitOptions.None)[0]; ;
                    customers = customers.Replace("\",\"","").Replace("[","").Replace("\n","").Replace("]","\r\n").Replace("\"","'");

                    string SelectStatement = "SELECT ";
                    SelectStatement += settings.showCustomerNumber == true ? "CustomerNo," : "";
                    SelectStatement += settings.showCustomerName == true ? "CustomerName," : "";
                    SelectStatement += settings.showCustomerAddress == true ? "AddressLine1,AddressLine2,AddressLine3,City,State,ZipCode,CountryCode," : "";
                    SelectStatement += settings.showCustomerPhone == true ? "TelephoneNo,TelephoneExt," : "";
                    SelectStatement += settings.showCustomerEmail == true ? "EmailAddress," : "";
                    SelectStatement += settings.showCustomerURL == true ? "URLAddress," : "";
                    SelectStatement += settings.showDateLastPurchased == true ? "DateLastPayment,":"";
                    SelectStatement += settings.showDateAccountCreated == true ? "DateEstablished,":"";
                    SelectStatement += settings.showSalesPersonNo == true ? "SalespersonNo," : "";
                    SelectStatement = SelectStatement.TrimEnd(',');
                    SelectStatement += " ";
                    bool hasColumns = false;
                    List<string> columnResults = new List<string>();
                    List<string> rowResults = new List<string>();
                    foreach (string customer in customers.Split(new string[] { "\r\n" }, StringSplitOptions.None))
                    {
                        string cust = customer.Replace("'","").Replace(",","");
                        string query = SelectStatement + "FROM AR_Customer WHERE CustomerNo = " +"'" + cust + "'";
                        string CustomerResults = Connectivity.MASCalls.Retrieve(query);

                        if (hasColumns == false)
                        {
                            string columnData = CustomerResults.Split(new string[] { "\"rows\":[" }, StringSplitOptions.None)[0];
                            columnData = columnData.Split('[')[1].Split(']')[0];
                            columnData = columnData.TrimStart('{');
                            List<string> columnsRaw = new List<string>(columnData.Split(new string[] { "\"name\":" }, StringSplitOptions.None));

                            foreach (string str in columnsRaw)
                            {
                                string col = str.Split(',')[0].Replace("\"", "");
                                if (col.Length > 0)
                                {
                                    columnResults.Add(col);
                                }
                            }
                            hasColumns = true;
                        }


                        string rowData = CustomerResults.Split(new string[] { "\"rows\":[" }, StringSplitOptions.None)[1];
                        if (rowData.Length > 5)
                        {
                            rowData.TrimEnd('}').TrimEnd(']');
                            int colCount = columnResults.Count;
                            rowData = rowData.Replace("],[", "]\r\n[").Replace("[","").Replace("]","").Replace("\n","").Replace("}","");
                            rowData = rowData.Replace("\",\"", "|").Replace("\"", "").Replace(",null,","||").Replace("null,","|").Replace("null","");
                            rowData += "|";
                            rowResults.Add(rowData);
                        }

                    }

                    p2Results.Columns = columnResults;
                    p2Results.Rows = rowResults;
                    return p2Results;
                     
                }
                else
                {
                    //query sucess; no results.
                    return new P2LinecodeSalesInfoResults();
                }
            }
            else
            {                
                //query failure; check query string.
                return new P2LinecodeSalesInfoResults();
            }


            //now with list of customers... get the rest of the info needed!



            return new P2LinecodeSalesInfoResults();
            

        }

        public static P2LinecodeSalesInfoResults JSONtoResults(string JSON)
        {
            P2LinecodeSalesInfoResults res = new P2LinecodeSalesInfoResults();
            return new P2LinecodeSalesInfoResults();

        }

    }
}
