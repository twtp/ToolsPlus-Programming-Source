using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace EBayUnreferencedTransactionReport
{
    class Program
    {

        public static string FullReport = "SELECT ID,ItemID,TransactionID,Quantity,SalesOrderNo,BuyerID,TimeInserted FROM EBayTransactionLog WHERE SalesOrderNo LIKE 'E-%' ORDER BY TimeInserted DESC";
        public static string TodaysReport = "SELECT ID,ItemID,TransactionID,Quantity,SalesOrderNo,BuyerID,TimeInserted FROM EBayTransactionLog WHERE SalesOrderNo LIKE 'E-%' AND TimeInserted>GETDATE() ORDER BY TimeInserted DESC";
        public static string WeeksReport = "SELECT ID,ItemID,TransactionID,Quantity,SalesOrderNo,BuyerID,TimeInserted FROM EBayTransactionLog WHERE SalesOrderNo LIKE 'E-%' AND TimeInserted>DATEADD(dd,-7,GETDATE()) ORDER BY TimeInserted DESC";
        
        public struct EBayTransactionLog
        {
            public int ID;
            public string ItemID;
            public string TransactionID;
            public int Quantity;
            public string SalesOrderNo;
            public string BuyerID;
            public DateTime TimeInserted;
        }
        public static Dictionary<int, string> ReportBase = new Dictionary<int, string>();

        public static void Main(string[] args)
        {

            ReportBase.Add(0, "Today's Report");
            ReportBase.Add(1, "Weekly Report");
            ReportBase.Add(2, "Full Report");
            int ReportToUse = 0;

            List<string> results = new List<string>();
            try
            {
                if (args[0].Length > 0)
                {
                    try
                    {
                        ReportToUse = int.Parse(args[0]);
                    }
                    catch
                    {
                        Console.WriteLine("Eh, invalid report type. \r\nEmpty or 0 is Today's Report\r\n1 is Weekly Report\r\n2 is Full Report\r\n\r\nDefaulting to Today's Report");
                        //Console.ReadKey();
                    }
                }
            }
            catch
            {
                Console.WriteLine("Eh, invalid report type. \r\nEmpty or 0 is Today's Report\r\n1 is Weekly Report\r\n2 is Full Report\r\n\r\nDefaulting to Today's Report");
                
            }
            //ReportToUse = 2;
            if (ReportToUse == 0)
            {
                results = Connectivity.SQLCalls.sqlQuery(TodaysReport);
            }
            else if (ReportToUse == 1)
            {
                results = Connectivity.SQLCalls.sqlQuery(WeeksReport);
            }
            else if (ReportToUse == 2)
            {                
                results = Connectivity.SQLCalls.sqlQuery(FullReport);
            }
            else
            {
                Console.WriteLine("Eh, invalid report type. \r\nEmpty or 0 is Today's Report\r\n1 is Weekly Report\r\n2 is Full Report\r\n\r\nDefaulting to Today's Report");
                //Console.ReadKey();
            }
            
            
            if (results.Count > 0)
            {
                List<EBayTransactionLog> EBayUnreferencedTransactions = new List<EBayTransactionLog>();
                foreach (string Line in results)
                {
                    EBayTransactionLog tmpLog = new EBayTransactionLog();
                    tmpLog.ID = int.Parse(Line.Split('|')[0]);
                    tmpLog.ItemID = Line.Split('|')[1];
                    tmpLog.TransactionID = Line.Split('|')[2];
                    tmpLog.Quantity = int.Parse(Line.Split('|')[3]);
                    tmpLog.SalesOrderNo = Line.Split('|')[4];
                    tmpLog.BuyerID = Line.Split('|')[5];
                    tmpLog.TimeInserted = DateTime.Parse(Line.Split('|')[6]);
                    EBayUnreferencedTransactions.Add(tmpLog);
                    
                }
                GenerateReport(EBayUnreferencedTransactions,ReportBase[ReportToUse]);
            }
            //Console.ReadKey();
        }
        public static void GenerateReport(List<EBayTransactionLog> Transactions, string ReportType)
        {
            //string EmailOutput = "";
            string html = "<html><body><center><h2>EBay Unreferenced Transactions Report</h2><br><h3><b>" + ReportType + "</b></h3><br><table border=\"1\" style=\"width:95%\"><tr><td><b>DB ID</b></td><td><b>ItemID</b></td><td><b>Transaction ID</b></td><td><b>Quantity</b></td><td><b>Sales Order</b></td><td><b>Buyer ID</b></td><td><b>Time Inserted</b></td></tr>";

            foreach (EBayTransactionLog etl in Transactions)
            {
                html += "<tr><td>" + etl.ID.ToString() + "</td><td>" + etl.ItemID + "</td><td>" + etl.TransactionID + "</td><td>" + etl.Quantity.ToString() + "</td><td>" + etl.SalesOrderNo + "</td><td>" + etl.BuyerID + "</td><td>" + etl.TimeInserted.ToString() + "</td></tr>";    
                //string ItemNumber = "";
                //List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM InventoryComponentMap WHERE ComponentID="+)
                //EmailOutput += etl.TimeInserted.ToString() + "\t" + etl.TransactionID + "\t" + etl.BuyerID + "\t\t" + etl.ItemID + "\t" + etl.Quantity.ToString() + "\r\n\r\n";
            }
            html += "</table></center></body></html>";
            Connectivity.EmailCalls.EmailAddresses = new string[] {"tom@tools-plus.com","eric@tools-plus.com"};
            //Connectivity.EmailCalls.Subject = "EBay Unreferenced Transaction Error";
            Connectivity.EmailCalls.SendEmail2("EBay Unreferenced Transaction Error",html,true);


        }
    }
}
