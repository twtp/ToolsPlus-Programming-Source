using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LivePriceComparer
{
    public partial class Form1
    {
        private struct Competitor
        {
            public string CompetitorName;
            public bool Enabled;
            public bool Blacklisted;
            public bool Whitelisted;
            public bool AddTax;
        }
        private static Dictionary<string, Competitor> CompetitorsData = new Dictionary<string, Competitor>();
        private static Dictionary<string, Competitor> GetCompetitorData()
        {
            Dictionary<string, Competitor> tmpDict = new Dictionary<string, Competitor>();
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT CompetitorName,Enabled,Blacklisted,Whitelisted,AddTax FROM Competitors");
            if (qr.Rows.Count > 0)
            {
                foreach (string row in qr.Rows)
                {
                    Competitor tmpComp = new Competitor();
                    tmpComp.CompetitorName = row.Split('|')[0].ToLower();
                    tmpComp.Enabled = (row.Split('|')[1] == "True" ? true : false);
                    tmpComp.Blacklisted = (row.Split('|')[2] == "True" ? true : false);
                    tmpComp.Whitelisted = (row.Split('|')[3] == "True" ? true : false);
                    tmpComp.AddTax = (row.Split('|')[4] == "True" ? true : false);
                    tmpDict.Add(tmpComp.CompetitorName, tmpComp);
                }
            }
            return tmpDict;
        }




    }
}
