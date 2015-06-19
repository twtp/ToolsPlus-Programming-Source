using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EBayLimitsInfo
{
    public partial class Form1 : Form
    {
        public const string VERSION = "1.2";

        public struct EBayAPI_AccessRule
        {
            public string APICallName;
            public bool APICountsTowardsAggregate;
            public long APIDailyHardLimit;
            public long APIDailySoftLimit;
            public long APIDailyUsage;
            public long APIHourlyHardLimit;
            public long APIHourlySoftLimit;
            public long APIHourlyUsage;
            public int APIPeriod;
            public long APIPeriodHardLimit;
            public long APIPeriodSoftLimit;
            public long APIPeriodUsage;
            public DateTime APILastModifiedTime;
            public string APIRuleCurrentStatus;
            public string APIRuleStatus;

        }
        public struct EBayAPI_GetApiAccessRulesResponse
        {
            public string EBayXMLVersion;
            public string EBayEncoding;
            public DateTime EBayTimeStamp;
            public string EBayAcknowledgement;
            public string EBayAPIVersion;
            public string EBayAPIBuild;
            public List<EBayAPI_AccessRule> AccessRules;


            public void EBayCreateAccessResponse(string response)
            {
               
                    AccessRules = new List<EBayAPI_AccessRule>();
                    EBayXMLVersion = response.Split(new string[] { "<?xml version=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    EBayEncoding = response.Split(new string[] { " encoding=\"" }, StringSplitOptions.None)[1].Split('"')[0];
                    EBayTimeStamp = DateTime.Parse(response.Split(new string[] { "<Timestamp>" }, StringSplitOptions.None)[1].Split(new string[] { "</Timestamp>" }, StringSplitOptions.None)[0]);
                    EBayAcknowledgement = response.Split(new string[] { "<Ack>" }, StringSplitOptions.None)[1].Split(new string[] { "</Ack>" }, StringSplitOptions.None)[0];

                    if (EBayAcknowledgement.ToUpper() == "SUCCESS")
                    {
                        EBayAPIVersion = response.Split(new string[] { "<Version>" }, StringSplitOptions.None)[1].Split(new string[] { "</Version>" }, StringSplitOptions.None)[0];
                        EBayAPIBuild = response.Split(new string[] { "<Build>" }, StringSplitOptions.None)[1].Split(new string[] { "</Build>" }, StringSplitOptions.None)[0];


                        foreach (string AccessRule in response.Split(new string[] { "<ApiAccessRule>" }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            if (AccessRule.Contains("<CallName>") == true)
                            {
                                string tmpAccessRule = AccessRule.Split(new string[] { "</ApiAccessRule>" }, StringSplitOptions.None)[0];

                                EBayAPI_AccessRule newRules = new EBayAPI_AccessRule();
                                newRules.APICallName = tmpAccessRule.Split(new string[] { "<CallName>" }, StringSplitOptions.None)[1].Split(new string[] { "</CallName>" }, StringSplitOptions.None)[0];
                                newRules.APICountsTowardsAggregate = Convert.ToBoolean(tmpAccessRule.Split(new string[] { "<CountsTowardAggregate>" }, StringSplitOptions.None)[1].Split(new string[] { "</CountsTowardAggregate>" }, StringSplitOptions.None)[0]);
                                newRules.APIDailyHardLimit = long.Parse(tmpAccessRule.Split(new string[] { "<DailyHardLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</DailyHardLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIDailySoftLimit = long.Parse(tmpAccessRule.Split(new string[] { "<DailySoftLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</DailySoftLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIDailyUsage = long.Parse(tmpAccessRule.Split(new string[] { "<DailyUsage>" }, StringSplitOptions.None)[1].Split(new string[] { "</DailyUsage>" }, StringSplitOptions.None)[0]);
                                newRules.APIHourlyHardLimit = long.Parse(tmpAccessRule.Split(new string[] { "<HourlyHardLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</HourlyHardLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIHourlySoftLimit = long.Parse(tmpAccessRule.Split(new string[] { "<HourlySoftLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</HourlySoftLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIHourlyUsage = long.Parse(tmpAccessRule.Split(new string[] { "<HourlyUsage>" }, StringSplitOptions.None)[1].Split(new string[] { "</HourlyUsage>" }, StringSplitOptions.None)[0]);
                                newRules.APIPeriod = int.Parse(tmpAccessRule.Split(new string[] { "<Period>" }, StringSplitOptions.None)[1].Split(new string[] { "</Period>" }, StringSplitOptions.None)[0]);
                                newRules.APIPeriodHardLimit = long.Parse(tmpAccessRule.Split(new string[] { "<PeriodicHardLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</PeriodicHardLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIPeriodSoftLimit = long.Parse(tmpAccessRule.Split(new string[] { "<PeriodicSoftLimit>" }, StringSplitOptions.None)[1].Split(new string[] { "</PeriodicSoftLimit>" }, StringSplitOptions.None)[0]);
                                newRules.APIPeriodUsage = long.Parse(tmpAccessRule.Split(new string[] { "<PeriodicUsage>" }, StringSplitOptions.None)[1].Split(new string[] { "</PeriodicUsage>" }, StringSplitOptions.None)[0]);
                                newRules.APILastModifiedTime = DateTime.Parse(tmpAccessRule.Split(new string[] { "<ModTime>" }, StringSplitOptions.None)[1].Split(new string[] { "</ModTime>" }, StringSplitOptions.None)[0]);
                                newRules.APIRuleCurrentStatus = tmpAccessRule.Split(new string[] { "<RuleCurrentStatus>" }, StringSplitOptions.None)[1].Split(new string[] { "</RuleCurrentStatus>" }, StringSplitOptions.None)[0];
                                newRules.APIRuleStatus = tmpAccessRule.Split(new string[] { "<RuleStatus>" }, StringSplitOptions.None)[1].Split(new string[] { "</RuleStatus>" }, StringSplitOptions.None)[0];

                                AccessRules.Add(newRules);
                            }                         

                        }
                        

                    }
                    else
                    {
                        MessageBox.Show("EBay returned an error. We should probably parse it...");
                    }


                

            }

        }


        public const string EBayAPIURL = "https://api.ebay.com/ws/api.dll";
        public const string ContentType = "application/x-www-form-urlencoded";
        public const string UserAgent = "ToolsPlus EBay Interface v0.1";
        public const string X_EBAY_API_COMPATIBILITY_LEVEL = "791";
        public const string X_EBAY_API_DEV_NAME = "18b4e21d-d682-4afe-b8d4-bb14c5cf62ad";
        public const string X_EBAY_API_APP_NAME = "ToolsPlu-e1bd-4921-ac6b-14497e74488a";
        public const string X_EBAY_API_CERT_NAME = "e54d8580-281d-4a70-89ea-884c77e3704c";
        public const string X_EBAY_API_SITEID = "0";
        public const string X_EBAY_API_CALL_NAME = "GetApiAccessRules";
        public const string Auth_Token = "AgAAAA**AQAAAA**aAAAAA**4l7QUg**nY+sHZ2PrBmdj6wVnY+sEZ2PrA2dj6AFkoCgDJOKogudj6x9nY+seQ**+4cBAA**AAMAAA**+9vd97XY1RzNQaC3Ny/35At4dmh29/7mJimJm7iynlxb9R5IPHYGl1/ekf0mmYQ/XivwI8VBQlkPWVFVi8w6AIXNpkZsCBUnsFYSmShD2oeRZO1UpMQnjKCv3VmYXMrb+LO9NZAHh0NwgdC97zVbS3fsJIMUll1O1ct4IcibXDQ+LunjRpiHW2j7kEILWR7j7e2nxLRNG7lSTNTI5WnvsXcs+CgilfGnli358i+KaWNN6lRDPOP8PrHaX1yjhJSQZ3ZC0mjf6oEyYEOl0O9Igmf0U4mfP980iRYWR9kInH8l2wEuKtyrGOx6zX4Pgrw//kWvOgxKng9rzE3X3+UtP4lTGNLGbSfQ0AmL1a+fQ0wmHBYSsi3r1a/PiifWEHZyoySEKw2mUgB9R3dUMQk+zTsRsB4HRDM3y6Ygapa1MAOYjqPTJCYhns+W2MPuXYtDcdKG46exHx2h1XjHOkbWby2QPlrDcOhXMdSPeE6c5y7KB0RPNp5VtHgkxZVsseWUXMZtowJ05Glv+cfhVMveJPgHkbNU7Ql0aW5hgthWTjD1xFTeF3v19BbnwVhvq5aC1uJuYEKc06YEb9qJ/AZ/Cj8X/4arrVzIBvY7ri0famvgvQAXuOo2DLA8ElVyL/c5XdRJrVfBJ58gvY0SCm/pQZdRmsPWnlGRQzJLxfSLwdXdlbjef8DxTu7+8FMB43ScCZGcI1lkgbuhDP4qC1ywZ/SYpO+N0/r81VSQEj8sYCLd1W+S3KQEybOtkKd2XKJa";


        public const string XML_GetApiAccessRulesRequest = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n" +
            
            "<GetApiAccessRulesRequest xmlns=\"urn:ebay:apis:eBLBaseComponents\">\r\n" +
            "<RequesterCredentials>\r\n" +
            " <eBayAuthToken>" + Auth_Token + "</eBayAuthToken>\r\n" +
            "</RequesterCredentials>\r\n" +
            " <ErrorLanguage>en_US</ErrorLanguage>\r\n" +
            "</GetApiAccessRulesRequest>\r\n";


        public EBayAPI_GetApiAccessRulesResponse APIAccessResponse = new EBayAPI_GetApiAccessRulesResponse();



        public void SetupListView()
        {

            resultsLVI.Items.Clear();
            resultsLVI.Clear();
            resultsLVI.Refresh();

            ColumnHeader CallName = new ColumnHeader();
            ColumnHeader CountsTowardsAggregate = new ColumnHeader();
            ColumnHeader DailyHardLimit = new ColumnHeader();
            ColumnHeader DailySoftLimit = new ColumnHeader();
            ColumnHeader DailyUsage = new ColumnHeader();
            ColumnHeader HourlyHardLimit= new ColumnHeader();
            ColumnHeader HourlySoftLimit = new ColumnHeader();
            ColumnHeader HourlyUsage = new ColumnHeader();
            ColumnHeader Period = new ColumnHeader();
            ColumnHeader PeriodicHardLimit = new ColumnHeader();
            ColumnHeader PeriodicSoftLimit = new ColumnHeader();
            ColumnHeader PeriodicUsage = new ColumnHeader();
            ColumnHeader ModifiedTime = new ColumnHeader();
            ColumnHeader RuleCurrentStatus = new ColumnHeader();
            ColumnHeader RuleStatus = new ColumnHeader();


            CallName.Text = "API Function";
            CountsTowardsAggregate.Text = "Is Aggregate";
            DailyHardLimit.Text = "Daily Hard Limit";
            DailySoftLimit.Text = "Daily Soft Limit";
            DailyUsage.Text = "Daily Usage";
            HourlyHardLimit.Text = "Hourly Hard Limit";
            HourlySoftLimit.Text = "Hourly Soft Limit";
            HourlyUsage.Text = "Hourly Usage";
            Period.Text = "Period";
            PeriodicHardLimit.Text = "Period Hard Limit";
            PeriodicSoftLimit.Text = "Period Soft Limit";
            PeriodicUsage.Text = "Period Usage";
            ModifiedTime.Text = "Last Modified";
            RuleCurrentStatus.Text = "Current Status";
            RuleStatus.Text = "Status";

            resultsLVI.Columns.Add(CallName);
            resultsLVI.Columns.Add(CountsTowardsAggregate);
            resultsLVI.Columns.Add(DailyHardLimit);
            resultsLVI.Columns.Add(DailySoftLimit);
            resultsLVI.Columns.Add(DailyUsage);
            resultsLVI.Columns.Add(HourlyHardLimit);
            resultsLVI.Columns.Add(HourlySoftLimit);
            resultsLVI.Columns.Add(HourlyUsage);
            resultsLVI.Columns.Add(Period);
            resultsLVI.Columns.Add(PeriodicHardLimit);
            resultsLVI.Columns.Add(PeriodicSoftLimit);
            resultsLVI.Columns.Add(PeriodicUsage);
            resultsLVI.Columns.Add(ModifiedTime);
            resultsLVI.Columns.Add(RuleCurrentStatus);
            resultsLVI.Columns.Add(RuleStatus);

            resultsLVI.Refresh();


        }

        public void PopulateInformation()
        {
            SetupListView();
            xmlVersionLbl.Text = APIAccessResponse.EBayXMLVersion;
            encodingLbl.Text = APIAccessResponse.EBayEncoding;
            timeStampLbl.Text = APIAccessResponse.EBayTimeStamp.ToString();
            acknowledgeLbl.Text = APIAccessResponse.EBayAcknowledgement;
            apiVersionLbl.Text = APIAccessResponse.EBayAPIVersion;
            apiBuildLbl.Text = APIAccessResponse.EBayAPIBuild;

            

            foreach(EBayAPI_AccessRule ebayAR in APIAccessResponse.AccessRules)
            {
                ListViewItem tmpLVI = new ListViewItem();
                tmpLVI.Text = ebayAR.APICallName;
                tmpLVI.SubItems.Add(ebayAR.APICountsTowardsAggregate.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIDailyHardLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIDailySoftLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIDailyUsage.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIHourlyHardLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIHourlySoftLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIHourlyUsage.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIPeriod.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIPeriodHardLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIPeriodSoftLimit.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIPeriodUsage.ToString());
                tmpLVI.SubItems.Add(ebayAR.APILastModifiedTime.ToString());
                tmpLVI.SubItems.Add(ebayAR.APIRuleCurrentStatus);
                tmpLVI.SubItems.Add(ebayAR.APIRuleStatus);
                resultsLVI.Items.Add(tmpLVI);
            }

        }


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            versionLbl.Text = "Version: " + VERSION;

            SetupListView();
            URLLbl.Text = EBayAPIURL;
            contentTypeLbl.Text = ContentType;
            userAgentLbl.Text = UserAgent;
            compatibilityLevelLbl.Text = X_EBAY_API_COMPATIBILITY_LEVEL;
            devNameLbl.Text = X_EBAY_API_DEV_NAME;
            appNameLbl.Text = X_EBAY_API_APP_NAME;
            certNameLbl.Text = X_EBAY_API_CERT_NAME;
            siteIDLbl.Text = X_EBAY_API_SITEID;
            callNameLbl.Text = X_EBAY_API_CALL_NAME;
            authTokenLbl.Text = Auth_Token;

        }


        public string HTTP_POST()
        {
            string Url = EBayAPIURL;
            string Data = XML_GetApiAccessRulesRequest;

            string Out = String.Empty;
            System.Net.WebRequest req = System.Net.WebRequest.Create(Url);

            try
            {
                req.Method = "POST";
                req.Timeout = 100000;
                req.ContentType = ContentType;
                req.Headers.Add("UserAgent: " + UserAgent + "\n");
                req.Headers.Add("X-EBAY-API-COMPATIBILITY-LEVEL: " + X_EBAY_API_COMPATIBILITY_LEVEL + "\n");
                req.Headers.Add("X-EBAY-API-DEV-NAME: " + X_EBAY_API_DEV_NAME + "\n");
                req.Headers.Add("X-EBAY-API-APP-NAME: " + X_EBAY_API_APP_NAME + "\n");
                req.Headers.Add("X-EBAY-API-CERT-NAME: " + X_EBAY_API_CERT_NAME + "\n");
                req.Headers.Add("X-EBAY-API-SITEID: " + X_EBAY_API_SITEID + "\n");
                req.Headers.Add("X-EBAY-API-CALL-NAME: " + X_EBAY_API_CALL_NAME + "\n");


                byte[] sentData = Encoding.UTF8.GetBytes(Data);
                req.ContentLength = sentData.Length;
                using (System.IO.Stream sendStream = req.GetRequestStream())
                {
                    sendStream.Write(sentData, 0, sentData.Length);
                    sendStream.Close();
                }
                System.Net.WebResponse res = req.GetResponse();
                System.IO.Stream ReceiveStream = res.GetResponseStream();
                string str = "";
                using (System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8))
                {
                    Char[] read = new Char[256];
                    int count = sr.Read(read, 0, 256);

                    while (count > 0)
                    {
                        str = new String(read, 0, count);
                        Out += str;
                        count = sr.Read(read, 0, 256);
                    }
                }
                //MessageBox.Show("The Server Responded With: " + Out);

                webBrowser1.DocumentText = Out;

                APIAccessResponse.EBayCreateAccessResponse(Out);
                PopulateInformation();


            }
            catch (ArgumentException ex)
            {
                Out = string.Format("HTTP_ERROR :: The second HttpWebRequest object has raised an Argument Exception as 'Connection' Property is set to 'Close' :: {0}", ex.Message);
            }
            catch (Exception ex)
            {
                Out = string.Format("HTTP_ERROR :: WebException raised! :: {0}", ex.Message);
            }
            catch
            {
                Out = string.Format("DAMNIT ERROR");
            }

            return Out;
        }

        


        private void button1_Click(object sender, EventArgs e)
        {
            HTTP_POST();
            //APIAccessResponse.EBayCreateAccessResponse(textBox1.Text);
            //PopulateInformation();
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void resultsLVI_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void viewBtn_Click(object sender, EventArgs e)
        {
            if (viewBtn.Text == "Raw Data")
            {
                viewBtn.Text = "Results";
                webBrowser1.Top = resultsLVI.Top;
                webBrowser1.Left = resultsLVI.Left;
                webBrowser1.Width = resultsLVI.Width;
                webBrowser1.Height = resultsLVI.Height;
                webBrowser1.Visible = true;
                webBrowser1.BringToFront();
            }
            else
            {
                viewBtn.Text = "Raw Data";
                webBrowser1.Visible = false;
                resultsLVI.BringToFront();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.BringToFront();
            this.Focus();
            SendKeys.Send("%{PRTSC}");
        }


    }
}
