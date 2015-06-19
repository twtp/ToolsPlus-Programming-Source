using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace ShippingMonitor
{
    public static class WillCall
    {



        public static void LoadWillCalls(ref Label willCallLbl, ref Timer willCallFlasherTmr)
        {
            SQLCalls newCall = new SQLCalls();
            List<string> willCallList = newCall.sqlQuery("SELECT ID,TransactionID FROM WhseTaskWillCallHeader WHERE TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");

            if (willCallList.Count > 0)
            {
                foreach (string willCall in willCallList)
                {

                    List<string> willCallInfo = newCall.sqlQuery("SELECT ComponentID,QuantityRequired,QuantityPicked FROM WhseTaskWillCallLines WHERE HeaderID=" + willCall.Split('|')[0]);

                    if (willCallInfo.Count > 0)
                    {
                        int totComponentsRequired = 0;
                        int totComponentsPicked = 0;
                        foreach (string str in willCallInfo)
                        {
                            totComponentsRequired += int.Parse(str.Split('|')[1]);
                            totComponentsPicked += int.Parse(str.Split('|')[2]);
                        }

                        if (totComponentsPicked < totComponentsRequired)
                        {
                            willCallLbl.Text = "Will Call " + willCall.Split('|')[1] + "\r\nQty: ";
                            willCallLbl.Text += totComponentsRequired.ToString() + "\r\nPicked: " + totComponentsPicked.ToString();
                            willCallLbl.Visible = true;
                            willCallLbl.BringToFront();
                            willCallFlasherTmr.Enabled = true;
                            //break;
                        }

                    }


                }
            }
            else
            {
                willCallLbl.Visible = false;
                willCallFlasherTmr.Enabled = false;
            }

            if (willCallLbl.Visible == false)
            {
                willCallFlasherTmr.Enabled = false;
            }

        }


        public static void willCallFlasherTmr_Tick(ref Timer willCallFlasherTmr, ref Label willCallLbl)
        {
            //


            if (willCallFlasherTmr.Interval == 3000)
            {
                willCallFlasherTmr.Interval = 300;
                willCallLbl.Visible = true;
            }
            else
            {
                if (willCallFlasherTmr.Interval == 301)
                {
                    willCallFlasherTmr.Interval = 302;
                    willCallLbl.Visible = true;
                }
                else if (willCallFlasherTmr.Interval == 300)
                {
                    willCallFlasherTmr.Interval = 301;
                    willCallLbl.Visible = false;
                }
                else if (willCallFlasherTmr.Interval == 302)
                {
                    willCallFlasherTmr.Interval = 303;
                    willCallLbl.Visible = false;
                }
                else
                {
                    willCallFlasherTmr.Interval = 3000;
                    willCallLbl.Visible = true;
                }
            }


        }


    }
}
