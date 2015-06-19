using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SearsConnectorView
{
    public partial class SearsConnectorFrm : Form
    {
        public SearsConnectorFrm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ResetForm();
            exampleLVIClear();
            attributesLVIClear();
            connectorsCmb.Items.Clear();
            List<string> Categories = Connectivity.SQLCalls.sqlQuery("SELECT ID,Name FROM SearsCategories");
            foreach (string cat in Categories)
            {
                connectorsCmb.Items.Add(cat.Split('╝')[1]);
                
            }
        }
        private void ResetForm()
        {
            attributesLVIClear();
            //listBox1.Items.Clear();
            listBox2.Items.Clear();
            attributeNameLbl.Text = "Attribute Name: ";
            attributeTypeLbl.Text = "Attribute Type: ";
            customValueLbl.Text = "Allows Custom Value: ";
            multipleChoiceLbl.Text = "Multiple Choice: ";
            trademarkLbl.Text = "Trademark: ";
            connectorLbl.Text = "Internal Connector: ";
        }
        private void exampleLVIClear()
        {
            exampleLVI.Items.Clear();
            exampleLVI.Columns.Clear();
            exampleLVI.Clear();
            ColumnHeader AttributeCol = new ColumnHeader();
            ColumnHeader ValueCol = new ColumnHeader();

            AttributeCol.Text=  "Attribute";
            AttributeCol.Width = exampleLVI.Width / 3;
            ValueCol.Text = "Value";
            ValueCol.Width = exampleLVI.Width / 2;

            exampleLVI.Columns.Add(AttributeCol);
            exampleLVI.Columns.Add(ValueCol);

        }
        private void attributesLVIClear()
        {
            attributesLVI.Items.Clear();
            attributesLVI.Columns.Clear();
            attributesLVI.Clear();

            ColumnHeader CategoryID = new ColumnHeader();
            ColumnHeader AttributeName = new ColumnHeader();

            CategoryID.Text = "CategoryID";
            CategoryID.Width = attributesLVI.Width / 3;
            AttributeName.Text = "Attribute";
            AttributeName.Width = attributesLVI.Width / 2;

            attributesLVI.Columns.Add(CategoryID);
            attributesLVI.Columns.Add(AttributeName);

        }
        private void connectorsCmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetForm();
            /*string[] commonAttribs = new string[] {"Item Id","Action Flag","Title","Short Description","Manufacturer Model #","Standard Price",
                "Brand Name","Shipping Length","Shipping Width","Shipping Height","Shipping Weight","Product Image URL","Food Item",
                "Requires Refridgeration","Frozen Item","Alcohol","Tobacco","Action Flag","Variation Group ID","Long Description",
                "Seller Categories","UPC","Sale Price","Sale Start Date","Sale End Data","Promotional Text","Shipping Override",
                "Shipping Override Start Date","Shipping Override End Date","Free Shipping Promotional Text","Low Inventory Alert",
                "MAP Price Indicator","Shipping Restrictions","Shipping Cost Ground","Shipping Cost 2 day","Shipping Cost Next Day",
                "Choking Hazard: small parts","Choking Hazard: balloons","Choking Hazard: small ball","Choking Hazard: contains small ball",
                "Choking Hazard: contains marble","Choking Hazard: other","Safety Warning: other","No Warning","Energy Star Compliant",
                "Good Housekeeping","Hazardous Material","California Emissions","Web Exclusive","Mature Content","Swatch Image URL",
                "Feature Image URL #1", "Feature Image URL #2", "Feature Image URL #3", "Feature Image URL #4", "Feature Image URL #5",
                "Feature Image URL #6"};
            listBox1.Items.AddRange(commonAttribs);

            */


            List<string> catID = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID FROM SearsCategories WHERE Name='" + connectorsCmb.Items[connectorsCmb.SelectedIndex].ToString().Replace("'","''") + "'");
            if (catID.Count > 0)
            {
                string CategoryID = catID[0].Split('╝')[0];
                List<string> attributesList = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID,AttributeName,AttributeType FROM SearsCategoriesLines WHERE CategoryID=" + CategoryID + " OR CategoryID=0");

                foreach (string attribute in attributesList)
                {
                    if (checkBox1.Checked == true)
                    {
                        if (attribute.Split('╝')[2].ToUpper().Contains("REQUIRED") == true)
                        {
                            //listBox1.Items.Add(attribute.Split('╝')[0]);
                            ListViewItem newItm = new ListViewItem();
                            newItm.Text = attribute.Split('╝')[0];
                            newItm.SubItems.Add(attribute.Split('╝')[1]);
                            attributesLVI.Items.Add(newItm);
                        }
                        
                    }
                    else
                    {
                        //listBox1.Items.Add(attribute.Split('╝')[0]);
                            ListViewItem newItm = new ListViewItem();
                            newItm.Text = attribute.Split('╝')[0];
                            newItm.SubItems.Add(attribute.Split('╝')[1]);
                            attributesLVI.Items.Add(newItm);
                    }
                }



            }
            else
            {
                MessageBox.Show("Could not find the Category ID from the Category Name. Damn, what a bummer.");
            }

            

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           

        }

        private void selectBtn_Click(object sender, EventArgs e)
        {  
            List<string> ToParseToXML = new List<string>();
            exampleLVIClear();
            foreach (ListViewItem lvi in attributesLVI.Items)
            {
                string CatID = lvi.Text;
                string AttribName = lvi.SubItems[1].Text;               
                string ConnectorName = "";
                bool gotConnector = false;
              
                List<string> Connector = Connectivity.SQLCalls.sqlQuery("SELECT Connector FROM SearsCategoriesLines WHERE CategoryID=" + CatID + " AND AttributeName='" + AttribName + "'");
                if (Connector.Count > 0)
                {
                    ConnectorName = Connector[0].Split('╝')[0];
                    if (ConnectorName.Length > 0)
                    {
                        gotConnector = true;
                        if (gotConnector == true)
                        {

                            if (ConnectorName.ToUpper().Contains("SELECT") == true)
                            {
                                List<string> getSQLValue = Connectivity.SQLCalls.sqlQuery(ConnectorName.Replace("{ITEM_NUMBER}", exampleItemTxt.Text));
                                if (getSQLValue.Count > 0)
                                {
                                    ListViewItem newLVI = new ListViewItem();
                                    string TagName = lvi.SubItems[1].Text;
                                    string TValue = getSQLValue[0].Split('╝')[0];
                                    newLVI.Text = TagName;
                                    newLVI.SubItems.Add(TValue);
                                    ToParseToXML.Add("{" + TagName + "}{" + TValue + "}");                                    
                                    exampleLVI.Items.Add(newLVI);


                                }
                                else
                                {
                                    ListViewItem newLVI = new ListViewItem();
                                    string TagName = lvi.SubItems[1].Text;
                                    string TValue = "No Connector Found";
                                    newLVI.Text = TagName;
                                    newLVI.SubItems.Add(TValue);
                                    ToParseToXML.Add("{" + TagName + "}{" + TValue + "}");
                                    exampleLVI.Items.Add(newLVI);

                                }
                            }
                            else
                            {
                                //use the actual value
                                ListViewItem newLVI = new ListViewItem();
                                string TagName = lvi.SubItems[1].Text;
                                string TValue = ConnectorName;
                                newLVI.Text = TagName;
                                newLVI.SubItems.Add(TValue);
                                ToParseToXML.Add("{" + TagName + "}{" + TValue + "}");
                                exampleLVI.Items.Add(newLVI);
                            }

                           



                        }
                        else
                        {
                            ListViewItem newLVI = new ListViewItem();
                            string TagName = lvi.SubItems[1].Text;
                            string TValue = "Connector Param Empty";
                            newLVI.Text = TagName;
                            newLVI.SubItems.Add(TValue);
                            ToParseToXML.Add("{" + TagName + "}{" + TValue + "}");
                            exampleLVI.Items.Add(newLVI);

                        }

                    }
                    else
                    {
                        gotConnector = false;
                        ListViewItem newLVI = new ListViewItem();
                        string TagName = lvi.SubItems[1].Text;
                        string TValue = "Couldn't Find Connector";
                        newLVI.Text = TagName;
                        newLVI.SubItems.Add(TValue);
                        ToParseToXML.Add("{" + TagName + "}{" + TValue + "}");
                        exampleLVI.Items.Add(newLVI);
                      
                    }
                }
                else
                {
                    //cannot find attribute...
                    gotConnector = false;
                    ListViewItem newLVI = new ListViewItem();
                    string TagName = lvi.SubItems[1].Text;
                    string TValue = "Couldn't Find Attribute";
                    newLVI.Text = TagName;
                    newLVI.SubItems.Add(TValue);
                    exampleLVI.Items.Add(newLVI);
                 
                }

            



            }

            List<SearsOffload.SearsOffloadItemData> OffloadData = SearsOffload.CreateSearsItemData(ToParseToXML);
            string xml = SearsOffload.GetXMLHeader() + SearsOffload.AddItemsToXMLSheet(OffloadData) + SearsOffload.GetXMLFooter();
            //MessageBox.Show(xml);
            System.IO.File.WriteAllText(@"C:\Users\tomwestbrook\documents\WiserFeed\xmlTmp.xml", xml);
            webBrowser1.Navigate(@"C:\Users\tomwestbrook\documents\WiserFeed\xmlTmp.xml");
        }

        private void attributesLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            attributeNameLbl.Text = "Attribute Name: ";
            attributeTypeLbl.Text = "Attribute Type: ";
            customValueLbl.Text = "Allows Custom Value: ";
            multipleChoiceLbl.Text = "Multiple Choice: ";
            trademarkLbl.Text = "Trademark: ";
            connectorLbl.Text = "Internal Connector: ";
            internalConnectorTxt.Text = "";
            if (attributesLVI.SelectedIndices.Count > 0)
            {
                string attribute = attributesLVI.Items[attributesLVI.SelectedIndices[0]].SubItems[1].Text;
                string catID = attributesLVI.Items[attributesLVI.SelectedIndices[0]].Text;
                //List<string> attributeDetails = Connectivity.SQLCalls.sqlQuery("SELECT AttributeName,AttributeType,AcceptsCustomValue,MultipleChoice,Trademark,AcceptedValues,Connector FROM SearsCategoriesLines WHERE AttributeName='" + listBox1.Items[listBox1.SelectedIndex] + "'");
                List<string> attributeDetails = Connectivity.SQLCalls.sqlQuery("SELECT AttributeName,AttributeType,AcceptsCustomValue,MultipleChoice,Trademark,AcceptedValues,Connector FROM SearsCategoriesLines WHERE AttributeName='" + attribute + "' AND CategoryID=" + catID);

                if (attributeDetails.Count > 0)
                {
                    attributeNameLbl.Text += attributeDetails[0].Split('╝')[0];
                    attributeTypeLbl.Text += attributeDetails[0].Split('╝')[1];
                    customValueLbl.Text += attributeDetails[0].Split('╝')[2].ToUpper() == "TRUE" ? "YES." : "NO.";
                    multipleChoiceLbl.Text += attributeDetails[0].Split('╝')[3].ToUpper() == "TRUE" ? "YES." : "NO.";
                    trademarkLbl.Text += attributeDetails[0].Split('╝')[4];
                    internalConnectorTxt.Text += attributeDetails[0].Split('╝')[6];
                    foreach (string values in attributeDetails[0].Split('╝')[5].Split(','))
                    {
                        listBox2.Items.Add(values.Trim());
                    }




                }
                else
                {
                    MessageBox.Show("Could not find the attribute information from the attribute name. Damnit, another freakin bummer!");
                }
            }

        }

        private void exampleItemTxt_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
