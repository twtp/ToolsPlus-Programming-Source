using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace ProjectTrueValue
{
    public partial class debugFrm : Form
    {
        public int StartIndex = 0;
        public bool OffsetStart = true;
        public int TotalSeconds = 120;
        public int CurrentSeconds = 0;
        public int BatchSize = 50;
        public struct TVItemImportData
        {
            public int ItemID;
            public string ItemNumber;
            public string ProductLine;
            public string PrimaryVendor;
            public string WhseCode;
            public string Component;
            public int ComponentID;
            public string Desc2;
            public int AvailabilityLimit;
            public long EBayStoreCategoryID;
            public int ItemStatus;
            public int WebsiteID;
        }
        public Dictionary<string, TVItemImportData> TVItems = new Dictionary<string, TVItemImportData>();

        public struct TVCatalogItem
        {
            public string item_nbr;
            public decimal srp_cost;
            public decimal member_cost;
            public decimal ds_cost;
            public string short_description;
            public string vendor_id;
            public string dtp_code;
            public string class_code;
            public string subclass_code;
            public string vendor_name;
            public string upc;
            public string long_description;
            public decimal weight;
            public decimal length;
            public decimal width;
            public decimal height;
            public decimal pack_weight;
            public decimal pack_length;
            public decimal pack_width;
            public decimal pack_height;
            public int retail_pack_qty;
            public int member_pack_qty;
            public string member_pack_type;
            public string member_break_pack;
            public string model;
            public string item_picture_id;
            public string country_code;
            public string to_be_discontinued;
            public string retail_uom;
            public string Edit_Divisor;
            public string Exclusive_Brand_Code;
        }

        public struct TVInventoryItem
        {
            public string rdc_nbr;
            public string item_nbr;
            public int Inventory;
        }


        public debugFrm()
        {
            InitializeComponent();
        }



        // Code for Catalog
        private void loadCatelogBtn_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                catalogPathLbl.Text = openFileDialog1.FileName;
                processCatalogBtn.Enabled = true;
            }


        }
        private void processCatalogBtn_Click(object sender, EventArgs e)
        {
            SetupCatalogView();

            System.Threading.Thread t = new System.Threading.Thread(ProcessCatalogXML);
            t.Start();
            
            
        }
        private void ProcessCatalogXML()
        {
            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnode;
            int i = 0;
            string str = null;
            FileStream fs = new FileStream(catalogPathLbl.Text, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);
            xmlnode = xmldoc.GetElementsByTagName("Item");
            //MessageBox.Show("Node Count: " + xmlnode.Count.ToString());
            for (i = 0; i <= xmlnode.Count - 1; i++)
            {
                TVCatalogItem tvc = new TVCatalogItem();
                ListViewItem lvi = new ListViewItem();
                lvi.Text = i.ToString();

                foreach (XmlAttribute xmlattr in xmlnode[i].Attributes)
                {
                    switch(xmlattr.Name)
                    {
                        case "item_nbr":
                            tvc.item_nbr = xmlattr.Value;
                            break;
                        case "srp_cost":
                            tvc.srp_cost = decimal.Parse(xmlattr.Value);
                            break;
                        case "member_cost":
                            tvc.member_cost = decimal.Parse(xmlattr.Value);
                            break;
                        case "ds_cost":
                            tvc.ds_cost = decimal.Parse(xmlattr.Value);
                            break;
                        case "short_description":
                            tvc.short_description = xmlattr.Value;
                            break;
                        case "vendor_id":
                            tvc.vendor_id = xmlattr.Value;
                            break;
                        case "dtp_code":
                            tvc.dtp_code = xmlattr.Value;
                            break;
                        case "class_code":
                            tvc.class_code = xmlattr.Value;
                            break;
                        case "subclass_code":
                            tvc.subclass_code = xmlattr.Value;
                            break;
                        case "vendor_name":
                            tvc.vendor_name = xmlattr.Value;
                            break;
                        case "upc":
                            tvc.upc = xmlattr.Value;
                            break;
                        case "long_description":
                            tvc.long_description = xmlattr.Value;
                            break;
                        case "weight":
                            tvc.weight = decimal.Parse(xmlattr.Value);
                            break;
                        case "length":
                            tvc.length = decimal.Parse(xmlattr.Value);
                            break;
                        case "width":
                            tvc.width = decimal.Parse(xmlattr.Value);
                            break;
                        case "height":
                            tvc.height = decimal.Parse(xmlattr.Value);
                            break;
                        case "pack_weight":
                            tvc.pack_weight = decimal.Parse(xmlattr.Value);
                            break;
                        case "pack_length":
                            tvc.pack_length = decimal.Parse(xmlattr.Value);
                            break;
                        case "pack_width":
                            tvc.pack_width = decimal.Parse(xmlattr.Value);
                            break;
                        case "pack_height":
                            tvc.pack_height = decimal.Parse(xmlattr.Value);
                            break;
                        case "retail_pack_qty":
                            tvc.retail_pack_qty = int.Parse(xmlattr.Value);
                            break;
                        case "member_pack_qty":
                            tvc.member_pack_qty = int.Parse(xmlattr.Value);
                            break;
                        case "member_pack_type":
                            tvc.member_pack_type = xmlattr.Value;
                            break;
                        case "member_break_pack":
                            tvc.member_break_pack = xmlattr.Value;
                            break;
                        case "model":
                            tvc.model = xmlattr.Value;
                            break;
                        case "item_picture_id":
                            tvc.item_picture_id = xmlattr.Value;
                            break;
                        case "country_code":
                            tvc.country_code = xmlattr.Value;
                            break;
                        case "to_be_discontinued":
                            tvc.to_be_discontinued = xmlattr.Value;
                            break;
                        case "Edit_Divisor":
                            tvc.Edit_Divisor = xmlattr.Value;
                            break;
                        case "Exclusive_Brand_Code":
                            tvc.Exclusive_Brand_Code = xmlattr.Value;
                            break;

                    }
                    //lvi.SubItems.Add(xmlattr.Value);
                }

                string sqlInsert = "INSERT INTO TrueValueCatalog (item_nbr,srp_cost,member_cost,ds_cost,short_description,vendor_id,dtp_code,class_code,subclass_code,vendor_name,upc,long_description,weight,length,width,height,pack_weight,pack_length,pack_width,pack_height,retail_pack_qty,member_pack_qty,member_pack_type,member_break_pack,model,item_picture_id,country_code,to_be_discontinued,retail_uom,Edit_Divisor,Exclusive_Brand_Code) VALUES('{item_nbr}',{srp_cost},{member_cost},{ds_cost},'{short_description}',{vendor_id},'{dtp_code}',{class_code},{subclass_code},'{vendor_name}','{upc}','{long_description}',{weight},{length},{width},{height},{pack_weight},{pack_length},{pack_width},{pack_height},{retail_pack_qty},{member_pack_qty},'{member_pack_type}','{member_break_pack}','{model}',{item_picture_id},'{country_code}','{to_be_discontinued}','{retail_uom}','{Edit_Divisor}','{Exclusive_Brand_Code}')";



                for (int myI = 0; myI <= 30; myI++)
                {
                    switch (myI)
                    {
                        case 0:
                            lvi.SubItems.Add((tvc.item_nbr == null ? "" : tvc.item_nbr));
                            sqlInsert = sqlInsert.Replace("{item_nbr}", (tvc.item_nbr == null ? "" : tvc.item_nbr).Replace("'","''"));
                            break;
                        case 1:
                            lvi.SubItems.Add((tvc.srp_cost == null ? "" : tvc.srp_cost.ToString()));
                            sqlInsert = sqlInsert.Replace("{srp_cost}", (tvc.srp_cost == null ? "" : tvc.srp_cost.ToString()).Replace("'", "''"));
                            break;
                        case 2:
                            lvi.SubItems.Add((tvc.member_cost == null ? "" : tvc.member_cost.ToString()));
                            sqlInsert = sqlInsert.Replace("{member_cost}", (tvc.member_cost == null ? "" : tvc.member_cost.ToString()).Replace("'", "''"));
                            break;
                        case 3:
                            lvi.SubItems.Add((tvc.ds_cost == null ? "" : tvc.ds_cost.ToString()));
                            sqlInsert = sqlInsert.Replace("{ds_cost}", (tvc.ds_cost == null ? "" : tvc.ds_cost.ToString()).Replace("'", "''"));
                            break;
                        case 4:
                            lvi.SubItems.Add((tvc.short_description == null ? "" : tvc.short_description));
                            sqlInsert = sqlInsert.Replace("{short_description}", (tvc.short_description == null ? "" : tvc.short_description).Replace("'", "''"));
                            break;
                        case 5:
                            lvi.SubItems.Add((tvc.vendor_id == null ? "" : tvc.vendor_id));
                            sqlInsert = sqlInsert.Replace("{vendor_id}", (tvc.vendor_id == null ? "" : tvc.vendor_id).Replace("'", "''"));
                            break;
                        case 6:
                            lvi.SubItems.Add((tvc.dtp_code == null ? "" : tvc.dtp_code));
                            sqlInsert = sqlInsert.Replace("{dtp_code}", (tvc.dtp_code == null ? "" : tvc.dtp_code).Replace("'", "''"));
                            break;
                        case 7:
                            lvi.SubItems.Add((tvc.class_code == null ? "" : tvc.class_code));
                            sqlInsert = sqlInsert.Replace("{class_code}", (tvc.class_code == null ? "" : tvc.class_code).Replace("'", "''"));
                            break;
                        case 8:
                            lvi.SubItems.Add((tvc.subclass_code == null ? "" : tvc.subclass_code));
                            sqlInsert = sqlInsert.Replace("{subclass_code}", (tvc.subclass_code == null ? "" : tvc.subclass_code).Replace("'", "''"));
                            break;
                        case 9:
                            lvi.SubItems.Add((tvc.vendor_name == null ? "" : tvc.vendor_name));
                            sqlInsert = sqlInsert.Replace("{vendor_name}", (tvc.vendor_name == null ? "" : tvc.vendor_name).Replace("'", "''"));
                            break;
                        case 10:
                            lvi.SubItems.Add((tvc.upc == null ? "" : tvc.upc));
                            sqlInsert = sqlInsert.Replace("{upc}", (tvc.upc == null ? "" : tvc.upc).Replace("'", "''"));
                            break;
                        case 11:
                            lvi.SubItems.Add((tvc.long_description == null ? "" : tvc.long_description));
                            sqlInsert = sqlInsert.Replace("{long_description}", (tvc.long_description == null ? "" : tvc.long_description).Replace("'", "''"));
                            break;
                        case 12:
                            lvi.SubItems.Add((tvc.weight == null ? "" : tvc.weight.ToString()));
                            sqlInsert = sqlInsert.Replace("{weight}", (tvc.weight == null ? "" : tvc.weight.ToString()).Replace("'", "''"));
                            break;
                        case 13:
                            lvi.SubItems.Add((tvc.length== null ? "" : tvc.length.ToString()));
                            sqlInsert = sqlInsert.Replace("{length}", (tvc.length == null ? "" : tvc.length.ToString()).Replace("'", "''"));
                            break;
                        case 14:
                            lvi.SubItems.Add((tvc.width == null ? "" : tvc.width.ToString()));
                            sqlInsert = sqlInsert.Replace("{width}", (tvc.width == null ? "" : tvc.width.ToString()).Replace("'", "''"));
                            break;
                        case 15:
                            lvi.SubItems.Add((tvc.height == null ? "" : tvc.height.ToString()));
                            sqlInsert = sqlInsert.Replace("{height}", (tvc.height == null ? "" : tvc.height.ToString()).Replace("'", "''"));
                            break;
                        case 16:
                            lvi.SubItems.Add((tvc.pack_weight == null ? "" : tvc.pack_weight.ToString()));
                            sqlInsert = sqlInsert.Replace("{pack_weight}", (tvc.pack_weight == null ? "" : tvc.pack_weight.ToString()).Replace("'", "''"));
                            break;
                        case 17:
                            lvi.SubItems.Add((tvc.pack_length == null ? "" : tvc.pack_length.ToString()));
                            sqlInsert = sqlInsert.Replace("{pack_length}", (tvc.pack_length == null ? "" : tvc.pack_length.ToString()).Replace("'", "''"));
                            break;
                        case 18:
                            lvi.SubItems.Add((tvc.pack_width == null ? "" : tvc.pack_width.ToString()));
                            sqlInsert = sqlInsert.Replace("{pack_width}", (tvc.pack_width == null ? "" : tvc.pack_width.ToString()).Replace("'", "''"));
                            break;
                        case 19:
                            lvi.SubItems.Add((tvc.pack_height == null ? "" : tvc.pack_height.ToString()));
                            sqlInsert = sqlInsert.Replace("{pack_height}", (tvc.pack_height == null ? "" : tvc.pack_height.ToString()).Replace("'", "''"));
                            break;
                        case 20:
                            lvi.SubItems.Add((tvc.retail_pack_qty == null ? "" : tvc.retail_pack_qty.ToString()));
                            sqlInsert = sqlInsert.Replace("{retail_pack_qty}", (tvc.retail_pack_qty == null ? "" : tvc.retail_pack_qty.ToString()).Replace("'", "''"));
                            break;
                        case 21:
                            lvi.SubItems.Add((tvc.member_pack_qty == null ? "" : tvc.member_pack_qty.ToString()));
                            sqlInsert = sqlInsert.Replace("{member_pack_qty}", (tvc.member_pack_qty == null ? "" : tvc.member_pack_qty.ToString()).Replace("'", "''"));
                            break;
                        case 22:
                            lvi.SubItems.Add((tvc.member_pack_type == null ? "" : tvc.member_pack_type));
                            sqlInsert = sqlInsert.Replace("{member_pack_type}", (tvc.member_pack_type == null ? "" : tvc.member_pack_type).Replace("'", "''"));
                            break;
                        case 23:
                            lvi.SubItems.Add((tvc.member_break_pack == null ? "" : tvc.member_break_pack));
                            sqlInsert = sqlInsert.Replace("{member_break_pack}", (tvc.member_break_pack == null ? "" : tvc.member_break_pack).Replace("'", "''"));
                            break;
                        case 24:
                            lvi.SubItems.Add((tvc.model == null ? "" : tvc.model));
                            sqlInsert = sqlInsert.Replace("{model}", (tvc.model == null ? "" : tvc.model).Replace("'", "''"));
                            break;
                        case 25:
                            lvi.SubItems.Add((tvc.item_picture_id == null ? "" : tvc.item_picture_id));
                            sqlInsert = sqlInsert.Replace("{item_picture_id}", (tvc.item_picture_id == null ? "" : tvc.item_picture_id).Replace("'", "''"));
                            break;
                        case 26:
                            lvi.SubItems.Add((tvc.country_code == null ? "" : tvc.country_code));
                            sqlInsert = sqlInsert.Replace("{country_code}", (tvc.country_code == null ? "" : tvc.country_code).Replace("'", "''"));
                            break;
                        case 27:
                            lvi.SubItems.Add((tvc.to_be_discontinued == null ? "" : tvc.to_be_discontinued));
                            sqlInsert = sqlInsert.Replace("{to_be_discontinued}", (tvc.to_be_discontinued == null ? "" : tvc.to_be_discontinued).Replace("'", "''"));
                            break;
                        case 28:
                            lvi.SubItems.Add((tvc.retail_uom == null ? "" : tvc.retail_uom));
                            sqlInsert = sqlInsert.Replace("{retail_uom}", (tvc.retail_uom == null ? "" : tvc.retail_uom).Replace("'", "''"));
                            break;
                        case 29:
                            lvi.SubItems.Add((tvc.Edit_Divisor == null ? "" : tvc.Edit_Divisor));
                            sqlInsert = sqlInsert.Replace("{Edit_Divisor}", (tvc.Edit_Divisor == null ? "" : tvc.Edit_Divisor).Replace("'", "''"));
                            break;
                        case 30:
                            lvi.SubItems.Add((tvc.Exclusive_Brand_Code == null ? "" : tvc.Exclusive_Brand_Code));
                            sqlInsert = sqlInsert.Replace("{Exclusive_Brand_Code}", (tvc.Exclusive_Brand_Code == null ? "" : tvc.Exclusive_Brand_Code).Replace("'", "''"));
                            break;
                    }
                }


                Connectivity.SQLCalls.sqlQuery(sqlInsert);


                    if (catalogLV.InvokeRequired == true)
                    {
                        catalogLV.Invoke(new MethodInvoker(() =>
                            {
                                catalogLV.Items.Add(lvi);
                                catalogCountLbl.Text = i.ToString() + "/" + xmlnode.Count.ToString();
                                catalogCountLbl.Refresh();
                            }));
                    }
                    else
                    {

                        catalogLV.Items.Add(lvi);
                    }
                //str = xmlnode[i].ChildNodes.Item(0).InnerText.Trim() + " " + xmlnode[i].ChildNodes.Item(1).InnerText.Trim() + " " + xmlnode[i].ChildNodes.Item(2).InnerText.Trim();
            }
            

            
        }
        private void SetupCatalogView()
        {
            catalogLV.Items.Clear();
            catalogLV.Columns.Clear();
            catalogLV.Clear();
            catalogLV.FullRowSelect = true;
            catalogLV.GridLines = true;
            catalogLV.View = View.Details;

            ColumnHeader col_LVID = new ColumnHeader();
            ColumnHeader col_item_nbr = new ColumnHeader();
            ColumnHeader col_srp_cost = new ColumnHeader();
            ColumnHeader col_member_cost = new ColumnHeader();
            ColumnHeader col_ds_cost = new ColumnHeader();
            ColumnHeader col_short_description = new ColumnHeader();
            ColumnHeader col_vendor_id = new ColumnHeader();
            ColumnHeader col_dtp_code = new ColumnHeader();
            ColumnHeader col_class_code = new ColumnHeader();
            ColumnHeader col_subclass_code = new ColumnHeader();
            ColumnHeader col_vendor_name = new ColumnHeader();
            ColumnHeader col_upc = new ColumnHeader();
            ColumnHeader col_long_description = new ColumnHeader();
            ColumnHeader col_weight = new ColumnHeader();
            ColumnHeader col_length = new ColumnHeader();
            ColumnHeader col_width = new ColumnHeader();
            ColumnHeader col_height = new ColumnHeader();
            ColumnHeader col_pack_weight = new ColumnHeader();
            ColumnHeader col_pack_length = new ColumnHeader();
            ColumnHeader col_pack_width = new ColumnHeader();
            ColumnHeader col_pack_height = new ColumnHeader();
            ColumnHeader col_retail_pack_qty = new ColumnHeader();
            ColumnHeader col_member_pack_qty = new ColumnHeader();
            ColumnHeader col_member_pack_type = new ColumnHeader();
            ColumnHeader col_member_break_pack = new ColumnHeader();
            ColumnHeader col_model = new ColumnHeader();
            ColumnHeader col_item_picture_id = new ColumnHeader();
            ColumnHeader col_country_code = new ColumnHeader();
            ColumnHeader col_to_be_discontinued = new ColumnHeader();
            ColumnHeader col_retail_uom = new ColumnHeader();
            ColumnHeader col_Edit_Divisor = new ColumnHeader();
            ColumnHeader col_Exclusive_Brand_Code = new ColumnHeader();

            col_LVID.Text = "LVID";
            col_LVID.Width = 10;
            col_item_nbr.Text = "Item Number";
            col_item_nbr.Width = 25;
            col_srp_cost.Text = "SRP Cost";
            col_srp_cost.Width = 15;
            col_member_cost.Text = "Member Cost";
            col_member_cost.Width = 15;
            col_ds_cost.Text = "DS Cost";
            col_ds_cost.Width = 15;
            col_short_description.Text = "Short Description";
            col_short_description.Width = 50;
            col_vendor_id.Text = "Vendor ID";
            col_vendor_id.Width = 20;
            col_dtp_code.Text = "DTP Code";
            col_dtp_code.Width = 15;
            col_class_code.Text = "Class Code";
            col_class_code.Width = 15;
            col_subclass_code.Text = "Subclass Code";
            col_subclass_code.Width = 15;
            col_vendor_name.Text = "Vendor Name";
            col_vendor_name.Width = 20;
            col_upc.Text = "UPC";
            col_upc.Width = 30;
            col_long_description.Text = "Long Description";
            col_long_description.Width = 60;
            col_weight.Text = "Weight";
            col_weight.Width = 15;
            col_length.Text = "Length";
            col_length.Width = 15;
            col_width.Text = "Width";
            col_width.Width = 15;
            col_height.Text = "Height";
            col_height.Width = 15;
            col_pack_weight.Text = "Pack Weight";
            col_pack_weight.Width = 15;
            col_pack_length.Text = "Pack Length";
            col_pack_length.Width = 15;
            col_pack_width.Text = "Pack Width";
            col_pack_width.Width = 15;
            col_pack_height.Text = "Pack Height";
            col_pack_height.Width = 15;
            col_retail_pack_qty.Text = "Retail Pack Qty";
            col_retail_pack_qty.Width = 15;
            col_member_pack_qty.Text = "Member Pack Qty";
            col_member_pack_qty.Width = 15;
            col_member_pack_type.Text = "Member Pack Type";
            col_member_pack_type.Width = 15;
            col_member_break_pack.Text = "Member Break Pack";
            col_member_break_pack.Width = 15;
            col_model.Text = "Model";
            col_model.Width = 20;
            col_item_picture_id.Text = "Item Picture ID";
            col_item_picture_id.Width = 15;
            col_country_code.Text = "Country Code";
            col_country_code.Width = 15;
            col_to_be_discontinued.Text = "To Be D/C";
            col_to_be_discontinued.Width = 15;
            col_retail_uom.Text = "Retail UOM";
            col_retail_uom.Width = 20;
            col_Edit_Divisor.Text = "Edit Divisor";
            col_Edit_Divisor.Width = 15;
            col_Exclusive_Brand_Code.Text = "Exclusive Brand Code";
            col_Exclusive_Brand_Code.Width = 15;


            catalogLV.Columns.AddRange(new ColumnHeader[] {col_LVID, col_item_nbr, col_srp_cost, col_member_cost, col_ds_cost, col_short_description, col_vendor_id, col_dtp_code, col_class_code, col_subclass_code, col_vendor_name, col_upc, col_long_description, col_weight, col_length, col_width, col_height, col_pack_weight, col_pack_length, col_pack_width, col_pack_height, col_retail_pack_qty, col_member_pack_qty, col_member_pack_type, col_member_break_pack, col_model, col_item_picture_id, col_country_code, col_to_be_discontinued, col_retail_uom, col_Edit_Divisor, col_Exclusive_Brand_Code });





        }


        //


















        //Code for Inventory
        private void loadInventoryBtn_Click(object sender, EventArgs e)
        {
            
            
            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                inventoryPathLbl.Text = openFileDialog1.FileName;
                processInventoryBtn.Enabled = true;
            }
        }

        private void processInventoryBtn_Click(object sender, EventArgs e)
        {
            SetupInventoryView();

            System.Threading.Thread t = new System.Threading.Thread(ProcessInventoryXML);
            t.Start();
            
        }

        private void SetupInventoryView()
        {
            inventoryLV.Items.Clear();
            inventoryLV.Columns.Clear();
            inventoryLV.Clear();
            inventoryLV.View = View.Details;
            inventoryLV.GridLines = true;
            inventoryLV.FullRowSelect = true;


            ColumnHeader col_LVID = new ColumnHeader();
            ColumnHeader col_rdc_nbr = new ColumnHeader();
            ColumnHeader col_item_nbr = new ColumnHeader();
            ColumnHeader col_Inventory = new ColumnHeader();

            col_LVID.Text = "LVID";
            col_LVID.Width = 30;
            col_rdc_nbr.Text = "rdc_nbr";
            col_rdc_nbr.Width = 50;
            col_Inventory.Text = "Inventory";
            col_Inventory.Width = 30;

            inventoryLV.Columns.AddRange(new ColumnHeader[] { col_LVID, col_rdc_nbr, col_Inventory });

        }

        private void ProcessInventoryXML()
        {
            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnode;
            int i = 0;
            string str = null;
            FileStream fs = new FileStream(inventoryPathLbl.Text, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);
            xmlnode = xmldoc.GetElementsByTagName("Item");
            //MessageBox.Show("Node Count: " + xmlnode.Count.ToString());
            for (i = 0; i <= xmlnode.Count - 1; i++)
            {
                TVInventoryItem tvi = new TVInventoryItem();
                ListViewItem lvi = new ListViewItem();
                lvi.Text = i.ToString();
                
                foreach (XmlAttribute xmlattr in xmlnode[i].Attributes)
                {
                    switch (xmlattr.Name)
                    {
                        case "rdc_nbr":
                            tvi.rdc_nbr = xmlattr.Value;
                            break;
                        case "item_nbr":
                            tvi.item_nbr = xmlattr.Value;
                            break;
                        case "Inventory":
                            tvi.Inventory = int.Parse(xmlattr.Value);
                            break;
                    }
                }

                string sqlInsert = "INSERT INTO TrueValueInventory (rdc_nbr,item_nbr,Inventory) VALUES('{rdc_nbr}','{item_nbr}',{Inventory})";

                for (int myI = 0; myI <= 2; myI++)
                {
                    switch (myI)
                    {
                        case 0:
                            lvi.SubItems.Add((tvi.rdc_nbr == null ? "" : tvi.rdc_nbr));
                            sqlInsert = sqlInsert.Replace("{rdc_nbr}", (tvi.rdc_nbr == null ? "" : tvi.rdc_nbr).Replace("'","''"));
                            break;
                        case 1:
                            lvi.SubItems.Add((tvi.item_nbr == null ? "" : tvi.item_nbr));
                            sqlInsert = sqlInsert.Replace("{item_nbr}", (tvi.item_nbr == null ? "" : tvi.item_nbr).Replace("'", "''"));
                            break;
                        case 2:
                            lvi.SubItems.Add((tvi.Inventory == null ? "" : tvi.Inventory.ToString()));
                            sqlInsert = sqlInsert.Replace("{Inventory}",(tvi.Inventory == null ? "" : tvi.Inventory.ToString()).Replace("'","''"));
                            break;
                    }
                }


                Connectivity.SQLCalls.sqlQuery(sqlInsert);


                if (catalogLV.InvokeRequired == true)
                {
                    catalogLV.Invoke(new MethodInvoker(() =>
                    {
                        inventoryLV.Items.Add(lvi);
                        processedLbl.Text = i.ToString() + "/" + xmlnode.Count.ToString();
                        processedLbl.Refresh();
                    }));
                }
                else
                {

                    inventoryLV.Items.Add(lvi);
                }


            }
        }


        




        //
















        //Code for adding items into poinv



        private void getItemsBtn_Click(object sender, EventArgs e)
        {
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT item_nbr,model FROM TrueValueCatalog ORDER BY item_nbr ASC");
            Connectivity.SQLCalls.QueryResults qr2 = Connectivity.SQLCalls.sqlQuery("SELECT PrimaryVendorNumber,AvailLimitDefault FROM ProductLine WHERE PRoductLine='T-V'");
            Connectivity.SQLCalls.QueryResults qr3 = Connectivity.SQLCalls.sqlQuery("SELECT TOP 1 EBayStoreCategoryID, COUNT(*) AS TotalOccurrences FROM PartNumbers WHERE LEFT(ItemNumber,3)='T-V' AND EBayStoreCategoryID IS NOT NULL AND EBayStoreCategoryID IN (SELECT CategoryID FROM EBayCategory WHERE CategoryType=1) GROUP BY EBayStoreCategoryID ORDER BY TotalOccurrences DESC");

            if (qr.Rows != null & qr2.Rows != null)
            {
                if (qr.Rows.Count > 0 & qr2.Rows.Count > 0)
                {
                    itemsInTVCatalogLbl.Text = qr.Rows.Count.ToString() + " items found.";

                    TVItems = new Dictionary<string, TVItemImportData>();

                    foreach (string str in qr.Rows)
                    {

                        string TVI = "T-V" + str.Split('|')[0];
                        TVItemImportData tviid = new TVItemImportData();
                        tviid.ProductLine = "T-V";
                        tviid.PrimaryVendor = qr2.Rows[0].Split('|')[0];
                        tviid.ItemNumber = TVI;
                        tviid.Component = TVI;
                        tviid.EBayStoreCategoryID = (qr3.Rows.Count > 0 ? long.Parse(qr3.Rows[0].Split('|')[0]) : 0);
                        tviid.ItemStatus = 4;
                        tviid.WhseCode = "000";
                        tviid.AvailabilityLimit = int.Parse(qr2.Rows[0].Split('|')[1]);
                        tviid.Desc2 = qr.Rows[0].Split('|')[1];
                        tviid.WebsiteID = 1;
                        tviid.ComponentID = 0;
                        tviid.ItemID = 0;

                        TVItems.Add(tviid.ItemNumber, tviid);

                        processedLbl.Text = "Processed " + TVItems.Count.ToString() + "/" + qr.Rows.Count.ToString();
                        processedLbl.Refresh();
                    }

                    //MessageBox.Show("Time to verify data in TVItems dictionary...");
                    ProcessTVItems();

                }
                else
                {
                    itemsInTVCatalogLbl.Text = "No items found in table.";
                    return;
                }
            }
            else
            {
                itemsInTVCatalogLbl.Text = "Error. Could not communicate with table 'TrueValueCatalog'";
                return;
            }

        }
        private void ProcessTVItems()
        {
            int count = 1;
            foreach (KeyValuePair<string, TVItemImportData> kvp in TVItems)
            {
                bool canAddItem = true;
                //step 1: does item already exist?
                Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT ItemID FROM InventoryMaster WHERE ItemNumber='" + kvp.Key + "'");
                if (qr.Rows != null)
                {
                    if (qr.Rows.Count > 0)
                    {
                        canAddItem = false;
                    }
                }
                else
                {
                    canAddItem = false;
                }
                //step 2: does linecode exist?
                Connectivity.SQLCalls.QueryResults qr2 = Connectivity.SQLCalls.sqlQuery("SELECT ProductLine FROM ProductLine WHERE ProductLine='" + kvp.Value.ProductLine + "'");
                if (qr2.Rows != null)
                {
                    if (qr2.Rows.Count == 0)
                    {
                        canAddItem = false;
                    }
                }
                else
                {
                    canAddItem = false;
                }

                //step 3: is itemnumber length between 4 and 27 chars?
                if (kvp.Value.ItemNumber.Length > 27 | kvp.Value.ItemNumber.Length < 4)
                {
                    canAddItem = false;
                }

                if (canAddItem == true)
                {
                    //AddItemToDB(kvp);
                    processingToFileLbl.Text = "Processed To File " + count.ToString() + "/" + TVItems.Count.ToString();
                    processingToFileLbl.Refresh();
                    AddItemToFile(kvp);
                }

                count++;
            }
        }


        private void AddItemToFile(KeyValuePair<string,TVItemImportData> kvp)
        {
            string line ="";
            if (System.IO.File.Exists("c:\\users\\tomwestbrook\\desktop\\TV_To_Be_Imported.csv") ==false)
            {
                
                line += "ItemID,ItemNumber,ItemStatus,WebsiteID,WhseCode,ProductLine,PrimaryVendor,AvailabilityLimit,Component,ComponentID,Desc2,EBayStoreCategoryID,\r\n";
                System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\TV_To_Be_Imported.csv",line);
            }

            line += kvp.Value.ItemID + "," + kvp.Value.ItemNumber + "," + kvp.Value.ItemStatus + "," + kvp.Value.WebsiteID + "," + kvp.Value.WhseCode + "," + kvp.Value.ProductLine + "," + kvp.Value.PrimaryVendor + "," + kvp.Value.AvailabilityLimit + "," + kvp.Value.Component + "," + kvp.Value.ComponentID + "," + kvp.Value.Desc2 + "," + kvp.Value.EBayStoreCategoryID + ",\r\n";
            System.IO.File.AppendAllText("c:\\users\\tomwestbrook\\desktop\\TV_To_Be_Imported.csv",line);
            
        }

        private void AddItemToDB(KeyValuePair<string,TVItemImportData> kvp)
        {
            int ItemID = 0;
            int ComponentID = 0;

            Connectivity.SQLCalls.sqlQuery("INSERT INTO InventoryMaster (ItemNumber,WebsiteID,ProductLine,PrimaryVendor) VALUES('" + kvp.Value.ItemNumber + "'," + kvp.Value.WebsiteID.ToString() + ",'" + kvp.Value.ProductLine + "','" + kvp.Value.PrimaryVendor + "')");
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT ItemID FROM InventoryMaster WHERE ItemNumber='" + kvp.Value.ItemNumber + "'");
            if (qr.Rows != null)
            {
                if (qr.Rows.Count > 0)
                {
                    ItemID = int.Parse(qr.Rows[0].Split('|')[0]);
                }
                else
                {
                    MessageBox.Show("Error 1");
                }
            }
            else
            {
                MessageBox.Show("Error 2");
            }

            Connectivity.SQLCalls.sqlQuery("INSERT INTO InventoryQuantities (ItemID,ItemNumber,WhseCode) VALUES(" + ItemID.ToString() + ",'" + kvp.Value.ItemNumber + "','" + kvp.Value.WhseCode + "')");
            Connectivity.SQLCalls.sqlQuery("INSERT INTO InventoryComponents (Component) VALUES('" + kvp.Value.ItemNumber + "'");
            Connectivity.SQLCalls.QueryResults qr2 = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID FROM InventoryComponents WHERE Component='" + kvp.Value.ItemNumber + "'");
            if (qr2.Rows != null)
            {
                if (qr2.Rows.Count > 0)
                {
                    ComponentID = int.Parse(qr2.Rows[0].Split('|')[0]);
                }
                else
                {
                    MessageBox.Show("Error 3");
                }
            }
            else
            {
                MessageBox.Show("Error 4");
            }
            Connectivity.SQLCalls.sqlQuery("INSERT INTO InventoryComponentMap (ItemID,ItemNumber,ComponentID) VALUES(" + ItemID.ToString() + ",'" + kvp.Value.ItemNumber + "'," + ComponentID.ToString() + ")");
            Connectivity.SQLCalls.sqlQuery("INSERT INTO PartNumbers (ItemNumber,Desc2,AvailabilityLimit) VALUES ('" + kvp.Value.ItemNumber + "','" + kvp.Value.Desc2 + "'," + kvp.Value.AvailabilityLimit.ToString() + ")");
            if (kvp.Value.EBayStoreCategoryID > 0)
            {
                Connectivity.SQLCalls.sqlQuery("UPDATE PartNumbers SET EBayStoreCategoryID=" + kvp.Value.EBayStoreCategoryID.ToString() + " WHERE ItemNumber='" + kvp.Value.ItemNumber + "'");
            }
            if (kvp.Value.ItemStatus > 0)
            {
                Connectivity.SQLCalls.sqlQuery("UPDATE InventoryMaster SET ItemStatus=" + kvp.Value.ItemStatus.ToString() + " WHERE ItemNumber='" + kvp.Value.ItemNumber + "'");
            }

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }
        


        public struct ShippingData
        {
            public string ItemNumber;
            public decimal Weight;
            public decimal Length;
            public decimal Width;
            public decimal Height;
        }

        private void updateShippingDataBtn_Click(object sender, EventArgs e)
        {
            StartIndex = int.Parse(textBox1.Text);
            CurrentSeconds = 200;
            timer1.Enabled = true;


        }

        private void debugFrm_Load(object sender, EventArgs e)
        {
            SetupShippingLV();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {            
            label4.Text = "Waiting " + (TotalSeconds - CurrentSeconds).ToString() + " secs to process more..";
            label4.Refresh();
            CurrentSeconds++;
            if (CurrentSeconds >= TotalSeconds)
            {
                timer1.Enabled = false;
                CurrentSeconds = 0;
                OffsetStart = true;
                ProcessItems();
                StartIndex += 50;
                timer1.Enabled = true;
            }
        }
        private void SetupShippingLV()
        {
            shippingLV.Items.Clear();
            shippingLV.Columns.Clear();
            shippingLV.Clear();
            shippingLV.View = View.Details;
            shippingLV.FullRowSelect = true;
            shippingLV.GridLines = true;
            ColumnHeader colID = new ColumnHeader();
            ColumnHeader colItem = new ColumnHeader();
            ColumnHeader colWeight = new ColumnHeader();
            ColumnHeader colLength = new ColumnHeader();
            ColumnHeader colWidth = new ColumnHeader();
            ColumnHeader colHeight = new ColumnHeader();
            ColumnHeader colUPS = new ColumnHeader();
            ColumnHeader colUSPS = new ColumnHeader();
            ColumnHeader colStatus = new ColumnHeader();
            colID.Text = "ID";
            colID.Width = 1;
            colItem.Text = "Item Number";
            colItem.Width = 100;
            colWeight.Text = "Weight";
            colWeight.Width = 50;
            colLength.Text = "Length";
            colLength.Width = 50;
            colWidth.Text = "Width";
            colWidth.Width = 50;
            colHeight.Text = "Height";
            colHeight.Width = 50;
            colUPS.Text = "UPS Rate";
            colUPS.Width = 50;
            colUSPS.Text = "USPS Rate";
            colUSPS.Width = 50;
            colStatus.Text = "Status";
            colStatus.Width = 100;
            shippingLV.Columns.AddRange(new ColumnHeader[] { colID, colItem, colWeight, colLength, colWidth, colHeight, colUPS, colUSPS,colStatus });
            shippingLV.Refresh();
        }
            
        private void ProcessItems()
        {
            string UPSShippingRate = "";
            string USPSShippingRate = "";
            decimal alteredWeight = 0;

            List<ShippingData> ShipData = new List<ShippingData>();
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT item_nbr, weight, length, width, height FROM TrueValueCatalog ORDER BY item_nbr ASC");
            if (qr.Rows != null)
            {
                if (qr.Rows.Count > 0)
                {
                    foreach (string str in qr.Rows)
                    {
                        ShippingData sd = new ShippingData();
                        sd.ItemNumber = str.Split('|')[0];
                        sd.Weight = decimal.Parse(str.Split('|')[1]);
                        sd.Length = decimal.Parse(str.Split('|')[2]);
                        sd.Width = decimal.Parse(str.Split('|')[3]);
                        sd.Height = decimal.Parse(str.Split('|')[4]);
                        ShipData.Add(sd);
                    }
                }
            }


            if (ShipData.Count > 0)
            {
                int Count = 0;
                foreach (ShippingData ship in ShipData)
                {
                    bool didJob = false;
                    if (Count >= StartIndex)
                    {
                        if (OffsetStart == true)
                        {
                            if (Count <= (StartIndex + BatchSize))
                            {
                                Count++;
                                label4.Text = "Working on " + Count.ToString() + "/" + ShipData.Count.ToString();
                                label4.Refresh();
                                alteredWeight = ship.Weight > 0 ? ship.Weight : 0.9M;
                                string url = "http://10.0.50.16/whse/freightquote.plex?type=box&boxweight0=" + Math.Ceiling(alteredWeight).ToString() + "&boxlength0=" + Math.Ceiling(ship.Length).ToString() + "&boxwidth0=" + Math.Ceiling(ship.Width).ToString() + "&boxheight0=" + Math.Ceiling(ship.Height).ToString() + "&zipcode=15001&state=PA&country=US&residential=1&carrier0=UPS&carrier1=USPS";
                                //string url = "http://10.0.50.16/whse/freightquote.plex?type=box&boxweight0=11&boxlength0=36&boxwidth0=7&boxheight0=3&zipcode=15001&state=PA&country=US&residential=1&carrier0=UPS&carrier1=USPS";


                                string results = Connectivity.HTTPCalls.HTTP_GET(url);

                                List<ShippingRates> Rates = ParseShippingRates(results);
                                decimal UPSRate = 0;
                                decimal USPSRate = 0;
                                foreach (ShippingRates sr in Rates)
                                {
                                    if (sr.ServiceProvider == "UPS" & sr.ServiceName == "Ground")
                                    {
                                        if (sr.Negotiated_Rate < UPSRate | UPSRate <= 0)
                                        {
                                            UPSRate = sr.Negotiated_Rate;
                                        }
                                    }
                                    if (sr.ServiceProvider == "USPS" & sr.ServiceName.Length==19 & sr.ServiceName.Contains("Priority Mail")==true)                                    
                                    {
                                        if (sr.Negotiated_Rate < USPSRate | USPSRate <=0)
                                        {
                                            USPSRate = sr.Negotiated_Rate;
                                        }
                                    }
                                    //overwrite USPSRate if it has first class available...
                                    if (sr.ServiceProvider == "USPS" & sr.ServiceName == "First-Class Mail Parcel")
                                    {
                                        if (sr.Negotiated_Rate < USPSRate | USPSRate <= 0)
                                        {
                                            USPSRate = sr.Negotiated_Rate;
                                        }
                                    }
                                }

                                UPSShippingRate = UPSRate.ToString();
                                USPSShippingRate = USPSRate.ToString();
                           //     if (results.Contains("<ResponseStatusDescription>Success") == true)
                           //     {
                           //         if (results.Contains("TP_7") == true | USPSRatesCheck(results) == true)
                           //         {
                                       /* //fully sucessful.. parse!
                                        string USPSdata = "";
                                        string UPSdata = "";
                                        try
                                        {
                                            USPSdata = results.Split(new string[] { "</RateV4Request>" }, StringSplitOptions.None)[1].Split(new string[] { "<?xml version" }, StringSplitOptions.None)[0];
                                        }
                                        catch { }
                                        try
                                        {
                                            UPSdata = results.Split(new string[] { "</RatingServiceSelectionRequest>" }, StringSplitOptions.None)[1].Split(new string[] { "<?xml version" }, StringSplitOptions.None)[0];
                                        }
                                        catch { }
                                       
                                        //string USPSShippingRate = USPSdata.Split(new string[] { "TP_7\":{\"tp_rate\":" }, StringSplitOptions.None)[1].Split(',')[0];\
                                        USPSShippingRate = "";
                                        try
                                        {
                                            USPSShippingRate = USPSdata.Split(new string[] { "2-Day" }, StringSplitOptions.None)[1].Split(new string[] { "negotiated_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                                        }
                                        catch
                                        {
                                            try
                                            {
                                                USPSShippingRate = USPSdata.Split(new string[] { "Standard Post" }, StringSplitOptions.None)[1].Split(new string[] { "negotiated_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                                            }
                                            catch { }
                                        }
                                        UPSShippingRate = UPSdata.Split(new string[] { "03\":{\"tp_rate\":" }, StringSplitOptions.None)[1].Split(',')[0];*/
                                        decimal tmp = 0;

                                        

                                        if ((decimal.TryParse(USPSShippingRate, out tmp) | USPSRatesCheck(results)==true) & decimal.TryParse(UPSShippingRate, out tmp))
                                        {
                                            bool didUpdate = false;
                                            Connectivity.SQLCalls.QueryResults test = Connectivity.SQLCalls.sqlQuery("SELECT item_nbr FROM TrueValueCosting WHERE item_nbr='" + ship.ItemNumber + "'");
                                            if (test.Rows.Count > 0)
                                            {
                                                Connectivity.SQLCalls.sqlQuery("UPDATE TrueValueCosting SET ups_estimated_shipping_cost=" + UPSShippingRate + ", usps_estimated_shipping_cost=" + (USPSRatesCheck(results) == false ? "NULL" : USPSShippingRate) + " WHERE item_nbr='" + ship.ItemNumber + "'");
                                                didJob = true;
                                                didUpdate = true;
                                            }
                                            else
                                            {
                                                Connectivity.SQLCalls.sqlQuery("INSERT INTO TrueValueCosting (item_nbr,ups_estimated_shipping_cost,usps_estimated_shipping_cost) VALUES('" + ship.ItemNumber + "'," + UPSShippingRate + "," + (USPSRatesCheck(results) == false ? "NULL" : USPSShippingRate) + ")");
                                                didJob = true;
                                            }

                                            ListViewItem lvi = new ListViewItem();
                                            lvi.Text = Count.ToString();
                                            lvi.SubItems.Add(ship.ItemNumber);
                                            lvi.SubItems.Add(alteredWeight.ToString("0.00"));
                                            lvi.SubItems.Add(ship.Length.ToString("0.00"));
                                            lvi.SubItems.Add(ship.Width.ToString("0.00"));
                                            lvi.SubItems.Add(ship.Height.ToString("0.00"));
                                            lvi.SubItems.Add(UPSShippingRate);
                                            lvi.SubItems.Add(USPSShippingRate);
                                            lvi.SubItems.Add((didUpdate == true ? "Updated" : "Inserted"));
                                            shippingLV.Items.Add(lvi);
                                            shippingLV.Refresh();
                                            progressBar1.Value = (int)Math.Ceiling(((decimal)Count / (decimal)ShipData.Count) * 100);
                                            label6.Text = (((decimal)Count / (decimal)ShipData.Count) * 100).ToString("0.0000") + "%";
                                            label6.Refresh();
                                            progressBar1.Refresh();
                                            shippingLV.EnsureVisible(shippingLV.Items.Count - 1);
                                        }




                                 //   }
                               // }
                            }
                            else
                            {
                                //dont error the rest of the items!
                                didJob = true;
                            }
                        }
                    }
                    else
                    {
                        Count++;
                        didJob = true;
                    }

                    if (didJob == false)
                    {
                        //this item didnt process... flag it
                        ListViewItem lvi = new ListViewItem();
                        lvi.BackColor = Color.Red;
                        lvi.ForeColor = Color.White;
                        lvi.Text = Count.ToString();
                        lvi.SubItems.Add(ship.ItemNumber);
                        lvi.SubItems.Add(alteredWeight.ToString("0.00"));
                        lvi.SubItems.Add(ship.Length.ToString("0.00"));
                        lvi.SubItems.Add(ship.Width.ToString("0.00"));
                        lvi.SubItems.Add(ship.Height.ToString("0.00"));
                        lvi.SubItems.Add(UPSShippingRate);
                        lvi.SubItems.Add(USPSShippingRate);
                        lvi.SubItems.Add("Error");
                        shippingLV.Items.Add(lvi);
                        shippingLV.Refresh();
                        progressBar1.Value = (int)Math.Ceiling(((decimal)Count / (decimal)ShipData.Count) * 100);
                        label6.Text = (((decimal)Count / (decimal)ShipData.Count) * 100).ToString("0.0000") + "%";
                        label6.Refresh();
                        progressBar1.Refresh();
                        shippingLV.EnsureVisible(shippingLV.Items.Count - 1);

                    }
                }


            }


        }

        //

        private bool USPSRatesCheck(string results)
        {
            switch (results)
            {
                case "package is too large to be mailed":
                    return false;                    
                case "Selected items are over the USPS 70 lb weight limit":
                    return false;
                default:
                    return true;                    
            }
        }
    
        public List<ShippingRates> ParseShippingRates(string results)
        {
            List<ShippingRates> tmpList = new List<ShippingRates>();

            string UPSData = results.Split(new string[]{"\"state_desc\":\""},StringSplitOptions.None)[1].Split(new string[]{"}}"},StringSplitOptions.None)[0];
            string USPSData = results.Split(new string[] { "\"state_desc\":\""}, StringSplitOptions.None)[2].Split(new string[] { "}}" }, StringSplitOptions.None)[0];
            if (UPSData.StartsWith("success")==true)
            {
                int count = 0;
                foreach (string str in UPSData.Split(new string[]{"\"tp_rate\":"},StringSplitOptions.None))
                {
                    if (count > 0)
                    {
                        try
                        {
                            string tp_rate = str.Split(',')[0];
                            string negotiated_rate = str.Split(new string[] { "negotiated_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            string list_rate = str.Split(new string[] { "list_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            string service_name = str.Split(new string[] { "service\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            ShippingRates tmp = new ShippingRates();
                            tmp.TP_Rate = decimal.Parse(tp_rate);
                            tmp.Negotiated_Rate = decimal.Parse(negotiated_rate);
                            tmp.List_Rate = decimal.Parse(list_rate);
                            tmp.ServiceProvider = "UPS";
                            tmp.ServiceName = service_name;
                            tmpList.Add(tmp);
                        }
                        catch { }
                    }
                    count++;
                }
            }
            if (USPSData.StartsWith("success")==true)
            {               
                int count = 0;
                foreach (string str in USPSData.Split(new string[] {"\"tp_rate\":"},StringSplitOptions.None))
                {
                    if (count > 0)
                    {
                        try
                        {
                            string tp_rate = str.Split(',')[0];
                            string negotiated_rate = str.Split(new string[] { "negotiated_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            string list_rate = str.Split(new string[] { "list_rate\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            string service_name = str.Split(new string[] { "service\":\"" }, StringSplitOptions.None)[1].Split('"')[0];
                            ShippingRates tmp = new ShippingRates();
                            tmp.TP_Rate = decimal.Parse(tp_rate);
                            tmp.Negotiated_Rate = decimal.Parse(negotiated_rate);
                            tmp.List_Rate = decimal.Parse(list_rate);
                            tmp.ServiceName = service_name;
                            tmp.ServiceProvider = "USPS";
                            tmpList.Add(tmp);
                        }
                        catch { }
                    }
                    count++;
                }
            }

            return tmpList;

        }

        public struct ShippingRates
        {
            public string ServiceProvider;
            public string ServiceName;
            public decimal TP_Rate;
            public decimal Negotiated_Rate;
            public decimal List_Rate;
        }
    }
}
