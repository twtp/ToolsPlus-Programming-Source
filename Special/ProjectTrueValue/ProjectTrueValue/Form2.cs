using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectTrueValue
{
    public partial class tvFrm : Form
    {

        public string BaseQuery = "";

        public tvFrm()
        {
            InitializeComponent();
        }

        private void tvFrm_Load(object sender, EventArgs e)
        {
            SetupBaseQuery();
            SetupCatalogView();
        }

        private void SetupBaseQuery()
        {
            BaseQuery = "SELECT item_nbr,TotalQty,srp_cost,member_cost,ds_cost,short_description,vendor_id,dtp_code,class_code,subclass_code,vendor_name,upc,long_description, weight, length, width, height, pack_weight, pack_length, pack_width, pack_height,retail_pack_qty,member_pack_qty,member_pack_type,member_break_pack,model,item_picture_id,country_code,to_be_discontinued,retail_uom, Edit_Divisor,Exclusive_Brand_Code FROM ";
            BaseQuery += "( SELECT TrueValueInventory.item_nbr AS ItemNumber, SUM(TrueValueInventory.Inventory) AS TotalQty FROM TrueValueInventory GROUP BY TrueValueInventory.item_nbr ) A ";
            BaseQuery += "INNER JOIN TrueValueCatalog ON TrueValueCatalog.item_nbr = ItemNumber ";
            BaseQuery += "WHERE ";
        }

        private void SetupCatalogView()
        {
            fbfLV.Items.Clear();
            fbfLV.Columns.Clear();
            fbfLV.Clear();
            fbfLV.FullRowSelect = true;
            fbfLV.GridLines = true;
            fbfLV.View = View.Details;

            ColumnHeader col_LVID = new ColumnHeader();
            ColumnHeader col_item_nbr = new ColumnHeader();
            ColumnHeader col_tot_qty = new ColumnHeader();
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
            col_LVID.Width = 1;
            col_item_nbr.Text = "Item Number";
            col_item_nbr.Width = 100;
            col_tot_qty.Text = "Total Qty";
            col_tot_qty.Width = 100;
            col_srp_cost.Text = "SRP Cost";
            col_srp_cost.Width = 100;
            col_member_cost.Text = "Member Cost";
            col_member_cost.Width = 100;
            col_ds_cost.Text = "DS Cost";
            col_ds_cost.Width = 100;
            col_short_description.Text = "Short Description";
            col_short_description.Width = 100;
            col_vendor_id.Text = "Vendor ID";
            col_vendor_id.Width = 100;
            col_dtp_code.Text = "DTP Code";
            col_dtp_code.Width = 100;
            col_class_code.Text = "Class Code";
            col_class_code.Width = 100;
            col_subclass_code.Text = "Subclass Code";
            col_subclass_code.Width = 100;
            col_vendor_name.Text = "Vendor Name";
            col_vendor_name.Width = 100;
            col_upc.Text = "UPC";
            col_upc.Width = 100;
            col_long_description.Text = "Long Description";
            col_long_description.Width = 100;
            col_weight.Text = "Weight";
            col_weight.Width = 100;
            col_length.Text = "Length";
            col_length.Width = 100;
            col_width.Text = "Width";
            col_width.Width = 100;
            col_height.Text = "Height";
            col_height.Width = 100;
            col_pack_weight.Text = "Pack Weight";
            col_pack_weight.Width = 100;
            col_pack_length.Text = "Pack Length";
            col_pack_length.Width = 100;
            col_pack_width.Text = "Pack Width";
            col_pack_width.Width = 100;
            col_pack_height.Text = "Pack Height";
            col_pack_height.Width = 100;
            col_retail_pack_qty.Text = "Retail Pack Qty";
            col_retail_pack_qty.Width = 100;
            col_member_pack_qty.Text = "Member Pack Qty";
            col_member_pack_qty.Width = 100;
            col_member_pack_type.Text = "Member Pack Type";
            col_member_pack_type.Width = 100;
            col_member_break_pack.Text = "Member Break Pack";
            col_member_break_pack.Width = 100;
            col_model.Text = "Model";
            col_model.Width = 100;
            col_item_picture_id.Text = "Item Picture ID";
            col_item_picture_id.Width = 100;
            col_country_code.Text = "Country Code";
            col_country_code.Width = 100;
            col_to_be_discontinued.Text = "To Be D/C";
            col_to_be_discontinued.Width = 100;
            col_retail_uom.Text = "Retail UOM";
            col_retail_uom.Width = 100;
            col_Edit_Divisor.Text = "Edit Divisor";
            col_Edit_Divisor.Width = 100;
            col_Exclusive_Brand_Code.Text = "Exclusive Brand Code";
            col_Exclusive_Brand_Code.Width = 100;



            fbfLV.Columns.AddRange(new ColumnHeader[] { col_LVID, col_item_nbr,col_tot_qty, col_srp_cost, col_member_cost, col_ds_cost, col_short_description, col_vendor_id, col_dtp_code, col_class_code, col_subclass_code, col_vendor_name, col_upc, col_long_description, col_weight, col_length, col_width, col_height, col_pack_weight, col_pack_length, col_pack_width, col_pack_height, col_retail_pack_qty, col_member_pack_qty, col_member_pack_type, col_member_break_pack, col_model, col_item_picture_id, col_country_code, col_to_be_discontinued, col_retail_uom, col_Edit_Divisor, col_Exclusive_Brand_Code });

                     
            fbfLV.Refresh();



        }

        private void button2_Click(object sender, EventArgs e)
        {
            string filepath = "";
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                filepath = saveFileDialog1.FileName;                
            }
            else
            {
                MessageBox.Show("Aborted Saving.");
                return;
            }

            //MessageBox.Show("Gotta give the programmer more time!!!");
            if (filepath.Length > 0)
            {
                
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(filepath))
                {
                    string tFile = "";
                    int count = 1;
                    foreach (ListViewItem lvi in fbfLV.Items)
                    {

                        if (count == 1)
                        {
                            statusLbl.Text = "Writing header. Please wait...";
                            statusLbl.Refresh();

                            string header = "";
                            foreach(ColumnHeader col in fbfLV.Columns)
                            {
                                if (col.Text != "LVID")
                                {
                                    header += col.Text + "\t";
                                }
                            }
                            sw.WriteLine(header);
                            

                        }

                        statusLbl.Text = "Writing line " + count.ToString() + "/" + fbfLV.Items.Count.ToString();
                        statusLbl.Refresh();
                        string tmp = "";
                        foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                        {
                            tmp += lvsi.Text + "\t";
                        }
                        sw.WriteLine(tmp);
                        tmp = "";
                        count++;
                    }
                }
                MessageBox.Show("Complete! File saved to \"" + filepath + "\"");

            }
            
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetupBaseQuery();
            string WhereClause = "";
            if (textBox1.Text.Length>0)
            {
                WhereClause += "AND srp_cost" + textBox1.Text + " ";
            }
            if (textBox27.Text.Length>0)
            {
                WhereClause += "AND TotalQty" + textBox27.Text + " ";
            }
            if (textBox2.Text.Length>0)
            {
                WhereClause += "AND member_cost" + textBox2.Text + " ";
            }
            if (textBox4.Text.Length>0)
            {
                WhereClause += "AND ds_cost" + textBox4.Text + " ";
            }
            if (textBox3.Text.Length>0)
            {
                WhereClause += "AND vendor_id" + textBox3.Text + " "; 
            }
            if (textBox8.Text.Length>0)
            {
                WhereClause += "AND dtp_code LIKE '%" + textBox8.Text + "%' ";
            }
            if (textBox7.Text.Length>0)
            {
                WhereClause += "AND class_code" + textBox7.Text + " ";
            }
            if (textBox6.Text.Length>0)
            {
                WhereClause += "AND subclass_code" + textBox6.Text + " ";
            }
            if (textBox5.Text.Length>0)
            {
                WhereClause += "AND vendor_name LIKE '%" + textBox5.Text + "%' ";
            }
            if (textBox12.Text.Length>0)
            {
                WhereClause += "AND upc LIKE '%" + textBox12.Text + "%' ";
            }
            if (textBox11.Text.Length>0)
            {
                WhereClause += "AND weight" + textBox11.Text + " ";
            }
            if (textBox10.Text.Length>0)
            {
                WhereClause += "AND length" + textBox10.Text + " ";
            }
            if (textBox9.Text.Length>0)
            {
                WhereClause += "AND width" + textBox9.Text + " ";
            }
            if (textBox16.Text.Length>0)
            {
                WhereClause += "AND height" + textBox16.Text + " ";
            }
            if (textBox15.Text.Length>0)
            {
                WhereClause += "AND pack_weight" + textBox15.Text + " ";
            }
            if (textBox14.Text.Length>0)
            {
                WhereClause += "AND pack_length" + textBox14.Text + " ";
            }
            if (textBox13.Text.Length>0)
            {
                WhereClause += "AND pack_width" + textBox13.Text + " ";
            }
            if (textBox20.Text.Length>0)
            {
                WhereClause += "AND pack_height" + textBox20.Text + " ";
            }
            if (textBox19.Text.Length>0)
            {
                WhereClause += "AND retail_pack_qty" + textBox19.Text + " ";
            }
            if (textBox18.Text.Length>0)
            {
                WhereClause += "AND member_pack_qty" + textBox18.Text + " ";
            }
            if (textBox17.Text.Length>0)
            {
                WhereClause += "AND member_break_type" + textBox17.Text + " ";
            }
            if (textBox24.Text.Length>0)
            {
                WhereClause += "AND member_break_pack" + textBox24.Text + " ";
            }
            if (textBox23.Text.Length>0)
            {
                WhereClause += "AND country_code LIKE '%" + textBox23.Text + "%' ";
            }
            if (textBox22.Text.Length>0)
            {
                WhereClause += "AND to_be_discontinued LIKE '%" + textBox22.Text + "%' ";
            }
            if (textBox21.Text.Length>0)
            {
                WhereClause += "AND retail_uom LIKE '%" + textBox21.Text + "%' ";
            }
            if (textBox26.Text.Length>0)
            {
                WhereClause += "AND Edit_Divisor LIKE '%" + textBox26.Text + "%' ";
            }
            if (textBox25.Text.Length>0)
            {
                WhereClause += "AND Exclusive_Brand_Code LIKE '%" + textBox25.Text + "%' ";
            }
            if (textBox30.Text.Length>0)
            {
                WhereClause += "AND short_description LIKE '%" + textBox30.Text + "%' ";
            }
            if (textBox29.Text.Length>0)
            {
                WhereClause += "AND long_description LIKE '%" + textBox29.Text + "%' ";
            }
            if (textBox28.Text.Length>0)
            {
                WhereClause += "AND model LIKE '%" + textBox28.Text + "%' ";
            }
            
            
            
            
            if (WhereClause.StartsWith("AND ") == true)
            {
                WhereClause = WhereClause.Substring(4);
            }          

            string FinalQuery = BaseQuery + WhereClause;
            RunQuery(FinalQuery);

        }


        private void RunQuery(string TrueValueQuery)
        {
            SetupCatalogView();
            statusLbl.Text = "Processing query, please wait...";
            statusLbl.Refresh();
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery(TrueValueQuery);
            if (qr.Rows.Count > 0)
            {
                statusLbl.Text = qr.Rows.Count.ToString() + " rows.";
                statusLbl.Refresh();


                int count = 1;
                foreach (string row in qr.Rows)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = count.ToString();
                    lvi.SubItems.Add(row.Split('|')[0]);
                    lvi.SubItems.Add(row.Split('|')[1]);
                    lvi.SubItems.Add(row.Split('|')[2]);
                    lvi.SubItems.Add(row.Split('|')[3]);
                    lvi.SubItems.Add(row.Split('|')[4]);
                    lvi.SubItems.Add(row.Split('|')[5]);
                    lvi.SubItems.Add(row.Split('|')[6]);
                    lvi.SubItems.Add(row.Split('|')[7]);
                    lvi.SubItems.Add(row.Split('|')[8]);
                    lvi.SubItems.Add(row.Split('|')[9]);
                    lvi.SubItems.Add(row.Split('|')[10]);
                    lvi.SubItems.Add(row.Split('|')[11]);
                    lvi.SubItems.Add(row.Split('|')[12]);
                    lvi.SubItems.Add(row.Split('|')[13]);
                    lvi.SubItems.Add(row.Split('|')[14]);
                    lvi.SubItems.Add(row.Split('|')[15]);
                    lvi.SubItems.Add(row.Split('|')[16]);
                    lvi.SubItems.Add(row.Split('|')[17]);
                    lvi.SubItems.Add(row.Split('|')[18]);
                    lvi.SubItems.Add(row.Split('|')[19]);
                    lvi.SubItems.Add(row.Split('|')[20]);
                    lvi.SubItems.Add(row.Split('|')[21]);
                    lvi.SubItems.Add(row.Split('|')[22]);
                    lvi.SubItems.Add(row.Split('|')[23]);
                    lvi.SubItems.Add(row.Split('|')[24]);
                    lvi.SubItems.Add(row.Split('|')[25]);
                    lvi.SubItems.Add(row.Split('|')[26]);
                    lvi.SubItems.Add(row.Split('|')[27]);
                    lvi.SubItems.Add(row.Split('|')[28]);
                    lvi.SubItems.Add(row.Split('|')[29]);
                    lvi.SubItems.Add(row.Split('|')[30]);
                    fbfLV.Items.Add(lvi);
                        
                    count++;
                }


            }
            else
            {
                statusLbl.Text = "0 Rows.";
                statusLbl.Refresh();
            }
        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
