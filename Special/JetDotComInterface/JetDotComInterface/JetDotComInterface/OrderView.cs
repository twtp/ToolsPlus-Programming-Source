using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JetDotComInterface
{
    public partial class OrderView : Form
    {
        public OrderView()
        {
            InitializeComponent();
        }

        private void customerPhoneLbl_Click(object sender, EventArgs e)
        {

        }

        private void OrderView_Load(object sender, EventArgs e)
        {
            //ClearOrderLVI();
        }
        private void ClearOrderLVI()
        {
            orderLVI.Items.Clear();
            orderLVI.Columns.Clear();
            orderLVI.Clear();


            ImageList newIL = new ImageList();
            newIL.Images.Add(Properties.Resources.check_3d);
            newIL.Images.Add(Properties.Resources.x_3d);
            newIL.Images.Add(Properties.Resources.warning);
            orderLVI.SmallImageList = newIL;


            ColumnHeader ItemFillable = new ColumnHeader();
            ColumnHeader ItemNumberCol = new ColumnHeader();
            ColumnHeader QuantityCol = new ColumnHeader();
            ColumnHeader CostCol = new ColumnHeader();
            ColumnHeader TaxCol = new ColumnHeader();
            ColumnHeader QuantityOnHandCol = new ColumnHeader();

            ItemFillable.Width = 25;
            ItemFillable.Text = "";
            orderLVI.Columns.Add(ItemFillable);
            ItemNumberCol.Width = orderLVI.Width / 6;
            ItemNumberCol.Text = "Item Number";
            orderLVI.Columns.Add(ItemNumberCol);
            QuantityCol.Width = orderLVI.Width / 8;
            QuantityCol.Text = "Qty";
            orderLVI.Columns.Add(QuantityCol);
            CostCol.Width = orderLVI.Width / 6;
            CostCol.Text = "Cost Per Unit";
            orderLVI.Columns.Add(CostCol);
            TaxCol.Width = orderLVI.Width / 6;
            TaxCol.Text = "Tax";
            orderLVI.Columns.Add(TaxCol);
            QuantityOnHandCol.Width = orderLVI.Width / 6;
            QuantityOnHandCol.Text = "QOH";
            orderLVI.Columns.Add(QuantityOnHandCol);



        }
        public void CalculateOrderSummary()
        {
            int partiallyFillableLines = 0;
            int nonFillableLines = 0;
            int fillableLines = 0;
            int totQty = 0;
            decimal totCost = 0;
            decimal totTax = 0;
            decimal totShipping = 0;
            decimal totShippingTax = 0;

            foreach (ListViewItem lvi in orderLVI.Items)
            {
               
                totQty += int.Parse(lvi.SubItems[2].Text);
               // totCost += decimal.Parse(lvi.SubItems[3].Text);
               // totTax += decimal.Parse(lvi.SubItems[4].Text);
                if (int.Parse(lvi.SubItems[2].Text) <= int.Parse(lvi.SubItems[5].Text))
                {
                    fillableLines++;
                }
                if (int.Parse(lvi.SubItems[2].Text) > int.Parse(lvi.SubItems[5].Text))
                {
                    partiallyFillableLines++;
                }
                if (int.Parse(lvi.SubItems[5].Text) <= 0)
                {
                    nonFillableLines++;
                }

            }

            if (nonFillableLines > 0 & partiallyFillableLines == 0 & fillableLines == 0)
            {
                orderFillableLbl.Text = "This order is not fillable";
            }
            if (partiallyFillableLines > 0 & nonFillableLines == 0)
            {
                orderFillableLbl.Text = "This order is partially fillable";
            }
            if (nonFillableLines == 0 & partiallyFillableLines == 0 & fillableLines > 0)
            {
                orderFillableLbl.Text = "This order is fillable.";
            }

            differentItemsLbl.Text = "Different Items: " + orderLVI.Items.Count;
            totQtyLbl.Text = "Total Qty: " + totQty.ToString();
            totCostLbl.Text = "Total Cost: $" + totCost.ToString("#.##");
            totTaxLbl.Text = "Total Tax: $" + totTax.ToString("#.##");
            //totShippingLbl.Text = "Total Shipping: $" + 
        }
        public void LoadCustomerData(Form1.JetSalesOrder jetSalesOrder)
        {
            ClearForm();
            orderNumberLbl.Text = "Order Number: " + jetSalesOrder.TPSalesOrderNumber;
            jetOrderNumberLbl.Text = "Jet Order Number: " + jetSalesOrder.JetSalesOrderNumber;
            timeOfOrderLbl.Text = "Time Of Order: " + jetSalesOrder.OrderPlacedDate.ToString();
            orderTransmissionTimeLbl.Text = "Order Transmission Time: " + jetSalesOrder.OrderTransmissionDate.ToString();
            customerNameLbl.Text = "Customer Name: " + jetSalesOrder.BuyerName;
            customerPhoneLbl.Text = "Customer Phone: " + jetSalesOrder.BuyerPhoneNumber;
            recipientNameLbl.Text = "Recipient Name: " + jetSalesOrder.ShippingToLocation.ShipTo.Name;
            recipientPhoneLbl.Text = "Recipient Phone: " + jetSalesOrder.ShippingToLocation.ShipTo.PhoneNumber;
            recipientAddress1Lbl.Text = "Recipient Address: " + jetSalesOrder.ShippingToLocation.ShipToAddress.Address1;
            recipientAddress2Lbl.Text = jetSalesOrder.ShippingToLocation.ShipToAddress.Address2;
            recipientCityLbl.Text = "Recipient City: " + jetSalesOrder.ShippingToLocation.ShipToAddress.City;
            recipientStateLbl.Text = "Recipient State: " + jetSalesOrder.ShippingToLocation.ShipToAddress.State;
            recipientZipcodeLbl.Text = "Recipient Zipcode: " + jetSalesOrder.ShippingToLocation.ShipToAddress.Zipcode;

            foreach (Form1.OrderItem itm in jetSalesOrder.OrderedItems)
            {
                ListViewItem newItem = new ListViewItem();
                ListViewItem.ListViewSubItem itemNumber = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem qty = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem cost = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem tax = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem qoh = new ListViewItem.ListViewSubItem();

                itemNumber.Text = itm.TPItemNumber;
                qty.Text = itm.RequestedQty.ToString();
                cost.Text = itm.ItemPricing.BasePrice.ToString("#.##");
                tax.Text = itm.ItemPricing.ItemTax.ToString("#.##");
                List<string> getQOH = Connectivity.SQLCalls.sqlQuery("SELECT QuantityOnHand FROM vInventorySpreadsheet WHERE ITEMNUMBER='" + itm.TPItemNumber + "'");
                int qtyOnHand = 0;
                if (getQOH.Count > 0)
                {
                    qtyOnHand = int.Parse(getQOH[0].Split('|')[0]);
                    qoh.Text = qtyOnHand.ToString();
                }
                else
                {
                    qoh.Text = "0";
                }

                if (itm.ItemPricing.BasePrice > 0)
                {
                    if (qtyOnHand >= itm.RequestedQty)
                    {
                        newItem.ImageIndex = 0;
                    }
                    else
                    {
                        newItem.ImageIndex = 1;
                    }
                }
                else
                {
                    newItem.ImageIndex = 2;
                }
               


                newItem.SubItems.Add(itemNumber);
                newItem.SubItems.Add(qty);
                newItem.SubItems.Add(cost);
                newItem.SubItems.Add(tax);
                newItem.SubItems.Add(qoh);
                orderLVI.Items.Add(newItem);


            }

        }
        public void ClearForm()
        {
            ClearOrderLVI();
            orderNumberLbl.Text = "Order Number: ";
            jetOrderNumberLbl.Text = "Jet Order Number: ";
            timeOfOrderLbl.Text = "Time Of Order: ";
            orderTransmissionTimeLbl.Text = "Order Transmission Time: ";
            customerNameLbl.Text = "Customer Name: ";
            customerPhoneLbl.Text = "Customer Phone: ";
            recipientAddress1Lbl.Text = "Recipient Address: ";
            recipientAddress2Lbl.Text = "";
            recipientCityLbl.Text = "Recipient City: ";
            recipientStateLbl.Text = "Recipient State: ";
            recipientZipcodeLbl.Text = "Recipient Zipcode: ";

        }
    }
}
