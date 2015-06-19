using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;

namespace JetDotComInterface
{
    public partial class ItemViewer : Form
    {
        public ItemViewer()
        {
            InitializeComponent();
        }
        public void PopulateInfo(Form1.ItemInfo itemData)
        {
            ClearForm();
           
            excludeFromFeeAdjustmentsLbl.Text += itemData.ExcludeFromFeeAdjustments.ToString();
            fulfillmentTimeLbl.Text += itemData.FulfillmentTime.ToString();
            heightLbl.Text += itemData.PackageHeightInches.ToString("#.##") + "in";
            itemBrandLbl.Text += itemData.Brand;
            jetBrowseNodeIDLbl.Text += itemData.JetBrowseNodeID;
            jetRetailSkuLbl.Text += itemData.JetRetailSKU;
            lengthLbl.Text += itemData.PackageLengthInches.ToString("#.##") + "in";
            manufacturerLbl.Text += itemData.Manufacturer;
            jetPriceLbl.Text += "$" + itemData.JetPrice.ToString();
            mapPriceLbl.Text += "$" + itemData.MAPPrice.ToString("#.##");
            mapTypeLbl.Text += itemData.MAPType;
            merchantIDLbl.Text += itemData.MerchantID;
            merchantSkuIDLbl.Text += itemData.MerchantSKUID;
            merchantSkuLbl.Text += itemData.MerchantSKU;
            msrpLbl.Text += "$" + itemData.MSRP.ToString("#.##");
            multipackQtyLbl.Text += itemData.MultipackQuantity.ToString();
            noReturnFeeAdjustmentLbl.Text += itemData.NoReturnFeeAdjustment.ToString();
            numberOfUnitsForPricePerUnitLbl.Text += itemData.NumberOfUnitsForPricePerUnit.ToString("#.##");
            typeOfUnitsForPricePerUnitLbl.Text += itemData.TypeOfUnitsForPricePerUnit;
            productDescriptionLbl.Text += itemData.ProductDescription;
            productTitleLbl.Text += itemData.ProductTitle;
            prop65Lbl.Text += itemData.Proposition65.ToString();
            shipsAloneLbl.Text += itemData.ShipsAlone.ToString();
            weightLbl.Text += itemData.ShippingWeight.ToString("#.##") + "lbs";
            widthLbl.Text += itemData.PackageWidthInches.ToString("#.##") + "in";
            try
            {
                pictureBox2.Image = GetImageFromUrl(itemData.MainImageURL);
            }
            catch { }
            try
            {
                int lblTop = fulfillmentCentersLbl.Bottom + 5;
                foreach (Form1.FulfillmentCenterPrice fcp in itemData.FulfillmentCenterPrices)
                {
                    Label fulfillmentLbl = new Label();
                    fulfillmentLbl.Parent = pricingInfoBox;
                    fulfillmentLbl.Top = lblTop;
                    fulfillmentLbl.Left = noReturnFeeAdjustmentLbl.Left;                    
                    fulfillmentLbl.Text = fcp.FulfillmentCenterName + ": $" + fcp.FulfillmentCenterItemPrice.ToString("#.##");
                    foreach (Form1.FulfillmentCenterInventory fci in itemData.FulfillmentCenterInventories)
                    {
                        if (fci.FulfillmentCenterID == fcp.FulfillmentCenterID)
                        {
                            fulfillmentLbl.Text += "    QOH: " + fci.FulfillmentCenterQuantity.ToString();
                        }
                    }
                    fulfillmentLbl.AutoSize = true;
                    fulfillmentLbl.Visible = true;
                    lblTop += fulfillmentLbl.Height + 5;
                }
            }
            catch { }

            if (itemData.Bullets.Count > 0)
            {
                bullet1Lbl.Text = "1. " + itemData.Bullets[0];
            }
            if (itemData.Bullets.Count > 1)
            {
                bullet2Lbl.Text = "2. " + itemData.Bullets[1];
            }
            if (itemData.Bullets.Count > 2)
            {
                bullet3Lbl.Text = "3. " + itemData.Bullets[2];
            }
            if (itemData.Bullets.Count > 3)
            {
                bullet4Lbl.Text = "4. " +itemData.Bullets[3];
            }
            if (itemData.Bullets.Count > 4)
            {
                bullet5Lbl.Text = "5. " + itemData.Bullets[4];
            }

            int labelTop = productCodeLbl.Bottom + 5;
            try
            {
                foreach (Form1.StandardProductCode spc in itemData.ProductBarcodes)
                {
                    Label barcodeLbl = new Label();
                    barcodeLbl.Parent = productCodeBox;
                    barcodeLbl.AutoSize = true;
                    barcodeLbl.Top = labelTop;
                    barcodeLbl.Left = productCodeLbl.Left;
                    barcodeLbl.Text = spc.StandardProductCodeNumber;
                    Label barcodeTypeLbl = new Label();
                    barcodeTypeLbl.Parent = productCodeBox;
                    barcodeTypeLbl.AutoSize = true;
                    barcodeTypeLbl.Top = labelTop;
                    barcodeTypeLbl.Left = codeTypeLbl.Left;
                    barcodeTypeLbl.Text = spc.StandardProductCodeType;
                    labelTop += barcodeLbl.Height + 5;
                }
           
            ToolTip mapTypeTip = new ToolTip();
            mapTypeTip.AutoPopDelay = 30000;
            mapTypeTip.InitialDelay = 200;
            mapTypeTip.ReshowDelay = 500;
            mapTypeTip.SetToolTip(mapQuestionPic,Form1.GetMAPTypeDetails(itemData.MAPType));
            mapQuestionPic.Left = mapTypeLbl.Right + 5;
            mapQuestionPic.Top = mapTypeLbl.Top;
            }
            catch { }
        }
        public static Image GetImageFromUrl(string url)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);

            using (HttpWebResponse httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse())
            {
                using (Stream stream = httpWebReponse.GetResponseStream())
                {
                    return Image.FromStream(stream);
                }
            }
        }
        private void ItemViewer_Load(object sender, EventArgs e)
        {
           
            ClearForm();
        }
        private void ClearForm()
        {
            heightLbl.Parent = pictureBox1;
          //  widthLbl.Parent = pictureBox1;
          

            excludeFromFeeAdjustmentsLbl.Text = "Exclude Fee Adjustments: ";
            fulfillmentTimeLbl.Text = "Fulfillment Time: ";
            heightLbl.Text = "Height: ";
            itemBrandLbl.Text = "Item Brand: ";
            jetBrowseNodeIDLbl.Text = "Jet Browse Node ID: ";
            jetRetailSkuLbl.Text = "Jet Retail SKU: ";
            lengthLbl.Text = "Length: ";
            manufacturerLbl.Text = "Manufacturer: ";
            mapPriceLbl.Text = "MAP Price: ";
            mapTypeLbl.Text = "MAP Type: ";
            jetPriceLbl.Text = "Jet.com Price: ";
            merchantIDLbl.Text = "Merchant ID: ";
            merchantSkuIDLbl.Text = "Merchant SKU ID: ";
            merchantSkuLbl.Text = "Merchant SKU: ";
            msrpLbl.Text = "MSRP: ";
            multipackQtyLbl.Text = "Multipack Quantity: ";
            noReturnFeeAdjustmentLbl.Text = "No Return Fee Adjustment: ";
            numberOfUnitsForPricePerUnitLbl.Text = "Number Of Units For Price Per Unit: ";
            typeOfUnitsForPricePerUnitLbl.Text = "Type Of Unit For Price Per Unit: ";
            productDescriptionLbl.Text = "Product Description: ";
            productTitleLbl.Text = "Product Title: ";
            prop65Lbl.Text = "Proposition 65: ";
            shipsAloneLbl.Text = "Ships Alone: ";
            weightLbl.Text = "Weight: ";
            widthLbl.Text = "Width: ";
            bullet1Lbl.Text = "";
            bullet2Lbl.Text = "";
            bullet3Lbl.Text = "";
            bullet4Lbl.Text = "";
            bullet5Lbl.Text = "";
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void mapTypeLbl_Click(object sender, EventArgs e)
        {
          
        }
    }
}
