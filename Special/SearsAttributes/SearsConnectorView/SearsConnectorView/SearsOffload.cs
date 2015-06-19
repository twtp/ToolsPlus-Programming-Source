using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SearsConnectorView
{
    public static class SearsOffload
    {
        public struct SearsOffloadItemData
        {
            public string ItemNumber;
            public string ItemTitle;
            public string ItemShortDescription;
            public string ItemLongDescription;
            public string ItemUPC;
            public string ItemClassID;
            public string ItemMPN;
            public string ItemStandardPrice;
            public string ItemMAPType;
            public string ItemBrand;
            public string ItemShippingLength;
            public string ItemShippingWidth;
            public string ItemShippingHeight;
            public string ItemShippingWeight;
            public string ItemRequiresRefridgeration;
            public string ItemRequiresFreezing;
            public string ItemContainsAlcohol;
            public string ItemContainsTobacco;
            public string ItemImageURL;
            public List<SearsItemAttribute> ItemAttributes;
        }
        public struct SearsItemAttribute
        {
            public string ItemAttributeID;
            public string ItemAttributeValueID;
            public string ItemAttributeValue;
        }

        public static string GetXMLHeader()
        {
            string XML_Header = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n";
            XML_Header += "<catalog-feed xmlns=\"http://seller.marketplace.sears.com/catalog/v17\" xsi:schemaLocation=\"http://seller.marketplace.sears.com/catalog/v17 ../../../../../rest/catalog/import/v17/lmp-item.xsd\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\r\n";
            XML_Header += " <fbm-catalog>\r\n";
            return XML_Header;
        }
        public static string GetXMLFooter()
        {
            string XML_Footer = " </fbm-catalog>\r\n</catalog-feed>\r\n";
            return XML_Footer;
        }
        public static List<SearsOffloadItemData> CreateSearsItemData(List<string> ItemsForSears)
        {
            List<SearsOffloadItemData> SOID_List = new List<SearsOffloadItemData>();
            SearsOffloadItemData soid = new SearsOffloadItemData();
            foreach (string itm in ItemsForSears)
            {
                string TagName = itm.Split('{')[1].Split('}')[0];
                string TValue = itm.Split('{')[2].Split('}')[0];
                
                if (TagName == "Short Description")
                {
                    soid.ItemShortDescription = TValue;
                }
                if (TagName == "Manufacturer Model #")
                {
                    soid.ItemMPN = TValue;
                }
                if (TagName == "Standard Price")
                {
                    soid.ItemStandardPrice = TValue;
                }
                if (TagName == "Brand Name")
                {
                    soid.ItemBrand = TValue;
                }
                if (TagName == "Shipping Length")
                {
                    soid.ItemShippingLength = TValue;
                }
                if (TagName == "Shipping Width")
                {
                    soid.ItemShippingWidth = TValue;
                }
                if (TagName == "Shipping Height")
                {
                    soid.ItemShippingHeight = TValue;
                }
                if (TagName == "Shipping Weight")
                {
                    soid.ItemShippingWeight = TValue;
                }
                if (TagName == "Product Image URL")
                {
                    soid.ItemImageURL = TValue;
                }
                if (TagName == "Item Id")
                {
                    soid.ItemNumber = TValue;
                }
                if (TagName == "UPC")
                {
                    soid.ItemUPC = TValue;
                }
              
               
                
            }
            soid.ItemContainsAlcohol = "false";
            soid.ItemContainsTobacco = "false";
            soid.ItemRequiresFreezing = "false";
            soid.ItemRequiresRefridgeration = "false";
            soid.ItemLongDescription = soid.ItemShortDescription;

            SOID_List.Add(soid);
            return SOID_List;
        }
        public static string AddItemsToXMLSheet(List<SearsOffloadItemData> ItemsInformation)
        {            
            
            string XML_Items = "  <items>\r\n";


            foreach (SearsOffloadItemData soid in ItemsInformation)
            {
                string xml_Item = "   <item item-id=\"" + soid.ItemNumber + "\">\r\n";
                xml_Item += "    <title>" + soid.ItemTitle + "</title>\r\n";
                xml_Item += "    <short-desc>" + soid.ItemShortDescription + "</short-desc>\r\n";
                xml_Item += "    <mature-content>false</mature-content>\r\n";
                xml_Item += "    <upc>" + soid.ItemUPC + "</upc>\r\n";
                xml_Item += "    <item-class id=\"" + soid.ItemClassID + "\"></item-class>\r\n";

                //insert Seller tags here...

                xml_Item += "    <long-desc>" + soid.ItemLongDescription + "</long-desc>\r\n";
                xml_Item += "    <model-number>" + soid.ItemMPN + "</model-number>\r\n";
                xml_Item += "    <standard-price>" + soid.ItemStandardPrice + "</standard-price>\r\n";

                //insert sale data here...

                xml_Item += "    <ship-to-store-method>\r\n     <ship-to-store-eligible>false</ship-to-store-eligible>\r\n    </ship-to-store-method>\r\n";
                
                //insert MAP price indicator here..

                xml_Item += "    <brand>" + soid.ItemBrand + "</brand>\r\n";
                xml_Item += "    <shipping-length>" + soid.ItemShippingLength + "</shipping-length>\r\n";
                xml_Item += "    <shipping-width>" + soid.ItemShippingWidth + "</shipping-width>\r\n";
                xml_Item += "    <shipping-height>" + soid.ItemShippingHeight + "</shipping-height>\r\n";
                xml_Item += "    <shipping-weight>" + soid.ItemShippingWeight + "</shipping-weight>\r\n";
                xml_Item += "    <local-marketplace-flags>\r\n";
                xml_Item += "     <is-restricted>false</is-restricted>\r\n";
                xml_Item += "     <perishable>false</perishable>\r\n";
                xml_Item += "     <requires-refridgeration>" + soid.ItemRequiresRefridgeration + "</requires-refridgeration>\r\n";
                xml_Item += "     <requires-freezing>" + soid.ItemRequiresFreezing + "</requires-freezing>\r\n";
                xml_Item += "     <contains-alcohol>" + soid.ItemContainsAlcohol + "</contains-alcohol>\r\n";
                xml_Item += "     <contains-tobacco>" + soid.ItemContainsTobacco + "</contains-tobacco>\r\n";
                xml_Item += "    </local-marketplace-flags>\r\n";
                xml_Item += "    <image-url>\r\n";
                xml_Item += "     <url>" + soid.ItemImageURL + "</url>\r\n";
                xml_Item += "    </image-url>\r\n";
               
                
                //attributes :( :( :(


                xml_Item += "   </item>\r\n";

                XML_Items += xml_Item;    
            }




            XML_Items += "  </items>\r\n";
            return XML_Items;
        }

    }
}
