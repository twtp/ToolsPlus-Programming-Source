using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace MAS_Variance_Fixer
{
    public partial class Form1 : Form
    {   
        public Form1()
        {
            InitializeComponent();
            
        }
        

        Dictionary<string, bool> MasKits = new Dictionary<string, bool>();
        Dictionary<string, IM_ItemWarehouse> IM_ItemWarehouse_Items = new Dictionary<string, IM_ItemWarehouse>();
        Dictionary<string, SO_InvoiceHeader> SO_NonXDInvoices = new Dictionary<string, SO_InvoiceHeader>();
        Dictionary<string, SO_InvoiceDetail> SO_InvoiceDetails = new Dictionary<string, SO_InvoiceDetail>();
        Dictionary<string, P2_TransactionHeader> P2_NonVTransactions = new Dictionary<string, P2_TransactionHeader>();
        Dictionary<string, P2_TransactionDetail> P2_TransactionDetails = new Dictionary<string, P2_TransactionDetail>();
        Dictionary<string, SO_SalesOrderHeader> SO_SalesOrders = new Dictionary<string, SO_SalesOrderHeader>();
        Dictionary<string, SO_SalesOrderDetail> SO_SalesDetails = new Dictionary<string, SO_SalesOrderDetail>();
        Dictionary<string, IM_ItemCost> IM_ItemCost_Items = new Dictionary<string, IM_ItemCost>();
        Dictionary<string, IM_OverDist> IM_OverDist_Items = new Dictionary<string,IM_OverDist>();
        Dictionary<string, WhseMasVariance> WhseMasVariance_Items = new Dictionary<string, WhseMasVariance>();
        Dictionary<int, int> WarehouseQuantity = new Dictionary<int, int>();
        Dictionary<int, int> WillCallComponents = new Dictionary<int, int>();
        Dictionary<int, int> ComponentExpected = new Dictionary<int, int>();
        Dictionary<int, string> ComponentName = new Dictionary<int, string>();
        List<string> NonZeroWarehouseItems = new List<string>();


        public struct MASColumnHeader
        {
            public string Name;
            public int TypeID;
            public int Size;
        } 
        public struct IM_ItemWarehouse
        {
            public string ItemCode;
            public string WarehouseCode;
            public int QuantityOnHand;
            public int QuantityOnPurchaseOrder;
            public int QuantityOnSalesOrder;
            public int QuantityOnBackOrder;
            public int QuantityInShipping;
        }
        public struct SO_InvoiceHeader
        {
            public string InvoiceNo;
        }
        public struct SO_InvoiceDetail
        {
            public string ItemCode;
            public string WarehouseCode;
            public int QuantityOrdered;
            public int QuantityShipped;
            public int QuantityBackordered;
            public string InvoiceNo;
        }
        public struct P2_TransactionHeader
        {
            public string TransactionNo;
        }
        public struct P2_TransactionDetail
        {
            public string TransactionNo;
            public string ItemCode;
            public string WarehouseCode;
            public int QuantityOrdered;
            public int QuantityShipped;
            public int QuantityBackordered;            
        }
        public struct SO_SalesOrderHeader
        {
            public string SalesOrderNo;
            public string OrderType;
        }
        public struct SO_SalesOrderDetail
        {
            public string SalesOrderNo;
            public string ItemCode;
            public string WarehouseCode;
            public int QuantityOrdered;
            public int QuantityShipped;
            public int QuantityBackordered;
        }
        public struct IM_ItemCost
        {
            public string ItemCode;
            public string WarehouseCode;
            public int QuantityOnHand;
        }
        public struct IM_OverDist
        {
            public string ItemCode;
            public string WarehouseCode;
            public string ReceiptNo;
            public decimal UnitCost;
            public int QuantityOnHand;
            public decimal ExtendedCost;
        }
        public struct WhseMasVariance
        {
            public string ItemCode;
            public int WarehouseQty;
            public int MASQty;
            public int Variance;
            public int WillCallsOutstanding;
        }






        private Dictionary<string,bool> GetMASKits()
        {
            string MASKits = Connectivity.MASCalls.Retrieve("SELECT SalesKitNo FROM IM_SalesKitHeader WHERE SalesKitNo IS NOT NULL");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();

            MAStoLVI(MASKits, ref tmpView);
            Dictionary<string, bool> tmpDict = new Dictionary<string, bool>();
            foreach (ListViewItem lvi in tmpView.Items)
            {
                tmpDict.Add(lvi.Text, true);
            }
            return tmpDict;

        }
        private Dictionary<string,IM_ItemWarehouse> GetItemWarehouseItems()
        {
            string MASKits = Connectivity.MASCalls.Retrieve("SELECT ItemCode, WarehouseCode, QuantityOnHand, QuantityOnPurchaseOrder, QuantityOnSalesOrder, QuantityOnBackOrder, QuantityInShipping FROM IM_ItemWarehouse");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();


            MAStoLVI(MASKits, ref tmpView);
            Dictionary<string, IM_ItemWarehouse> tmpDict = new Dictionary<string, IM_ItemWarehouse>();
            foreach (ListViewItem lvi in tmpView.Items)
            {
                IM_ItemWarehouse tmpItm = new IM_ItemWarehouse();
                tmpItm.ItemCode = lvi.Text;
                string tmpName = lvi.Text;
                tmpItm.WarehouseCode = lvi.SubItems[1].Text;
                tmpItm.QuantityOnHand = (int)(decimal.Parse(lvi.SubItems[2].Text));
                tmpItm.QuantityOnPurchaseOrder = (int)(decimal.Parse(lvi.SubItems[3].Text));
                tmpItm.QuantityOnSalesOrder = (int)(decimal.Parse(lvi.SubItems[4].Text));
                tmpItm.QuantityOnBackOrder = (int)(decimal.Parse(lvi.SubItems[5].Text));
                tmpItm.QuantityInShipping = (int)(decimal.Parse(lvi.SubItems[6].Text));
                if (tmpItm.ItemCode == "FRE91-102" | tmpItm.ItemCode == "BOSMFN-201")
                {
                    if (tmpItm.QuantityInShipping < 0)
                    {

                    }
                }
                
                if (tmpItm.WarehouseCode != "000")
                {
                    tmpName = "(" + tmpItm.WarehouseCode + ")" + tmpItm.ItemCode;
                    NonZeroWarehouseItems.Add(tmpName);
                }

               // if (MasKits.ContainsKey(tmpItm.ItemCode) == false)
              //  { 
                    //if 2 keys are the same, merge???


                if (tmpDict.ContainsKey(tmpName) == false)
                {
                    tmpDict.Add(tmpName, tmpItm);
                }
                else
                {

                    int QOH = tmpDict[lvi.Text].QuantityOnHand + tmpItm.QuantityOnHand;
                    int QOSO = tmpDict[lvi.Text].QuantityOnSalesOrder + tmpItm.QuantityOnSalesOrder;
                    int QOPO = tmpDict[lvi.Text].QuantityOnPurchaseOrder + tmpItm.QuantityOnPurchaseOrder;
                    int QOBO = tmpDict[lvi.Text].QuantityOnBackOrder + tmpItm.QuantityOnBackOrder;
                    int QIS = tmpDict[lvi.Text].QuantityInShipping + tmpItm.QuantityInShipping;
                    tmpItm.QuantityOnHand = QOH;
                    tmpItm.QuantityOnSalesOrder = QOSO;
                    tmpItm.QuantityOnPurchaseOrder = QOPO;
                    tmpItm.QuantityOnBackOrder = QOBO;
                    tmpItm.QuantityInShipping = QIS;

                    tmpDict.Remove(tmpName);
                    tmpDict.Add(tmpName, tmpItm);
                }

                    
               // }
            }
            return tmpDict;
        }
        private Dictionary<string, SO_InvoiceHeader> GetNonXDInvoices()
        {
            string MASKits = Connectivity.MASCalls.Retrieve("SELECT InvoiceNo FROM SO_InvoiceHeader WHERE InvoiceType<>'XD'");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();

            MAStoLVI(MASKits, ref tmpView);
            Dictionary<string,SO_InvoiceHeader> tmpDict = new Dictionary<string, SO_InvoiceHeader>();
            foreach (ListViewItem lvi in tmpView.Items)
            {              

                SO_InvoiceHeader soTmp = new SO_InvoiceHeader();
                soTmp.InvoiceNo = lvi.Text;               
                tmpDict.Add(lvi.Text, soTmp);
            }
            return tmpDict;
        }
        private Dictionary<string, SO_InvoiceDetail> GetInvoicedItems()
        {
            Dictionary<string, SO_InvoiceDetail> tmpDict = new Dictionary<string, SO_InvoiceDetail>();
            foreach (KeyValuePair<string,SO_InvoiceHeader> kvp in SO_NonXDInvoices)
            {
                string query = "SELECT ItemCode, WarehouseCode, QuantityShipped, QuantityBackordered FROM SO_InvoiceDetail WHERE InvoiceNo='{INVOICENO}' AND ItemType='1' AND WarehouseCode='000'";
                query = query.Replace("{INVOICENO}", kvp.Key.ToString());
                if (kvp.Key.ToString() == "0071879")
                {

                }
                string MASKits = Connectivity.MASCalls.Retrieve(query);
                ListView tmpView = new ListView();
                tmpView.View = View.Details;
                tmpView.FullRowSelect = true;
                tmpView.GridLines = true;
                tmpView.Items.Clear();
                tmpView.Columns.Clear();
                tmpView.Clear();
                tmpView.Refresh();

                MAStoLVI(MASKits, ref tmpView);
               
                foreach (ListViewItem lvi in tmpView.Items)
                {
                    SO_InvoiceDetail soTmp = new SO_InvoiceDetail();
                    soTmp.InvoiceNo = kvp.Key.ToString();
                    soTmp.ItemCode = lvi.Text;
                    if (soTmp.ItemCode == "FRE91-102" | soTmp.ItemCode == "BOSMFN-201")
                    {
                        if (soTmp.QuantityShipped < 0)
                        {

                        }
                    }

                    soTmp.WarehouseCode = lvi.SubItems[1].Text;
                    if (soTmp.InvoiceNo.Length > 0)
                    {
                        soTmp.QuantityShipped = (int)(decimal.Parse(lvi.SubItems[2].Text));
                    }
                    else
                    {
                        soTmp.QuantityOrdered += (int)(decimal.Parse(lvi.SubItems[2].Text));
                        soTmp.QuantityBackordered += (int)(decimal.Parse(lvi.SubItems[3].Text));
                        soTmp.QuantityShipped += (int)(decimal.Parse(lvi.SubItems[2].Text));

                    }
                   

                   // if (MasKits.ContainsKey(soTmp.ItemCode) == false)
                   // {
                    if (MasKits.ContainsKey(soTmp.ItemCode) == false)
                    {

                        if (tmpDict.ContainsKey(lvi.Text)==false)
                        {
                            tmpDict.Add(lvi.Text, soTmp);
                        }
                        else
                        {
                            //int QOH = tmpDict[lvi.Text].QuantityOnHand + tmpItm.QuantityOnHand;
                            //int QOSO = tmpDict[lvi.Text].QuantityOnSalesOrder + tmpItm.QuantityOnSalesOrder;
                            //int QOPO = tmpDict[lvi.Text].QuantityOnPurchaseOrder + tmpItm.QuantityOnPurchaseOrder;
                            int QOBO = tmpDict[lvi.Text].QuantityBackordered + soTmp.QuantityBackordered;
                            int QIS = tmpDict[lvi.Text].QuantityShipped + soTmp.QuantityShipped;

                            soTmp.QuantityBackordered = QOBO;
                            soTmp.QuantityShipped = QIS;

                            tmpDict.Remove(lvi.Text);
                            tmpDict.Add(lvi.Text, soTmp);
                        }

                    }
                   // }

                }
            }



            return tmpDict;
        }
        private Dictionary<string, P2_TransactionHeader> GetNonVTransactions()
        {
            Dictionary<string, P2_TransactionHeader> tmpDict = new Dictionary<string, P2_TransactionHeader>();
           
                string MASKits = Connectivity.MASCalls.Retrieve("SELECT TransactionNo FROM P2_TransactionHeader WHERE TransactionCode IN ('01','02') AND TransactionStatus<>'V' AND WarehouseCode='000'");
                ListView tmpView = new ListView();
                tmpView.View = View.Details;
                tmpView.FullRowSelect = true;
                tmpView.GridLines = true;
                tmpView.Items.Clear();
                tmpView.Columns.Clear();
                tmpView.Clear();
                tmpView.Refresh();

                MAStoLVI(MASKits, ref tmpView);

                foreach (ListViewItem lvi in tmpView.Items)
                {
                    P2_TransactionHeader p2Tmp = new P2_TransactionHeader();
                    p2Tmp.TransactionNo = lvi.Text;
                    tmpDict.Add(lvi.Text, p2Tmp);
                }  

            return tmpDict;
        }
        private Dictionary<string, P2_TransactionDetail> GetP2TransactionItems()
        {

            Dictionary<string, P2_TransactionDetail> tmpDict = new Dictionary<string, P2_TransactionDetail>();
            foreach (KeyValuePair<string, P2_TransactionHeader> kvp in P2_NonVTransactions)
            {
                string query = "SELECT ItemCode, WarehouseCode, QuantityShipped, QuantityBackordered FROM P2_TransactionDetail WHERE TransactionNo='{TRANSACTIONNO}' AND ItemType='1'";
                query = query.Replace("{TRANSACTIONNO}", kvp.Key.ToString());
                string MASKits = Connectivity.MASCalls.Retrieve(query);
                ListView tmpView = new ListView();
                tmpView.View = View.Details;
                tmpView.FullRowSelect = true;
                tmpView.GridLines = true;
                tmpView.Items.Clear();
                tmpView.Columns.Clear();
                tmpView.Clear();
                tmpView.Refresh();

                MAStoLVI(MASKits, ref tmpView);

                foreach (ListViewItem lvi in tmpView.Items)
                {
                    P2_TransactionDetail soTmp = new P2_TransactionDetail();
                    soTmp.TransactionNo = kvp.Key.ToString();
                    soTmp.ItemCode = lvi.Text;

                    if (soTmp.ItemCode == "FRE91-102" | soTmp.ItemCode == "BOSMFN-201")
                    {
                        if (soTmp.QuantityShipped < 0)
                        {

                        }
                    }
                    soTmp.WarehouseCode = lvi.SubItems[1].Text;
                    soTmp.QuantityOrdered = (int)(decimal.Parse(lvi.SubItems[2].Text));
                    soTmp.QuantityShipped = (int)(decimal.Parse(lvi.SubItems[2].Text));
                    soTmp.QuantityBackordered = (int)(decimal.Parse(lvi.SubItems[3].Text));
                   
                    if (MasKits.ContainsKey(soTmp.ItemCode) == false)
                    {
                        if (tmpDict.ContainsKey(lvi.Text) == false)
                        {
                            
                                tmpDict.Add(lvi.Text, soTmp);
                        }
                        else
                        {
                                int QOBO = tmpDict[lvi.Text].QuantityBackordered + soTmp.QuantityBackordered;
                                int QIS = tmpDict[lvi.Text].QuantityShipped + soTmp.QuantityShipped;
                                int QOSO = tmpDict[lvi.Text].QuantityOrdered + soTmp.QuantityOrdered;
                                soTmp.QuantityBackordered = QOBO;
                                soTmp.QuantityShipped = QIS;
                                soTmp.QuantityOrdered = QOSO;

                                tmpDict.Remove(lvi.Text);
                                tmpDict.Add(lvi.Text, soTmp);
                         }
                     }

                    

                }
            }



            return tmpDict;

        }
        private Dictionary<string, SO_SalesOrderHeader> GetSalesOrders()
        {
            Dictionary<string, SO_SalesOrderHeader> tmpDict = new Dictionary<string, SO_SalesOrderHeader>();

            string MASKits = Connectivity.MASCalls.Retrieve("SELECT SalesOrderNo, OrderType FROM SO_SalesOrderHeader WHERE WarehouseCode='000'");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();

            MAStoLVI(MASKits, ref tmpView);

            foreach (ListViewItem lvi in tmpView.Items)
            {
                
                    SO_SalesOrderHeader soTmp = new SO_SalesOrderHeader();
                    soTmp.SalesOrderNo = lvi.Text;
                    soTmp.OrderType = lvi.SubItems[1].Text;
                    tmpDict.Add(lvi.Text, soTmp);
                
            }
            

            return tmpDict;
        }
        private Dictionary<string, SO_SalesOrderDetail> GetSalesOrderItems()
        {
            Dictionary<string, SO_SalesOrderDetail> tmpDict = new Dictionary<string, SO_SalesOrderDetail>();
            foreach (KeyValuePair<string, SO_SalesOrderHeader> kvp in SO_SalesOrders)
            {
                if (kvp.Key.ToString() == "0071879")
                {

                }
                string query = "SELECT ItemCode, WarehouseCode, QuantityOrdered, QuantityShipped, QuantityBackordered FROM SO_SalesOrderDetail WHERE SalesOrderNo='{SALESORDERNO}' AND ItemType='1'";
                query = query.Replace("{SALESORDERNO}", kvp.Key.ToString());
                string MASKits = Connectivity.MASCalls.Retrieve(query);
                ListView tmpView = new ListView();
                tmpView.View = View.Details;
                tmpView.FullRowSelect = true;
                tmpView.GridLines = true;
                tmpView.Items.Clear();
                tmpView.Columns.Clear();
                tmpView.Clear();
                tmpView.Refresh();

                MAStoLVI(MASKits, ref tmpView);

                foreach (ListViewItem lvi in tmpView.Items)
                {
                    SO_SalesOrderDetail soTmp = new SO_SalesOrderDetail();
                    soTmp.SalesOrderNo = kvp.Key.ToString();
                    soTmp.ItemCode = lvi.Text;
                    if (soTmp.ItemCode == "FRE91-102" | soTmp.ItemCode == "BOSMFN-201")
                    {
                        if (soTmp.QuantityShipped < 0)
                        {

                        }
                    }
                    soTmp.WarehouseCode = lvi.SubItems[1].Text;
                    soTmp.QuantityOrdered = (int)(decimal.Parse(lvi.SubItems[2].Text));
                    //soTmp.QuantityShipped = (int)(decimal.Parse(lvi.SubItems[3].Text));
                    soTmp.QuantityBackordered = (int)(decimal.Parse(lvi.SubItems[4].Text));



                    //soTmp.QuantityOrdered = soTmp.QuantityOrdered - soTmp.QuantityShipped;
                    soTmp.QuantityOrdered -= (int)(decimal.Parse(lvi.SubItems[3].Text));
                   
                    
                    if (lvi.SubItems[1].Text.ToUpper() != "Q" & MasKits.ContainsKey(soTmp.ItemCode)==false)
                    {

                        //soTmp.QuantityOrdered = soTmp.QuantityOrdered - soTmp.QuantityShipped;

                        if (kvp.Value.OrderType.ToUpper() != "S")
                        {

                            soTmp.QuantityOrdered -= soTmp.QuantityBackordered;
                        }
                        else
                        {
                            soTmp.QuantityBackordered = 0;
                        }




                            if (tmpDict.ContainsKey(lvi.Text) == false)
                            {

                                tmpDict.Add(lvi.Text, soTmp);

                            }
                            else
                            {
                                int QOSO = tmpDict[lvi.Text].QuantityOrdered + soTmp.QuantityOrdered;
                                int QOBO = tmpDict[lvi.Text].QuantityBackordered + soTmp.QuantityBackordered;
                               // int QIS = tmpDict[lvi.Text].QuantityShipped + soTmp.QuantityShipped;

                                soTmp.QuantityOrdered = QOSO;
                                soTmp.QuantityBackordered = QOBO;
                                //soTmp.QuantityShipped = QIS;
                                //to my knowledge, Qty Shipped is not evaluated in this method...

                                tmpDict.Remove(lvi.Text);
                                tmpDict.Add(lvi.Text, soTmp);

                            }
                        

                    }
                    
                    






                       

                    

                }
            }



            return tmpDict;
        }
        private Dictionary<string,IM_ItemCost> GetItemCostItems()
        {
            string MASKits = Connectivity.MASCalls.Retrieve("SELECT ItemCode, WarehouseCode, SUM(QuantityOnHand) FROM IM_ItemCost GROUP BY ItemCode, WarehouseCode");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();


            MAStoLVI(MASKits, ref tmpView);
            Dictionary<string, IM_ItemCost> tmpDict = new Dictionary<string, IM_ItemCost>();
            foreach (ListViewItem lvi in tmpView.Items)
            {
                IM_ItemCost tmpItm = new IM_ItemCost();
                tmpItm.ItemCode = lvi.Text;
                string tmpNaming = lvi.Text;
                if (tmpItm.ItemCode == "PRZPR-2700")
                {

                }
                tmpItm.WarehouseCode = lvi.SubItems[1].Text;
                tmpItm.QuantityOnHand = (int)(decimal.Parse(lvi.SubItems[2].Text));

                
                if (tmpItm.WarehouseCode != "000")
                {
                    tmpNaming = "(" + tmpItm.WarehouseCode + ")" + tmpItm.ItemCode;
                    NonZeroWarehouseItems.Add(tmpNaming);
                }

                if (MasKits.ContainsKey(tmpNaming) == false)
                {
                    //if 2 keys are the same, merge???
                    try
                    {
                        tmpDict.Add(tmpNaming, tmpItm);
                    }
                    catch
                    {
                        int QOH = tmpDict[tmpNaming].QuantityOnHand + tmpItm.QuantityOnHand;

                        tmpItm.QuantityOnHand = QOH;


                        tmpDict.Remove(tmpNaming);
                        tmpDict.Add(tmpNaming, tmpItm);


                    }
                }
            }
            return tmpDict;
        }
        private Dictionary<string, IM_OverDist> GetItemCostOverDistItems()
        {
             string MASKits = Connectivity.MASCalls.Retrieve("SELECT ItemCode, WarehouseCode, ReceiptNo, UnitCost, QuantityOnHand, ExtendedCost FROM IM_ItemCost WHERE UnitCost*QuantityOnHand<>ExtendedCost");
            ListView tmpView = new ListView();
            tmpView.View = View.Details;
            tmpView.FullRowSelect = true;
            tmpView.GridLines = true;
            tmpView.Items.Clear();
            tmpView.Columns.Clear();
            tmpView.Clear();
            tmpView.Refresh();


            MAStoLVI(MASKits, ref tmpView);
            Dictionary<string, IM_OverDist> tmpDict = new Dictionary<string, IM_OverDist>();
            foreach (ListViewItem lvi in tmpView.Items)
            {
                IM_OverDist tmpItm = new IM_OverDist();
                tmpItm.ItemCode = lvi.Text;
                tmpItm.WarehouseCode = lvi.SubItems[1].Text;
                tmpItm.ReceiptNo = lvi.SubItems[2].Text;
                tmpItm.UnitCost = decimal.Parse(lvi.SubItems[3].Text);
                tmpItm.QuantityOnHand = (int)(decimal.Parse(lvi.SubItems[4].Text));
                tmpItm.ExtendedCost = decimal.Parse(lvi.SubItems[5].Text);

                
                    if (tmpItm.ItemCode=="JOR3718-HD")
                    {

                    }
                if (MasKits.ContainsKey(tmpItm.ItemCode) == false)
                { 
                    //if 2 keys are the same, merge???
                    try
                    {
                        tmpDict.Add(lvi.Text, tmpItm);
                    }
                    catch
                    {
                        tmpDict.Remove(lvi.Text);
                        tmpDict.Add(lvi.Text,tmpItm);
                    }
                }
            }
            return tmpDict;
        }
        private Dictionary<string, WhseMasVariance> GetWhseMasVarianceItems()

        {
            List<string> WhseItems = Connectivity.SQLCalls.sqlQuery("SELECT InventoryMaster.ItemNumber, InventoryMaster.Hide, InventoryComponentMap.ComponentID, InventoryComponentMap.Quantity, InventoryComponents.Component FROM InventoryMaster INNER JOIN InventoryComponentMap ON InventoryMaster.ItemID=InventoryComponentMap.ItemID INNER JOIN InventoryComponents ON InventoryComponentMap.ComponentID=InventoryComponents.ID WHERE InventoryMaster.ItemNumber<'Z%'");


            Dictionary<string, WhseMasVariance> tmpDict = new Dictionary<string, WhseMasVariance>();

  
            foreach (string str in WhseItems)
            {
                if (str.Split('|')[0] == "NOV57074")
                {

                }
                if (str.Split('|')[0] == "NOV57080")
                {

                }
                   
                 if (IM_ItemWarehouse_Items.ContainsKey(str.Split('|')[0]) ==true)
                 {
                     string ItemNumber = str.Split('|')[0];
                     bool Hidden = str.Split('|')[1] == "1" ? true : false;
                     int CompID = int.Parse(str.Split('|')[2]);
                     int Quantity = int.Parse(str.Split('|')[3]) * IM_ItemWarehouse_Items[ItemNumber].QuantityOnHand;
                     string CompName = str.Split('|')[4];
                     int existingQty = ComponentExpected.ContainsKey(int.Parse(str.Split('|')[2])) == true ? ComponentExpected[CompID] : 0;                     
                     if (CompID == 44297)
                     {

                     }
                     if (ComponentName.ContainsKey(CompID) == false)
                     {
                         ComponentName.Add(CompID, CompName);
                     }
                     try
                     {
                         ComponentExpected.Remove(CompID);
                     }
                     catch { }
                     ComponentExpected.Add(CompID, existingQty + Quantity);




                    /* if (ComponentExpected.ContainsKey(int.Parse(str.Split('|')[2])) == true)
                     {
                         KeyValuePair<int, int> tmpVal = new KeyValuePair<int, int>(int.Parse(str.Split('|')[2]), ComponentExpected[int.Parse(str.Split('|')[2])] + (int.Parse(str.Split('|')[3]) * IM_ItemWarehouse_Items[str.Split('|')[0]].QuantityOnHand));
                         ComponentExpected.Remove(int.Parse(str.Split('|')[2]));
                         ComponentExpected.Add(tmpVal.Key, tmpVal.Value);
                     }
                     else
                     {
                         KeyValuePair<int, int> tmpVal = new KeyValuePair<int, int>(int.Parse(str.Split('|')[2]), int.Parse(str.Split('|')[3]));
                         ComponentExpected.Add(tmpVal.Key, tmpVal.Value);
                     }
                     if (ComponentName.ContainsKey(int.Parse(str.Split('|')[2])) == true)
                     {
                         KeyValuePair<int, string> tmpVal = new KeyValuePair<int, string>(int.Parse(str.Split('|')[2]),str.Split('|')[4]);
                         ComponentName.Remove(int.Parse(str.Split('|')[2]));
                         ComponentName.Add(tmpVal.Key, tmpVal.Value);
                     }
                     else
                     {
                         KeyValuePair<int, string> tmpVal = new KeyValuePair<int, string>(int.Parse(str.Split('|')[2]), str.Split('|')[4]);
                         ComponentName.Add(tmpVal.Key, tmpVal.Value);
                     }*/

                 }
                 else
                 {
                    if (str.Split('|')[1] != "1")
                    {
                        //warn missing item in IM_ItemWarehouse
                    }
                    //IM_ItemWarehouse_Items[str.Split('|')[0]].QuantityOnHand = 0;
                 }
            }
            
            List<string> tmpWillCall = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID, QuantityRequired-QuantityFilled FROM WhseTaskWillCallLines WHERE QuantityFilled<>QuantityRequired");
            foreach (string wc in tmpWillCall)
            {
                int CID = int.Parse(wc.Split('|')[0]);
                int QTY = int.Parse(wc.Split('|')[1]);

                if (WillCallComponents.ContainsKey(CID) == true)
                {
                    QTY += WillCallComponents[CID];
                    WillCallComponents.Remove(CID);
                    WillCallComponents.Add(CID, QTY);
                }
                else
                {
                    WillCallComponents.Add(CID, QTY);
                }
            }

            List<string> tmpWhseQty = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID, SUM(Quantity) AS TotalQuantity FROM LocationContents GROUP BY ComponentID");
            foreach (string wq in tmpWhseQty)
            {
                int CID = int.Parse(wq.Split('|')[0]);
                int WhseQty = int.Parse(wq.Split('|')[1]);
                WarehouseQuantity.Add(CID, WhseQty);
            }







            return tmpDict; 
        }

    
        public void MAStoLVI(string responseData, ref ListView outputLV)
        {
            string ColumnData = "";
            string RowData = "";

            try
            {
                ColumnData = responseData.Split('[')[1].Split(']')[0];
                RowData = responseData.Split(new string[] { "rows\":[" }, StringSplitOptions.None)[1].Split(new string[] { "]}" }, StringSplitOptions.None)[0];
            }
            catch { //MessageBox.Show("The last query errored as it returned nothing.");
            return; }

            string[] Columns = ColumnData.Split(new string[] { "{\"name\":\"" }, StringSplitOptions.RemoveEmptyEntries);
            for (int x = 0; x < Columns.GetLength(0); x++)
            {
                Columns[x] = Columns[x].Split('}')[0];
            }

            string[] Rows = RowData.Split(new string[] { "[" }, StringSplitOptions.RemoveEmptyEntries);
            for (int y = 0; y < Rows.GetLength(0); y++)
            {
                Rows[y] = Rows[y].Split(']')[0];
            }

            List<MASColumnHeader> finalColumns = new List<MASColumnHeader>();
            List<List<string>> finalRows = new List<List<string>>();
            foreach (string data in Columns)
            {
                MASColumnHeader mch = new MASColumnHeader();
                string Name, TypeID, Size = "";
                Name = data.Split('"')[0];
                TypeID = data.Split(new string[] { "typeID\":" }, StringSplitOptions.None)[1].Split(',')[0];
                Size = data.Split(new string[] { "size\":" }, StringSplitOptions.None)[1].Split(',')[0];
                mch.Name = Name;
                mch.TypeID = int.Parse(TypeID);
                mch.Size = int.Parse(Size);
                finalColumns.Add(mch);
            }
            foreach (string data in Rows)
            {
                List<string> fields = new List<string>();
                bool splitOffset = false;
                foreach (string field in data.Split(','))
                {
                    if (splitOffset == false)
                    {
                        splitOffset = true;
                        fields.Add(field.Replace("\"", ""));
                    }
                    else
                    {
                        fields.Add(field.Replace("\"", ""));
                    }
                }
                finalRows.Add(fields);
            }

            List<ListViewItem> resultsRows = new List<ListViewItem>();
            List<ColumnHeader> resultsColumns = new List<ColumnHeader>();

            outputLV.Items.Clear();
            outputLV.Columns.Clear();
            outputLV.Clear();
            outputLV.Refresh();
            outputLV.GridLines = true;
            outputLV.FullRowSelect = true;
            outputLV.View = View.Details;

            foreach (MASColumnHeader mch in finalColumns)
            {
                ColumnHeader tmpCol = new ColumnHeader();
                tmpCol.Text = mch.Name;
                tmpCol.Tag = mch;
                tmpCol.AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize);
                resultsColumns.Add(tmpCol);
            }
            foreach (List<string> fieldList in finalRows)
            {
                ListViewItem tmpLVI = new ListViewItem();
                bool firstCol = true;
                foreach (string fld in fieldList)
                {
                    if (firstCol == false)
                    {
                        tmpLVI.SubItems.Add(fld);
                    }
                    else
                    {
                        tmpLVI.Text = fld;
                        firstCol = false;
                    }

                }
                resultsRows.Add(tmpLVI);
            }

            foreach (ColumnHeader ch in resultsColumns)
            {
                outputLV.Columns.Add(ch);
            }
            foreach (ListViewItem lvi in resultsRows)
            {
                outputLV.Items.Add(lvi);
            }


        }

        public void SetupLVI()
        {
            if (this.InvokeRequired == true)
            {
                masLVI.Invoke((MethodInvoker)(() =>
                {
                    masLVI.Items.Clear();
                    masLVI.Columns.Clear();
                    masLVI.Clear();
                    masLVI.View = View.Details;
                    masLVI.FullRowSelect = true;
                    masLVI.GridLines = true;
                    masLVI.Refresh();

                }));

                whseLVI.Invoke((MethodInvoker)(() =>
                    {
                        whseLVI.Items.Clear();
                        whseLVI.Columns.Clear();
                        whseLVI.Clear();
                        whseLVI.View = View.Details;
                        whseLVI.FullRowSelect = true;
                        whseLVI.GridLines = true;
                        whseLVI.Refresh();

                    }));
                overdistLVI.Invoke((MethodInvoker)(() =>
                    {

                        overdistLVI.Items.Clear();
                        overdistLVI.Columns.Clear();
                        overdistLVI.Clear();
                        overdistLVI.View = View.Details;
                        overdistLVI.FullRowSelect = true;
                        overdistLVI.GridLines = true;
                        overdistLVI.Refresh();

                    }));

            }
            else
            {
                masLVI.Items.Clear();
                masLVI.Columns.Clear();
                masLVI.Clear();
                whseLVI.Items.Clear();
                whseLVI.Columns.Clear();
                whseLVI.Clear();
                masLVI.View = View.Details;
                masLVI.FullRowSelect = true;
                masLVI.GridLines = true;
                masLVI.Refresh();
                whseLVI.View = View.Details;
                whseLVI.FullRowSelect = true;
                whseLVI.GridLines = true;
                whseLVI.Refresh();

                overdistLVI.Items.Clear();
                overdistLVI.Columns.Clear();
                overdistLVI.Clear();
                overdistLVI.View = View.Details;
                overdistLVI.FullRowSelect = true;
                overdistLVI.GridLines = true;
                overdistLVI.Refresh();
            }


            ColumnHeader warehouseCol = new ColumnHeader();
            ColumnHeader fieldCol = new ColumnHeader();
            ColumnHeader itemCol = new ColumnHeader();
            ColumnHeader inventoryQtyCol = new ColumnHeader();
            ColumnHeader moduleQtyCol = new ColumnHeader();

            ColumnHeader componentCol = new ColumnHeader();
            ColumnHeader masQtyCol = new ColumnHeader();
            ColumnHeader whseQtyCol = new ColumnHeader();
            ColumnHeader varianceCol = new ColumnHeader();
            ColumnHeader willcallsCol = new ColumnHeader();

            ColumnHeader itemcodeCol = new ColumnHeader();
            ColumnHeader whseCodeCol = new ColumnHeader();
            ColumnHeader receiptCol = new ColumnHeader();
            ColumnHeader qtyCol = new ColumnHeader();
            ColumnHeader unitcostCol = new ColumnHeader();
            ColumnHeader extensionCol = new ColumnHeader();
            ColumnHeader expectedCol = new ColumnHeader();



            warehouseCol.Text = "Whse";
            fieldCol.Text = "Problem";
            itemCol.Text = "Item Number";
            inventoryQtyCol.Text = "Inventory Qty";
            moduleQtyCol.Text = "Module Qty";

            componentCol.Text = "Component";
            masQtyCol.Text = "MAS Qty";
            whseQtyCol.Text = "Whse Qty";
            varianceCol.Text = "Variance";
            willcallsCol.Text = "Will Calls Outstanding";

            itemcodeCol.Text = "ItemCode";
            whseCodeCol.Text = "Whse";
            receiptCol.Text = "Receipt No";
            qtyCol.Text = "Qty";
            unitcostCol.Text = "Unit Cost";
            extensionCol.Text = "Extension";
            expectedCol.Text = "Expected";

            warehouseCol.Width = masLVI.ClientSize.Width / 12;
            fieldCol.Width = masLVI.ClientSize.Width / 12;
            itemCol.Width = masLVI.ClientSize.Width / 6;
            inventoryQtyCol.Width = masLVI.ClientSize.Width / 12;
            moduleQtyCol.Width = masLVI.ClientSize.Width / 12;

            componentCol.Width = (int)((float)(whseLVI.ClientSize.Width / 7));
            masQtyCol.Width = whseLVI.ClientSize.Width / 14;
            whseQtyCol.Width = whseLVI.ClientSize.Width / 14;
            varianceCol.Width = whseLVI.ClientSize.Width / 14;
            willcallsCol.Width = (int)((float)(whseLVI.ClientSize.Width / 7));

            itemcodeCol.Width = overdistLVI.ClientSize.Width / 12;
            whseCodeCol.Width = overdistLVI.ClientSize.Width / 12;
            receiptCol.Width = overdistLVI.ClientSize.Width / 12;
            qtyCol.Width = overdistLVI.ClientSize.Width / 20;
            unitcostCol.Width = overdistLVI.ClientSize.Width / 20;
            extensionCol.Width = overdistLVI.ClientSize.Width / 16;
            expectedCol.Width = overdistLVI.ClientSize.Width / 16;

            if (this.InvokeRequired == true)
            {
                masLVI.Invoke((MethodInvoker)(() =>
                    {

                        masLVI.Columns.Add(warehouseCol);
                        masLVI.Columns.Add(fieldCol);
                        masLVI.Columns.Add(itemCol);
                        masLVI.Columns.Add(inventoryQtyCol);
                        masLVI.Columns.Add(moduleQtyCol);
                        masLVI.Refresh();

                    }));

                whseLVI.Invoke((MethodInvoker)(() =>
                    {
                        whseLVI.Columns.Add(componentCol);
                        whseLVI.Columns.Add(masQtyCol);
                        whseLVI.Columns.Add(whseQtyCol);
                        whseLVI.Columns.Add(varianceCol);
                        whseLVI.Columns.Add(willcallsCol);
                        whseLVI.Refresh();

                    }));

                overdistLVI.Invoke((MethodInvoker)(() =>
                    {
                        overdistLVI.Columns.Add(itemcodeCol);
                        overdistLVI.Columns.Add(whseCodeCol);
                        overdistLVI.Columns.Add(receiptCol);
                        overdistLVI.Columns.Add(qtyCol);
                        overdistLVI.Columns.Add(unitcostCol);
                        overdistLVI.Columns.Add(extensionCol);
                        overdistLVI.Columns.Add(expectedCol);
                        overdistLVI.Refresh();

                    }));

            }
            else
            {
                masLVI.Columns.Add(warehouseCol);
                masLVI.Columns.Add(fieldCol);
                masLVI.Columns.Add(itemCol);
                masLVI.Columns.Add(inventoryQtyCol);
                masLVI.Columns.Add(moduleQtyCol);

                whseLVI.Columns.Add(componentCol);
                whseLVI.Columns.Add(masQtyCol);
                whseLVI.Columns.Add(whseQtyCol);
                whseLVI.Columns.Add(varianceCol);
                whseLVI.Columns.Add(willcallsCol);

                overdistLVI.Columns.Add(itemcodeCol);
                overdistLVI.Columns.Add(whseCodeCol);
                overdistLVI.Columns.Add(receiptCol);
                overdistLVI.Columns.Add(qtyCol);
                overdistLVI.Columns.Add(unitcostCol);
                overdistLVI.Columns.Add(extensionCol);
                overdistLVI.Columns.Add(expectedCol);

                masLVI.Refresh();
                whseLVI.Refresh();
                overdistLVI.Refresh();
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.ClientSize = new System.Drawing.Size(label1.Left,this.ClientSize.Height);
            this.BringToFront();
            this.Refresh();
            Shown += Form1_Shown;
            //PrefetchData();
        }

        void Form1_Shown(object sender, EventArgs e)
        {
            this.Refresh();
            groupBox2.Refresh();
            //new System.Threading.Thread(() =>
            //{
                //Type myType = this.GetType();
                //MethodInfo prefetchMethod = myType.GetMethod("PrefetchData");                
                //prefetchMethod.Invoke(this, null);
                
                //prefetchMethod = myType.GetMethod("button2_Click(sender, e)");
                //prefetchMethod.Invoke(this, null);

            //}).Start();
            PrefetchData();
            button2_Click(sender, e);
            //prefetchMethod = myType.GetMethod("FixInformation");
            //prefetchMethod.Invoke(this, null);
            FixInformation();

        }

        private void FixInformation()
        {
            ColumnHeader resolution1Col = new ColumnHeader();
            resolution1Col.Text = "Resolution";
            resolution1Col.Width = masLVI.ClientSize.Width / 2;
            masLVI.Columns.Add(resolution1Col);
            ColumnHeader resolution2Col = new ColumnHeader();
            resolution2Col.Text = "Resolution";
            resolution2Col.Width = overdistLVI.ClientSize.Width / 2;     
            overdistLVI.Columns.Add(resolution2Col);
            ColumnHeader resolution3Col = new ColumnHeader();
            resolution3Col.Text = "Resolution";
            resolution3Col.Width = masLVI.ClientSize.Width / 2;         
            whseLVI.Columns.Add(resolution3Col);


            foreach (ListViewItem lvi in masLVI.Items)
            {
                FixItemWarehouseCheck(lvi);
            }
            foreach (ListViewItem lvi in overdistLVI.Items)
            {
                FixOverdist(lvi);
            }
            foreach (ListViewItem lvi in whseLVI.Items)
            {
                FixWarehouseMASVariance(lvi);
            }
        }

        public void PrefetchData()
        {
            SetupLVI();

            if (this.InvokeRequired == true)
            {
                statusLbl.Invoke((MethodInvoker)(() =>
                    {
                        ////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of MAS Kits...";
                        statusLbl.Refresh();
                        MasKits = GetMASKits();
                        ////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of IM_ItemWarehouse Item Details...";
                        statusLbl.Refresh();
                        IM_ItemWarehouse_Items = GetItemWarehouseItems();
                        ////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of SO_InvoiceHeader Info...";
                        statusLbl.Refresh();
                        SO_NonXDInvoices = GetNonXDInvoices();
                        ////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of SO_InvoiceDetails Items...";
                        statusLbl.Refresh();
                        SO_InvoiceDetails = GetInvoicedItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of P2_TransactionHeader Info...";
                        statusLbl.Refresh();
                        P2_NonVTransactions = GetNonVTransactions();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of P2_TransactionDetails Items...";
                        statusLbl.Refresh();
                        P2_TransactionDetails = GetP2TransactionItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of SO_SalesOrderHeader Info...";
                        statusLbl.Refresh();
                        SO_SalesOrders = GetSalesOrders();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of SO_SalesOrderDetails Items...";
                        statusLbl.Refresh();
                        SO_SalesDetails = GetSalesOrderItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of IM_ItemCost Items...";
                        statusLbl.Refresh();
                        IM_ItemCost_Items = GetItemCostItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of Item Overdists...";
                        statusLbl.Refresh();
                        IM_OverDist_Items = GetItemCostOverDistItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Caching List Of Whse/MAS Variances...";
                        statusLbl.Refresh();
                        WhseMasVariance_Items = GetWhseMasVarianceItems();
                        /////////////////////////////////////////////////
                        statusLbl.Text = "Status: Ready to Rock And Roll!";
                        statusLbl.Refresh();

                    }));

            }
            else
            {
                ////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of MAS Kits...";
                statusLbl.Refresh();
                MasKits = GetMASKits();
                ////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of IM_ItemWarehouse Item Details...";
                statusLbl.Refresh();
                IM_ItemWarehouse_Items = GetItemWarehouseItems();
                ////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of SO_InvoiceHeader Info...";
                statusLbl.Refresh();
                SO_NonXDInvoices = GetNonXDInvoices();
                ////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of SO_InvoiceDetails Items...";
                statusLbl.Refresh();
                SO_InvoiceDetails = GetInvoicedItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of P2_TransactionHeader Info...";
                statusLbl.Refresh();
                P2_NonVTransactions = GetNonVTransactions();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of P2_TransactionDetails Items...";
                statusLbl.Refresh();
                P2_TransactionDetails = GetP2TransactionItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of SO_SalesOrderHeader Info...";
                statusLbl.Refresh();
                SO_SalesOrders = GetSalesOrders();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of SO_SalesOrderDetails Items...";
                statusLbl.Refresh();
                SO_SalesDetails = GetSalesOrderItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of IM_ItemCost Items...";
                statusLbl.Refresh();
                IM_ItemCost_Items = GetItemCostItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of Item Overdists...";
                statusLbl.Refresh();
                IM_OverDist_Items = GetItemCostOverDistItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Caching List Of Whse/MAS Variances...";
                statusLbl.Refresh();
                WhseMasVariance_Items = GetWhseMasVarianceItems();
                /////////////////////////////////////////////////
                statusLbl.Text = "Status: Ready to Rock And Roll!";
                statusLbl.Refresh();
            }
        }
        private void QuantityOnSalesOrder(string ItemNumber)
        {
            if (MasKits.ContainsKey(ItemNumber)==true)
            {
                return;
            }
            if (ItemNumber=="BCHADS181-102")
            {

            }
            int Ordered = 0;
            int Backordered = 0;
            int Shipped = 0;

            int IM_Ordered = 0;
            int IM_Backordered = 0;
            int IM_Shipped = 0;


            if (SO_InvoiceDetails.ContainsKey(ItemNumber) == true)
            {
                Backordered += SO_InvoiceDetails[ItemNumber].QuantityBackordered;
                Shipped += SO_InvoiceDetails[ItemNumber].QuantityShipped;
                Ordered += SO_InvoiceDetails[ItemNumber].QuantityOrdered;
            }

            if (P2_TransactionDetails.ContainsKey(ItemNumber) == true)
            {
                Backordered += P2_TransactionDetails[ItemNumber].QuantityBackordered;
                Shipped += P2_TransactionDetails[ItemNumber].QuantityShipped;
                Ordered += P2_TransactionDetails[ItemNumber].QuantityOrdered;

            }
            if (SO_SalesDetails.ContainsKey(ItemNumber) == true)
            {
                Backordered += SO_SalesDetails[ItemNumber].QuantityBackordered;
                Shipped += SO_SalesDetails[ItemNumber].QuantityShipped;
                Ordered += SO_SalesDetails[ItemNumber].QuantityOrdered;
            }
            if (IM_ItemWarehouse_Items.ContainsKey(ItemNumber) == true)
            {
                IM_Ordered = IM_ItemWarehouse_Items[ItemNumber].QuantityOnSalesOrder;
                IM_Backordered = IM_ItemWarehouse_Items[ItemNumber].QuantityOnBackOrder;
                IM_Shipped = IM_ItemWarehouse_Items[ItemNumber].QuantityInShipping;
            }  

            if (IM_Ordered != Ordered)
            { 
                if (ItemNumber == "D-HDWHT51411")
                {

                }
                ListViewItem lvi = new ListViewItem();
                lvi.Text = IM_ItemWarehouse_Items[ItemNumber].WarehouseCode.ToString();
                lvi.SubItems.Add("Qty On Sales Order");
                lvi.SubItems.Add(ItemNumber);
                lvi.SubItems.Add(IM_Ordered.ToString());
                lvi.SubItems.Add(Ordered.ToString());

                masLVI.Items.Add(lvi);
                masLVI.Refresh();
            }

            //MessageBox.Show("Ordered: " + Ordered.ToString() + " / Backordered: " + Backordered.ToString() + " / Shipped: " + Shipped.ToString() + "\r\nIM_ItemWarehouse shows....Ordered: " + IM_Ordered.ToString() + " / Backordred: " + IM_Backordered.ToString() + " / Shipped: " + IM_Shipped.ToString());
           // MessageBox.Show("Inventory Qty (IM_ItemWarehouse): " + IM_Ordered.ToString() + "\r\n" + "Module Qty: " + Ordered.ToString());

        }
        private void QuantityInShipping(string ItemNumber)
        {
            if (MasKits.ContainsKey(ItemNumber) == true)
            {
                return;
            }
            if (ItemNumber == "FRE91-102" | ItemNumber == "BOSMFN-201")
            {
            }
            int Ordered = 0;
            int Backordered = 0;
            int Shipped = 0;

            int IM_Ordered = 0;
            int IM_Backordered = 0;
            int IM_Shipped = 0;


            if (SO_InvoiceDetails.ContainsKey(ItemNumber)==true)
            {
                Backordered += SO_InvoiceDetails[ItemNumber].QuantityBackordered;
                Shipped += SO_InvoiceDetails[ItemNumber].QuantityShipped;
                Ordered += SO_InvoiceDetails[ItemNumber].QuantityOrdered;
            }
            if (P2_TransactionDetails.ContainsKey(ItemNumber)==true)
            {
                Backordered += P2_TransactionDetails[ItemNumber].QuantityBackordered;
                Shipped += P2_TransactionDetails[ItemNumber].QuantityShipped;
                Ordered += P2_TransactionDetails[ItemNumber].QuantityOrdered;
            }
            if (SO_SalesDetails.ContainsKey(ItemNumber)==true)
            {
                Backordered += SO_SalesDetails[ItemNumber].QuantityBackordered;
                Shipped += SO_SalesDetails[ItemNumber].QuantityShipped;
                Ordered += SO_SalesDetails[ItemNumber].QuantityOrdered;
            }
            if (IM_ItemWarehouse_Items.ContainsKey(ItemNumber)==true)
            {
                IM_Ordered = IM_ItemWarehouse_Items[ItemNumber].QuantityOnSalesOrder;
                IM_Backordered = IM_ItemWarehouse_Items[ItemNumber].QuantityOnBackOrder;
                IM_Shipped = IM_ItemWarehouse_Items[ItemNumber].QuantityInShipping;
            }


            if (Shipped != IM_Shipped)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = IM_ItemWarehouse_Items[ItemNumber].WarehouseCode.ToString();
                lvi.SubItems.Add("Qty In Shipping");
                lvi.SubItems.Add(ItemNumber);
                lvi.SubItems.Add(IM_Shipped.ToString());
                lvi.SubItems.Add(Shipped.ToString());                
                masLVI.Items.Add(lvi);
                masLVI.Refresh();
            }

            //MessageBox.Show("Ordered: " + Ordered.ToString() + " / Backordered: " + Backordered.ToString() + " / Shipped: " + Shipped.ToString() + "\r\nIM_ItemWarehouse shows....Ordered: " + IM_Ordered.ToString() + " / Backordred: " + IM_Backordered.ToString() + " / Shipped: " + IM_Shipped.ToString());
            //MessageBox.Show("Qty In Shipping (IM_ItemWarehouse): " + Shipped.ToString() + "\r\n" + "Module Shipped Qty: " + IM_Shipped.ToString());

        }
        private void InventoryCosting(string ItemNumber)
        {
            if (MasKits.ContainsKey(ItemNumber) == true)
            {
                return;
            }
            if (ItemNumber == "BCHADS181-102")
            {
            }
            if (IM_ItemWarehouse_Items.ContainsKey(ItemNumber) == true & IM_ItemCost_Items.ContainsKey(ItemNumber)==true)
            {
                int IM_Cost = IM_ItemCost_Items[ItemNumber].QuantityOnHand;
                int IM_Warehouse = IM_ItemWarehouse_Items[ItemNumber].QuantityOnHand;

                if (IM_Cost != IM_Warehouse)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = IM_ItemWarehouse_Items[ItemNumber].WarehouseCode.ToString();
                    lvi.SubItems.Add("Inventory Costing");
                    lvi.SubItems.Add(ItemNumber);
                    lvi.SubItems.Add(IM_Warehouse.ToString());
                    lvi.SubItems.Add(IM_Cost.ToString());
                    masLVI.Items.Add(lvi);
                    masLVI.Refresh();
                }
            }
            //if (IM_ItemWarehouse_Items.ContainsKey(ItemNumber) == true)
            //{
            //    MessageBox.Show("ItemWarehouse has the item");
            //}
            //else
            //{
            //   MessageBox.Show("ItemWarehouse does not have item.");
            //}
            //if (IM_ItemCost_Items.ContainsKey(ItemNumber)==true)
            //{
            //    MessageBox.Show("Item Cost has the item.");
            //}
            //else
            //{
            //    MessageBox.Show("Item Cost does not have the item.");
            //}
           
        }
        private void ItemOverDist(string ItemNumber)
        {

            if (IM_OverDist_Items.ContainsKey(ItemNumber) == true)
            {
                string item = IM_OverDist_Items[ItemNumber].ItemCode;
                string warehouse = IM_OverDist_Items[ItemNumber].WarehouseCode;
                string receipt = IM_OverDist_Items[ItemNumber].ReceiptNo;
                string unitcost = IM_OverDist_Items[ItemNumber].UnitCost.ToString();
                string qoh = IM_OverDist_Items[ItemNumber].QuantityOnHand.ToString();
                string ext = IM_OverDist_Items[ItemNumber].ExtendedCost.ToString();

                if (decimal.Parse(ext) != 0)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = ItemNumber;
                    lvi.SubItems.Add(IM_ItemCost_Items[ItemNumber].WarehouseCode.ToString());
                    lvi.SubItems.Add(receipt.ToUpper());
                    lvi.SubItems.Add(qoh);
                    lvi.SubItems.Add(unitcost);
                    lvi.SubItems.Add(ext);
                    lvi.SubItems.Add("0"); //???
                    overdistLVI.Items.Add(lvi);
                    overdistLVI.Refresh();
                }

              

                //MessageBox.Show("Item " + item + " from warehouse " + warehouse + " has an " + receipt + " receipt. Unit Cost is " + unitcost + ", Quantity On Hand is " + qoh + " and the Extended Cost (Which should be 0) is: " + ext);

            }
            else
            {
               // MessageBox.Show("The item " + ItemNumber + " actually looks to be fine.");
            }

        }
        private void WhseMasVariances(string ItemNumber)
        {
            //MessageBox.Show("Debuggin... " + ItemNumber + " in MAS has qty: " + WhseMasVariance_Items[ItemNumber].MASQty.ToString());
            
           // foreach (KeyValuePair<int,string> kvp in ComponentName)
           // {
          
             int CID = ComponentName.FirstOrDefault(x=>x.Value == ItemNumber).Key;
                //if (kvp.Value.ToUpper() == ItemNumber.ToUpper())
               // {
                    //int CID = kvp.Key;
               
                        if (ComponentExpected.ContainsKey(CID) == true & WarehouseQuantity.ContainsKey(CID) == true)
                        {
                            if (ComponentExpected[CID] == WarehouseQuantity[CID])
                            {
                                //ok
                                //MessageBox.Show("This item has been determined OK.");
                            }
                            else
                            {
                                string PartName = "";
                                int Expected = 0; //mas qty
                                int WhseQty = 0;
                                int Willcalls = 0;
                                PartName = ComponentName[CID];
                                Expected = ComponentExpected[CID];
                                WhseQty = WarehouseQuantity[CID];
                                try
                                {
                                    Willcalls = WillCallComponents[CID];
                                }
                                catch { }


                                if (Expected != WhseQty)
                                {
                                    ListViewItem lvi = new ListViewItem();
                                    lvi.Text = PartName;
                                    lvi.SubItems.Add(Expected.ToString());
                                    lvi.SubItems.Add(WhseQty.ToString());
                                    lvi.SubItems.Add((WhseQty - Expected).ToString());
                                    lvi.SubItems.Add(Willcalls.ToString());
                                    whseLVI.Items.Add(lvi);
                                    whseLVI.Refresh();
                                }


                                //MessageBox.Show(PartName + " is expected to have " + Expected.ToString() + " qty (from MAS), but the Warehouse shows " + WhseQty.ToString() + ". There is " + Willcalls.ToString() + " on Willcall.");
                                //break;

                            }
                        }

                        
 
               // }
               // else
              //  {
                    //couldn't find value
              //  }
            //}
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton6.Checked == true)
            {
                WhseMasVariances(textBox1.Text);
            }
            else if (radioButton5.Checked == true)
            {
                ItemOverDist(textBox1.Text);
            }
            else if (radioButton3.Checked == true)
            {
                QuantityOnSalesOrder(textBox1.Text);
            }
            else if (radioButton2.Checked == true)
            {
                QuantityInShipping(textBox1.Text);
            }
            else if (radioButton1.Checked == true)
            {
                InventoryCosting(textBox1.Text);
            }
            else
            {
                //MessageBox.Show("Only Variances On 'Sales Orders' and 'In Shipping' Can Be Fixed At The Moment!");
                return;
            }
            MessageBox.Show("Command Complete");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            statusLbl.Text = "Status: Producing Report...";
            statusLbl.Refresh();
            SetupLVI();
            long counter = 0;

            List<string> items = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM InventoryMaster WHERE Hide=0");


            //foreach (KeyValuePair<string,IM_ItemWarehouse> kvp in IM_ItemWarehouse_Items)\
            foreach (string itm in items)
            {
                string itemNumber = itm.Split('|')[0];
                QuantityOnSalesOrder(itemNumber);
                QuantityInShipping(itemNumber);
                InventoryCosting(itemNumber);
                ItemOverDist(itemNumber);
                WhseMasVariances(itemNumber);
                counter++;
                statusLbl.Text = "Status: Processing " + counter.ToString() + " of " + (items.Count +NonZeroWarehouseItems.Count).ToString();
                statusLbl.Refresh();
            }
            foreach (string itm in NonZeroWarehouseItems)
            {
                //string itemNumber = "(900)" + itm.Split('|')[0];
                QuantityOnSalesOrder(itm);
                QuantityInShipping(itm);
                InventoryCosting(itm);
                ItemOverDist(itm);
                WhseMasVariances(itm);
                counter++;
                statusLbl.Text = "Status: Processing 900s..." + counter.ToString() + " of " + (items.Count + NonZeroWarehouseItems.Count).ToString();
                statusLbl.Refresh();
            }




            label3.Text = masLVI.Items.Count.ToString();
            label4.Text = overdistLVI.Items.Count.ToString();
            label5.Text = whseLVI.Items.Count.ToString();
        }

        private void masLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (int index in overdistLVI.SelectedIndices)
            {
                overdistLVI.Items[index].Selected = false;
            }
            foreach (int index in whseLVI.SelectedIndices)
            {
                whseLVI.Items[index].Selected = false;
            }
        }

        private void overdistLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (int index in masLVI.SelectedIndices)
            {
                masLVI.Items[index].Selected = false;
            }
            foreach (int index in whseLVI.SelectedIndices)
            {
                whseLVI.Items[index].Selected = false;
            }

        }

        private void whseLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (int index in masLVI.SelectedIndices)
            {
                masLVI.Items[index].Selected = false;
            }
            foreach (int index in overdistLVI.SelectedIndices)
            {
                overdistLVI.Items[index].Selected = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //find out which lv's item got selected...
            int line = -1;
            if (masLVI.SelectedIndices.Count>0)
            {
                line = masLVI.SelectedIndices[0];
                FixItemWarehouseCheck(masLVI.Items[line]);      
                
            }
            if (overdistLVI.SelectedIndices.Count > 0)
            {
                line = overdistLVI.SelectedIndices[0];
                FixOverdist(overdistLVI.Items[line]);               
            }
            if (whseLVI.SelectedIndices.Count>0)
            {
                line = whseLVI.SelectedIndices[0];
                //FixWarehouseMASVariance(ref whseLVI.Items[line]);          
            }


        }

        private void FixWarehouseMASVariance(ListViewItem variance)
        {
            bool fixfound = false;
            //MessageBox.Show("Let's take a peek at our Warehouse/MAS variance...");
            string ItemNumber = variance.Text;
            int MASQty = int.Parse(variance.SubItems[1].Text);
            int WHSEQty = int.Parse(variance.SubItems[2].Text);
            int ReportedVariance = int.Parse(variance.SubItems[3].Text);
            int WillCalls = int.Parse(variance.SubItems[4].Text);

            if (ReportedVariance == WillCalls)
            {
                foreach (ListViewItem lvi in whseLVI.Items)
                {
                    if (lvi == variance)
                    {
                        lvi.BackColor = Color.Green;
                        lvi.SubItems.Add("Whoa, so the variance is " + ReportedVariance.ToString() + " and we also have " + WillCalls.ToString() + " for this item. I am willing to bet this is OK and will dissapear as soon as the willcall is completed.");
                        fixfound = true;
                    }
                }
                //MessageBox.Show("Whoa, so the variance is " + ReportedVariance.ToString() + " and we also have " + WillCalls.ToString() + " for this item. I am willing to bet this is OK and will dissapear as soon as the willcall is completed.");
                return;
            }
            
            //now lets try to test item stuck in tote...
            Dictionary<int, int> TotesWithItem = ItemCountsInTotes(ItemNumber);
            int totQty = 0;
            foreach (KeyValuePair<int,int> tData in TotesWithItem)
            {
                totQty += tData.Value;
            }
            if (totQty == ReportedVariance)
            {
                //looks to be a stuck item in a tote...
                string response = ItemNumber + " happens to be stuck in tote(s) ";
                foreach (KeyValuePair<int,int> tData in TotesWithItem)
                {
                    if (tData.Value > 0)
                    {
                        response += "#" + tData.Key.ToString() + " (with qty of " + tData.Value.ToString() + "), ";
                    }
                }
                response = response.Trim(' ').Trim(',');
                response += ". The variance of this item was " + ReportedVariance.ToString() + " and the total qty in the tote(s) equals " + totQty.ToString() + ".";
                response += " Make sure item has shipped and clear tote.";
                foreach (ListViewItem lvi in whseLVI.Items)
                {
                    if (lvi == variance)
                    {
                        lvi.BackColor = Color.Yellow;
                        lvi.ForeColor = Color.Black;
                        lvi.SubItems.Add(response);
                        fixfound = true;
                    }
                }
                

                //MessageBox.Show(response);
                return;
            }

            //now lets see if the item was on a will call that has already been cleared (today)...

            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID,QuantityFilled FROM WhseTaskWillCallLines INNER JOIN WhseTaskWillCallHeader ON HeaderID = WhseTaskWillCallHeader.ID WHERE ComponentID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + ItemNumber + "') AND TimeInserted>DATEADD(dd,-1,GETDATE())");
            Dictionary<int,int> Quantities = new Dictionary<int,int>();
            foreach (string res in results)
            {
                Quantities.Add(int.Parse(res.Split('|')[0]), int.Parse(res.Split('|')[1]));
            }

            foreach (KeyValuePair<int,int> kvp in Quantities)
            {
                if (kvp.Value == -1 * ReportedVariance)
                {
                    foreach (ListViewItem itm in whseLVI.Items)
                    {
                        if (itm == variance)
                        {
                            itm.BackColor = Color.Green;
                            itm.SubItems.Add("Item on filled WillCall w/ " + kvp.Value.ToString() + " qty, variance is " + ReportedVariance.ToString() + ". This variance should dissapear at closing today.");
                            fixfound = true;
                        }
                    }
                    //MessageBox.Show("After some digging, I found that this item was on a willcall that filled " + kvp.Value.ToString() + " quantity, which our variance happens to be " + ReportedVariance.ToString() + ". This variance should dissapear as it will be a variance until the register is closed.");
                    return;
                }

            }


            //run check to see if qty in shipping staging is > than qty in shipping
            int MASQtyInShipping = IM_ItemWarehouse_Items[ItemNumber].QuantityInShipping;
            
            List<string> cID = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + ItemNumber + "'");
            if (cID.Count > 0)
            {
                string ID = cID[0].Split('|')[0];
                List<string> qtyInShipping = Connectivity.SQLCalls.sqlQuery("SELECT Quantity FROM LocationContents WHERE LocationID=10488");
                if (qtyInShipping.Count>0)
                {
                    int WhseShippingQty = int.Parse(qtyInShipping[0].Split('|')[0]);
                    if (WhseShippingQty > MASQtyInShipping)
                    {
                        foreach (ListViewItem itm in whseLVI.Items)
                        {
                            if (itm == variance)
                            {
                                itm.BackColor = Color.Red;
                                itm.ForeColor = Color.White;
                                string line = "Whse has " + WhseShippingQty.ToString() + " in shipping staging, but MAS says there is only " + MASQtyInShipping.ToString();
                                if (WhseShippingQty < System.Math.Abs(ReportedVariance))
                                {
                                    line += " Also, there must be another issue.";
                                }

                                itm.SubItems.Add(line);
                                fixfound = true;
                            }
                        }
                        return;
                    }
                }
            }
         

            if (fixfound == false)
            {
                foreach (ListViewItem itm in whseLVI.Items)
                {
                    if (itm == variance)
                    {
                        itm.BackColor = Color.White;
                        itm.SubItems.Add("Tell Tom we have a new kind of variance.");                     
                    }
                }
            }



        }
        private void FixItemWarehouseCheck(ListViewItem variance)
        {

            string salesOrderNo = "N/A";
            if (SO_SalesDetails.ContainsKey(variance.SubItems[2].Text) == true)
            {
                salesOrderNo = SO_SalesDetails[variance.SubItems[2].Text].SalesOrderNo;
            }
            foreach (ListViewItem lvi in masLVI.Items)
            {
                bool foundissue = false;
                if (variance == lvi)
                {
                    string variancedItem = variance.SubItems[2].Text;
                    string varianceType = variance.SubItems[1].Text;
                    int varianceAmount = int.Parse(variance.SubItems[4].Text);

                    if (varianceType == "Qty On Sales Order")
                    {
                        foreach (ListViewItem prob in masLVI.Items)
                        {
                            if (prob.SubItems[2].Text == variancedItem)
                            {
                                if (prob.SubItems[4].Text == varianceAmount.ToString())
                                {
                                    if (prob.SubItems[1].Text == "Qty In Shipping")
                                    {
                                        lvi.BackColor = Color.Green;
                                        lvi.SubItems.Add("Has matching In Shipping variance... Will dissapear in a good way.");
                                        foundissue = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (varianceType == "Qty In Shipping")
                    {
                        foreach (ListViewItem prob in masLVI.Items)
                        {
                            if (prob.SubItems[2].Text == variancedItem)
                            {
                                if (prob.SubItems[4].Text == varianceAmount.ToString())
                                {
                                    if (prob.SubItems[1].Text == "Qty On Sales Order")
                                    {
                                        lvi.BackColor = Color.Green;
                                        lvi.SubItems.Add("Has matching On Sales Order variance... Will dissapear in a good way.");
                                        foundissue = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }


                    //if variance is found in both in shipping and sales order and qtys are equal, it should dissapear
                    if (foundissue==false)
                    {
                        lvi.SubItems.Add("Tell Tom we have a new kind of variance. Order #" + salesOrderNo);
                    }
                }
            }
            
            //MessageBox.Show("Cannot fix IM_ItemWarehouse checks yet. So use your brain!");
        }
        private void FixOverdist(ListViewItem variance)
        {
            foreach (ListViewItem lvi in overdistLVI.Items)
            {
                if (variance == lvi)
                {
                    string ReceiptNo = lvi.SubItems[2].Text;
                    int Quantity = int.Parse(lvi.SubItems[3].Text);
                    decimal Extension = decimal.Parse(lvi.SubItems[5].Text);
                    decimal Expected = decimal.Parse(lvi.SubItems[6].Text);
                    if (Quantity == 0 & Expected == 0 & Extension != 0)
                    {
                        lvi.BackColor = Color.Red;
                        lvi.ForeColor = Color.White;
                        lvi.SubItems.Add("Need to go into IM_ItemCost and set the extended amount to 0 on receipt #" + ReceiptNo + ".");
                    }
                    else
                    {

                        lvi.SubItems.Add("Tell Tom we have a new kind of variance.");
                    }
                }
            }
           // MessageBox.Show("Cannot fix overdists yet. Use your brain!");
        }








       
        public struct ToteData
        {
            public int ToteNumber;
            public int ComponentID;
            public int Quantity;
            public DateTime LastFilled;
        }

        private Dictionary<int,int> ItemCountsInTotes(string ItemNumber)
        {
            Dictionary<int,int> ReturnData = new Dictionary<int,int>();
            List<ToteData> ToteData = new List<Form1.ToteData>();
            ToteData tmpData = new ToteData();


            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT Division1,LocationContents.ComponentID,LocationContents.Quantity,LocationContents.LastInventoriedDate FROM LocationContents INNER JOIN LocationMaster ON LocationMaster.ID=LocationContents.LocationID INNER JOIN InventoryComponentMap ON InventoryComponentMap.ComponentID=LocationContents.ComponentID WHERE LocationTypeID=11 AND InventoryComponentMap.ItemNumber='" + ItemNumber + "'");
            foreach (string res in results)
            {
                int ToteNumber=int.Parse(res.Split('|')[0]);
                int ComponentID= int.Parse(res.Split('|')[1]);
                int Quantity = int.Parse(res.Split('|')[2]);
                DateTime ToteWasFilled = DateTime.Parse(res.Split('|')[3]);
                tmpData.ToteNumber = ToteNumber;
                tmpData.ComponentID = ComponentID;
                tmpData.Quantity = Quantity;
                tmpData.LastFilled = ToteWasFilled;
                ToteData.Add(tmpData);
            }
            foreach (ToteData td in ToteData)
            {
                //if (td.LastFilled.Date < DateTime.Now.Date.AddDays(-1))
                 if (td.LastFilled.Date < DateTime.Now.Date.AddDays(0))
                 {
                     ReturnData.Add(td.ToteNumber,td.Quantity);
                 }
            }
            return ReturnData;

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
