using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Runtime.Serialization;

using TP.Object.FrontEnd;

namespace PoInvClient
{
  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {
      try
      {
        /*var item = Get<TP.Object.Item>("/item/123?preload=components,barcodes");
        var comp = FeComponent.FromBase(item.ItemComponents.First().Component);
        MessageBox.Show(comp.Barcodes.First().Code);*/

        /*var item = Get<TP.Object.Item>("/item/123?preload=components");
        var comp = FeComponent.FromBase(item.ItemComponents.First().Component);
        comp.FillBarcodeList();
        MessageBox.Show(comp.Barcodes.First().Code);*/
        
        /*var l1 = FeInventoryLocation.GetByLocationID(504, preloadContents: true);
        var l2 = FeInventoryLocation.GetByLocationID(506);
        var l3 = FeInventoryLocation.GetByLocationID(507);
        var l4 = FeInventoryLocation.GetByLocationID(508, preloadContents: true);
        MessageBox.Show(l1.CanSpanTo(l2).ToString());
        MessageBox.Show(l1.CanSpanTo(l3).ToString());*/
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    }

    /*private T Get<T>(string url)
    {

      WebRequest req = WebRequest.Create("http://dev-03:1234" + url);
      WebResponse resp = req.GetResponse();

      using (System.IO.StreamReader r = new System.IO.StreamReader(resp.GetResponseStream(), Encoding.UTF8))
      {
        string xml = r.ReadToEnd();

        try
        {
          T item = TP.Object.Serializer.DeserializeDataContractToObject<T>(xml, Encoding.UTF8);
          return item;
        }
        catch (System.Runtime.Serialization.SerializationException ex)
        {
          if (ex.Message.Contains("Encountered 'Element'  with name 'ErrorResponse'"))
          {
            try
            {
              TP.Object.ErrorResponse err = TP.Object.Serializer.DeserializeDataContractToObject<TP.Object.ErrorResponse>(xml, Encoding.UTF8);
              throw new Exception(err.Message);
            }
            catch
            {
              throw;
            }
          }
          else
          {
            throw;
          }
        }
        finally
        {
          resp.Close();
        }
      }
    }*/
  }
}
