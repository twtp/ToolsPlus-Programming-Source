using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.FrontEnd
{
  public static class Extensions
  {
    private const string SERVER_PROTOCOL = "http";
    private const string SERVER_NAME = "dev-03";
    private const string SERVER_PORT = "1234";


    public static string assemblePreloads(params string[] preloads)
    {
      if (preloads == null || preloads.Length == 0) { return string.Empty; }
      return "?preload=" + string.Join(",", preloads);
    }

    public static T Get<T>(params object[] urlParts)
    {
      return Get<T>("/" + string.Join("/", urlParts));
    }

    public static T Get<T>(string[] preloads, params object[] urlParts)
    {
      return Get<T>("/" + string.Join("/", urlParts) + assemblePreloads(preloads));
    }

    public static T Get<T>(string url)
    {
      System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(SERVER_PROTOCOL + "://" + SERVER_NAME + ":" + SERVER_PORT + url);
      req.AutomaticDecompression = System.Net.DecompressionMethods.Deflate | System.Net.DecompressionMethods.GZip;
      System.Net.WebResponse resp = req.GetResponse();

      string xml = null;
      using (System.IO.StreamReader r = new System.IO.StreamReader(resp.GetResponseStream(), Encoding.UTF8))
      {
        xml = r.ReadToEnd();
      }

      return doDeserialize<T>(xml);
    }




    public static TReturn Post<T, TReturn>(string serializableName, T serializableParam, params object[] urlParts) where T : TP.Object.BaseBusinessObject
    {
      var foo = new Dictionary<string, string>();
      return Post<T, TReturn>(serializableName, serializableParam, foo, string.Join("/", urlParts));
    }

    public static TReturn Post<T, TReturn>(string serializableName, T serializableParam, Dictionary<string, string> stringParams, params object[] urlParts) where T : TP.Object.BaseBusinessObject
    {
      return Post<T, TReturn>(serializableName, serializableParam, stringParams, string.Join("/", urlParts));
    }

    public static TReturn Post<T, TReturn>(string serializableName, T serializableParam, Dictionary<string, string> stringParams, string url) where T : TP.Object.BaseBusinessObject
    {
      stringParams.Add(serializableName, Serializer.ByteArrayToString(Serializer.SerializeObjectToDataContract(serializableParam)));
      return Post<TReturn>(stringParams, url);
    }

    public static TReturn Post<TReturn>(Dictionary<string, string> stringParams, params object[] urlParts)
    {
      return Post<TReturn>(stringParams, string.Join("/", urlParts));
    }

    public static TReturn Post<TReturn>(Dictionary<string, string> stringParams, string url)
    {
      var payload = string.Join("&", from k in stringParams.Keys select string.Format("{0}={1}", Uri.EscapeDataString(k).Replace("%20", "+"), Uri.EscapeDataString(stringParams[k]).Replace("%20", "+")));
      var bytes = Encoding.ASCII.GetBytes(payload);

      if (url.StartsWith("/") == false) { url = "/" + url; }

      var req = System.Net.WebRequest.Create(SERVER_PROTOCOL + "://" + SERVER_NAME + ":" + SERVER_PORT + url);
      req.Method = "POST";
      req.ContentType = "application/x-www-form-urlencoded";
      req.ContentLength = bytes.Length;

      using (var s = req.GetRequestStream())
      {
        s.Write(bytes, 0, bytes.Length);
      }

      var resp = req.GetResponse();
      string xml = null;
      using (System.IO.StreamReader r = new System.IO.StreamReader(resp.GetResponseStream(), Encoding.UTF8))
      {
        xml = r.ReadToEnd();
      }

      return doDeserialize<TReturn>(xml);
    }

    public static bool PostIsSuccessful<T>(T serializableParam, Dictionary<string, string> stringParams, params object[] urlParts) where T : TP.Object.BaseBusinessObject
    {
      TP.Base.SuccessResponse r = Post<T, TP.Base.SuccessResponse>("TP.Object.SuccessResponse", serializableParam, stringParams, urlParts);
      return true;
    }



    private static T doDeserialize<T>(string xml)
    {
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
            TP.Base.ErrorResponse err = TP.Object.Serializer.DeserializeDataContractToObject<TP.Base.ErrorResponse>(xml, Encoding.UTF8);
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
    }

  }
}
