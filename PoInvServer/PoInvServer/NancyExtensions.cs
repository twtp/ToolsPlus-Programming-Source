using Nancy;
using System.Collections.Generic;

namespace PoInvServer
{
  public static class NancyExtensions
  {
    public static Response Respond<T>(T serializable) where T : class
    {
      if (serializable == null)
      {
        return new Response {StatusCode = HttpStatusCode.NoContent};
      }

      var bytes = TP.Object.Serializer.SerializeObjectToDataContract(serializable);
      if (bytes == null)
      {
        return new Response {StatusCode = HttpStatusCode.InternalServerError};
      }

      return new Response
      {        
        StatusCode = HttpStatusCode.OK,
        ContentType = "application/xml",
        Contents = s => s.Write(bytes, 0, bytes.Length)
      };
    }

    [System.Diagnostics.DebuggerStepThrough()]
    public static bool HasParameter(DynamicDictionary reqForm, string paramName)
    {
      return reqForm.ContainsKey(paramName);
    }

    [System.Diagnostics.DebuggerStepThrough()]
    public static T GetDbObjectParameter<T>(DynamicDictionary reqForm, string paramName) where T : TP.Object.BaseBusinessObject
    {
      if (reqForm.ContainsKey(paramName) == false)
      {
        throw new System.ArgumentException("expected a '" + paramName + "' datacontract serialization");
      }

      return TP.Object.Serializer.DeserializeDataContractToObject<T>(reqForm[paramName].Value);
    }

    [System.Diagnostics.DebuggerStepThrough()]
    public static T GetBasicParameter<T>(DynamicDictionary reqForm, string paramName, T defaultTo = default(T)) where T : System.IConvertible
    {
      if (reqForm.ContainsKey(paramName) == false || (string)reqForm[paramName].Value == string.Empty)
      {
        return defaultTo;
      }

      try
      {
        return System.Convert.ChangeType(reqForm[paramName].Value, typeof(T));
      }
      catch (System.Exception)
      {
        return defaultTo;
      }
    }

    [System.Diagnostics.DebuggerStepThrough()]
    public static string[] GetPreloads(DynamicDictionary reqQuery)
    {
      return reqQuery.ContainsKey("preload") ? ((string)reqQuery["preload"]).Split(',') : new string[] {};
    }

    [System.Diagnostics.DebuggerStepThrough()]
    public static List<System.Tuple<T1, T2>> GetPairedValueList<T1, T2>(DynamicDictionary reqQuery, string field1Name, string field2Name, T1 field1Default = default(T1), T2 field2Default = default(T2)) where T1 : System.IConvertible where T2 : System.IConvertible
    {
      var retval = new List<System.Tuple<T1, T2>>();
      bool paramsFinished = false;
      int dx = 1;
      while (paramsFinished == false)
      {
        T1 k = NancyExtensions.GetBasicParameter<T1>(reqQuery, field1Name + dx.ToString(), field1Default);
        T2 v = NancyExtensions.GetBasicParameter<T2>(reqQuery, field2Name + dx.ToString(), field2Default);
        if (object.Equals(k, field1Default) && object.Equals(v, field2Default))
        {
          paramsFinished = true;
        }
        else if (object.Equals(k, field1Default) || object.Equals(v, field2Default))
        {
          throw new System.ArgumentException("invalid arguments, must have both " + field1Name + " and " + field2Name + " for each specified level");
        }
        else
        {
          retval.Add(new System.Tuple<T1, T2>(k, v));
          dx++;
        }
      }
      return retval;
    }
  }


  public class CustomBootstrapper : DefaultNancyBootstrapper
  {
    protected override void ApplicationStartup(Nancy.TinyIoc.TinyIoCContainer container, Nancy.Bootstrapper.IPipelines pipelines)
    {
      base.ApplicationStartup(container, pipelines);
      
      pipelines.BeforeRequest += (ctx) =>
      {
        System.Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss.fff}   {1,13}   {2}", System.DateTime.Now, ctx.Request.UserHostAddress, ctx.Request.Url));
        foreach (var key in ctx.Request.Form) {
          System.Console.WriteLine(string.Format("  POST: {0} => {1}", key, ctx.Request.Form[key]));
        }
        return null;
      };

      pipelines.AfterRequest += (ctx) =>
      {

        var ms = new System.IO.MemoryStream();
        ctx.Response.Contents.Invoke(ms);
        ms.Position = 0;
        string respContent = System.Text.Encoding.ASCII.GetString(ms.ToArray());
        respContent = respContent.Replace(" xmlns:i=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"TP.Object\"", "");
        if (respContent.Length > 3)
        {
          respContent = respContent.Substring(3);
          if (respContent.Length > 100)
          {
            respContent = respContent.Substring(0, 100);
          }
        }

        System.Tuple<bool, long> gzip = GzipCompressor.TryCompression(ctx);
        System.Console.WriteLine(" " + ctx.Response.StatusCode.ToString() + ": " + respContent);
      };
    }
  }


  public static class GzipCompressor
  {
    public static System.Tuple<bool, long> TryCompression(NancyContext ctx)
    {
      if (GzipCompressor.IsAcceptableToClient(ctx.Request.Headers.AcceptEncoding))
      {
        if (GzipCompressor.IsAcceptableForResponse(ctx.Response.ContentType))
        {
          return GzipCompressor.ReprocessStream(ctx);
        }
      }
      return new System.Tuple<bool, long>(false, -1);
    }

    private static bool IsAcceptableToClient(IEnumerable<string> clientAcceptEncodings)
    {
      foreach (string enc in clientAcceptEncodings)
      {
        if (enc.Contains("gzip"))
        {
          return true;
        }
      }
      return false;
    }

    private static bool IsAcceptableForResponse(string responseContentType)
    {
      return responseContentType == "application/xml";
    }

    private static System.Tuple<bool, long> ReprocessStream(NancyContext ctx)
    {
      var ms = new System.IO.MemoryStream();
      ctx.Response.Contents.Invoke(ms);
      ms.Position = 0;
      long uncompressedLength = ms.Length;
      if (ms.Length < 4096)
      {
        ctx.Response.Contents = s =>
        {
          ms.CopyTo(s);
          s.Flush();
        };
        return new System.Tuple<bool, long>(false, uncompressedLength);
      }
      else
      {
        ctx.Response.Headers["Content-Encoding"] = "gzip";
        ctx.Response.Contents = s =>
        {
          var gzip = new System.IO.Compression.GZipStream(s, System.IO.Compression.CompressionMode.Compress, true);
          ms.CopyTo(gzip);
          gzip.Close();
          s.Flush();
        };
        return new System.Tuple<bool, long>(true, uncompressedLength);
      }
    }
  }


}

