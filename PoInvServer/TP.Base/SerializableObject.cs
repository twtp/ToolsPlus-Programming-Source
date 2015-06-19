using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;

namespace TP.Base
{
  [DataContract(Name = "SerializableObject", Namespace = "TP.Object")]
  public abstract class SerializableObject
  {
    // nothing here, just a marker that it can/should be serialized
  }

  [DataContract(Name = "ErrorResponse", Namespace = "TP.Object")]
  public class ErrorResponse : SerializableObject
  {
    public ErrorResponse(Exception ex) : this(ex.Message, ex.StackTrace) { }
    public ErrorResponse(string message, string stackTrace = null)
    {
      this.Message = message;
      this.StackTrace = stackTrace;
    }

    [DataMember]
    public string Message { get; set; }
    [DataMember]
    public string StackTrace { get; set; }
  }

  [DataContract(Name = "SuccessResponse", Namespace = "TP.Object")]
  public class SuccessResponse : SerializableObject
  {
    public SuccessResponse(string message)
    {
      this.Message = message;
    }

    [DataMember]
    public string Message { get; set; }
  }
}
