using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.Serialization;

namespace TP.Object.Process.MASExport
{
  [DataContract(Name = "ProcessJobDefinition", Namespace = "TP.Object")]
  public class JobDefinition<T> : TP.Object.BaseBusinessObject where T : IMASExportObject
  {
    private const string DATA_FILE_DIR = "\\\\toolsplus04\\databases\\mastest\\mas90-signs\\export\\";
    private const string MUTEX_FILE_DIR = "\\\\toolsplus04\\databases\\mastest\\mas90-signs\\export\\";

    [DataMember]
    public string JobFriendlyName { get; set; }
    [DataMember]
    public string JobCompiledName { get; set; }
    [DataMember]
    public string DataFileName { get; set; }
    [DataMember]
    public string MutexFileName { get; set; }
    [DataMember]
    public IEnumerable<T> DataFileContent { get; set; }
    [DataMember]
    public DisplayMode DisplayMode { get; set; } // DISPLAY or MANUAL

    public string DataFileURI { get { return DATA_FILE_DIR + this.DataFileName; } }
    public string MutexFileURI { get { return MUTEX_FILE_DIR + this.MutexFileName; } }
  }

  [DataContract(Name = "ProcessJobErrorReport", Namespace = "TP.Object")]
  public class JobErrorReport
  {
    [DataMember]
    public string Key { get; set; }
    [DataMember]
    public string Field { get; set; }
    [DataMember]
    public string Value { get; set; }

    public override string ToString()
    {
      if (this.Field == "CREATE")
      {
        return string.Format("{0}: didn't create record", this.Key);
      }
      else
      {
        return string.Format("{0}: didn't set '{1}' to '{2}'", this.Key, this.Field, this.Value);
      }
    }
  }

  public enum DisplayMode // stringified in TP.O.FE.P.MASExport.doExport<T>()
  {
    Display,
    Manual,
  };

  public class StatusChangeEventArgs : EventArgs
  {
    public string Module { get; private set; }
    public string[] Message { get; private set; }
    public StatusChangeEventArgs(string module, params string[] message)
    {
      this.Module = module;
      this.Message = message;
    }
    public StatusChangeEventArgs(string module, IEnumerable<string> message) : this(module, message.ToArray()) { }
  }
}
