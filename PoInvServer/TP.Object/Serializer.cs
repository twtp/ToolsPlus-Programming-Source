using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.Serialization;

namespace TP.Object
{
  public class Serializer
  {

    public static byte[] SerializeObjectToXml<T>(T obj) { return SerializeObjectToXml(obj, Encoding.UTF8); }
    public static byte[] SerializeObjectToXml<T>(T obj, Encoding encoding)
    {
      try
      {
        var ms = new MemoryStream();
        using (var tw = new XmlTextWriter(ms, encoding))
        {
          var xmlSerializer = new XmlSerializer(typeof(T));
          xmlSerializer.Serialize(tw, obj);
          ms = (MemoryStream)tw.BaseStream;
        }

        return ms.ToArray();
      }
      catch
      {
        return null;
      }
    }

    public static byte[] SerializeObjectToDataContract<T>(T obj) { return SerializeObjectToDataContract(obj, Encoding.UTF8); }
    public static byte[] SerializeObjectToDataContract<T>(T obj, Encoding encoding)
    {
      try
      {
        DataContractSerializer s = new DataContractSerializer(typeof(T));
        var ms = new MemoryStream();
        using (var tw = new XmlTextWriter(ms, encoding))
        {
          s.WriteObject(tw, obj);
          ms = (MemoryStream)tw.BaseStream;
        }

        return ms.ToArray();
      }
      catch
      {
        return null;
      }
    }

    public static T DeserializeDataContractToObject<T>(System.IO.Stream s) { return DeserializeDataContractToObject<T>(s, Encoding.UTF8); }
    public static T DeserializeDataContractToObject<T>(System.IO.Stream s, Encoding encoding)
    {
      using (StreamReader r = new StreamReader(s, encoding))
      {
        return DeserializeDataContractToObject<T>(r.ReadToEnd(), encoding);
      }
    }
    public static T DeserializeDataContractToObject<T>(string s) { return DeserializeDataContractToObject<T>(s, Encoding.UTF8); }
    public static T DeserializeDataContractToObject<T>(string s, Encoding encoding)
    {
      using (MemoryStream ms = new MemoryStream(encoding.GetBytes(s)))
      {
        return _doDeserialize<T>(ms);
      }
    }
    private static T _doDeserialize<T>(System.IO.Stream s)
    {
      DataContractSerializer dcs = new DataContractSerializer(typeof(T));
      return (T)dcs.ReadObject(s);
    }


    public static string ByteArrayToString(byte[] byteArray) { return ByteArrayToString(byteArray, Encoding.UTF8); }
    public static string ByteArrayToString(byte[] byteArray, Encoding encoding)
    {
      return encoding.GetString(byteArray);
    }
  }



  public class ObjectDumper
  {
    private int _level;
    private readonly int _indentSize;
    private readonly StringBuilder _stringBuilder;
    private readonly System.Collections.Generic.List<int> _hashListOfFoundElements;

    private ObjectDumper(int indentSize)
    {
      _indentSize = indentSize;
      _stringBuilder = new StringBuilder();
      _hashListOfFoundElements = new System.Collections.Generic.List<int>();
    }

    public static string Dump(object element)
    {
      return Dump(element, 2);
    }

    public static string Dump(object element, int indentSize)
    {
      var instance = new ObjectDumper(indentSize);
      return instance.DumpElement(element);
    }

    private string DumpElement(object element)
    {
      if (element == null || element is ValueType || element is string)
      {
        Write(FormatValue(element));
      }
      else
      {
        var objectType = element.GetType();
        if (!typeof(System.Collections.IEnumerable).IsAssignableFrom(objectType))
        {
          Write("{{{0}}}", objectType.FullName);
          _hashListOfFoundElements.Add(element.GetHashCode());
          _level++;
        }

        var enumerableElement = element as System.Collections.IEnumerable;
        if (enumerableElement != null)
        {
          foreach (object item in enumerableElement)
          {
            if (item is System.Collections.IEnumerable && !(item is string))
            {
              _level++;
              DumpElement(item);
              _level--;
            }
            else
            {
              if (!AlreadyTouched(item))
                DumpElement(item);
              else
                Write("{{{0}}} <-- bidirectional reference found", item.GetType().FullName);
            }
          }
        }
        else
        {
          System.Reflection.MemberInfo[] members = element.GetType().GetMembers(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
          foreach (var memberInfo in members)
          {
            var fieldInfo = memberInfo as System.Reflection.FieldInfo;
            var propertyInfo = memberInfo as System.Reflection.PropertyInfo;

            if (fieldInfo == null && propertyInfo == null)
              continue;

            var type = fieldInfo != null ? fieldInfo.FieldType : propertyInfo.PropertyType;
            object value = fieldInfo != null
                               ? fieldInfo.GetValue(element)
                               : propertyInfo.GetValue(element, null);

            if (type.IsValueType || type == typeof(string))
            {
              Write("{0}: {1}", memberInfo.Name, FormatValue(value));
            }
            else
            {
              var isEnumerable = typeof(System.Collections.IEnumerable).IsAssignableFrom(type);
              Write("{0}: {1}", memberInfo.Name, isEnumerable ? "..." : "{ }");

              var alreadyTouched = !isEnumerable && AlreadyTouched(value);
              _level++;
              if (!alreadyTouched)
                DumpElement(value);
              else
                Write("{{{0}}} <-- bidirectional reference found", value.GetType().FullName);
              _level--;
            }
          }
        }

        if (!typeof(System.Collections.IEnumerable).IsAssignableFrom(objectType))
        {
          _level--;
        }
      }

      return _stringBuilder.ToString();
    }

    private bool AlreadyTouched(object value)
    {
      if (value == null)
      {
        return false;
      }
      var hash = value.GetHashCode();
      for (var i = 0; i < _hashListOfFoundElements.Count; i++)
      {
        if (_hashListOfFoundElements[i] == hash)
          return true;
      }
      return false;
    }

    private void Write(string value, params object[] args)
    {
      var space = new string(' ', _level * _indentSize);

      if (args != null)
        value = string.Format(value, args);

      _stringBuilder.AppendLine(space + value);
    }

    private string FormatValue(object o)
    {
      if (o == null)
        return ("null");

      if (o is DateTime)
        return (((DateTime)o).ToShortDateString());

      if (o is string)
        return string.Format("\"{0}\"", o);

      if (o is ValueType)
        return (o.ToString());

      if (o is System.Collections.IEnumerable)
        return ("...");

      return ("{ }");
    }
  }

}
