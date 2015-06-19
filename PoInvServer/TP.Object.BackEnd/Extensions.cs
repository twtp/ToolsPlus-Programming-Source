using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using TP.Database;

namespace TP.Object.BackEnd
{
  public static class Extensions
  {
    
    public static string PropertyDbFieldName<T>(this T obj, string property) where T : TP.Object.BaseBusinessObject
    {
      var att = TP.Object.AttributeHelper.GetAttribute<T, TP.Object.FieldNameAttribute>(property);
      return att == null ? null : att.FieldName;
    }

    public static bool PropertyIsPrimaryKey<T>(this T obj, string property) where T : TP.Object.BaseBusinessObject
    {
      var att = TP.Object.AttributeHelper.GetAttribute<T, TP.Object.PrimaryKeyAttribute>(property);
      return att == null ? false : att.IsPrimaryKey;
    }

    public static bool PropertyIsEditable<T>(this T obj, string property) where T : TP.Object.BaseBusinessObject
    {
      var att = TP.Object.AttributeHelper.GetAttribute<T, TP.Object.IsModifiableAttribute>(property);
      return att == null ? true : att.IsModifiable;
    }

    public static System.Data.SqlDbType PropertySqlType<T>(this T obj, string property) where T : TP.Object.BaseBusinessObject
    {
      var att = TP.Object.AttributeHelper.GetAttribute<T, TP.Object.SqlDbTypeAttribute>(property);
      return att.SqlDbType;
    }

    public static bool PropertyValueIsValid<T>(this T obj, string property, object newValue) where T : TP.Object.BaseBusinessObject
    {
      return TP.Object.AttributeHelper.Validate<T>(property, newValue);
    }

    internal static ObjectChanges PropertiesModified<T>(this T obj, T orig) where T : TP.Object.BaseBusinessObject
    {
      ObjectChanges retval = new ObjectChanges();
      Type t = typeof(T);

      foreach (string prop in TP.Object.AttributeHelper.GetProperties<T>())
      {
        object origValue = t.GetProperty(prop).GetValue(orig, null);
        object thisValue = t.GetProperty(prop).GetValue(obj, null);
        //bool equal = origValue == null ? thisValue == null : origValue.Equals(thisValue);
        bool equal = object.Equals(origValue, thisValue);

        if (obj.PropertyIsPrimaryKey(prop))
        {
          retval.AddKey(
              obj.PropertyDbFieldName(prop),
              obj.PropertySqlType(prop),
              origValue
            );
        }

        if (equal == false)
        {
          if (obj.PropertyIsEditable(prop))
          {
            if (obj.PropertyValueIsValid(prop, thisValue))
            {
              retval.AddField(
                  obj.PropertyDbFieldName(prop),
                  obj.PropertySqlType(prop),
                  thisValue
                );
            }
            else
            {
              throw new ArgumentException(string.Format("Value '{0}' of property '{1}' is invalid", thisValue, prop));
            }
          }
          else
          {
            throw new ArgumentException("attempt to change unmodifiable field '" + prop + "'");
          }
        }
      }

      return retval;
    }



    internal class ObjectChanges
    {
      private List<string> fields = new List<string>();
      private List<System.Data.SqlDbType> types = new List<System.Data.SqlDbType>();
      private List<object> values = new List<object>();
      private List<string> keyFields = new List<string>();
      private List<System.Data.SqlDbType> keyTypes = new List<System.Data.SqlDbType>();
      private List<object> keyValues = new List<object>();

      public int AddField(string fieldName, System.Data.SqlDbType type, object value)
      {
        if (this.fields.Contains(fieldName) == false)
        {
          this.fields.Add(fieldName);
          this.types.Add(type);
          this.values.Add(value);
        }
        return this.fields.Count;
      }
      public int AddKey(string fieldName, System.Data.SqlDbType type, object value)
      {
        this.keyFields.Add(fieldName);
        this.keyTypes.Add(type);
        this.keyValues.Add(value);
        return this.fields.Count;
      }

      public int Count { get { return this.fields.Count; } }

      public string UpdateFieldString
      {
        get
        {
          List<string> temp = new List<string>();
          for (int i = 0; i < this.fields.Count; i++)
          {
            temp.Add(string.Format("{0}=@f{1}", this.fields[i], i));
          }
          return string.Join(", ", temp.ToArray());
        }
      }
      public string UpdateWhereClause
      {
        get
        {
          List<string> temp = new List<string>();
          for (int i = 0; i < this.keyFields.Count; i++)
          {
            temp.Add(string.Format("{0}=@pk{1}", this.keyFields[i], i));
          }
          return string.Join(" AND ", temp.ToArray());
        }
      }
      public QueryParameter[] UpdateValues
      {
        get
        {
          List<QueryParameter> temp = new List<QueryParameter>();
          for (int i = 0; i < this.fields.Count; i++)
          {
            temp.Add(new QueryParameter(string.Format("@f{0}", i), this.types[i], this.values[i]));
          }
          for (int i = 0; i < this.keyFields.Count; i++)
          {
            temp.Add(new QueryParameter(string.Format("@pk{0}", i), this.keyTypes[i], this.keyValues[i]));
          }
          return temp.ToArray();
        }
      }

      /*public string InsertFieldString
      {
        get
        {
          return string.Join(", ", this.fields.ToArray());
        }
      }
      public string InsertValuesString
      {
        get
        {
          List<string> temp = new List<string>();
          for (int i = 0; i < this.fields.Count; i++)
          {
            temp.Add(string.Format("@pk{1}", this.fields[i], i));
          }
          return string.Join(", ", temp.ToArray());
        }
      }
      public QueryParameter[] InsertValues
      {
        get
        {
          List<QueryParameter> temp = new List<QueryParameter>();
          for (int i = 0; i < this.fields.Count; i++)
          {
            temp.Add(new QueryParameter(string.Format("@f{0}", i), this.types[i], this.values[i]));
          }
          return temp.ToArray();
        }
      }*/
    }
  }
}
