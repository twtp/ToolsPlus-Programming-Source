using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace TP.Object
{

  #region Objects

  [DataContract(Name = "BaseBusinessObject", Namespace = "TP.Object")]
  public abstract class BaseBusinessObject : TP.Base.SerializableObject
  {
    protected string nullableString(object field)
    {
      if (field == DBNull.Value)
      {
        return null;
      }
      return (string)field;
    }

    protected Nullable<T> nullable<T>(object field) where T : struct
    {
      if (field == DBNull.Value)
      {
        return null;
      }
      return (T)field;
    }

    public virtual string DumpSelf() { return "NOT IMPLEMENTED, HashCode=" + this.GetHashCode().ToString(); }
  }

  [DataContract(Name = "ObjectResponse", Namespace = "TP.Object")]
  public class ObjectResponse<T> : TP.Base.SerializableObject
  {
    [DataMember]
    public bool IsSuccessful { get; protected set; }
    [DataMember]
    public T ReturnObject { get; protected set; }
    [DataMember]
    public string ErrorMessage { get; protected set; }

    public ObjectResponse(T returnObject)
    {
      this.IsSuccessful = true;
      this.ReturnObject = returnObject;
      this.ErrorMessage = string.Empty;
    }

    public ObjectResponse(string errorMessage)
    {
      this.IsSuccessful = false;
      this.ReturnObject = default(T);
      this.ErrorMessage = errorMessage;
    }
  }

  #endregion


  #region Enums

  /*[DataContract(Name = "ItemCostType")]
  public enum ItemCostType // table ItemCostType
  {
    [EnumMember]
    StdCost = 1,
    [EnumMember]
    NewCost = 2,
    [EnumMember]
    ListPrice = 3,
    [EnumMember]
    MAPP = 4,
  };*/

  [DataContract(Name = "ApplicationModule")]
  public enum ApplicationModule // table ApplicationModule
  {
    [EnumMember]
    OtherOrUnknown = 1,
    [EnumMember]
    PoinvInventoryMaintenance = 2,
    [EnumMember]
    PoinvGeneralSpreadsheetImport = 3,
    [EnumMember]
    PoinvNewProductsImport = 4,
  };

  [DataContract(Name = "ProductLineContactInfoType")]
  public enum ProductLineContactInfoType // no table
  {
    [EnumMember]
    GeneralContact = 0,
    [EnumMember]
    Payables = 1,
    [EnumMember]
    Purchasing = 2,
    [EnumMember]
    Service = 3,
    [EnumMember]
    Parts = 4,
    [EnumMember]
    Rebates = 5,
  };

  [DataContract(Name = "AddressType")]
  public enum AddressZoningType // no table
  {
    [EnumMember]
    Commercial = 0,
    [EnumMember]
    Residential = 1,
  };

  [DataContract(Name = "DropshipChargeCalculationType")]
  public enum DropshipChargeCalculationType // no table
  {
    [EnumMember]
    Percentage = 0,
    [EnumMember]
    Flat = 1,
  };

  #endregion


  #region Interfaces

  public interface IMASExportObject
  {
    // public static void ToCSVHeader(System.Text.StringBuilder sb);
    void ToCSVLine(StringBuilder sb);
  }

  public static class IMASExportObjectExtensions
  {
    private const string CSV_QUOTE_CHAR = "\"";
    private const string CSV_ESCAPED_QUOTE = "\"\"";
    private static readonly char[] CSV_ESCAPED_CHARS = { ',', '"', '\n' };
    public static string csvEscape<T>(this T obj, string s)
    {
      if (s == null) { return string.Empty; }
      if (s.IndexOfAny(CSV_ESCAPED_CHARS) == -1) { return s; }
      return CSV_QUOTE_CHAR + s.Replace(CSV_QUOTE_CHAR, CSV_ESCAPED_QUOTE) + CSV_QUOTE_CHAR;
    }
    public static string csvEscape<T>(this T obj, bool? s)
    {
      return s.HasValue && s.Value == true ? "Y" : s.HasValue ? "N" : "";
    }
    public static string csvEscape<T>(this T obj, int? s) { return obj.csvEscape(s.ToString()); }
    public static string csvEscape<T>(this T obj, decimal? s) { return obj.csvEscape(s.ToString()); }
  }

  public interface IExportableToMAS
  {
    bool FlagExported(Func<bool> recomparer, params string[] options);
  }

  #endregion


  #region Attributes

  public static class AttributeHelper
  {
    
    private static Dictionary<Type, TypeCache> cache = new Dictionary<Type, TypeCache>();

    public static TAttribute GetAttribute<TObject, TAttribute>(string propName) where TObject : BaseBusinessObject where TAttribute : IDatabaseField
    {
      Type o = typeof(TObject);
      if (!cache.ContainsKey(o))
      {
        cache.Add(o, new TypeCache(o));
      }

      if (!cache[o].ContainsKey(propName))
      {
        return default(TAttribute);
      }

      Type a = typeof(TAttribute);
      if (!cache[o][propName].ContainsKey(a))
      {
        return default(TAttribute);
      }

      return (TAttribute)cache[o][propName][a];
    }

    public static bool HasAttribute<TObject, TAttribute>(string propName) where TObject : BaseBusinessObject where TAttribute : IDatabaseField
    {
      TAttribute attr = GetAttribute<TObject, TAttribute>(propName);
      return !attr.Equals(default(TAttribute));
    }

    public static IEnumerable<string> GetProperties<TObject>()
    {
      Type o = typeof(TObject);
      if (!cache.ContainsKey(o))
      {
        cache.Add(o, new TypeCache(o));
      }

      return cache[o].Keys;
    }

    public static bool Validate<TObject>(string propName, object value) where TObject : BaseBusinessObject
    {
      if (value == null)
      {
        NullableAttribute nullable = AttributeHelper.GetAttribute<TObject, NullableAttribute>(propName);
        return nullable != null && nullable.IsNullable;
      }
      else
      {
        SqlDbTypeAttribute sqlType = AttributeHelper.GetAttribute<TObject, SqlDbTypeAttribute>(propName);
        switch (sqlType.SqlDbType)
        {
          case System.Data.SqlDbType.VarChar:
          case System.Data.SqlDbType.Char:
          case System.Data.SqlDbType.Text:
          case System.Data.SqlDbType.NVarChar:
          case System.Data.SqlDbType.NChar:
          case System.Data.SqlDbType.NText:
            return ValidateString<TObject>(propName, (string)value);
          case System.Data.SqlDbType.Int:
          case System.Data.SqlDbType.SmallInt:
          case System.Data.SqlDbType.TinyInt:
            return ValidateInteger<TObject>(propName, (int)value);
          case System.Data.SqlDbType.Decimal:
          case System.Data.SqlDbType.Money:
          case System.Data.SqlDbType.Real:
          case System.Data.SqlDbType.Float:
          case System.Data.SqlDbType.SmallMoney:
            return ValidateDecimal<TObject>(propName, (decimal)value);
          case System.Data.SqlDbType.Date:
          case System.Data.SqlDbType.DateTime:
          case System.Data.SqlDbType.DateTime2:
          case System.Data.SqlDbType.SmallDateTime:
            // no validation checks for datetime
            // maybe to do? future only, past only
            return true;
          case System.Data.SqlDbType.Bit:
            // no validation checks for boolean
            return true;
          default:
            return true;
        }
      }
    }

    public static bool ValidateString<TObject>(string propName, string value) where TObject : BaseBusinessObject
    {
      IDatabaseFieldConstraint c;
      
      c = AttributeHelper.GetAttribute<TObject, StringMaximumLengthAttribute>(propName);
      if (c != null && c.IsValid(value) == false)
      {
        return false;
      }

      return true;
    }

    public static bool ValidateInteger<TObject>(string propName, int? value) where TObject : BaseBusinessObject
    {
      IDatabaseFieldConstraint c;

      c = AttributeHelper.GetAttribute<TObject, IntegerMinimumAttribute>(propName);
      if (c != null && c.IsValid(value) == false)
      {
        return false;
      }
      c = AttributeHelper.GetAttribute<TObject, IntegerMaximumAttribute>(propName);
      if (c != null && c.IsValid(value) == false)
      {
        return false;
      }

      return true;
    }

    public static bool ValidateDecimal<TObject>(string propName, decimal? value) where TObject : BaseBusinessObject
    {
      IDatabaseFieldConstraint c;

      c = AttributeHelper.GetAttribute<TObject, DecimalMinimumAttribute>(propName);
      if (c != null && c.IsValid(value) == false)
      {
        return false;
      }
      c = AttributeHelper.GetAttribute<TObject, DecimalMaximumAttribute>(propName);
      if (c != null && c.IsValid(value) == false)
      {
        return false;
      }

      return true;
    }


    private class TypeCache : Dictionary<string, PropCache>
    {
      public TypeCache(Type obj) : base()
      {
        foreach (System.Reflection.PropertyInfo pi in obj.GetProperties())
        {
          IDatabaseField[] attrs = (IDatabaseField[])pi.GetCustomAttributes(typeof(IDatabaseField), true);
          if (attrs.Length > 0)
          {
            this.Add(pi.Name, new PropCache(attrs));
          }
        }
      }
    }

    private class PropCache : Dictionary<Type, IDatabaseField>
    {
      public PropCache(IDatabaseField[] attrs) : base()
      {
        foreach (IDatabaseField attr in attrs)
        {
          this.Add(attr.GetType(), attr);
        }
      }
    }
  }



  public interface IDatabaseField
  {
    // nothing, just a marker that it's a db-handled property
  }

  public interface IDatabaseFieldConstraint
  {
    bool IsValid(object value);
  }



  public class FieldNameAttribute : Attribute, IDatabaseField
  {
    public FieldNameAttribute(string fieldName)
    {
      this.FieldName = fieldName;
    }
    public string FieldName { get; private set; }
  }

  public class PrimaryKeyAttribute : Attribute, IDatabaseField
  {
    public PrimaryKeyAttribute(bool primaryKey)
    {
      this.IsPrimaryKey = primaryKey;
    }
    public bool IsPrimaryKey { get; private set; }
  }
  
  public class IsModifiableAttribute : Attribute, IDatabaseField
  {
    public IsModifiableAttribute(bool isModifiable)
    {
      this.IsModifiable = isModifiable;
    }
    public bool IsModifiable { get; private set; }
  }

  public class SqlDbTypeAttribute : Attribute, IDatabaseField
  {
    public SqlDbTypeAttribute(System.Data.SqlDbType sqlDbType)
    {
      this.SqlDbType = sqlDbType;
    }
    public System.Data.SqlDbType SqlDbType { get; private set; }
  }


  public class NullableAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public NullableAttribute(bool isNullable)
    {
      this.IsNullable = isNullable;
    }
    public bool IsNullable { get; private set; }

    public bool IsValid(object value)
    {
      return value != null || this.IsNullable;
    }
  }

  public class StringMaximumLengthAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public StringMaximumLengthAttribute(int maxLength)
    {
      this.MaxLength = maxLength;
    }
    public int MaxLength { get; private set; }

    public bool IsValid(object value)
    {
      return ((string)value).Length <= this.MaxLength;
    }
  }

  public class StringMinimumLengthAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public StringMinimumLengthAttribute(int minLength)
    {
      this.MinLength = minLength;
    }
    public int MinLength { get; private set; }

    public bool IsValid(object value)
    {
      return ((string)value).Length >= this.MinLength;
    }
  }

  public class IntegerMinimumAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public IntegerMinimumAttribute(int minimum)
    {
      this.Minimum = minimum;
    }
    public int Minimum { get; private set; }

    public bool IsValid(object value)
    {
      return ((int)value) >= this.Minimum;
    }
  }

  public class IntegerMaximumAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public IntegerMaximumAttribute(int maximum)
    {
      this.Maximum = maximum;
    }
    public int Maximum { get; private set; }

    public bool IsValid(object value)
    {
      return ((int)value) <= this.Maximum;
    }
  }

  public class DecimalMinimumAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public DecimalMinimumAttribute(int minimum)
    {
      this.Minimum = minimum;
    }
    public int Minimum { get; private set; }

    public bool IsValid(object value)
    {
      return ((decimal)value) >= this.Minimum;
    }
  }

  public class DecimalMaximumAttribute : Attribute, IDatabaseField, IDatabaseFieldConstraint
  {
    public DecimalMaximumAttribute(int maximum)
    {
      this.Maximum = maximum;
    }
    public int Maximum { get; private set; }

    public bool IsValid(object value)
    {
      return ((decimal)value) <= this.Maximum;
    }
  }

  #endregion

}
