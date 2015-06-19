using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.Attributes {

  [AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
  public abstract class GenericEnumAttribute : System.Attribute {
    public abstract object TheAttribute { get; }

    // Given an enum, attribute, and attribute value, get the enumtype
    private static Dictionary<Type, Dictionary<Type, Dictionary<object, object>>> enumAttrValCache = new Dictionary<Type, Dictionary<Type, Dictionary<object, object>>>();
    public static TEnum GetEnumWith<TEnum, TAttr>(object val) where TAttr : GenericEnumAttribute /*where TEnum : Enum*/ {
      Type enumType = typeof(TEnum);
      Type attrType = typeof(TAttr);
      if (!enumAttrValCache.ContainsKey(enumType)) {
        enumAttrValCache.Add(enumType, new Dictionary<Type, Dictionary<object, object>>());

        foreach (TEnum enumItem in Enum.GetValues(enumType)) {
          foreach (object checkAttr in enumType.GetField(enumItem.ToString()).GetCustomAttributes(false)) {
            GenericEnumAttribute gea = checkAttr as GenericEnumAttribute;
            if (gea != null) {
              Type geaType = gea.GetType();
              if (!enumAttrValCache[enumType].ContainsKey(geaType)) {
                enumAttrValCache[enumType].Add(geaType, new Dictionary<object, object>());
              }
              enumAttrValCache[enumType][geaType].Add(gea.TheAttribute, enumItem);
            }
          }
        }
      }

      if (enumAttrValCache.ContainsKey(enumType)) {
        if (enumAttrValCache[enumType].ContainsKey(attrType)) {
          if (enumAttrValCache[enumType][attrType].ContainsKey(val)) {
            return (TEnum)enumAttrValCache[enumType][attrType][val];
          }
        }
      }
      return default(TEnum);
    }
  }

  [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
  public abstract class GenericClassAttribute : System.Attribute {
    public abstract object TheAttribute { get; }

    // Given a typeof(class) and attribute, return the value of the attribute
    // as an object (to be cast correctly). For an AllowMultiple=true, return an
    // array of attributes (again, to be cast).
    private static Dictionary<Type, Dictionary<Type, object>> attrCache = new Dictionary<Type,Dictionary<Type,object>>();
    public static object GetAttributeValue<TAttr>(Type classType) where TAttr: GenericClassAttribute {
      if (!attrCache.ContainsKey(classType)) {
        attrCache.Add(classType, new Dictionary<Type, object>());
      }
      Type attrType = typeof(TAttr);
      if (!attrCache[classType].ContainsKey(attrType)) {
        TAttr[] list = (TAttr[])classType.GetCustomAttributes(attrType, false);
        switch (list.Length) {
          case 0:
            attrCache[classType].Add(attrType, null);
            break;
          case 1:
            attrCache[classType].Add(attrType, list[0].TheAttribute);
            break;
          default:
            attrCache[classType].Add(attrType, list);
            break;
        }
      }
      return attrCache[classType][attrType];
    }

    // given an assembly and attribute, return the class(es) that have it
    private static Dictionary<System.Reflection.Assembly, Dictionary<Type, List<Type>>> assAttrCache = new Dictionary<System.Reflection.Assembly, Dictionary<Type, List<Type>>>();
    public static List<Type> GetClassesWith<TAttr>(System.Reflection.Assembly assembly) where TAttr: GenericClassAttribute {
      if (!assAttrCache.ContainsKey(assembly)) {
        assAttrCache.Add(assembly, new Dictionary<Type, List<Type>>());

        foreach (Type checkType in assembly.GetTypes()) {
          foreach (object checkAttr in checkType.GetCustomAttributes(false)) {
            GenericClassAttribute gca = checkAttr as GenericClassAttribute;
            if (gca != null) {
              Type gcaType = gca.GetType();
              if (!assAttrCache[assembly].ContainsKey(gcaType)) {
                assAttrCache[assembly].Add(gcaType, new List<Type>());
              }
              assAttrCache[assembly][gcaType].Add(checkType);
            }
          }
        }

      }

      Type thisType = typeof(TAttr);
      if (assAttrCache[assembly].ContainsKey(thisType)) {
        return assAttrCache[assembly][thisType];
      }
      else {
        return null;
      }
    }

    // given an assembly, attribute, and value, return the class(es) that have both
    private static Dictionary<System.Reflection.Assembly, Dictionary<Type, Dictionary<object, List<Type>>>> assAttrValCache = new Dictionary<System.Reflection.Assembly, Dictionary<Type, Dictionary<object, List<Type>>>>();
    public static List<Type> GetClassesWith<TAttr>(object val, System.Reflection.Assembly assembly) where TAttr : GenericClassAttribute {
      if (!assAttrValCache.ContainsKey(assembly)) {
        assAttrValCache.Add(assembly, new Dictionary<Type, Dictionary<object, List<Type>>>());

        foreach (Type checkType in assembly.GetTypes()) {
          foreach (object checkAttr in checkType.GetCustomAttributes(false)) {
            GenericClassAttribute gca = checkAttr as GenericClassAttribute;
            if (gca != null) {
              Type gcaType = gca.GetType();
              if (!assAttrValCache[assembly].ContainsKey(gcaType)) {
                assAttrValCache[assembly].Add(gcaType, new Dictionary<object, List<Type>>());
              }
              if (!assAttrValCache[assembly][gcaType].ContainsKey(gca.TheAttribute)) {
                assAttrValCache[assembly][gcaType].Add(gca.TheAttribute, new List<Type>());
              }
              assAttrValCache[assembly][gcaType][gca.TheAttribute].Add(checkType);
            }
          }
        }

      }

      Type thisType = typeof(TAttr);
      if (assAttrValCache[assembly].ContainsKey(thisType)) {
        if (assAttrValCache[assembly][thisType].ContainsKey(val)) {
          return assAttrValCache[assembly][thisType][val];
        }
      }
      return null;
    }

  }
  
}
