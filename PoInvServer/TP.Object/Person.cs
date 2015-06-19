using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "Person", Namespace = "TP.Object")]
  public class Person : BaseBusinessObject
  {
    #region Constructors

    protected Person() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int PersonID { get; protected set; }

    [DataMember]
    [FieldName("NTUsername")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string NTUsername { get; protected set; }

    [DataMember]
    [FieldName("ShortName")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(50)]
    public string ShortName { get; protected set; }

    [DataMember]
    [FieldName("Deleted")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsDeleted { get; protected set; }

    [DataMember]
    [FieldName("SpecialUser")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsSpecialUser { get; protected set; }

    #endregion


    #region Links

    // TODO: Person has-many ModulePermission

    #endregion

    public bool IsValidNTUsername() { return IsValidNTUsername(this.NTUsername); }
    public static bool IsValidNTUsername(string ntUsername)
    {
      return AttributeHelper.Validate<Person>("NTUsername", ntUsername) &&
             System.Text.RegularExpressions.Regex.IsMatch(ntUsername, @"^[a-z]+$");
    }



  }
}
