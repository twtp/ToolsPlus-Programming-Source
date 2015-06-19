using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "Computer", Namespace = "TP.Object")]
  public class Computer : BaseBusinessObject
  {
    #region Constructors

    protected Computer() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ComputerID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ComputerID { get; protected set; }

    [DataMember]
    [FieldName("ComputerName")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(15)]
    public string ComputerName { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion


  }
}
