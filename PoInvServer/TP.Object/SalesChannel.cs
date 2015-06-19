using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "SalesChannel", Namespace = "TP.Object")]
  public class SalesChannel : BaseBusinessObject
  {

    #region Constructors

    protected SalesChannel() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SalesChannelID { get; protected set; }

    [DataMember]
    [FieldName("BackendID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int BackendID { get; protected set; }

    [DataMember]
    [FieldName("Name")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [IsModifiable(false)]
    [StringMaximumLength(30)]
    public string Name { get; protected set; }

    #endregion


    #region Links

    // TODO: SalesChannel has-one SalesChannelBackend

    // TODO: SalesChannel has-many SalesChannelAttribute

    #endregion



  }
}
