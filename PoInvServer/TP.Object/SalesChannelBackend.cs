using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  /*
  [DataContract(Name = "SalesChannelBackend", Namespace = "TP.Object")]
  public class SalesChannelBackend : BaseBusinessObject
  {

    #region Constructors

    protected SalesChannelBackend() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
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

    // TODO: SalesChannel has-many SalesChannelDefaultAttribute

    #endregion


  }*/

  [DataContract(Name = "SalesChannelBackend")]
  public enum SalesChannelBackend // table SalesChannelBackend
  {
    [EnumMember]
    [Backend(typeof(PointOfSaleChannel))]
    PointOfSale = 1,
    [EnumMember]
    [Backend(typeof(YahooStoreChannel))]
    YahooStore = 2,
    [EnumMember]
    [Backend(typeof(EBayStoreChannel))]
    EBayStore = 3,
  };

  public class BackendAttribute : Attribute, IDatabaseField
  {
    public BackendAttribute(Type backendType)
    {
      this.BackendType = backendType;
    }
    public Type BackendType { get; private set; }

    private ISalesChannelBackend _backend = null;
    public ISalesChannelBackend Backend {
      get {
        if (this._backend == null)
        {
          this._backend = (ISalesChannelBackend)Activator.CreateInstance(this.BackendType);
        }
        return this._backend;
      }
    }
  }


  public interface ISalesChannelBackend
  {
    // web offload
    // what else?
  }

  public class PointOfSaleChannel : ISalesChannelBackend
  {

  }

  public class YahooStoreChannel : ISalesChannelBackend
  {

  }

  public class EBayStoreChannel : ISalesChannelBackend
  {

  }

}
