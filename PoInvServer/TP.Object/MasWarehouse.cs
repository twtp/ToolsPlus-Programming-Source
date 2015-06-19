using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "MasWarehouse", Namespace = "TP.Object")]
  public class MasWarehouse : BaseBusinessObject, IExportableToMAS
  {
    #region Constructors

    protected MasWarehouse() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("WarehouseCode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string WarehouseCode { get; protected set; }

    [DataMember]
    [FieldName("WarehouseDesc")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Description { get; protected set; }

    [DataMember]
    [FieldName("WarehouseName")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Name { get; protected set; }

    [DataMember]
    [FieldName("WarehouseAddress1")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Address1 { get; protected set; }

    [DataMember]
    [FieldName("WarehouseAddress2")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Address2 { get; protected set; }

    [DataMember]
    [FieldName("WarehouseAddress3")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string Address3 { get; protected set; }

    [DataMember]
    [FieldName("WarehouseCity")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string City { get; protected set; }

    [DataMember]
    [FieldName("WarehouseState")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(2)]
    public string State { get; protected set; }

    [DataMember]
    [FieldName("WarehouseZipCode")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(10)]
    public string PostalCode { get; protected set; }

    [DataMember]
    [FieldName("WarehouseCountryCode")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(3)]
    public string CountryCode { get; protected set; }

    [DataMember]
    [FieldName("ExportRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    #endregion


    #region Links

    // none

    #endregion

    public override string ToString()
    {
      return string.Format("{0} ({1})", this.Description, this.WarehouseCode);
    }


    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }
  }
}
