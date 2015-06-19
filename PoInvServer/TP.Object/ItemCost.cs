using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemCost", Namespace = "TP.Object")]
  public class ItemCost : BaseBusinessObject
  {
    #region Constructors

    protected ItemCost() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TupleID { get; protected set; }

    [DataMember]
    [FieldName("ItemID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ItemID { get; protected set; }

    /*[DataMember]
    [FieldName("ItemCostTypeID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public ItemCostType CostType { get; protected set; }*/
    [DataMember]
    [FieldName("ItemCostType")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(12)]
    public string CostType { get; protected set; }

    [DataMember]
    [FieldName("CostValue")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal CostValue { get; protected set; }

    [DataMember]
    [FieldName("TimeEffective")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime TimeEffective { get; protected set; }

    [DataMember]
    [FieldName("PersonID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int PersonID { get; protected set; }

    [DataMember]
    [FieldName("ApplicationModuleID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public ApplicationModule ApplicationModuleID { get; protected set; }

    #endregion


    #region Links

    // ItemCost has-one Person
    protected Person _person;
    [DataMember(EmitDefaultValue = false)]
    public Person Person { get { return this._person; } private set { this._person = value; } }
    public virtual bool FillPerson() { throw new ApplicationException("kaboom"); }

    #endregion
  }
}
