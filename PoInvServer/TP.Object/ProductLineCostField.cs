using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{

  [DataContract(Name = "ProductLineCostField", Namespace = "TP.Object")]
  public class ProductLineCostField : BaseBusinessObject
  {

    #region Constructors

    protected ProductLineCostField() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int TupleID { get; protected set; }

    [DataMember]
    [FieldName("ProductLineID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int ProductLineID { get; protected set; }

    [DataMember]
    [FieldName("Name")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(20)]
    public string Name { get; protected set; }

    [DataMember]
    [FieldName("Description")]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(99)]
    public string Description { get; protected set; }

    [DataMember]
    [FieldName("IsDeleted")]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsDeleted { get; protected set; }

    #endregion


    #region Links

    // ProductLineCostField has-one ProductLine
    protected ProductLine _productline;
    [DataMember(EmitDefaultValue = false)]
    public ProductLine ProductLine { get { return this._productline; } private set { this._productline = value; } }
    public virtual bool FillProductLine() { throw new ApplicationException("kaboom"); }

    // ProductLineCostField has-many ItemCost
    /*protected IEnumerable<ItemCost> _itemcostcurrent;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<ItemCost> CurrentItemCosts { get { return this._itemcostcurrent; } private set { this._itemcostcurrent = value; } }
    public virtual bool FillCurrentItemCosts() { throw new ApplicationException("kaboom"); }*/

    #endregion


  }
}
