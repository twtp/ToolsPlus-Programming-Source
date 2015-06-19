using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemKitContent", Namespace = "TP.Object")]
  public class ItemKitContent : BaseBusinessObject, IExportableToMAS
  {
    #region Constructors

    protected ItemKitContent() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("SalesKitNo")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string ItemNumber { get; protected set; }

    [DataMember]
    [FieldName("ComponentItemCode")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.VarChar)]
    [StringMaximumLength(30)]
    public string SubItemNumber { get; protected set; }

    [DataMember]
    [FieldName("QuantityPerAssembly")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int Quantity { get; protected set; }

    [DataMember]
    [FieldName("CoreComponent")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool IsCoreComponent { get; protected set; }

    [DataMember]
    [FieldName("ExportRequired")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Bit)]
    public bool ExportRequired { get; protected set; }

    #endregion


    #region Links

    // ItemKitContent has-one Item (as Item)
    protected Item _item;
    [DataMember(EmitDefaultValue = false)]
    public Item Item { get { return this._item; } private set { this._item = value; } }
    public virtual bool FillItem() { throw new ApplicationException("kaboom"); }

    // ItemKitContent has-one Item (as SubItem)
    protected Item _subitem;
    [DataMember(EmitDefaultValue = false)]
    public Item SubItem { get { return this._subitem; } private set { this._subitem = value; } }
    public virtual bool FillSubItem() { throw new ApplicationException("kaboom"); }

    #endregion



    public override string ToString()
    {
      return string.Format("{0} => {1}x {2}", this.ItemNumber, this.Quantity, this.SubItemNumber);
    }


    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }
  }
}
