using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "ItemPrice", Namespace = "TP.Object")]
  public class ItemPrice : BaseBusinessObject, IExportableToMAS
  {
    protected const int MAXIMUM_QUANTITY = 999999;
    protected const int MAXIMUM_LEVELS = 5; // this is also assumed in places: fields region, MakeList(), ApplyList(), DbItemPrice ctor, vItemPriceCurrent

    #region Constructors

    protected ItemPrice() { }

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

    [DataMember]
    [FieldName("SalesChannelID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int SalesChannelID { get; protected set; }

    [DataMember]
    [FieldName("TimeEffective")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.DateTime)]
    public DateTime TimeEffective { get; protected set; }

    [DataMember]
    [FieldName("MaximumQuantity1")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(0)]
    [IntegerMaximum(99999999)]
    public int MaximumQuantity1 { get; protected set; }

    [DataMember]
    [FieldName("PriceValue1")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal PriceValue1 { get; protected set; }

    [DataMember]
    [FieldName("MaximumQuantity2")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(0)]
    [IntegerMaximum(99999999)]
    public int MaximumQuantity2 { get; protected set; }

    [DataMember]
    [FieldName("PriceValue2")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal PriceValue2 { get; protected set; }

    [DataMember]
    [FieldName("MaximumQuantity3")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(0)]
    [IntegerMaximum(99999999)]
    public int MaximumQuantity3 { get; protected set; }

    [DataMember]
    [FieldName("PriceValue3")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal PriceValue3 { get; protected set; }

    [DataMember]
    [FieldName("MaximumQuantity4")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(0)]
    [IntegerMaximum(99999999)]
    public int MaximumQuantity4 { get; protected set; }

    [DataMember]
    [FieldName("PriceValue4")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal PriceValue4 { get; protected set; }

    [DataMember]
    [FieldName("MaximumQuantity5")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    [IntegerMinimum(0)]
    [IntegerMaximum(99999999)]
    public int MaximumQuantity5 { get; protected set; }

    [DataMember]
    [FieldName("PriceValue5")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Decimal)]
    [DecimalMinimum(0)]
    public decimal PriceValue5 { get; protected set; }

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

    // TODO: ItemPrice has-one SalesChannel

    // ItemPrice has-one Item
    protected Item _item;
    [DataMember(EmitDefaultValue = false)]
    public Item Item { get { return this._item; } private set { this._item = value; } }
    public virtual bool FillItem() { throw new ApplicationException("kaboom"); }

    // ItemPrice has-one Person
    protected Person _person;
    [DataMember(EmitDefaultValue = false)]
    public Person Person { get { return this._person; } private set { this._person = value; } }
    public virtual bool FillPerson() { throw new ApplicationException("kaboom"); }

    #endregion


    public List<Tuple<int, decimal>> MakeList()
    {
      return MakeList(this.MaximumQuantity1, this.PriceValue1, this.MaximumQuantity2, this.PriceValue2, this.MaximumQuantity3, this.PriceValue3, this.MaximumQuantity4, this.PriceValue4, this.MaximumQuantity5, this.PriceValue5);
    }
    public static List<Tuple<int, decimal>> MakeList(int q1, decimal p1, int q2, decimal p2, int q3, decimal p3, int q4, decimal p4, int q5, decimal p5)
    {
      List<Tuple<int, decimal>> retval = new List<Tuple<int, decimal>>();
      retval.Add(new Tuple<int, decimal>(q1, p1));
      if (q1 == MAXIMUM_QUANTITY) { return retval; }
      retval.Add(new Tuple<int, decimal>(q2, p2));
      if (q2 == MAXIMUM_QUANTITY) { return retval; }
      retval.Add(new Tuple<int, decimal>(q3, p3));
      if (q3 == MAXIMUM_QUANTITY) { return retval; }
      retval.Add(new Tuple<int, decimal>(q4, p4));
      if (q4 == MAXIMUM_QUANTITY) { return retval; }
      retval.Add(new Tuple<int, decimal>(q5, p5));
      return retval;
    }

    public void ApplyList(List<Tuple<int, decimal>> priceList)
    {
      this.MaximumQuantity1 = priceList[0].Item1;
      this.PriceValue1 = priceList[0].Item2;
      this.MaximumQuantity2 = priceList.Count > 1 ? priceList[1].Item1 : 0;
      this.PriceValue2 = priceList.Count > 1 ? priceList[1].Item2 : 0;
      this.MaximumQuantity3 = priceList.Count > 2 ? priceList[2].Item1 : 0;
      this.PriceValue3 = priceList.Count > 2 ? priceList[2].Item2 : 0;
      this.MaximumQuantity4 = priceList.Count > 3 ? priceList[3].Item1 : 0;
      this.PriceValue4 = priceList.Count > 3 ? priceList[3].Item2 : 0;
      this.MaximumQuantity5 = priceList.Count > 4 ? priceList[4].Item1 : 0;
      this.PriceValue5 = priceList.Count > 4 ? priceList[4].Item2 : 0;
    }

    public decimal PriceFor(int quantity)
    {
      return PriceFor(quantity, this.MakeList());
    }
    public static decimal PriceFor(int quantity, List<Tuple<int, decimal>> priceList)
    {
      if (quantity > MAXIMUM_QUANTITY) { throw new ArgumentException("quantity exceeds supported amount!"); }
      if (priceList.Count == 0) { return 0; }
      for (int i = 0; i < priceList.Count; i++)
      {
        if (quantity < priceList[i].Item1)
        {
          return priceList[i].Item2;
        }
      }
      // should never get here
      throw new ArgumentException("quantity exceeds supported amount!");
    }

    public bool IsValid()
    {
      return IsValid(this.MakeList());
    }
    public static bool IsValid(List<Tuple<int, decimal>> priceList)
    {
      if (priceList.Count == 0) { return false; }

      int lastQ = priceList[0].Item1;
      decimal lastP = priceList[0].Item2;

      for (int i = 1; i < priceList.Count; i++)
      {
        if (lastQ > priceList[i].Item1)
        {
          return false;
        }
        if (lastP < priceList[i].Item2)
        {
          return false;
        }
        lastQ = priceList[i].Item1;
        lastP = priceList[i].Item2;
      }
      if (lastQ != MAXIMUM_QUANTITY)
      {
        return false;
      }
      return true;
    }



    public virtual bool FlagExported(Func<bool> recomparer, params string[] options) { throw new ApplicationException("kaboom"); }
  }
}
