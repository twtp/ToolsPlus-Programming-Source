using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace TP.Object
{
  [DataContract(Name = "InventoryLocation", Namespace = "TP.Object")]
  public class InventoryLocation : BaseBusinessObject, IComparable<InventoryLocation>
  {
    #region Constructors

    protected InventoryLocation() { }

    #endregion


    #region Fields

    [DataMember]
    [FieldName("ID")]
    [PrimaryKey(true)]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int LocationID { get; protected set; }

    [DataMember]
    [FieldName("LocationTypeID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int LocationTypeID { get; protected set; }

    [DataMember]
    [FieldName("WarehouseID")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int WarehouseID { get; protected set; }

    [DataMember]
    [FieldName("Division1")]
    [IsModifiable(false)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int Division1 { get; protected set; }

    [DataMember]
    [FieldName("Division2")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? Division2 { get; protected set; }

    [DataMember]
    [FieldName("Division3")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? Division3 { get; protected set; }

    [DataMember]
    [FieldName("Division4")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? Division4 { get; protected set; }

    [DataMember]
    [FieldName("Division5")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? Division5 { get; protected set; }

    [DataMember]
    [FieldName("LeftID")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? LeftID { get; protected set; }

    [DataMember]
    [FieldName("RightID")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? RightID { get; protected set; }

    [DataMember]
    [FieldName("FrontID")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? FrontID { get; protected set; }

    [DataMember]
    [FieldName("RearID")]
    [IsModifiable(false)]
    [Nullable(true)]
    [SqlDbType(System.Data.SqlDbType.Int)]
    public int? RearID { get; protected set; }

    // NOTE that LocationName and LocationTypeName are also DataMembers for the handhelds

    #endregion


    #region Links

    // InventoryLocation has-many InventoryLocationContent
    protected IEnumerable<InventoryLocationContent> _contents;
    [DataMember(EmitDefaultValue = false)]
    public IEnumerable<InventoryLocationContent> Contents { get { return this._contents; } private set { this._contents = value; } }
    public virtual bool FillContents() { throw new ApplicationException("kaboom"); }
    public virtual bool FillContents(System.Data.SqlClient.SqlTransaction txn) { throw new ApplicationException("kaboom"); }

    // (4) InventoryLocation has-one InventoryLocation
    protected InventoryLocation _neighbor_left, _neighbor_right, _neighbor_front, _neighbor_rear;
    [DataMember(EmitDefaultValue = false)]
    public InventoryLocation NeighborLeft { get { return this._neighbor_left; } private set { this._neighbor_left = value; } }
    [DataMember(EmitDefaultValue = false)]
    public InventoryLocation NeighborRight { get { return this._neighbor_right; } private set { this._neighbor_right = value; } }
    [DataMember(EmitDefaultValue = false)]
    public InventoryLocation NeighborFront { get { return this._neighbor_front; } private set { this._neighbor_front = value; } }
    [DataMember(EmitDefaultValue = false)]
    public InventoryLocation NeighborRear { get { return this._neighbor_rear; } private set { this._neighbor_rear = value; } }
    public virtual bool FillNeighbors() { throw new ApplicationException("kaboom"); }
    public virtual bool FillNeighbors(System.Data.SqlClient.SqlTransaction txn) { throw new ApplicationException("kaboom"); }

    #endregion


    [DataMember]
    public string LocationTypeName { get { return db[this.LocationTypeID].TypeName; } set { } }
    [DataMember]
    public string LocationName { get { return db[this.LocationTypeID].Name(this); } set { } }

    public bool IsValidPickZone { get { return db[this.LocationTypeID].IsValidPickZone; } }
    public bool IsValidPutZone { get { return db[this.LocationTypeID].IsValidPutZone; } }
    public bool IsTypeStandard { get { return db[this.LocationTypeID].IsTypeStandard; } }
    public bool IsTypePrepack { get { return db[this.LocationTypeID].IsTypePrepack; } }
    public bool IsTypePostpack { get { return db[this.LocationTypeID].IsTypePostpack; } }
    public bool IsStrictlyManaged { get { return db[this.LocationTypeID].IsStrictlyManaged; } }
    public bool IsTransactionRequired { get { return db[this.LocationTypeID].IsTransactionRequired; } }




    // For putting stuff away, we need to see if the positions are open and contiguous
    // The delegate stuff here is awful, but needed for sending in a fill that uses a transaction.
    public bool CanSpanTo(InventoryLocation end, Action<InventoryLocation> fillContents = null, Action<InventoryLocation> fillNeighbors = null)
    {
      if (fillContents == null) { fillContents = _ => _.FillContents(); }
      if (fillNeighbors == null) { fillNeighbors = _ => _.FillNeighbors(); }

      var test = GetPotentialSpanMembers(end, fillContents, fillNeighbors);
      return test != null;
    }
    public InventoryLocation[] GetPotentialSpanMembers(InventoryLocation end, Action<InventoryLocation> fillContents = null, Action<InventoryLocation> fillNeighbors = null)
    {
      if (fillContents == null) { fillContents = _ => _.FillContents(); }
      if (fillNeighbors == null) { fillNeighbors = _ => _.FillNeighbors(); }

      fillContents(this);
      if (this.Contents.Count() != 0)
      {
        return null;
      }

      if (this.LocationID == end.LocationID) { return new InventoryLocation[] { this }; }
      var iters = db[this.LocationTypeID].SpanIteratorList;
      if (iters == null)
      {
        return null;
      }
      foreach (var iter in iters)
      {
        var span = this.span1DElements(end, iter, fillContents, fillNeighbors);
        if (span != null) { return span.ToArray(); }
      }
      return null;
    }

    public InventoryLocation[] GetCurrentSpanMembers(InventoryLocation end, Action<InventoryLocation> fillContents = null, Action<InventoryLocation> fillNeighbors = null, Func<int, IEnumerable<InventoryLocationContent>> getContentsByMasterID = null)
    {
      if (fillContents == null) { fillContents = _ => _.FillContents(); }
      if (fillNeighbors == null) { fillNeighbors = _ => _.FillNeighbors(); }
      if (getContentsByMasterID == null) { throw new ArgumentException("getContentsByMasterID is required"); }

      fillContents(this);
      if (this.Contents.Count() == 0)
      {
        return null;
      }

      int realEnd = GetSpanEndpointLocationID(fillContents, getContentsByMasterID);
      if (end.LocationID != realEnd)
      {
        throw new ArgumentException("incorrect span endpoint");
      }

      if (this.LocationID == end.LocationID) { return new InventoryLocation[] { this }; }
      var iters = db[this.LocationTypeID].SpanIteratorList;
      if (iters == null)
      {
        return null;
      }
      foreach (var iter in iters)
      {
        var span = this.span1DElements(end, iter, fillContents, fillNeighbors);
        if (span != null) { return span.ToArray(); }
      }
      return null;
    }

    private List<InventoryLocation> span1DElements(InventoryLocation other, Func<InventoryLocation, InventoryLocation> iter, Action<InventoryLocation> fillContents, Action<InventoryLocation> fillNeighbors)
    {
      if (this.LocationTypeID != other.LocationTypeID) { return null; }

      List<InventoryLocation> retval = new List<InventoryLocation>();
      if (this.LocationID == other.LocationID)
      {
        retval.Add(this);
        return retval;
      }
      
      /*InventoryLocation start, end, pointer;
      start = this < other ? this : other;
      end = this < other ? other : this;
      pointer = start;*/
      InventoryLocation start = this;
      InventoryLocation pointer = this;
      InventoryLocation end = other;

      while (true)
      {
        // no more neighbors in iter's direction
        if (pointer == null) { return null; }
        // must be empty, or contents paired with starting location
        fillContents(pointer);
        if (pointer.Contents.Count() > 0 && pointer.Contents.Any(_ => _.MasterID != start.LocationID)) { return null; }
        // add this location to the running list, fill its pointers, and continue
        retval.Add(pointer);
        if (pointer.LocationID == end.LocationID) { return retval; }
        fillNeighbors(pointer);
        pointer = iter(pointer);
      }
    }

    public int GetSpanEndpointLocationID(Action<InventoryLocation> fillContents, Func<int, IEnumerable<InventoryLocationContent>> getContentsByMasterID)
    {
      fillContents(this);

      if (this.Contents.Count() == 0)
      {
        throw new ArgumentException(string.Format("given location {0} is empty, no span exists", this.LocationID));
      }

      if (this.Contents.First().MasterID != this.LocationID)
      {
        throw new ArgumentException(string.Format("given location {0} is not the starting point of the span, {1} is", this.LocationID, this.Contents.First().MasterID));
      }

      // this assumes that the span starting point is the minimum location id,
      // and the span ending point is the maximum location id.
      var spanContents = getContentsByMasterID(this.LocationID);
      int lid = spanContents.Max(_ => _.LocationID);
      return lid;
    }





    // instead of a giant inheritance tree, add the differences to a
    // LocTypeDef class, and then lookup in a giant "db" for parameters.
    //
    // TODO: fillNeighbors() calls for iterator spanning are not transacted?
    // should they be? right now, iter is only called from span1DElements(),
    // and that fills the neighbors previously, so the call in the delegate is
    // a no-op.
    private static Dictionary<int, LocTypeDef> db = new Dictionary<int,LocTypeDef> {
        {0, new LocTypeDef { TypeName = "Palletized",
                             Name = _ => string.Format("{0} {1} {2} {3}", _.Division1, numberToLetter(_.Division2), _.Division3, numberToLetter(_.Division4)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = new Func<InventoryLocation, InventoryLocation>[] { l => { l.FillNeighbors(); return l.NeighborRight; }, l => { l.FillNeighbors(); return l.NeighborRear; } }
          }},
        {1, new LocTypeDef { TypeName = "Flow Rack",
                             Name = _ => string.Format("{0} {1} {2} {3}", _.Division1, numberToLetter(_.Division2), _.Division3, numberToLetter(_.Division4)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = new Func<InventoryLocation, InventoryLocation>[] { l => { l.FillNeighbors(); return l.NeighborRight; } }
          }},
        {2, new LocTypeDef { TypeName = "Horiz. Carousel",
                             Name = _ => string.Format("{0} {1} {2} {3}", _.Division2, _.Division3, _.Division4, numberToLetter(_.Division5)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = new Func<InventoryLocation, InventoryLocation>[] { l => { l.FillNeighbors(); return l.NeighborRight; } }
          }},
        {3, new LocTypeDef { TypeName = "Vert. Carousel",
                             Name = _ => string.Format("{0} {1} {2}", _.Division3, _.Division4, numberToLetter(_.Division5)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {4, new LocTypeDef { TypeName = "Longs",
                             Name = _ => string.Format("{0} {1} {2} {3}", _.Division1, numberToLetter(_.Division2), _.Division3, numberToLetter(_.Division4)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = new Func<InventoryLocation, InventoryLocation>[] { l => { l.FillNeighbors(); return l.NeighborRight; } }
          }},
        {5, new LocTypeDef { TypeName = "Unmanaged",
                             Name = _ => string.Format("{0}", numberToLetter(_.Division1)),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = false,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {6, new LocTypeDef { TypeName = "Transfer Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = false,
                             IsTypePostpack = true,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {7, new LocTypeDef { TypeName = "Receiving Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = false,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {8, new LocTypeDef { TypeName = "Shipping Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = false,
                             IsTypePostpack = true,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {9, new LocTypeDef { TypeName = "Will Call Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = false,
                             IsTypePostpack = true,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {10, new LocTypeDef { TypeName = "Vehicle",
                             Name = _ => string.Format("{0}", "Tools Plus Truck"),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {11, new LocTypeDef { TypeName = "Tote",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = true,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = true,
                             SpanIteratorList = null
          }},
        {12, new LocTypeDef { TypeName = "Holding",
                             Name = _ => string.Format("{0} {1}", _.Division1, numberToLetter(_.Division2)),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = true,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = true,
                             SpanIteratorList = new Func<InventoryLocation, InventoryLocation>[] { l => { l.FillNeighbors(); return l.NeighborRight; } }
          }},
        {13, new LocTypeDef { TypeName = "Backorder",
                             Name = _ => string.Format("{0} {1}", _.Division1, numberToLetter(_.Division2)),
                             IsValidPickZone = false,
                             IsValidPutZone = false,
                             IsTypeStandard = false,
                             IsTypePrepack = true,
                             IsTypePostpack = true,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = true,
                             SpanIteratorList = null
          }},
        {14, new LocTypeDef { TypeName = "Restocking Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = true,
                             IsValidPutZone = false,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {15, new LocTypeDef { TypeName = "Sales Order Staging",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = true,
                             IsValidPutZone = false,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {16, new LocTypeDef { TypeName = "Tunnel",
                             Name = _ => string.Format("{0}", _.Division1 == 1 ? "Left" : "Right"),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {17, new LocTypeDef { TypeName = "Upstairs",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }},
        {18, new LocTypeDef { TypeName = "Container",
                             Name = _ => string.Format("{0}", _.Division1),
                             IsValidPickZone = true,
                             IsValidPutZone = true,
                             IsTypeStandard = true,
                             IsTypePrepack = false,
                             IsTypePostpack = false,
                             IsStrictlyManaged = true,
                             IsTransactionRequired = false,
                             SpanIteratorList = null
          }}
      };

    private class LocTypeDef
    {
      public string TypeName { get; set; }
      public Func<InventoryLocation, string> Name { get; set; }
      public bool IsValidPickZone { get; set; }
      public bool IsValidPutZone { get; set; }
      public bool IsTypeStandard { get; set; }
      public bool IsTypePrepack { get; set; }
      public bool IsTypePostpack { get; set; }
      public bool IsStrictlyManaged { get; set; }
      public bool IsTransactionRequired { get; set; }
      public Func<InventoryLocation, InventoryLocation>[] SpanIteratorList { get; set; }
    }

    private static string numberToLetter(int? number) {
      if (number.HasValue == false) { return string.Empty; }

      int dividend = number.Value;
      string retval = string.Empty;
      int modulo;

      while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        string currLetter = Convert.ToChar(65 + modulo).ToString();
        retval = currLetter + retval;
        dividend = (int)((dividend - modulo) / 26);
      }

      return retval;
    }

    #region IComparable<InventoryLocation> Members

    public int CompareTo(InventoryLocation other)
    {
      if ((object)other == null) { throw new ArgumentException("comparison is not a location"); }
      if (this.LocationID == other.LocationID) { return 0; }
      if (this.LocationTypeID != other.LocationTypeID) { return this.LocationTypeID.CompareTo(other.LocationTypeID); }

      if (this.Division1 != other.Division1) { return this.Division1.CompareTo(other.Division1); }
      if (this.Division2.Value != other.Division2.Value) { return this.Division2.Value.CompareTo(other.Division2.Value); }
      if (this.Division3.Value != other.Division3.Value) { return this.Division3.Value.CompareTo(other.Division3.Value); }
      if (this.Division4.Value != other.Division4.Value) { return this.Division4.Value.CompareTo(other.Division4.Value); }
      if (this.Division5.Value != other.Division5.Value) { return this.Division5.Value.CompareTo(other.Division5.Value); }
      throw new ApplicationException("location database is bad!");
    }

    public static bool operator <(InventoryLocation l1, InventoryLocation l2) { return l1.CompareTo(l2) < 0; }
    public static bool operator <=(InventoryLocation l1, InventoryLocation l2) { return l1.CompareTo(l2) <= 0; }
    public static bool operator >(InventoryLocation l1, InventoryLocation l2) { return l1.CompareTo(l2) > 0; }
    public static bool operator >=(InventoryLocation l1, InventoryLocation l2) { return l1.CompareTo(l2) >= 0; }

    #endregion

    public override string ToString()
    {
      return string.Format("{0} {1}", this.LocationTypeName, this.LocationName);
    }
  }
}
