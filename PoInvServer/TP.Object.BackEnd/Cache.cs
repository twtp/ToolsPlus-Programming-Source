using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace TP.Object.BackEnd
{
  // 1 key         - DbComponent
  // 1 key to list - DbItemComponent
  // 1 key 1 index - DbBarcode
  // 1 key 2 index - DbProductLineSalesRep
  // 2 key         - DbItem
  // 2 key 1 index - DbPerson

  internal class CacheEntry<T> where T : TP.Object.BaseBusinessObject
  {
    public T Value { get; set; }
    public int SecondsToLive { get; set; }
    public DateTime Expires { get; set; }
    public CacheEntry(T value, int secondsToLive)
    {
      this.Value = value;
      this.SecondsToLive = secondsToLive;
      this.Revalidate();
    }
    public bool IsExpired { get { return this.Expires < DateTime.Now; } }
    public void Revalidate() { this.Expires = DateTime.Now.AddSeconds(this.SecondsToLive); }
  }

  internal class CacheEntryDictionary<TKey, TObj> where TObj : TP.Object.BaseBusinessObject
  {
    public Dictionary<TKey, CacheEntry<TObj>> Value { get; set; }
    public int SecondsToLive { get; set; }
    public DateTime Expires { get; set; }
    public CacheEntryDictionary(int secondsToLive)
    {
      this.Value = new Dictionary<TKey,CacheEntry<TObj>>();
      this.SecondsToLive = secondsToLive;
      this.Revalidate();
    }
    public bool IsExpired { get { return this.Expires < DateTime.Now || this.Values.Any(_ => _.IsExpired == true); } }
    public void Revalidate() { this.Expires = DateTime.Now.AddSeconds(this.SecondsToLive); }

    // exposing dictionary methods
    public CacheEntry<TObj> this[TKey index] { get { return this.Value[index]; } }
    public void Add(TKey key, CacheEntry<TObj> value) { this.Value.Add(key, value); return; }
    public void Clear() { this.Value.Clear(); return; }
    public bool ContainsKey(TKey key) { return this.Value.ContainsKey(key); }
    public Dictionary<TKey, CacheEntry<TObj>>.KeyCollection Keys { get { return this.Value.Keys; } }
    public bool Remove(TKey key) { return this.Value.Remove(key); }
    public Dictionary<TKey, CacheEntry<TObj>>.ValueCollection Values { get { return this.Value.Values; } }
  }


  #region Primary Key Only

  internal abstract class CacheBase1Key<TObj, TKey> where TObj : TP.Object.BaseBusinessObject
  {
    private Dictionary<TKey, CacheEntry<TObj>> cacheKeyToObj = new Dictionary<TKey, CacheEntry<TObj>>();
    protected int cacheSecondsToLive = 3600;

    protected bool contains(TKey pk)
    {
      return cacheKeyToObj.ContainsKey(pk) && cacheKeyToObj[pk].IsExpired == false;
    }

    protected bool containsNonNull(TKey pk)
    {
      return cacheKeyToObj.ContainsKey(pk) && cacheKeyToObj[pk].IsExpired == false && cacheKeyToObj[pk] != null;
    }

    protected TObj retrieve(TKey pk)
    {
      return cacheKeyToObj[pk].Value; // note that this assumes .IsExpired==false has been tested!
    }

    protected void put(TKey pk, TObj obj)
    {
      if (cacheKeyToObj.ContainsKey(pk) == false || cacheKeyToObj[pk].IsExpired || object.Equals(cacheKeyToObj[pk], obj) == false)
      {
        cacheKeyToObj.Remove(pk);
        cacheKeyToObj.Add(pk, new CacheEntry<TObj>(obj, this.cacheSecondsToLive));
      }
    }

    protected void clear(TKey pk)
    {
      cacheKeyToObj.Remove(pk);
    }

    protected void wipe()
    {
      cacheKeyToObj.Clear();
    }

    protected IEnumerable<TKey> keys()
    {
      return cacheKeyToObj.Keys.Where(_ => cacheKeyToObj[_].IsExpired == false);
    }

    protected IEnumerable<TObj> values()
    {
      return cacheKeyToObj.Values.Where(_ => _.IsExpired == false).Select(_ => _.Value);
    }

    protected int count()
    {
      return this.keys().Count();
    }

    public string Dump()
    {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      sb.AppendLine("index by primary key:");
      foreach (TKey pk in cacheKeyToObj.Keys)
      {
        sb.AppendLine(string.Format("  {0} = {{ {1} }}", pk, cacheKeyToObj[pk].Value == null ? "null" : cacheKeyToObj[pk].Value.DumpSelf()));
      }
      return sb.ToString();
    }
  }

  internal class AutoCache1Key<TObj, TKey> : CacheBase1Key<TObj, TKey> where TObj : TP.Object.BaseBusinessObject
  {
    public Func<System.Data.DataRow, TObj> Constructor { get; set; } // _ => new DbItemCost(_)
    public Func<DataRow, TKey> Identifier { get; set; }              // _ => (int)_["ID"]
    public Func<TObj, TKey> PrimaryKey { get; set; }                 // _ => _.TupleID
    public int SecondsToLive { get { return base.cacheSecondsToLive; } set { this.cacheSecondsToLive = value; } }

    public TObj FetchByPrimaryKey(TKey tupleID, Func<DataTable> getter)
    {
      if (!base.contains(tupleID))
      {
        DataTable dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), obj);
        }
        else
        {
          base.put(tupleID, null);
        }
      }
      return base.retrieve(tupleID);
    }

    public ICollection<TObj> FetchNonIndexed(Func<DataTable> getter)
    {
      var retval = new List<TObj>();
      DataTable dt = getter();
      foreach (DataRow r in dt.Rows)
      {
        var tupleID = this.Identifier(r);
        var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
        base.put(this.PrimaryKey(obj), obj);
        retval.Add(obj);
      }
      return retval;
    }

    public void Clear(TKey pk)
    {
      base.clear(pk);
    }
  }

  #endregion


  #region Primary Key + Indexed Collection
  
  internal abstract class CacheBase1Key1Index<TObj, TKey, TIndex1> : CacheBase1Key<TObj, TKey> where TObj : TP.Object.BaseBusinessObject
  {
    private Dictionary<TIndex1, CacheEntryDictionary<TKey, TObj>> cacheIndexToKeyToListOfObj = new Dictionary<TIndex1, CacheEntryDictionary<TKey, TObj>>();

    protected bool containsIndex1(TIndex1 dx)
    {
      return cacheIndexToKeyToListOfObj.ContainsKey(dx) && cacheIndexToKeyToListOfObj[dx].IsExpired == false;
    }

    protected IEnumerable<TObj> retrieveIndex1(TIndex1 dx)
    {
      return cacheIndexToKeyToListOfObj[dx].Values.Select(_ => _.Value); // note that this assumes .IsExpired == false has been tested (which tests all subitems too)
    }

    protected void put(TKey pk, bool dxHasValue, TIndex1 dx, TObj obj)
    {
      base.put(pk, obj);
      
      if (dxHasValue == false) { return; }
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx) == false) { return; } // hasn't been instantiated, skip

      if (cacheIndexToKeyToListOfObj[dx].ContainsKey(pk) == false || cacheIndexToKeyToListOfObj[dx].IsExpired || cacheIndexToKeyToListOfObj[dx][pk].IsExpired || object.Equals(cacheIndexToKeyToListOfObj[dx][pk], obj) == false)
      {
        cacheIndexToKeyToListOfObj[dx].Remove(pk);
        cacheIndexToKeyToListOfObj[dx].Add(pk, new CacheEntry<TObj>(obj, this.cacheSecondsToLive));
      }
    }

    protected void startIndex1(TIndex1 dx)
    {
      if (!cacheIndexToKeyToListOfObj.ContainsKey(dx))
      {
        cacheIndexToKeyToListOfObj.Add(dx, new CacheEntryDictionary<TKey, TObj>(this.cacheSecondsToLive));
      }
      else
      {
        cacheIndexToKeyToListOfObj[dx].Clear();
      }
    }

    protected void clear(TKey pk, TIndex1 dx)
    {
      base.clear(pk);
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx))
      {
        cacheIndexToKeyToListOfObj.Remove(dx);
      }
    }

    public new string Dump()
    {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      sb.AppendLine("index 1:");
      foreach (TIndex1 dx in cacheIndexToKeyToListOfObj.Keys)
      {
        if (cacheIndexToKeyToListOfObj[dx].Keys.Count == 0)
        {
          sb.AppendLine(string.Format("  {0}/EMPTY", dx));
        }
        else
        {
          foreach (TKey pk in cacheIndexToKeyToListOfObj[dx].Keys)
          {
            sb.AppendLine(string.Format("  {0}/{1} = {{ {2} }}", dx, pk, cacheIndexToKeyToListOfObj[dx][pk].Value == null ? "null" : cacheIndexToKeyToListOfObj[dx][pk].Value.DumpSelf()));
          }
        }
      }
      return base.Dump() + sb.ToString();
    }
  }

  internal class AutoCache1Key1Index<TObj, TKey, TIndex1> : CacheBase1Key1Index<TObj, TKey, TIndex1> where TObj : TP.Object.BaseBusinessObject where TIndex1 : struct
  {
    public Func<System.Data.DataRow, TObj> Constructor { get; set; } // _ => new DbItemCost(_)
    public Func<System.Data.DataRow, TKey> Identifier { get; set; }  // _ => (int)_["ID"]
    public Func<TObj, TKey> PrimaryKey { get; set; }                 // _ => _.TupleID
    public Func<TObj, TIndex1> Index1 { get; set; }                  // _ => _.ItemID
    public int SecondsToLive { get { return base.cacheSecondsToLive; } set { this.cacheSecondsToLive = value; } }

    public TObj FetchByPrimaryKey(TKey tupleID, Func<System.Data.DataTable> getter)
    {
      if (!base.contains(tupleID))
      {
        var dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), true, this.Index1(obj), obj);
        }
        else
        {
          base.put(tupleID, null);
        }
      }
      return base.retrieve(tupleID);
    }

    public IEnumerable<TObj> FetchByIndex1(TIndex1 index1, Func<System.Data.DataTable> getter)
    {
      if (!base.containsIndex1(index1))
      {
        base.startIndex1(index1);
        System.Data.DataTable dt = getter();
        foreach (System.Data.DataRow r in dt.Rows)
        {
          var tupleID = this.Identifier(r);
          var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
          base.put(this.PrimaryKey(obj), true, this.Index1(obj), obj);
        }
      }
      return base.retrieveIndex1(index1);
    }

    public IEnumerable<TObj> FetchNonIndexed(Func<System.Data.DataTable> getter)
    {
      var retval = new List<TObj>();
      DataTable dt = getter();
      foreach (DataRow r in dt.Rows)
      {
        var tupleID = this.Identifier(r);
        var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
        base.put(this.PrimaryKey(obj), true, this.Index1(obj), obj);
        retval.Add(obj);
      }
      return retval;
    }

    public void Clear(TKey pk, TIndex1 dx)
    {
      base.clear(pk, dx);
    }
  }

  #endregion


  #region Primary Key + Indexed Collection + Indexed Collection

  internal abstract class CacheBase1Key2Index<TObj, TKey, TIndex1, TIndex2> : CacheBase1Key1Index<TObj, TKey, TIndex1> where TObj : TP.Object.BaseBusinessObject
  {
    private Dictionary<TIndex2, CacheEntryDictionary<TKey, TObj>> cacheIndexToKeyToListOfObj = new Dictionary<TIndex2, CacheEntryDictionary<TKey, TObj>>();

    protected bool containsIndex2(TIndex2 dx)
    {
      return cacheIndexToKeyToListOfObj.ContainsKey(dx) && cacheIndexToKeyToListOfObj[dx].IsExpired == false;
    }

    protected IEnumerable<TObj> retrieveIndex2(TIndex2 dx)
    {
      return cacheIndexToKeyToListOfObj[dx].Values.Select(_ => _.Value); // note that this assumes .IsExpired == false has been tested (which tests all subitems too)
    }

    protected void put(TKey pk, bool dx1HasValue, TIndex1 dx1, bool dx2HasValue, TIndex2 dx2, TObj obj)
    {
      base.put(pk, dx1HasValue, dx1, obj);

      if (dx2HasValue == false) { return; }
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx2) == false) { return; } // hasn't been instantiated, skip

      if (cacheIndexToKeyToListOfObj[dx2].ContainsKey(pk) == false || cacheIndexToKeyToListOfObj[dx2].IsExpired || cacheIndexToKeyToListOfObj[dx2][pk].IsExpired || object.Equals(cacheIndexToKeyToListOfObj[dx2][pk], obj) == false)
      {
        cacheIndexToKeyToListOfObj[dx2].Remove(pk);
        cacheIndexToKeyToListOfObj[dx2].Add(pk, new CacheEntry<TObj>(obj, this.cacheSecondsToLive));
      }
    }

    protected void startIndex2(TIndex2 dx2)
    {
      if (!cacheIndexToKeyToListOfObj.ContainsKey(dx2))
      {
        cacheIndexToKeyToListOfObj.Add(dx2, new CacheEntryDictionary<TKey,TObj>(this.cacheSecondsToLive));
      }
      else
      {
        cacheIndexToKeyToListOfObj[dx2].Clear();
      }
    }

    protected void clear(TKey pk, TIndex1 dx1, TIndex2 dx2)
    {
      base.clear(pk, dx1);
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx2))
      {
        cacheIndexToKeyToListOfObj.Remove(dx2);
      }
    }

    public new string Dump()
    {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      sb.AppendLine("index 2:");
      foreach (TIndex2 dx in cacheIndexToKeyToListOfObj.Keys)
      {
        if (cacheIndexToKeyToListOfObj[dx].Keys.Count == 0)
        {
          sb.AppendLine(string.Format("  {0}/EMPTY", dx));
        }
        else
        {
          foreach (TKey pk in cacheIndexToKeyToListOfObj[dx].Keys)
          {
            sb.AppendLine(string.Format("  {0}/{1} = {{ {2} }}", dx, pk, cacheIndexToKeyToListOfObj[dx][pk].Value == null ? "null" : cacheIndexToKeyToListOfObj[dx][pk].Value.DumpSelf()));
          }
        }
      }
      return base.Dump() + sb.ToString();
    }
  }

  internal class AutoCache1Key2Index<TObj, TKey, TIndex1, TIndex2> : CacheBase1Key2Index<TObj, TKey, TIndex1, TIndex2> where TObj : TP.Object.BaseBusinessObject /*where TIndex1 : struct where TIndex2 : struct*/
  {
    public Func<System.Data.DataRow, TObj> Constructor { get; set; } // _ => new DbItemCost(_)
    public Func<System.Data.DataRow, TKey> Identifier { get; set; }  // _ => (int)_["ID"]
    public Func<TObj, TKey> PrimaryKey { get; set; }                 // _ => _.TupleID
    public Func<TObj, TIndex1> Index1 { get; set; }                  // _ => _.ItemID
    public Func<TObj, TIndex2> Index2 { get; set; }                  // _ => _.ItemID
    public int SecondsToLive { get { return base.cacheSecondsToLive; } set { this.cacheSecondsToLive = value; } }

    public TObj FetchByPrimaryKey(TKey tupleID, Func<System.Data.DataTable> getter)
    {
      if (!base.contains(tupleID))
      {
        var dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), true, this.Index1(obj), true, this.Index2(obj), obj);
        }
        else
        {
          base.put(tupleID, null);
        }
      }
      return base.retrieve(tupleID);
    }

    public IEnumerable<TObj> FetchByIndex1(TIndex1 index1, Func<System.Data.DataTable> getter)
    {
      if (!base.containsIndex1(index1))
      {
        base.startIndex1(index1);
        System.Data.DataTable dt = getter();
        foreach (System.Data.DataRow r in dt.Rows)
        {
          var tupleID = this.Identifier(r);
          var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
          base.put(this.PrimaryKey(obj), true, this.Index1(obj), true, this.Index2(obj), obj);
        }
      }
      return base.retrieveIndex1(index1);
    }

    public IEnumerable<TObj> FetchByIndex2(TIndex2 index2, Func<System.Data.DataTable> getter)
    {
      if (!base.containsIndex2(index2))
      {
        base.startIndex2(index2);
        System.Data.DataTable dt = getter();
        foreach (System.Data.DataRow r in dt.Rows)
        {
          var tupleID = this.Identifier(r);
          var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
          base.put(this.PrimaryKey(obj), true, this.Index1(obj), true, this.Index2(obj), obj);
        }
      }
      return base.retrieveIndex2(index2);
    }

    public ICollection<TObj> FetchNonIndexed(Func<DataTable> getter)
    {
      var retval = new List<TObj>();
      DataTable dt = getter();
      foreach (DataRow r in dt.Rows)
      {
        var tupleID = this.Identifier(r);
        var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
        base.put(this.PrimaryKey(obj), true, this.Index1(obj), true, this.Index2(obj), obj);
        retval.Add(obj);
      }
      return retval;
    }

    public void Clear(TKey pk, TIndex1 dx1, TIndex2 dx2)
    {
      base.clear(pk, dx1, dx2);
    }
  }

  #endregion


  #region Primary Key + Alternate Key (possibly null)

  internal abstract class CacheBase2Key<TObj, TKey, TAlternate> : CacheBase1Key<TObj, TKey> where TObj : TP.Object.BaseBusinessObject
  {
    private Dictionary<TAlternate, CacheEntry<TObj>> cacheAlternateToObj = new Dictionary<TAlternate, CacheEntry<TObj>>();

    protected bool containsAlternate(TAlternate altPK)
    {
      return cacheAlternateToObj.ContainsKey(altPK) && cacheAlternateToObj[altPK].IsExpired == false;
    }

    protected bool containsNonNull(TAlternate altPK)
    {
      return cacheAlternateToObj.ContainsKey(altPK) && cacheAlternateToObj[altPK].IsExpired == false && cacheAlternateToObj[altPK] != null;
    }

    protected TObj retrieveAlternate(TAlternate altPK)
    {
      return cacheAlternateToObj[altPK].Value; // note that this assumes .IsExpired==false has been tested!
    }

    protected void put(TKey pk, TAlternate altPK, TObj obj)
    {
      base.put(pk, obj);

      if (altPK == null) { return; } // needed for purchase order numbers
      if (cacheAlternateToObj.ContainsKey(altPK) == false || cacheAlternateToObj[altPK].IsExpired || object.Equals(cacheAlternateToObj[altPK], obj) == false)
      {
        cacheAlternateToObj.Remove(altPK);
        cacheAlternateToObj.Add(altPK, new CacheEntry<TObj>(obj, this.cacheSecondsToLive));
      }
    }

    protected void clear(TKey pk, TAlternate altPK)
    {
      base.clear(pk);
      if (altPK != null)
      {
        cacheAlternateToObj.Remove(altPK); // needed for purchase order numbers
      }
    }

    protected void putAlternateMiss(TAlternate altPK)
    {
      cacheAlternateToObj.Remove(altPK);
      cacheAlternateToObj.Add(altPK, new CacheEntry<TObj>(null, this.cacheSecondsToLive));
    }

    public new string Dump()
    {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      sb.AppendLine("alternate 1:");
      foreach (TAlternate altPK in cacheAlternateToObj.Keys)
      {
        sb.AppendLine(string.Format("  {0} = {{ {1} }}", altPK, cacheAlternateToObj[altPK].Value == null ? "null" : cacheAlternateToObj[altPK].Value.DumpSelf()));
      }
      return base.Dump() + sb.ToString();
    }
  }

  internal class AutoCache2Key<TObj, TKey, TAlternate> : CacheBase2Key<TObj, TKey, TAlternate> where TObj : TP.Object.BaseBusinessObject
  {
    public Func<DataRow, TObj> Constructor { get; set; }     // _ => new DbItemCost(_)
    public Func<DataRow, TKey> Identifier { get; set; }      // _ => (int)_["ID"]
    public Func<TObj, TKey> PrimaryKey { get; set; }         // _ => _.TupleID
    public Func<TObj, TAlternate> AlternateKey { get; set; } // _ => _.PurchaseOrderNo
    public int SecondsToLive { get { return base.cacheSecondsToLive; } set { this.cacheSecondsToLive = value; } }

    public TObj FetchByPrimaryKey(TKey pk, Func<DataTable> getter)
    {
      if (!base.contains(pk))
      {
        DataTable dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), this.AlternateKey(obj), obj);
        }
        else
        {
          base.put(pk, null);
        }
      }
      return base.retrieve(pk);
    }

    public TObj FetchByAlternateKey(TAlternate altKey, Func<DataTable> getter)
    {
      if (!base.containsAlternate(altKey))
      {
        DataTable dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), this.AlternateKey(obj), obj);
        }
        else
        {
          base.putAlternateMiss(altKey);
        }
      }
      return base.retrieveAlternate(altKey);
    }

    public ICollection<TObj> FetchNonIndexed(Func<DataTable> getter)
    {
      var retval = new List<TObj>();
      DataTable dt = getter();
      foreach (DataRow r in dt.Rows)
      {
        var tupleID = this.Identifier(r);
        var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
        base.put(this.PrimaryKey(obj), this.AlternateKey(obj), obj);
        retval.Add(obj);
      }
      return retval;
    }

    public void Clear(TKey pk, TAlternate altKey)
    {
      base.clear(pk, altKey);
    }
  }

  #endregion



  #region Primary Key + Alternate Key (possibly null) + Indexed Collection

  internal abstract class CacheBase2Key1Index<TObj, TKey, TAlternate, TIndex1> : CacheBase2Key<TObj, TKey, TAlternate> where TObj : TP.Object.BaseBusinessObject /*where TIndex1 : struct*/
  {
    private Dictionary<TIndex1, CacheEntryDictionary<TKey, TObj>> cacheIndexToKeyToListOfObj = new Dictionary<TIndex1, CacheEntryDictionary<TKey, TObj>>();

    protected bool containsIndex1(TIndex1 dx)
    {
      return cacheIndexToKeyToListOfObj.ContainsKey(dx) && cacheIndexToKeyToListOfObj[dx].IsExpired == false;
    }

    protected IEnumerable<TObj> retrieveIndex1(TIndex1 dx)
    {
      return cacheIndexToKeyToListOfObj[dx].Values.Select(_ => _.Value); // note that this assumes .IsExpired == false has been tested (which tests all subitems too)
    }

    protected void put (TKey pk, TAlternate altPK, bool dx1HasValue, TIndex1 dx, TObj obj)
    {
      base.put(pk, altPK, obj);

      if (dx1HasValue == false) { return; }
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx) == false) { return; } // hasn't been instantiated, skip

      if (cacheIndexToKeyToListOfObj[dx].ContainsKey(pk) == false || cacheIndexToKeyToListOfObj[dx].IsExpired || cacheIndexToKeyToListOfObj[dx][pk].IsExpired || object.Equals(cacheIndexToKeyToListOfObj[dx][pk], obj) == false)
      {
        cacheIndexToKeyToListOfObj[dx].Remove(pk);
        cacheIndexToKeyToListOfObj[dx].Add(pk, new CacheEntry<TObj>(obj, this.cacheSecondsToLive));
      }
    }

    protected void startIndex1(TIndex1 dx)
    {
      if (!cacheIndexToKeyToListOfObj.ContainsKey(dx))
      {
        cacheIndexToKeyToListOfObj.Add(dx, new CacheEntryDictionary<TKey, TObj>(this.cacheSecondsToLive));
      }
      else
      {
        cacheIndexToKeyToListOfObj[dx].Clear();
      }
    }

    protected void clear(TKey pk, TAlternate altPK, TIndex1 dx)
    {
      base.clear(pk, altPK);
      if (cacheIndexToKeyToListOfObj.ContainsKey(dx))
      {
        cacheIndexToKeyToListOfObj.Remove(dx);
      }
    }

    public new string Dump()
    {
      System.Text.StringBuilder sb = new System.Text.StringBuilder();
      sb.AppendLine("index 1:");
      foreach (TIndex1 dx in cacheIndexToKeyToListOfObj.Keys)
      {
        if (cacheIndexToKeyToListOfObj[dx].Keys.Count == 0)
        {
          sb.AppendLine(string.Format("  {0}/EMPTY", dx));
        }
        else
        {
          foreach (TKey pk in cacheIndexToKeyToListOfObj[dx].Keys)
          {
            sb.AppendLine(string.Format("  {0}/{1} = {{ {2} }}", dx, pk, cacheIndexToKeyToListOfObj[dx][pk].Value == null ? "null" : cacheIndexToKeyToListOfObj[dx][pk].Value.DumpSelf()));
          }
        }
      }
      return base.Dump() + sb.ToString();
    }
  }

  internal class AutoCache2Key1Index<TObj, TKey, TAlternate, TIndex1> : CacheBase2Key1Index<TObj, TKey, TAlternate, TIndex1> where TObj : TP.Object.BaseBusinessObject /*where TIndex1 : struct*/
  {
    public Func<DataRow, TObj> Constructor { get; set; }     // _ => new DbItemCost(_)
    public Func<DataRow, TKey> Identifier { get; set; }      // _ => (int)_["ID"]
    public Func<TObj, TKey> PrimaryKey { get; set; }         // _ => _.TupleID
    public Func<TObj, TAlternate> AlternateKey { get; set; } // _ => _.PurchaseOrderNo
    public Func<TObj, TIndex1> Index1 { get; set; }          // _ => _.ItemID
    public int SecondsToLive { get { return base.cacheSecondsToLive; } set { this.cacheSecondsToLive = value; } }

    public TObj FetchByPrimaryKey(TKey pk, Func<DataTable> getter)
    {
      if (!base.contains(pk))
      {
        DataTable dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), this.AlternateKey(obj), true, this.Index1(obj), obj);
        }
        else
        {
          base.put(pk, null);
        }
      }
      return base.retrieve(pk);
    }

    public TObj FetchByAlternateKey(TAlternate altKey, Func<DataTable> getter)
    {
      if (!base.containsAlternate(altKey))
      {
        DataTable dt = getter();
        if (dt.Rows.Count == 1)
        {
          var obj = this.Constructor(dt.Rows[0]);
          base.put(this.PrimaryKey(obj), this.AlternateKey(obj), true, this.Index1(obj), obj);
        }
        else
        {
          base.putAlternateMiss(altKey);
        }
      }
      return base.retrieveAlternate(altKey);
    }

    public IEnumerable<TObj> FetchIndex1(TIndex1 dx1, Func<DataTable> getter)
    {
      if (!base.containsIndex1(dx1))
      {
        base.startIndex1(dx1);
        DataTable dt = getter();
        foreach (DataRow r in dt.Rows)
        {
          var tupleID = this.Identifier(r);
          var obj = base.containsNonNull(tupleID) ? base.retrieve(tupleID) : this.Constructor(r);
          base.put(this.PrimaryKey(obj), this.AlternateKey(obj), true, this.Index1(obj), obj);
        }
      }
      return base.retrieveIndex1(dx1);
    }

    public void Clear(TKey pk, TAlternate altKey, TIndex1 dx1)
    {
      base.clear(pk, altKey, dx1);
    }
  }

  #endregion

}
