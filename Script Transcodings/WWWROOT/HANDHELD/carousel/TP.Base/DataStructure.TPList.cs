using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  // this is basically a List<T> facade, but with a bunch of hidden
  // members that will make it read-only. the subclass can turn these
  // internal or public if required ("internal new void Add(" or
  // something), but that should be the only boilerplate stuff to do.
  public class TPList<T> : IEnumerable<T>, System.Collections.IEnumerable {

    protected List<T> _list = new List<T>();

    public T this[int index] {
      get {
        return this._list[index];
      }
      protected set {
        this._list[index] = value;
      }
    }

    protected void Add(T item) {
      this._list.Add(item);
    }

    protected void AddRange(IEnumerable<T> collection) {
      this._list.AddRange(collection);
    }

    protected void Clear() {
      this._list.Clear();
    }

    public bool Contains(T item) {
      return this._list.Contains(item);
    }

    public void CopyTo(T[] array, int arrayIndex) {
      this._list.CopyTo(array, arrayIndex);
    }

    public int Count {
      get { return this._list.Count; }
    }

    public T Find(Predicate<T> match) {
      return this._list.Find(match);
    }

    public int IndexOf(T item) {
      return this._list.IndexOf(item);
    }

    protected void Insert(int index, T item) {
      this._list.Insert(index, item);
    }

    protected bool Remove(T item) {
      return this._list.Remove(item);
    }

    protected void RemoveAt(int index) {
      this._list.RemoveAt(index);
    }

    public void Reverse() {
      this._list.Reverse();
    }
    public void Reverse(int index, int count) {
      this._list.Reverse(index, Count);
    }

    public void Sort() {
      this._list.Sort();
    }
    public void Sort(Comparison<T> comparison) {
      this._list.Sort(comparison);
    }
    public void Sort(IComparer<T> comparer) {
      this._list.Sort(comparer);
    }
    public void Sort(int index, int count, IComparer<T> comparer) {
      this._list.Sort(index, count, comparer);
    }

    public T[] ToArray() {
      return this._list.ToArray();
    }

    protected bool Replace(T oldItem, T newItem) {
      int dx = this.IndexOf(oldItem);
      if (dx == -1) {
        return false;
      }
      this.Insert(dx, newItem);
      this.Remove(oldItem);
      return true;
    }

    // called like:
    //   InventoryTransactionList list = InventoryTransactionList.FromEnumeration<InventoryTransactionList>(someDictionary.Values);
    public static TList FromEnumeration<TList>(IEnumerable<T> collection) where TList : TPList<T>, new() {
      TList retval = new TList();
      retval.AddRange(collection);
      return retval;
    }

    // called like:
    //   DropshipPurchaseOrderList list = DropshipPurchaseOrderList.FromEnumeration<DropshipPurchaseOrderList, ListViewItem>(this.OrdersListView.CheckedItems, delegate(ListViewItem lvi) { return (DropshipPurchaseOrder)lvi.Tag; });
    public delegate T Mutator<TOther>(TOther other);
    public static TList FromEnumeration<TList, TOther>(IEnumerable<TOther> collection, Mutator<TOther> mutator) where TList : TPList<T>, new() {
      TList retval = new TList();
      foreach (TOther other in collection) {
        retval.Add(mutator(other));
      }
      return retval;
    }
    public static TList FromEnumeration<TList, TOther>(System.Collections.IEnumerable collection, Mutator<TOther> mutator) where TList : TPList<T>, new() {
      TList retval = new TList();
      foreach (TOther other in collection) {
        retval.Add(mutator(other));
      }
      return retval;
    }

    // called like:
    //   IEnumerable<string> test = allSOs.Map<string>(delegate(SalesOrder so) { return so.SalesOrderNumber; });
    //   SalesOrderList test = allSOs.MapTP<SalesOrder>(delegate(SalesOrder so) { return --some alteration of so--; });
    public IEnumerable<TResult> Map<TResult>(Converter<T, TResult> converter) {
      return this._list.ConvertAll(converter);
    }
    public TResultList Map<TResultList, TResult>(Converter<T, TResult> converter) where TResultList : ICollection<TResult>, new() {
      TResultList retval = new TResultList();
      foreach (T t in this._list) {
        retval.Add(converter(t));
      }
      return retval;
    }
    public TResultList MapTP<TResultList, TResult>(Converter<T, TResult> converter) where TResultList : TPList<TResult>, new() {
      TResultList retval = new TResultList();
      foreach (T t in this._list) {
        retval.Add(converter(t));
      }
      return retval;
    }

    // called like:
    //   IEnumerable<SalesOrder> test = allSOs.Grep(delegate(SalesOrder so) { return so.IsActiveForPicking; });
    public TList Grep<TList>(Predicate<T> match) where TList : TPList<T>, new() {
      TList retval = new TList();
      foreach (T t in this._list) {
        if (match(t)) {
          retval.Add(t);
        }
      }
      return retval;
    }

    // called like
    //   decimal balance = this.Accumulate<decimal>(0m, delegate(ref decimal accum, OpenInvoice oi) { accum += oi.Balance; });
    public delegate void Accumulator<TResult>(ref TResult accum, T inc);
    public TResult Accumulate<TResult>(TResult start, Accumulator<TResult> accumulator) {
      TResult retval = start;
      foreach (T t in this) {
        accumulator(ref retval, t);
      }
      return retval;
    }





    #region IEnumerable<T> Members

    public IEnumerator<T> GetEnumerator() {
      return this._list.GetEnumerator();
    }

    #endregion

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      return this._list.GetEnumerator();
    }

    #endregion
  }
}
