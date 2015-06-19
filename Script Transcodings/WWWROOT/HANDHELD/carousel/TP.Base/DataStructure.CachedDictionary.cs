using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class CachedDictionary<TKey, TValue> : IDictionary<TKey, TValue>, ICollection<KeyValuePair<TKey, TValue>>, IEnumerable<KeyValuePair<TKey, TValue>>, System.Collections.IDictionary, System.Collections.ICollection, System.Collections.IEnumerable {
    
    public delegate TValue ValueGetterDelegate(TKey key);

    public CachedDictionary(ValueGetterDelegate valueGetter) {
      this.ValueGetter = valueGetter;
    }
    public CachedDictionary(ValueGetterDelegate valueGetter, int timeout, bool autoCleanup) {
      this.ValueGetter = valueGetter;
      this.timeout = timeout;
      this.autoCleanup = autoCleanup;
    }

    protected Dictionary<TKey, TValue> _dict = new Dictionary<TKey, TValue>();
    protected Dictionary<TKey, DateTime> _time = new Dictionary<TKey, DateTime>();
    protected object _lockObject = new object();

    protected int timeout = 5; // in minutes
    public int Timeout {
      get {
        return this.timeout;
      }
      set {
        this.timeout = value;
        if (this.autoCleanup) {
          this.autoCleanupTimer.Interval = this.timeout;
        }
      }
    }

    protected bool autoCleanup = false;
    protected System.Timers.Timer autoCleanupTimer = null;
    public bool AutoCleanup {
      get {
        return this.autoCleanup;
      }
      set {
        if (this.autoCleanup != value) {
          this.autoCleanup = value;
          if (this.autoCleanup) {
            this.autoCleanupTimer = new System.Timers.Timer();
            this.autoCleanupTimer.Interval = this.timeout;
            this.autoCleanupTimer.AutoReset = false;
            this.autoCleanupTimer.Elapsed += new System.Timers.ElapsedEventHandler(this.autoCleanupTimer_Elapsed);
            this.autoCleanupTimer.Start();
          }
          else {
            this.autoCleanupTimer.Stop();
            this.autoCleanupTimer.Elapsed -= new System.Timers.ElapsedEventHandler(this.autoCleanupTimer_Elapsed);
            this.autoCleanupTimer = null;
          }
        }
      }
    }

    protected void autoCleanupTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e) {
      Cleanup();
      this.autoCleanupTimer.Start();
    }

    protected ValueGetterDelegate valueGetter = null;
    public ValueGetterDelegate ValueGetter {
      get {
        return this.valueGetter;
      }
      internal set {
        this.valueGetter = value;
      }
    }

    public TValue this[TKey key] {
      get {
        lock (this._lockObject) {
          if (this._dict.ContainsKey(key) == false || this.IsExpired(key)) {
            this._dict.Remove(key);
            this._time.Remove(key);
            this._dict.Add(key, this.valueGetter(key));
            this._time.Add(key, DateTime.Now);
          }
          return this._dict[key];
        }
      }
      set {
        lock (this._lockObject) {
          this._dict.Remove(key);
          this._time.Remove(key);
          this._dict.Add(key, value);
          this._time.Add(key, DateTime.Now);
        }
      }
    }
    public object this[object key] {
      get {
        return this[(TKey)key];
      }
      set {
        this[(TKey)key] = (TValue)value;
      }
    }

    public bool IsExpired(TKey key) {
      return this._time[key].AddMinutes(this.timeout) < DateTime.Now;
    }

    public void Cleanup() {
      lock (this._lockObject) {
        List<TKey> keys = new List<TKey>(this._dict.Keys);
        foreach (TKey key in keys) {
          if (this.IsExpired(key)) {
            this._dict.Remove(key);
            this._time.Remove(key);
          }
        }
      }
    }

    public bool ContainsKey(TKey key) {
      lock (this._lockObject) {
        return this._dict.ContainsKey(key) && !this.IsExpired(key);
      }
    }
    public bool Contains(KeyValuePair<TKey, TValue> item) {
      lock (this._lockObject) {
        return this._dict.ContainsKey(item.Key) && !this.IsExpired(item.Key);
      }
    }
    public bool Contains(object key) {
      lock (this._lockObject) {
        return this._dict.ContainsKey((TKey)key) && !this.IsExpired((TKey)key);
      }
    }

    public int Count {
      get {
        this.Cleanup();
        return this._dict.Count;
      }
    }

    public void Add(TKey key, TValue value) {
      lock (this._lockObject) {
        this[key] = value;
      }
    }
    public void Add(KeyValuePair<TKey, TValue> item) {
      lock (this._lockObject) {
        this[item.Key] = item.Value;
      }
    }
    public void Add(object key, object value) {
      lock (this._lockObject) {
        this[(TKey)key] = (TValue)value;
      }
    }

    public bool Remove(TKey key) {
      lock (this._lockObject) {
        return this._dict.Remove(key) && this._time.Remove(key);
      }
    }
    public bool Remove(KeyValuePair<TKey, TValue> item) {
      lock (this._lockObject) {
        return this._dict.Remove(item.Key) && this._time.Remove(item.Key);
      }
    }
    public void Remove(object key) {
      lock (this._lockObject) {
        this._dict.Remove((TKey)key);
        this._time.Remove((TKey)key);
      }
    }
    public void Invalidate(TKey key) {
      this.Remove(key);
    }

    public void Clear() {
      lock (this._lockObject) {
        this._dict.Clear();
        this._time.Clear();
      }
    }
    public void Invalidate() {
      this.Clear();
    }

    public bool TryGetValue(TKey key, out TValue value) {
      lock (this._lockObject) {
        value = this[key]; // auto-get if not cached or missing
        return true;
      }
    }

    ICollection<TKey> IDictionary<TKey, TValue>.Keys {
      get {
        lock (this._lockObject) {
          this.Cleanup();
          return this._dict.Keys;
        }
      }
    }
    System.Collections.ICollection System.Collections.IDictionary.Keys {
      get {
        lock (this._lockObject) {
          this.Cleanup();
          return this._dict.Keys;
        }
      }
    }

    ICollection<TValue> IDictionary<TKey, TValue>.Values {
      get {
        lock (this._lockObject) {
          this.Cleanup();
          return this._dict.Values;
        }
      }
    }
    System.Collections.ICollection System.Collections.IDictionary.Values {
      get {
        lock (this._lockObject) {
          this.Cleanup();
          return this._dict.Values;
        }
      }
    }

    public bool IsReadOnly {
      get {
        return false;
      }
    }

    public bool IsFixedSize {
      get {
        return false;
      }
    }

    public bool IsSynchronized {
      get {
        return false;
      }
    }

    public object SyncRoot {
      get {
        return this._lockObject;
      }
    }

    IEnumerator<KeyValuePair<TKey, TValue>> IEnumerable<KeyValuePair<TKey, TValue>>.GetEnumerator() {
      lock (this._lockObject) {
        this.Cleanup();
        return this._dict.GetEnumerator();
      }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() {
      lock (this._lockObject) {
        this.Cleanup();
        return this._dict.GetEnumerator();
      }
    }

    System.Collections.IDictionaryEnumerator System.Collections.IDictionary.GetEnumerator() {
      lock (this._lockObject) {
        this.Cleanup();
        return this._dict.GetEnumerator();
      }
    }

    public void CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex) {
      throw new Exception("The method or operation is not implemented.");
    }

    public void CopyTo(Array array, int index) {
      throw new Exception("The method or operation is not implemented.");
    }

  }
}
