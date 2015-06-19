using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class Pair<T, T2> : IEquatable<Pair<T, T2>> {
    private T first;
    private T2 second;

    public Pair(T first, T2 second) {
      this.first = first;
      this.second = second;
    }

    public T First {
      get { return this.first; }
      set { this.first = value; }
    }
    public T2 Second {
      get { return this.second; }
      set { this.second = value; }
    }

    public bool Equals(Pair<T, T2> pair) {
      if ((object)pair == null) {
        return false;
      }
      return (this.first.Equals(pair.first) && this.second.Equals(pair.second));
    }

    public override bool Equals(object obj) {
      return base.Equals(obj);
    }

    public override int GetHashCode() {
      return base.GetHashCode();
    }

    public static bool operator ==(Pair<T, T2> x, Pair<T, T2> y) {
      if ((object)x == null && (object)y == null) {
        return true;
      }
      if ((object)x == null || (object)y == null) {
        return false;
      }
      return x.Equals(y);
    }

    public static bool operator !=(Pair<T, T2> x, Pair<T, T2> y) {
      if ((object)x == null && (object)y == null) {
        return false;
      }
      if ((object)x == null || (object)y == null) {
        return true;
      }
      return !x.Equals(y);
    }
  }
}
