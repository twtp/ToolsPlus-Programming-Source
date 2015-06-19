using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class Triplet<T, T2, T3> : IEquatable<Triplet<T, T2, T3>> {
    private T first;
    private T2 second;
    private T3 third;

    public Triplet(T first, T2 second, T3 third) {
      this.first = first;
      this.second = second;
      this.third = third;
    }

    public T First {
      get { return this.first; }
      set { this.first = value; }
    }
    public T2 Second {
      get { return this.second; }
      set { this.second = value; }
    }
    public T3 Third {
      get { return this.third; }
      set { this.third = value; }
    }

    public bool Equals(Triplet<T, T2, T3> triplet) {
      if ((object)triplet == null) {
        return false;
      }
      return (this.first.Equals(triplet.first) && this.second.Equals(triplet.second) && this.third.Equals(triplet.third));
    }

    public override bool Equals(object obj) {
      return base.Equals(obj);
    }

    public override int GetHashCode() {
      return base.GetHashCode();
    }

    public static bool operator ==(Triplet<T, T2, T3> x, Triplet<T, T2, T3> y) {
      if ((object)x == null && (object)y == null) {
        return true;
      }
      if ((object)x == null || (object)y == null) {
        return false;
      }
      return x.Equals(y);
    }

    public static bool operator !=(Triplet<T, T2, T3> x, Triplet<T, T2, T3> y) {
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
