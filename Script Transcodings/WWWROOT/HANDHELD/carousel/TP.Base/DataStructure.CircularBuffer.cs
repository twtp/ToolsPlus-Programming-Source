using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class CircularBuffer<T> : IEnumerable<T>, IEnumerable {

    private int head;
    private int tail;
    private int size;
    private T[] buffer;
    private bool overwriteAllowed;

    public CircularBuffer(int maxSize) : this(maxSize, false) { }
    public CircularBuffer(int maxSize, bool overwriteAllowed) {
      this.buffer = new T[maxSize];
      this.overwriteAllowed = overwriteAllowed;
      this.head = 0;
      this.tail = 0;
      this.size = 0;
    }

    public bool Add(T item) {
      if (this.IsFull && this.overwriteAllowed == false) {
        throw new OverflowException("overflow");
      }

      this.tail = (this.tail + 1) % this.buffer.Length;
      this.buffer[this.tail] = item;
      if (this.IsFull) {
        // overwriting the first element, so the front moves
        this.head = (this.head + 1) % this.buffer.Length;
      }
      else {
        this.size++;
      }
      return true;
    }

    public T Remove() {
      if (this.IsEmpty) {
        throw new OverflowException("underflow");
      }

      this.head = (this.head + 1) % this.buffer.Length;
      this.size--;
      return this.buffer[this.head];
    }

    public bool Empty() {
      this.buffer = new T[this.buffer.Length];
      this.head = 0;
      this.tail = 0;
      return true;
    }

    public bool IsEmpty { get { return this.size == 0; } }
    public bool IsFull { get { return this.size == this.buffer.Length; } }
    public int Size { get { return this.size; } }

    public T[] Clone() {
      T[] retval = new T[this.Size];
      if (this.Size == 0) {
        return retval;
      }

      // no copyrange function?
      if (this.head < this.tail) {
        for (int i = this.head + 1; i < this.tail + 1; i++) {
          retval[i - this.head - 1] = this.buffer[i];
        }
      }
      else {
        for (int i = this.head + 1; i < this.buffer.Length; i++) {
          retval[i - this.head - 1] = this.buffer[i];
        }
        for (int i = 0; i <= this.tail; i++) {
          retval[i + this.buffer.Length - this.head - 1] = this.buffer[i];
        }
      }

      return retval;
    }

    public IEnumerator<T> GetEnumerator() {
      int dx = this.head;
      do {
        dx = (dx + 1) % this.buffer.Length;
        yield return this.buffer[dx];
      } while (dx != this.tail);
    }

    #region IEnumerable<T> Members
    IEnumerator<T> IEnumerable<T>.GetEnumerator() {
      return GetEnumerator();
    }
    #endregion

    #region IEnumerable Members
    IEnumerator IEnumerable.GetEnumerator() {
      return (IEnumerator)GetEnumerator();
    }
    #endregion
  }
}
