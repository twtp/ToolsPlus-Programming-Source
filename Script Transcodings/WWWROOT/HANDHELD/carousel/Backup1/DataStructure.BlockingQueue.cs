using System;
using System.Collections.Generic;
using System.Text;

namespace TP.Base.DataStructure {
  public class BlockingQueue<T> {
    private Queue<T> queue = new Queue<T>();
    private bool closing = false;

    public void Enqueue(T item) {
      lock (queue) {
        queue.Enqueue(item);
        if (queue.Count == 1) {
          System.Threading.Monitor.PulseAll(queue);
        }
      }
    }

    public bool TryDequeue(out T value) {
      lock (queue) {
        while (queue.Count == 0) {
          if (closing) {
            value = default(T);
            return false;
          }
          else {
            //TP.Base.Logger.Log("BlockingQueue waiting for queue signal");
            System.Threading.Monitor.Wait(queue);
            //TP.Base.Logger.Log("BlockingQueue wait complete");
          }
        }
        value = queue.Dequeue();
        return true;
      }
    }

    public T Dequeue() {
      lock (queue) {
        while (queue.Count == 0) {
          //TP.Base.Logger.Log("BlockingQueue waiting for queue signal");
          System.Threading.Monitor.Wait(queue);
          //TP.Base.Logger.Log("BlockingQueue wait complete");
        }
        return queue.Dequeue();
      }
    }

    public void Close() {
      lock (queue) {
        closing = true;
        System.Threading.Monitor.PulseAll(queue);
      }
    }

    public int Count {
      get { return queue.Count; }
    }
  }
}
