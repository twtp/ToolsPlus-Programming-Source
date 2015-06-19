using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Object.BackEnd
{
  public interface IWarehouseTask
  {
    int TransactionTypeID { get; }
    int TransactionID { get; }
    string TransactionReference { get; }
  }

  public interface IWarehouseTaskLine
  {
    int HeaderID { get; }
    int LineID { get; }
    int ComponentID { get; }

    int QuantityPicked { get; }

    IWarehouseTask Header(System.Data.SqlClient.SqlTransaction txn);
    IWarehouseTaskLine UpdateQuantity(System.Data.SqlClient.SqlTransaction txn, int fromLocationID, int? toLocationID, int deltaQuantity);
    bool HasChangedFromDbCurrent(System.Data.SqlClient.SqlTransaction txn);
  }

}
