using System;
using Nancy.Hosting.Self;

namespace PoInvServer
{
  public class Program
  {
    private static readonly NancyHost host = new NancyHost(new Uri("http://localhost:1234"));

    public static void Main(string[] args)
    {
      
      /*var l1 = TP.Object.BackEnd.DbInventoryLocation.GetByID(504); // 1 F 4 A
      var l2 = TP.Object.BackEnd.DbInventoryLocation.GetByID(506); // 1 F 4 C
      var l3 = TP.Object.BackEnd.DbInventoryLocation.GetByID(507); // 1 F 4 D
      System.Diagnostics.Debug.WriteLine(l1.CanSpanTo(l2));
      System.Diagnostics.Debug.WriteLine(l1.CanSpanTo(l3));*/

      /*var v1 = TP.Object.BackEnd.DbVendor.GetByID("BLAC01");
      var v2 = TP.Object.BackEnd.DbVendor.GetByID("JET");
      var v1_1 = TP.Object.BackEnd.DbVendor.GetByID("BLAC01");
      System.Threading.Thread.Sleep(60000);
      v1_1 = TP.Object.BackEnd.DbVendor.GetByID("BLAC01");*/

      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(504, 504, 42224, 1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 A    -- create -- no span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(505, 506, 42224, 1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 B-C  -- create -- span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(512, 512, 42224, 1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 I    -- exists -- no span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(507, 508, 42224, 1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 D-E  -- exists -- span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(504, 504, 42224, -1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 A    -- delete  -- no span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(505, 506, 42224, -1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 B-C  -- partial -- span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(512, 512, 42224, -1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 I    -- delete  -- no span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(507, 508, 42224, -1, -1, -1, -1, null, "UNKWN", "FOOBAR", -1, -1); // 1 F 4 D-E  -- partial -- span
      //TP.Object.BackEnd.DbInventoryLocationContent.AddOrRemove(10622, 10622, 42224, 1, null, null, 2, -1, -1);

      /*TP.Object.BackEnd.DbInventoryLocationContent.MoveFromTo(
          10830, 10830,
          5499, 5499, // fail, span invalid
          38705,
          10,
          null,
          null,
          "briandonorfio",
          "dev-03",
          null
        );
      TP.Object.BackEnd.DbInventoryLocationContent.MoveFromTo(
          10830, 10830,
          5500, 5500, // fail, span invalid
          38705,
          10,
          null,
          null,
          "briandonorfio",
          "dev-03",
          null
        );*/
      //TP.Object.BackEnd.DbInventoryLocation.GetSpanEndpoint(5501);
      //TP.Object.BackEnd.DbInventoryLocation.GetSpanEndpoint(5500);
      //TP.Object.BackEnd.DbInventoryLocation.GetSpanEndpoint(5509);
      //TP.Object.BackEnd.DbInventoryLocation.GetSpanEndpoint(5547);

      /*var loc_a = TP.Object.BackEnd.DbInventoryLocation.GetByID(5499);
      var loc_b = TP.Object.BackEnd.DbInventoryLocation.GetByID(5500);
      var loc_c = TP.Object.BackEnd.DbInventoryLocation.GetByID(5501);
      System.Diagnostics.Debug.WriteLine(loc_a.GetCurrentSpanMembers(loc_a, null, null, () => TP.Object.BackEnd.DbInventoryLocationContent.GetByMasterID(loc_a.LocationID)).Length);
      System.Diagnostics.Debug.WriteLine(loc_a.GetCurrentSpanMembers(loc_b, null, null, () => TP.Object.BackEnd.DbInventoryLocationContent.GetByMasterID(loc_a.LocationID)).Length);
      System.Diagnostics.Debug.WriteLine(loc_a.GetCurrentSpanMembers(loc_c, null, null, () => TP.Object.BackEnd.DbInventoryLocationContent.GetByMasterID(loc_a.LocationID)).Length);*/



      host.Start();
      //TP.Base.Logger.Log("server listening on http://localhost:1234");
      
      Console.ReadLine();
    }

  }
}
