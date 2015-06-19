using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Nancy;

using TP.Object;
using TP.Object.BackEnd;

namespace PoInvServer
{
  public class DomainAuthentication : NancyModule
  {
    public DomainAuthentication() : base()
    {
      Post["/authenticate"] = _ =>
      {
        try
        {
          string username = NancyExtensions.GetBasicParameter<string>(Request.Form, "username", null);
          string password = NancyExtensions.GetBasicParameter<string>(Request.Form, "password", null);

          if (string.IsNullOrEmpty(username)) { return NancyExtensions.Respond(new ObjectResponse<Person>("username is required")); }
          if (string.IsNullOrEmpty(password)) { return NancyExtensions.Respond(new ObjectResponse<Person>("password is required")); }

          System.DirectoryServices.DirectoryEntry entry = new System.DirectoryServices.DirectoryEntry("LDAP://TOOLBOX", username, password);
          object nativeObject = entry.NativeObject;

          return NancyExtensions.Respond(new ObjectResponse<Person>(DbPerson.FindOrCreatePerson(username)));
        }
        catch (System.DirectoryServices.DirectoryServicesCOMException)
        {
          return NancyExtensions.Respond(new ObjectResponse<Person>("invalid username or password"));
        }
        catch (Exception ex)
        {
          return NancyExtensions.Respond(new ObjectResponse<Person>(ex.Message));
        }
      };
    }
  }
}
