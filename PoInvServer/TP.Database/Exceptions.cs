using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TP.Database
{
  public class MasNotInstalledException : Exception
  {
    public MasNotInstalledException(Exception ex) : base("MAS200 is not installed on this computer!", ex)
    {
      //nothing
    }
  }
}
