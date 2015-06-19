using System;
using System.Collections.Generic;
using System.Text;

namespace MAS200
{
    public class CommonCalls
    {


        /// <summary>
        /// Gets either commandline args, query args, or both and puts to a list.
        /// </summary>
        /// <param name="useCmdArgs"></param>
        /// <param name="useQryArgs"></param>
        /// <param name="cmdArgs"></param>
        /// <param name="qryArgs"></param>
        /// <returns></returns>
        public static List<string> GetRuntimeArguments(bool useCmdArgs, bool useQryArgs, string[] cmdArgs, string qryArgs)
        {
            List<string> Args = new List<string>();
            try
            {
                if (useCmdArgs == true)
                {
                    foreach (string ca in cmdArgs)
                    {
                        Args.Add(ca);
                    }

                }
            }
            catch
            {
                Args.Add("Error Collecting Command-line Arguments");
            }
            try
            {
                if (useQryArgs == true)
                {
                    foreach (string qa in qryArgs.Split('&'))
                    {
                        Args.Add(qa);
                    }
                }
            }
            catch
            {
                Args.Add("Error Collecting Query Arguments");
            }

            return Args;

        }




    }
}
