using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yvand.ClaimsProviders;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //var config = AzureCP.GetConfiguration();
            AzureCP cp = new AzureCP("AzureCPSE");
            cp.ValidateLocalConfiguration(null);
            //cp.Configuration.
        }
    }
}
