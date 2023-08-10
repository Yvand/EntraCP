using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yvand.ClaimsProviders;
using static Microsoft.SharePoint.Workflow.SPWorkflowAssociationCollection;
using Yvand.ClaimsProviders.Config;
using Microsoft.SharePoint.Administration;

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

            //SPFarm parent = SPFarm.Local;
            //object configuration = (object) parent.GetObject("AzureCPSEConfig", parent.Id, typeof(AADConf<IAADSettings>));
            //configuration = parent.GetObject(new Guid("4ea86a04-7030-4853-bf97-f778de32a274"));

            //Console.WriteLine("end");
        }
    }
}
