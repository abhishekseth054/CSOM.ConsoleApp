using SPO.ClientManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSOM.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SiteCollectionHelper sc = new SiteCollectionHelper();
            sc.ProvisionSiteCollection();

            Console.WriteLine("Test For Github");
            Console.ReadKey();
        }
    }
}
