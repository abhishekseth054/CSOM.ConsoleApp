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
            try
            {

                //SiteCollectionHelper.ProvisionSiteCollection();

                //UserProfile.UpdateUserprofileCustomAttribute();

                //Console.WriteLine("");
                //Console.WriteLine("*************** List and Library Creation Started *******************");
                //Console.WriteLine("");

                //ListHelper.ValidateAndCreateListAndLibrry();

                //Console.WriteLine("");
                //Console.WriteLine("*************** List and Library Created Successfully *******************");
                //Console.WriteLine("");

                //Console.WriteLine("*************** Content Type Creation Started *******************");
                //Console.WriteLine("");
                //ContentTypeHelper.ValidateAndCreateContentType();

                //Console.WriteLine("");
                //Console.WriteLine("*************** Content Type Created Successfully *******************");
                //Console.WriteLine("");

                //Console.WriteLine("*************** Site Column Creation Started *******************");
                //Console.WriteLine("");

                //SiteColumnHelper.ValidateAndCreateSiteColumn();

                //Console.WriteLine("");
                //Console.WriteLine("*************** Site Column Created Successfully *******************");
                //Console.WriteLine("");

                //Console.WriteLine("*************** Adding Site Column to Content Type Started *******************");
                //Console.WriteLine("");

                //ContentTypeHelper.ValidateAndAddSiteColumnToContentType();

                //Console.WriteLine("");
                //Console.WriteLine("*************** Adding Site Column to Content Type Completed Successfully *******************");
                //Console.WriteLine("");

                //Console.WriteLine("*************** Associating Content Type To List started *******************");
                //Console.WriteLine("");
                //ListHelper.ValidateAndAssociateContenTypeToList();

                //Console.WriteLine("");
                //Console.WriteLine("*************** Associating Content Type To List Completed Successfully *******************");
                //Console.WriteLine("");

                //FixLookupSiteColumn.UpdateLookUpSiteColumn();

                Console.WriteLine("*******============ COMPLETED ===============***********");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\n" + ex.StackTrace);
            }

            Console.ReadKey();
        }
    }
}
