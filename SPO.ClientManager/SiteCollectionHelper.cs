using System;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.Threading;

namespace SPO.ClientManager
{
    public class SiteCollectionHelper
    {
        public static void ProvisionSiteCollection()
        {
            try
            {
                //Check if Site Collection is exist
                var siteCtx = AuthHelper.GetClientContext();
                bool isSiteExist = false;

                // This will not check if Site Collection is deleted and available in Recycle bin
                if (siteCtx.WebExistsFullUrl(AuthHelper.siteUrl))
                {
                    isSiteExist = true;
                }

                //Create Context of SPO Admin
                var tenantCtx = AuthHelper.GetTenantContext();
                var tenant = new Tenant(tenantCtx);
                SpoOperation spo = null;

                if (isSiteExist)
                {
                    DeleteSiteCollectionAndRecycled(spo, tenantCtx, tenant);
                    RemoveSiteFromRecycleBin(spo, tenantCtx, tenant);
                }
                else
                {
                    CreateSiteCollection(spo, tenantCtx, tenant);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void DeleteSiteCollectionAndRecycled(SpoOperation spo, ClientContext tenantCtx, Tenant tenant)
        {
            // Remove SiteCollection, it will go to Recycle bin
            spo = tenant.RemoveSite(AuthHelper.siteUrl);
            tenantCtx.Load(tenant);

            //Get the IsComplete property to check if the Site Collection is been removed.
            tenantCtx.Load(spo, i => i.IsComplete);
            tenantCtx.ExecuteQuery();
            var msg = "Site Collection Recycling process...";
            WaitForOperation(tenantCtx, spo, msg);
        }

        private void CreateSiteCollection(SpoOperation spo, ClientContext tenantCtx, Tenant tenant)
        {
            var siteCreationProperties = new SiteCreationProperties();
            siteCreationProperties.Url = AuthHelper.siteUrl;
            siteCreationProperties.Title = "Site Created from Code";
            siteCreationProperties.Owner = AuthHelper.userName;

            //Assign Team Site Template of the SiteCollection.
            siteCreationProperties.Template = "STS#0";

            //Storage Limit in MB
            siteCreationProperties.StorageMaximumLevel = 100;

            //UserCode Resource Points Allowed
            siteCreationProperties.UserCodeMaximumLevel = 100;

            //Create the SiteCollection
            spo = tenant.CreateSite(siteCreationProperties);
            tenantCtx.Load(tenant);
            //Get the IsComplete property to check if the Site Collection creation is complete.
            tenantCtx.Load(spo, i => i.IsComplete);
            tenantCtx.ExecuteQuery();
            var msg = "Site Collection creation process...";
            WaitForOperation(tenantCtx, spo, msg);
        }

        private void RemoveSiteFromRecycleBin(SpoOperation spo, ClientContext tenantCtx, Tenant tenant)
        {
            //Removing Site Collection from Recycle bin
            spo = tenant.RemoveDeletedSite(AuthHelper.siteUrl);
            tenantCtx.Load(spo, i => i.IsComplete);
            tenantCtx.ExecuteQuery();
            var msg = "Removing Site Collection from Recycle bin process...";
            WaitForOperation(tenantCtx, spo, msg);
        }

        private void WaitForOperation(ClientContext tenantCtx, SpoOperation spo, string msg)
        {
            Console.WriteLine($"{msg} status: {"Waiting"}");

            while (!spo.IsComplete)
            {
                //Wait for 15 seconds and then try again
                Thread.Sleep(15000);
                //tenantCtx.Load(spo);
                spo.RefreshLoad();
                tenantCtx.ExecuteQuery();
            }
            Console.WriteLine($"{msg} status: {"Completed"}");


        }
    }
}
