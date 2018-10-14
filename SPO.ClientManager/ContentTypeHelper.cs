using SPO.ClientManager.Model;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager
{
    public class ContentTypeHelper
    {
        public static void ValidateAndCreateContentType()
        {
            var clientContext = AuthHelper.GetClientContext();
            Web oWeb = clientContext.Web;

            var contentTypeDetails = new Data().GetListInfo().Values.ToList();

            CreateContentType(clientContext, oWeb, contentTypeDetails,"List");

            var contentTypeDetailsForLib = new Data().GetLibInfo().Values.ToList();
            CreateContentType(clientContext, oWeb, contentTypeDetailsForLib, "Library");

            //contentTypeDetailsForLib.ForEach(x =>
            //{
            //    contentTypeDetails.Add(x);
            //});

        }

        private static void CreateContentType(ClientContext clientContext, Web oWeb, List<Dictionary<string, List<string>>> contentTypeDetails, string baseContentType)
        {
            for (int i = 0; i < contentTypeDetails.Count(); i++)
            {
                foreach (string contentTypeName in contentTypeDetails[i].Keys)
                {
                    int count = Helper.IsExist_Helper(clientContext, contentTypeName, "contenttype");

                    if (count > 0)
                    {
                        string contentTypeId = GetContentTypeIdByName(clientContext, contentTypeName);

                        DeleteContentType(clientContext, contentTypeId, contentTypeName);

                        CreateContentTypeAsperParentBaseType(clientContext, contentTypeName, baseContentType);
                    }
                    else
                    {
                        CreateContentTypeAsperParentBaseType(clientContext, contentTypeName, baseContentType);
                    }
                }
            }
        }

        private static void CreateContentTypeAsperParentBaseType(ClientContext clientContext, string contentTypeName, string baseContentType)
        {
            Web oWeb = clientContext.Web;

            if (baseContentType == "Library")
            {
                ContentType ct = oWeb.ContentTypes.GetById("0x0101");
                oWeb.ContentTypes.Add(new ContentTypeCreationInformation() { Name = contentTypeName, Group = "Phoenix", ParentContentType=ct });
            }
            else
            {
                ContentType ct = oWeb.ContentTypes.GetById("0x01");
                oWeb.ContentTypes.Add(new ContentTypeCreationInformation() { Name = contentTypeName, Group = "Phoenix", ParentContentType = ct });
            }

            clientContext.ExecuteQuery();

            Console.WriteLine(contentTypeName + ": Created");
        }

        public static void DeleteContentType(ClientContext clientContext, string contentTypeId, string contentTypeName)
        {
            ContentType ct = clientContext.Web.ContentTypes.GetById(contentTypeId);
            ct.DeleteObject();
            clientContext.ExecuteQuery();

            Console.WriteLine(contentTypeName + ": Deleted");
        }

        public static string GetContentTypeIdByName(ClientContext clientContext, string contentTypeName)
        {
            Web oWeb = clientContext.Web;

            var itemContentType = clientContext.LoadQuery(oWeb.ContentTypes.Where(ct => ct.Name == contentTypeName));
            clientContext.ExecuteQuery();
            var contentTypeDetails = itemContentType.FirstOrDefault();
            var contentTypeId = contentTypeDetails.Id.ToString();

            return contentTypeId;
        }

        public static void ValidateAndAddSiteColumnToContentType()
        {
            var clientContext = AuthHelper.GetClientContext();
            Web oWeb = clientContext.Web;

            var contentTypeDetails = new Data().GetListInfo().Values.ToList();
            var contentTypeDetailsForLib = new Data().GetLibInfo().Values.ToList();

            contentTypeDetailsForLib.ForEach(x =>
            {
                contentTypeDetails.Add(x);
            });

            foreach (var ctDetails in contentTypeDetails)
            {
                for (int i = 0; i < ctDetails.Count(); i++)
                {
                    var contentTypeName = ctDetails.ToList()[i].Key;

                    var siteColumnDetails = ctDetails.ToList()[i].Value;

                    foreach (string columnName in siteColumnDetails)
                    {
                        Console.WriteLine(columnName);

                        bool isExist = oWeb.FieldExistsByNameInContentType(contentTypeName, columnName);
                        if (isExist)
                        {
                            RemoveFieldFromContentType(clientContext, oWeb, columnName, contentTypeName);
                            AddFieldToContentType(clientContext, oWeb, columnName, contentTypeName);
                        }
                        else
                        {
                            AddFieldToContentType(clientContext, oWeb, columnName, contentTypeName);
                        }
                    }
                }
            }

        }

        private static void RemoveFieldFromContentType(ClientContext clientContext, Web oWeb, string columnName, string contentTypeName)
        {
            string contentTypeID = GetContentTypeIdByName(clientContext, contentTypeName);
            ContentType oContentType = oWeb.ContentTypes.GetById(contentTypeID);
            Guid fieldID = SiteColumnHelper.GetSiteColumnIDByName(clientContext, oWeb, columnName);

            var fieldLinks = oContentType.FieldLinks;
            var fieldLinkToRemove = fieldLinks.GetById(fieldID);
            fieldLinkToRemove.DeleteObject();
            oContentType.Update(true); //push changes
            clientContext.ExecuteQuery();

            Console.WriteLine(columnName + " Site Column has been removed from the " + contentTypeName + " Content Type");
        }

        private static void AddFieldToContentType(ClientContext clientContext, Web oWeb, string columnName, string contentTypeName)
        {
            Field fieldName = oWeb.Fields.GetByInternalNameOrTitle(columnName);
            string contentTypeID = GetContentTypeIdByName(clientContext, contentTypeName);
            ContentType oContentType = oWeb.ContentTypes.GetById(contentTypeID);
            FieldLinkCreationInformation fdCI = new FieldLinkCreationInformation();
            fdCI.Field = fieldName;

            oContentType.FieldLinks.Add(fdCI);

            oContentType.Update(true);
            clientContext.ExecuteQuery();

            Console.WriteLine(columnName + " Site Column has been Added to " + contentTypeName + " Content Type");

        }
    }
}
