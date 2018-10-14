using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Linq;
using System.Security;

namespace SPO.ClientManager
{
    public class Helper
    {
        public static dynamic IsExist_Helper(ClientContext context, String fieldToCheck, String type)
        {
            var isExist = 0;
            Web oWeb = context.Web;
            ListCollection listCollection = oWeb.Lists;
            ContentTypeCollection cntCollection = oWeb.ContentTypes;
            FieldCollection fldCollection = oWeb.Fields;

            switch (type)
            {
                case "list":
                    context.Load(listCollection, lsts => lsts.Include(list => list.Title).Where(list => list.Title == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = listCollection.Count;
                    break;
                case "contenttype":
                    context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name).Where(ct => ct.Name == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = cntCollection.Count;
                    break;
                case "contenttypeName":
                    context.Load(cntCollection, cntyp => cntyp.Include(ct => ct.Name, ct => ct.Id).Where(ct => ct.Name == fieldToCheck));
                    context.ExecuteQuery();
                    foreach (ContentType ct in cntCollection)
                    {
                        return ct.Id.ToString();
                    }
                    break;
                case "field":
                    FieldCollection fieldColl = oWeb.Fields;

                    var siteColumnDetails = context.LoadQuery(fieldColl.Where(ct => ct.InternalName == fieldToCheck));
                    context.ExecuteQuery();

                    var siteColumn = siteColumnDetails.FirstOrDefault();

                    if (siteColumn != null && !String.IsNullOrEmpty(siteColumn.InternalName))
                    {
                        isExist = 1;
                    }
                    else
                    {
                        isExist = 0;
                    }
                    break;
                case "listcntype":
                    List lst = context.Web.Lists.GetByTitle(fieldToCheck);
                    ContentTypeCollection lstcntype = lst.ContentTypes;
                    context.Load(lstcntype, lstc => lstc.Include(lc => lc.Name).Where(lc => lc.Name == fieldToCheck));
                    context.ExecuteQuery();
                    isExist = lstcntype.Count;
                    break;
            }
            return isExist;
        }
    }
}
