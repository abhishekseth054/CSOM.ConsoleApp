using Microsoft.SharePoint.Client.UserProfiles;
using System.DirectoryServices;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager
{
    public class UserProfile
    {

        public static void UpdateUserprofileCustomAttribute()
        {
            SyncFromADUserByWhenChanged();
        }

        private static void SyncFromADUserByWhenChanged()
        {
            int days = Convert.ToInt16(ConfigurationManager.AppSettings["NumberOfDays"]);
            string forestName = ConfigurationManager.AppSettings["ForestName"];
            string adminName = ConfigurationManager.AppSettings["Administrator"];
            string adminPassword = ConfigurationManager.AppSettings["AdminPassword"];
            string LdapPath = ConfigurationManager.AppSettings["LdapPath"];

            var lastModifiedDate = DateTime.Now.AddDays(-days).Date;
            DirectoryEntry searchRoot = new DirectoryEntry(LdapPath, adminName, adminPassword);
            DirectorySearcher search = new DirectorySearcher(searchRoot);

            search.Filter = "(&(objectClass=user)(objectCategory=person))";
            using (SearchResultCollection results = search.FindAll())
            {
                int successCount = 0;
                int failureCount = 0;
                foreach (SearchResult result in results)
                {
                    string userName = GetProperty(result, "userPrincipalName");
                    string samAccountName = GetProperty(result, "sAMAccountName");
                    //These AD attributs are custom created
                    string subDepartment = GetProperty(result, "SPSubDepartment");
                    string birthday = GetProperty(result, "SPBirthday");
                    string aniversaryDate = GetProperty(result, "SPAniversaryDate");

                    var modifiedDate = Convert.ToDateTime(GetProperty(result, "whenChanged"));

                    if (!string.IsNullOrEmpty(samAccountName)
                            && DateTime.Compare(modifiedDate, lastModifiedDate) >= 0)
                    {

                        Console.WriteLine(userName + " | " + samAccountName + " | " + subDepartment + " | " + birthday + 
                                                    " | " + aniversaryDate + " | " + modifiedDate);

                        bool res = UpdateProfileProperties(samAccountName, userName, subDepartment
                                                , aniversaryDate, birthday, modifiedDate);

                        if (res)
                        {
                            successCount += 1;
                        }
                        else
                        {
                            failureCount += 1;
                        }
                    }
                }
                Console.WriteLine(successCount + " User Profile Updated");
                Console.WriteLine(failureCount + " User Profile Updation Fail");
            }
        }

        private static string GetProperty(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        private static bool UpdateProfileProperties(string userAccount, string userName, string subDepartment
                                            , string Anniversarydate, string birthday, DateTime lastUpdated)
        {
            var SubDepartmentAttribute = ConfigurationManager.AppSettings["SubDepartmentAttribute"];
            var AnniversaryAttribute = ConfigurationManager.AppSettings["AnniversaryAttribute"];
            var BirthdayAttribute = ConfigurationManager.AppSettings["BirthdayAttribute"];

            var userAccountName = "i:0#.f|membership|" + AuthHelper.userName;

            var tenantCtx = AuthHelper.GetTenantContext();
            var peopleManager = new PeopleManager(tenantCtx);

            peopleManager.SetSingleValueProfileProperty(userAccountName, SubDepartmentAttribute, subDepartment);
            peopleManager.SetSingleValueProfileProperty(userAccountName, AnniversaryAttribute, Anniversarydate);
            peopleManager.SetSingleValueProfileProperty(userAccountName, BirthdayAttribute, birthday);

            tenantCtx.ExecuteQuery();
            return true;
        }
    }
}
