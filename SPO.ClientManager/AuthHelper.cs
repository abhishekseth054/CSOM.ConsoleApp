using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager
{
    public class AuthHelper
    {
        public static string adminSiteUrl = ConfigurationManager.AppSettings["SPAdminSiteURL"].ToString();
        public static string siteUrl = ConfigurationManager.AppSettings["SPSiteURL"].ToString();
        public static string userName = ConfigurationManager.AppSettings["UserName"].ToString();
        public static ClientContext GetClientContext()
        {

            //Console.WriteLine("Enter UserName");
            //string userName = Console.ReadLine();

            //  //For On-Premise

            //CredentialCache cc = new CredentialCache();
            //cc.Add(new Uri("SITEURL"), "NTLM", CredentialCache.DefaultNetworkCredentials);


            //ClientContext context = new ClientContext(SiteUrl)
            //{
            //    Credentials = cc,

            //    AuthenticationMode = ClientAuthenticationMode.Default

            //      // with Windows Authentication
            //    Credentials = CredentialCache.DefaultNetworkCredentials,

            //       // With custom Credential with Domain
            //    Credentials = new NetworkCredential("USERNAME", "PASSWORD", "DOMAINNAME")

            //       // With custom Credential without Domain
            //    Credentials = new NetworkCredential("USERNAMEt", "PASSWORD")

            //      // Form based Authentication
            //      AuthenticationMode = ClientAuthenticationMode.FormsAuthentication,
            //      FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo("USERNAME", "PASSWORD")

            //};

            //  //For Office 365 - SharePoint Online

            ClientContext context = new ClientContext(siteUrl)
            {
                Credentials = new SharePointOnlineCredentials(userName, GetPassword()),
            };

            return context;
        }

        public static ClientContext GetTenantContext()
        {
            ClientContext context = new ClientContext(adminSiteUrl)
            {
                Credentials = new SharePointOnlineCredentials(userName, GetPassword()),
            };

            return context;
        }

        private static SecureString GetPassword()
        {
            //Console.WriteLine("Enter Password");
            //string password = string.Empty;
            //ConsoleKeyInfo info = Console.ReadKey(true);
            //while (info.Key != ConsoleKey.Enter)
            //{
            //    if (info.Key != ConsoleKey.Backspace)
            //    {
            //        password += info.KeyChar;
            //        info = Console.ReadKey(true);
            //    }
            //    else if (info.Key == ConsoleKey.Backspace)
            //    {
            //        if (!string.IsNullOrEmpty(password))
            //        {
            //            password = password.Substring
            //            (0, password.Length - 1);
            //        }
            //        info = Console.ReadKey(true);
            //    }
            //}
            //for (int i = 0; i < password.Length; i++)
            //    Console.Write("*");

            var password = "Mango#123";

            SecureString s = new SecureString();
            foreach (char c in password)
            {
                s.AppendChar(c);
            }

            return s;
        }

    }
}
