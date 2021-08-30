
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using Microsoft.AspNetCore.Mvc;
using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using System.IdentityModel.Tokens;
using System.Web;

namespace ADB.UserProfile.Service.Controllers
{
    [ApiController]
    [Route("/api/[controller]")]
    public class UserProfileController : ControllerBase
    {

        public AuthenticationProviderParams ProviderParams;

        public UserProfileController()
        {
            ProviderParams = new AuthenticationProviderParams();
            ProviderParams.TenantId = "58ede2a3-9bf3-4920-ad14-d17b16b972cd";
            ProviderParams.ClientId = "17d1a726-36a1-4d02-8b36-b33aa5e8a424";
            ProviderParams.ClientThumbPrint = "651D96D07307790E0F4642D4F655980ABD7B6464";
            ProviderParams.AppScopes = new string[] { "https://v2smartsolutions.sharepoint.com/.default", "https://graph.microsoft.com/.default" };
            ProviderParams.GetCertificate();
        }

        [HttpGet]
        public string Get()
        {
            return "pass parameter group name as \"/api/UserProfile/<<GroupName>>\"";
        }

        [HttpGet("{groupName}")]
        public async Task<IEnumerable<string>> Get(string groupName)
        {
            IEnumerable<string> result = new List<string>();
            result = await this.GetData(groupName);
            return result;
        }


        private async Task<IEnumerable<string>> GetData(string groupName)
        {
            List<string> result = new List<string>();

            // string siteUrl = "https://v2smartsolutions.sharepoint.com/sites/Dev1";
            string siteUrl = "https://v2smartsolutions.sharepoint.com/sites/Dev1";

            //var certificate = GetCertificate(@"D:\Certificate\PowerCi.pfx", "acpass@123");

            var accessToken = await GetApplicationAuthenticatedClient(ProviderParams.ClientId, ProviderParams.ClientThumbPrint, new string[] { ProviderParams.AppScopes[0] }, ProviderParams.TenantId);

            try
            {
                using (var clientContext = GetClientContextWithAccessToken(siteUrl, accessToken))
                {
                    Group group = clientContext.Web.SiteGroups.GetByName(groupName);
                    clientContext.Load(group);
                    //clientContext.Load(clientContext.Web, w => w.Title);
                    clientContext.ExecuteQuery();
                    // var webTitle = clientContext.Web.Title;
                    UserCollection users = group.Users;
                    clientContext.Load(users);
                    clientContext.ExecuteQuery();
                    foreach (User user in group.Users)
                    {
                        switch (user.PrincipalType)
                        {
                            case Microsoft.SharePoint.Client.Utilities.PrincipalType.User:
                                if (!result.Contains(user.Title)) result.Add(user.Title);
                                break;
                            case Microsoft.SharePoint.Client.Utilities.PrincipalType.SecurityGroup:
                                List<string> securityGroupUsers = await GetSecurityGroupUsers(user.LoginName.Split("|").Last(), accessToken);
                                foreach (string sUser in securityGroupUsers) { if (!result.Contains(sUser)) result.Add(sUser); }
                                break;
                        }
                    }
                    // result.Add(webTitle);
                }
            }
            catch (System.Exception e)
            {
                result.Add(e.Message);
            }

            return result;
        }

        private async Task<List<string>> GetSecurityGroupUsers(string userId, string accessToken)
        {
            List<string> memberNames = new List<string>();
            // Microsoft.Graph.GraphServiceClient client = new Microsoft.Graph.GraphServiceClient(new AzureAuthenticationProvider(accessToken));
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
            this.ProviderParams.AppScopes = scopes;
            Microsoft.Graph.GraphServiceClient client = new Microsoft.Graph.GraphServiceClient("https://graph.microsoft.com/v1.0",
                new AuthenticationProvider(this.ProviderParams)
                );

            await GetGroupMembers(userId, memberNames, client);

            return memberNames;
        }

        private async Task GetGroupMembers(string userId, List<string> memberNames, Microsoft.Graph.GraphServiceClient client)
        {
            Microsoft.Graph.Group group = await client.Groups[userId].Request().GetAsync();

            // Microsoft.Graph.IGraphServiceGroupsCollectionPage group = await client.Groups.Request().Filter($"startswith(mail, '{userId}')").GetAsync();
            if (group != null)
            {
                Microsoft.Graph.IGroupMembersCollectionWithReferencesPage members = await client.Groups[userId].Members.Request().GetAsync();
                // IGroupMembersCollectionWithReferencesPage members = await group.Members.Select("displayName,mail,id").GetAsync();
                if (members != null)
                {
                    foreach (var item in members)
                    {
                        switch (item.ODataType)
                        {
                            case "#microsoft.graph.group":
                                await GetGroupMembers(item.Id, memberNames, client);
                                break;
                            case "#microsoft.graph.user":
                                Microsoft.Graph.User user = item as Microsoft.Graph.User;
                                if (!memberNames.Contains(user.DisplayName)) memberNames.Add(user.DisplayName);
                                break;
                            default:
                                memberNames.Add(item.Id);
                                break;
                        }
                    }
                }
            }

        }

        private static X509Certificate2 GetCertificate(string path, string password)
        {
            return new X509Certificate2(path, password, X509KeyStorageFlags.MachineKeySet);
        }
        internal static async Task<string> GetApplicationAuthenticatedClient(string clientId, string certThumprint, string[] scopes, string tenantId)
        {
            X509Certificate2 certificate = GetAppOnlyCertificate(certThumprint);
            IConfidentialClientApplication clientApp = ConfidentialClientApplicationBuilder
                                            .Create(clientId)
                                            .WithCertificate(certificate)
                                            .WithTenantId(tenantId)
                                            .Build();
            // .WithCertificate(certificate)
            Microsoft.Identity.Client.AuthenticationResult authResult = await clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
            // string accessToken = authResult.AccessToken;
            return authResult.AccessToken;
        }

        private static X509Certificate2 GetAppOnlyCertificate(string thumbPrint)
        {
            X509Certificate2 appOnlyCertificate = null;
            using (X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                certStore.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, thumbPrint, false);
                if (certCollection.Count > 0)
                {
                    appOnlyCertificate = certCollection[0];
                }
                certStore.Close();
                return appOnlyCertificate;
            }
        }
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };
            return clientContext;
        }


    }
}