using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using System.Security.Cryptography.X509Certificates;


namespace ADB.CopyDocument.Service
{
    public class CommonUtility
    {
        private static string clientId = "17d1a726-36a1-4d02-8b36-b33aa5e8a424";
        private static string certThumprint = "651D96D07307790E0F4642D4F655980ABD7B6464";
        private static string tenantId = "58ede2a3-9bf3-4920-ad14-d17b16b972cd";
        private static string[] scopes = new string[] { "https://v2smartsolutions.sharepoint.com/.default" };

        public static ClientContext GetClientContextWithAccessToken(string targetUrl)
        {
            Task<string> accessToken = GetApplicationAuthenticatedClient(clientId, certThumprint, scopes, tenantId);
            ClientContext clientContext = new ClientContext(targetUrl);
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken.Result;
                };
            return clientContext;
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

        internal static X509Certificate2 GetAppOnlyCertificate(string thumbPrint)
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


    }
}