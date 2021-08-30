using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using System.Security.Cryptography.X509Certificates;


namespace ADB.Permission.Service
{
    public class CommonUtility
    {
        private static string clientId = "9dad629b-2e4c-4255-825a-5688072bb8dd";
        private static string certThumprint = "C712EA824CCEEEA93A5E56ECC85CF1501DE338C5";
        private static string tenantId = "c8734771-c201-4e53-a6de-dec0ca209654";
        private static string[] scopes = new string[] {"https://abhijeeto365.sharepoint.com/.default" };

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