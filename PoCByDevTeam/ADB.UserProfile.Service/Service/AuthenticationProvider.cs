using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;


namespace ADB.UserProfile.Service
{
    public class AuthenticationProviderParams
    {
        public string ClientId { get; set; }
        public string ClientThumbPrint { get; set; }
        private X509Certificate2 _certificate;
        public X509Certificate2 Certificate
        {
            get
            {
                return this._certificate;
            }
        }
        public string TenantId { get; set; }
        public string[] AppScopes { get; set; }

        public void GetCertificate()
        {
            X509Certificate2 appOnlyCertificate = null;
            using (X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                certStore.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = certStore.Certificates.Find(X509FindType.FindByThumbprint, this.ClientThumbPrint, false);
                if (certCollection.Count > 0)
                {
                    appOnlyCertificate = certCollection[0];
                }
                certStore.Close();
                this._certificate = appOnlyCertificate;
            }
        }
    }
    public class AuthenticationProvider : IAuthenticationProvider
    {
        public AuthenticationProviderParams ProviderParams { get; set; }

        public AuthenticationProvider(AuthenticationProviderParams providerParams)
        {
            this.ProviderParams = providerParams;
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var clientApplication = ConfidentialClientApplicationBuilder.Create(this.ProviderParams.ClientId)
                .WithAuthority($"https://login.microsoftonline.com/{this.ProviderParams.TenantId}")
                .WithCertificate(this.ProviderParams.Certificate)
                .WithTenantId(this.ProviderParams.TenantId)
                .Build();

            var result = await clientApplication.AcquireTokenForClient(this.ProviderParams.AppScopes).ExecuteAsync();
            var authHeader = result.CreateAuthorizationHeader();

            request.Headers.Add("Authorization", authHeader);
        }


    }
}