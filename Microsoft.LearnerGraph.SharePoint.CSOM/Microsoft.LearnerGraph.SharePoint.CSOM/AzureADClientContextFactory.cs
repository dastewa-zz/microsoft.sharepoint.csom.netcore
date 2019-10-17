// <copyright file="AzureADClientContextFactory.cs" company="Microsoft">
//
// Copyright (c) Microsoft. All rights reserved.
//
// </copyright>

namespace Microsoft.LearnerGraph.SharePoint.CSOM
{
    using System;
    using System.Net.Http;
    using System.Security;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.SharePoint.Client;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// AzureADClientContextFactory constructs and returns a ClientContext
    /// for SharePoint that is authenticated with Azure Active Directory.
    /// I.e. the ClientContext is using a valid Azure AD token for it's
    /// requests.
    /// </summary>
    public class AzureADClientContextFactory
    {
        /// <summary>
        /// Backing field for HttpClient.
        /// </summary>
        private HttpClient httpClient;

        /// <summary>
        /// Gets or sets the httpClient.
        /// </summary>
        public HttpClient HttpClient
        {
            get
            {
                if (this.httpClient != null)
                {
                    return this.httpClient;
                }

                this.httpClient = new HttpClient();
                return this.httpClient;
            }

            set
            {
                this.httpClient = value ?? throw new ArgumentNullException(nameof(value));
            }
        }

        /// <summary>
        /// Gets a lock.
        /// </summary>
        private object Lock { get; } = new object();

        /// <summary>
        /// Gets the endpoint for token procurement.
        /// </summary>
        private string AzureADTokenEndpoint { get; } = "https://login.microsoftonline.com/common/oauth2/token";

        /// <summary>
        /// Gets or sets the Azure AD token.
        /// </summary>
        private string AzureADToken { get; set; }

        /// <summary>
        /// Gets the content type to use during the token procurement.
        /// </summary>
        private string TokenRequestContentType { get; } = "application/x-www-form-urlencoded";

        /// <summary>
        /// Gets the key for the token in the JSON object.
        /// </summary>
        private string TokenObjectKey { get; } = "access_token";

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated.</param>
        /// <param name="userPrincipalName">The user id.</param>
        /// <param name="userPassword">The user's password as a string.</param>
        /// <returns>Client context object.</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, string userPassword)
        {
            var spUri = new Uri(siteUrl);
            string resourceUri = spUri.Scheme + "://" + spUri.Authority;

            var clientContext = new ClientContext(siteUrl);
            clientContext.DisableReturnValueCache = true;

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                this.EnsureAzureADCredentialsToken(resourceUri, userPrincipalName, userPassword);
                args.WebRequest.Headers.Add(System.Net.HttpRequestHeader.Authorization, $"Bearer {this.AzureADToken}");
            };

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated.</param>
        /// <param name="userPrincipalName">The user id.</param>
        /// <param name="userPassword">The user's password as a secure string.</param>
        /// <returns>Client context object.</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, SecureString userPassword)
        {
            string password = new System.Net.NetworkCredential(string.Empty, userPassword).Password;
            return this.GetAzureADCredentialsContext(siteUrl, userPrincipalName, password);
        }

        /// <summary>
        /// The lease for the access token to track expiry during use.
        /// </summary>
        /// <param name="expiresOn">The datetime that the token expires.</param>
        /// <returns>A timespan representing when the lease duration.</returns>
        private static TimeSpan GetAccessTokenLease(DateTime expiresOn)
        {
            DateTime now = DateTime.UtcNow;
            DateTime expires = expiresOn.Kind == DateTimeKind.Utc ?
                expiresOn : TimeZoneInfo.ConvertTimeToUtc(expiresOn);
            TimeSpan lease = expires - now;
            return lease;
        }

        /// <summary>
        /// The token request body.
        /// </summary>
        /// <param name="resourceUri">The token endpoint.</param>
        /// <param name="userName">The AD username (email).</param>
        /// <param name="password">The AD password.</param>
        /// <returns>The request string.</returns>
        private static string BuildTokenRequestBodyAsString(string resourceUri, string userName, string password)
        {
            if (string.IsNullOrWhiteSpace(resourceUri))
            {
                throw new ArgumentException(nameof(resourceUri));
            }

            if (string.IsNullOrWhiteSpace(userName))
            {
                throw new ArgumentException(nameof(userName));
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException(nameof(password));
            }

            StringBuilder sb = new StringBuilder();

            sb.Append($"resource={resourceUri}");
            sb.Append($"&client_id=9bc3ab49-b65d-410a-85ad-de819febfddc&grant_type=password");
            sb.Append($"&username={userName}");
            sb.Append($"&password={password}");

            return sb.ToString();
        }

        /// <summary>
        /// Acquire a token from AD for the given user/password.
        /// </summary>
        /// <param name="resourceUri">Token endpoint.</param>
        /// <param name="userName">The username.</param>
        /// <param name="password">The password.</param>
        /// <returns>The access token.</returns>
        private async Task<string> AcquireTokenAsync(string resourceUri, string userName, string password)
        {
            StringContent requestBody = new StringContent(
                BuildTokenRequestBodyAsString(resourceUri, userName, password),
                Encoding.UTF8,
                this.TokenRequestContentType);

            string result = await this.HttpClient.PostAsync(this.AzureADTokenEndpoint, requestBody).ContinueWith<string>((response) =>
            {
                return response.Result.Content.ReadAsStringAsync().Result;
            });

            return JObject.Parse(result)[this.TokenObjectKey].Value<string>();
        }

        /// <summary>
        /// Renew the Azure AD token if it has expired.
        /// </summary>
        /// <param name="resourceUri">The token endpoint.</param>
        /// <param name="userPrincipalName">The AD username (AD).</param>
        /// <param name="userPassword">The AD password.</param>
        private void EnsureAzureADCredentialsToken(string resourceUri, string userPrincipalName, string userPassword)
        {
            if (this.AzureADToken == null)
            {
                lock (this.Lock)
                {
                    if (this.AzureADToken == null)
                    {
                        string accessToken = Task.Run(() => this.AcquireTokenAsync(resourceUri, userPrincipalName, userPassword)).GetAwaiter().GetResult();
                        ThreadPool.QueueUserWorkItem(obj =>
                        {
                            try
                            {
                                var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);

                                var lease = GetAccessTokenLease(token.ValidTo);
                                lease =
                                    TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
                                Thread.Sleep(lease);
                                this.AzureADToken = null;
                            }
                            catch (Exception)
                            {
                                this.AzureADToken = null;
                            }
                        });

                        this.AzureADToken = accessToken;
                    }
                }
            }
        }
    }
}
