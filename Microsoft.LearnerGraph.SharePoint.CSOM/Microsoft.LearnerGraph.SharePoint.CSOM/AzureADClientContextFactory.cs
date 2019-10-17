//  -----------------------------------------------------------------------
//  
//  <copyright file="Startup.cs" company="Microsoft">
//  
//  Copyright (c) Microsoft. All rights reserved.
// 
//  </copyright>
// 
//  -----------------------------------------------------------------------

namespace Microsoft.LearnerGraph.SharePoint.CSOM
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Net.Mime;
    using System.Text;
    using System.Threading.Tasks;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// AzureADClientContextFactory constructs and returns a ClientContext
    /// for SharePoint that is authenticated with Azure Active Directory.
    /// I.e. the ClientContext is using a valid Azure AD token for it's
    /// requests.
    /// </summary>
    public static class AzureADClientContextFactory
    {
        /// <summary>
        /// Endpoint for token procurement.
        /// </summary>
        private static string AzureADTokenEndpoint { get; } = "https://login.microsoftonline.com/common/oauth2/token";

        /// <summary>
        /// Content type to use during the token procurement.
        /// </summary>
        private static string TokenRequestContentType { get; } = "application/x-www-form-urlencoded";

        /// <summary>
        /// Key for the token in the JSON object.
        /// </summary>
        private static string TokenObjectKey { get; } = "access_token";

        /// <summary>
        /// Backing field for HttpClient
        /// </summary>
        private static HttpClient httpClient;

        /// <summary>
        /// HttpClient
        /// </summary>
        public static HttpClient HttpClient
        {
            get
            {
                if (httpClient != null)
                {
                    return httpClient;
                }

                httpClient = new HttpClient();
                return httpClient;
            }

            set
            {
                httpClient = value ?? throw new ArgumentNullException(nameof(value));
            }
        }

        private static string BuildTokenRequestBodyAsString(string resourceUri, string userName, string password)
        {
            if (String.IsNullOrWhiteSpace(resourceUri))
            {
                throw new ArgumentException(nameof(resourceUri));
            }

            if (String.IsNullOrWhiteSpace(userName))
            {
                throw new ArgumentException(nameof(userName));
            }

            if (String.IsNullOrWhiteSpace(password))
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
        /// <returns></returns>
        private static async Task<string> AcquireTokenAsync(string resourceUri, string userName, string password)
        {
            StringContent requestBody = new StringContent(
                BuildTokenRequestBodyAsString(resourceUri, userName, password),
                System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

            string result = await HttpClient.PostAsync(AzureADTokenEndpoint, requestBody).ContinueWith<string>((response) =>
            {
                return response.Result.Content.ReadAsStringAsync().Result;
            });

            return JObject.Parse(result)[TokenObjectKey].Value<string>();
        }
    }
}
