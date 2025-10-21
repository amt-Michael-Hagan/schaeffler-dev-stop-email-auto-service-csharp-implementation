using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using Newtonsoft.Json;

namespace EmailAutomationLegacy.Services
{
    public class TokenManager
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(TokenManager));
        
        private string _cachedToken;
        private DateTime _tokenExpiry;
        private readonly object _tokenLock = new object();

        public string GetAccessToken()
        {
            lock (_tokenLock)
            {
                // Return cached token if still valid (with 5 minute buffer)
                if (!string.IsNullOrEmpty(_cachedToken) && DateTime.UtcNow < _tokenExpiry.AddMinutes(-5))
                {
                    return _cachedToken;
                }

                // Request new token
                return RequestNewToken();
            }
        }

        private string RequestNewToken()
        {
            try
            {
                var tokenEndpoint = $"https://login.microsoftonline.com/{AppSettings.TenantId}/oauth2/v2.0/token";
                
                var postData = string.Format(
                    "client_id={0}&scope={1}&client_secret={2}&grant_type=client_credentials",
                    Uri.EscapeDataString(AppSettings.ClientId),
                    Uri.EscapeDataString("https://graph.microsoft.com/.default"),
                    Uri.EscapeDataString(AppSettings.ClientSecret)
                );

                var request = (HttpWebRequest)WebRequest.Create(tokenEndpoint);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.UserAgent = "EmailAutomation/1.0";

                // Write POST data
                var data = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = data.Length;
                
                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                // Get response
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        throw new InvalidOperationException($"Token request failed: {response.StatusCode}");
                    }

                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        var responseText = reader.ReadToEnd();
                        var tokenResponse = JsonConvert.DeserializeObject<TokenResponse>(responseText);
                        
                        if (tokenResponse?.AccessToken == null)
                        {
                            throw new InvalidOperationException("No access token received");
                        }

                        _cachedToken = tokenResponse.AccessToken;
                        _tokenExpiry = DateTime.UtcNow.AddSeconds(tokenResponse.ExpiresIn);
                        
                        log.Info($"Successfully obtained access token, expires at: {_tokenExpiry}");
                        return _cachedToken;
                    }
                }
            }
            catch (WebException ex)
            {
                string errorDetails = "Unknown error";
                if (ex.Response != null)
                {
                    using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                    {
                        errorDetails = reader.ReadToEnd();
                    }
                }
                
                log.Error($"Token request failed: {ex.Message}. Response: {errorDetails}", ex);
                throw new InvalidOperationException($"Authentication failed: {ex.Message}");
            }
            catch (Exception ex)
            {
                log.Error("Token request failed", ex);
                throw;
            }
        }

        private class TokenResponse
        {
            [JsonProperty("access_token")]
            public string AccessToken { get; set; }

            [JsonProperty("expires_in")]
            public int ExpiresIn { get; set; }

            [JsonProperty("token_type")]
            public string TokenType { get; set; }
        }
    }
}