using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using EmailAutomationLegacy.Models;

namespace EmailAutomationLegacy.Services
{
    public class GraphApiClient
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(GraphApiClient));
        private readonly TokenManager _tokenManager;

        public GraphApiClient(TokenManager tokenManager)
        {
            _tokenManager = tokenManager;
        }

        public T Get<T>(string endpoint)
        {
            return ExecuteRequest<T>("GET", endpoint, null);
        }

        public byte[] GetBytes(string endpoint)
        {
            for (int attempt = 0; attempt < AppSettings.RetryAttempts; attempt++)
            {
                try
                {
                    var request = CreateRequest("GET", endpoint);
                    
                    using (var response = (HttpWebResponse)request.GetResponse())
                    {
                        if (response.StatusCode != HttpStatusCode.OK)
                        {
                            throw new InvalidOperationException($"Request failed: {response.StatusCode}");
                        }

                        using (var stream = response.GetResponseStream())
                        {
                            var buffer = new byte[response.ContentLength];
                            var totalRead = 0;
                            
                            while (totalRead < buffer.Length)
                            {
                                var read = stream.Read(buffer, totalRead, buffer.Length - totalRead);
                                if (read == 0) break;
                                totalRead += read;
                            }
                            
                            return buffer;
                        }
                    }
                }
                catch (Exception ex) when (attempt < AppSettings.RetryAttempts - 1)
                {
                    log.Warn($"Request attempt {attempt + 1} failed, retrying: {ex.Message}");
                    Thread.Sleep(AppSettings.RetryDelayMs);
                }
            }
            
            throw new InvalidOperationException($"Failed to get bytes after {AppSettings.RetryAttempts} attempts");
        }

        private T ExecuteRequest<T>(string method, string endpoint, object body)
        {
            for (int attempt = 0; attempt < AppSettings.RetryAttempts; attempt++)
            {
                try
                {
                    var request = CreateRequest(method, endpoint);
                    
                    if (body != null && method != "GET")
                    {
                        var json = JsonConvert.SerializeObject(body);
                        var data = Encoding.UTF8.GetBytes(json);
                        request.ContentLength = data.Length;
                        request.ContentType = "application/json";
                        
                        using (var stream = request.GetRequestStream())
                        {
                            stream.Write(data, 0, data.Length);
                        }
                    }

                    using (var response = (HttpWebResponse)request.GetResponse())
                    {
                        //if (response.StatusCode != HttpStatusCode.OK)
                        //{
                        //    throw new InvalidOperationException($"Request failed: {response.StatusCode}");
                        //}

                        using (var reader = new StreamReader(response.GetResponseStream()))
                        {
                            var responseText = reader.ReadToEnd();
                            
                            if (typeof(T) == typeof(string))
                            {
                                return (T)(object)responseText;
                            }
                            
                            return JsonConvert.DeserializeObject<T>(responseText);
                        }
                    }
                }
                catch (WebException ex) when (attempt < AppSettings.RetryAttempts - 1)
                {
                    log.Warn($"Request attempt {attempt + 1} failed, retrying: {ex.Message}");
                    Thread.Sleep(AppSettings.RetryDelayMs);
                }
                catch (Exception ex) when (attempt < AppSettings.RetryAttempts - 1)
                {
                    log.Warn($"Request attempt {attempt + 1} failed, retrying: {ex.Message}");
                    Thread.Sleep(AppSettings.RetryDelayMs);
                }
            }
            
            throw new InvalidOperationException($"Failed to execute request after {AppSettings.RetryAttempts} attempts");
        }

        private HttpWebRequest CreateRequest(string method, string endpoint)
        {
            var url = endpoint.StartsWith("https://") ? endpoint : $"https://graph.microsoft.com/v1.0/{endpoint}";
            var request = (HttpWebRequest)WebRequest.Create(url);
            
            request.Method = method;
            request.Headers.Add("Authorization", $"Bearer {_tokenManager.GetAccessToken()}");
            request.UserAgent = "EmailAutomation/1.0";
            request.Accept = "application/json";
            
            return request;
        }

        public GraphResponse<T> GetPaged<T>(string endpoint)
        {
            var response = Get<GraphResponse<T>>(endpoint);
            return response;
        }
    }
}