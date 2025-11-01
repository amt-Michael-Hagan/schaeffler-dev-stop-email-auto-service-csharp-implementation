using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Threading;

namespace EmailAutomationLegacy.Services
{
    public class TokenManager
    {
        private readonly ClientSecretCredential _credential;
        private readonly string[] _scopes = new[] { "https://graph.microsoft.com/.default" };
        private AccessToken? _currentToken;
        private DateTimeOffset _tokenExpiresOn;

        public TokenManager()
        {
            _credential = new ClientSecretCredential(
                AppSettings.TenantId,
                AppSettings.ClientId,
                AppSettings.ClientSecret
            );
        }


        public string GetAccessToken()
        {
            // Return cached token if it's still valid
            if (_currentToken.HasValue && DateTimeOffset.UtcNow < _tokenExpiresOn)
            {
                return _currentToken.Value.Token;
            }

            // Request a new token
            _currentToken = _credential.GetToken(
                new TokenRequestContext(_scopes),
                CancellationToken.None
            );

            // Set token expiration (with 5 minute buffer)
            _tokenExpiresOn = _currentToken.Value.ExpiresOn.AddMinutes(-5);

            return _currentToken.Value.Token;
        }

        public virtual GraphServiceClient GetGraphClient()
        {
            return new GraphServiceClient(_credential, _scopes);
        }

    }
}