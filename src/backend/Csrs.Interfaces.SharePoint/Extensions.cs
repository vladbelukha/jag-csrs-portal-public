using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace Csrs.Interfaces
{
    public static class Extensions
    {
        /// <summary>
        /// Adds the required services to use SharePoint file management.
        /// When <c>UseSharePointOnlineContingency</c> is set to <c>true</c> in configuration,
        /// registers <see cref="SharePointOnlineFileManager"/> (Microsoft Graph API, SharePoint Online).
        /// Otherwise registers <see cref="SharePointFileManager"/> (SAML / FedAuth, on-premises).
        /// </summary>
        public static void AddSharePointIntegration(this IServiceCollection services, IConfiguration configuration)
        {
            bool useSpo = string.Equals(
                configuration["USE_SHAREPOINT_ONLINE_CONTINGENCY"] ?? configuration["UseSharePointOnlineContingency"],
                "true",
                StringComparison.OrdinalIgnoreCase);

            if (useSpo)
            {
                services.AddSingleton(GetSharePointOnlineConfiguration);
                services.AddSingleton<ISharePointOnlineAuthenticator, SharePointOnlineAuthenticator>();
                services.AddScoped<ISharePointFileManager, SharePointOnlineFileManager>();
            }
            else
            {
                services.AddSingleton(GetSharePointFileManagerConfiguration);
                // SAML services
                // token caches are singleton because they maintain a per instance prefix
                // that can be changed to effectively clear the cache
                services.AddSingleton<ITokenCache<SamlTokenParameters, string>, SamlTokenTokenCache>();
                services.AddTransient<ISamlAuthenticator, SamlAuthenticator>();
                services.AddScoped<ISharePointFileManager, SharePointFileManager>();
            }

            services.AddMemoryCache();
        }

        private static SharePointFileManagerConfiguration GetSharePointFileManagerConfiguration(IServiceProvider serviceProvider)
        {
            IConfiguration configuration = serviceProvider.GetService<IConfiguration>();

            var sharePointFileManagerConfiguration = new SharePointFileManagerConfiguration
            {
                ApiGatewayHost = configuration["APIGATEWAY_HOST"],
                ApiGatewayPolicy = configuration["APIGATEWAY_POLICY"],
                RelyingPartyIdentifier = configuration["RELYING_PARTY_IDENTIFIER"],
                AuthorizationUri = new Uri(configuration["AUTHORIZATION_URI"]),
                Resource = new Uri(configuration["RESOURCE"]),
                Username = configuration["SHAREPOINT_USERNAME"],
                Password = configuration["SHAREPOINT_PASSWORD"],
            };

            return sharePointFileManagerConfiguration;
        }

        private static SharePointOnlineConfiguration GetSharePointOnlineConfiguration(IServiceProvider serviceProvider)
        {
            IConfiguration configuration = serviceProvider.GetService<IConfiguration>();

            return new SharePointOnlineConfiguration
            {
                TenantId = configuration["SPO_TENANT_ID"],
                ClientId = configuration["SPO_CLIENT_ID"],
                ClientSecret = configuration["SPO_CLIENT_SECRET"],
                Resource = new Uri(configuration["SPO_RESOURCE"]),
            };
        }
    }
}

