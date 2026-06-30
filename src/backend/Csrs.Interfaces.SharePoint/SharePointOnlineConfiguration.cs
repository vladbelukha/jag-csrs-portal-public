using System;

namespace Csrs.Interfaces
{
    public class SharePointOnlineConfiguration
    {
        /// <summary>
        /// The Azure AD / Entra ID tenant identifier.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// The application (client) identifier registered in Azure AD.
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// The client secret for the registered application.
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// The SharePoint Online site URL (e.g. https://tenant.sharepoint.com/sites/csrs/).
        /// Used to resolve the site in Microsoft Graph and to build server-relative URLs.
        /// </summary>
        public Uri Resource { get; set; }

        /// <summary>
        /// Microsoft Graph API base URL.
        /// </summary>
        public Uri GraphBaseUri { get; set; } = new Uri("https://graph.microsoft.com/v1.0/");
    }
}
