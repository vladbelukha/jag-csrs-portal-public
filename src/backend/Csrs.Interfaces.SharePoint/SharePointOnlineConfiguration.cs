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
        /// The SharePoint Online site URL used as the resource base address.
        /// </summary>
        public Uri Resource { get; set; }
    }
}
