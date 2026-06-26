using System.Threading.Tasks;

namespace Csrs.Interfaces
{
    public interface ISharePointOnlineAuthenticator
    {
        /// <summary>
        /// Returns a valid OAuth2 bearer token for Microsoft Graph.
        /// Tokens are cached internally until they expire.
        /// </summary>
        Task<string> GetAccessTokenAsync();
    }
}
