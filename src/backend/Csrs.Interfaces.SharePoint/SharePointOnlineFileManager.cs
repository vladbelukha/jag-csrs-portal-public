using Microsoft.Extensions.Logging;
using Microsoft.Rest;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Csrs.Interfaces
{
    /// <summary>
    /// SharePoint file manager implementation for SharePoint Online (SPO).
    /// Uses Microsoft Graph API with OAuth2 client credentials (application permissions).
    /// </summary>
    public class SharePointOnlineFileManager : ISharePointFileManager
    {
        private const string DefaultDocumentListTitle = "Account";
        private const int MaxUrlLength = 256;

        private readonly SharePointOnlineConfiguration _configuration;
        private readonly ISharePointOnlineAuthenticator _authenticator;
        private readonly ILogger<SharePointOnlineFileManager> _logger;
        private readonly Dictionary<string, string> _driveIdCache = new(StringComparer.OrdinalIgnoreCase);

        private string _siteId;

        public SharePointOnlineFileManager(
            SharePointOnlineConfiguration configuration,
            ISharePointOnlineAuthenticator authenticator,
            ILogger<SharePointOnlineFileManager> logger)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _authenticator = authenticator ?? throw new ArgumentNullException(nameof(authenticator));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<List<FileDetailsList>> GetFileDetailsListInFolder(string listTitle, string folderName, string documentType)
        {
            folderName = FixFoldername(folderName);
            var driveId = await GetDriveIdAsync(listTitle);
            var fileDetailsList = new List<FileDetailsList>();

            string requestPath = string.IsNullOrEmpty(folderName)
                ? $"drives/{driveId}/root/children"
                : $"drives/{driveId}/{EncodeGraphPath(folderName)}/children";

            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, requestPath);

            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                _logger.LogInformation("The folder '{ListTitle}/{FolderName}' was not found.", listTitle, folderName);
                return fileDetailsList;
            }

            string responseContent = await ReadSuccessResponseAsync(response, HttpMethod.Get, requestPath);
            var children = ParseGraphCollection(responseContent);

            foreach (var item in children)
            {
                if (item["file"] == null)
                {
                    continue;
                }

                string name = item["name"]?.ToString();
                var fileDetails = new FileDetailsList
                {
                    Name = name,
                    Length = item["size"]?.ToString() ?? "0",
                    TimeCreated = item["createdDateTime"]?.ToString(),
                    TimeLastModified = item["lastModifiedDateTime"]?.ToString(),
                    ServerRelativeUrl = BuildServerRelativeUrl(listTitle, folderName, name),
                    DocumentType = Path.GetExtension(name)?.ToUpper(CultureInfo.InvariantCulture) ?? string.Empty,
                };
                fileDetailsList.Add(fileDetails);
            }

            return fileDetailsList;
        }

        public async Task CreateFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            var driveId = await GetDriveIdAsync(listTitle);

            var body = new JObject
            {
                ["name"] = folderName,
                ["folder"] = new JObject(),
                ["@microsoft.graph.conflictBehavior"] = "fail",
            };

            string requestPath = $"drives/{driveId}/root/children";
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Post, requestPath, body);

            if (!response.IsSuccessStatusCode)
            {
                await ThrowGraphExceptionAsync(response, HttpMethod.Post, requestPath);
            }
        }

        public async Task<object> CreateDocumentLibrary(string listTitle, string documentTemplateUrlTitle = null)
        {
            if (string.IsNullOrEmpty(documentTemplateUrlTitle))
            {
                documentTemplateUrlTitle = listTitle;
            }

            var body = new JObject
            {
                ["displayName"] = documentTemplateUrlTitle,
                ["list"] = new JObject { ["template"] = "documentLibrary" },
            };

            string requestPath = "lists";
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Post, requestPath, body);
            string responseContent = await ReadSuccessResponseAsync(response, HttpMethod.Post, requestPath, HttpStatusCode.Created);
            var createdList = JObject.Parse(responseContent);

            if (!string.Equals(listTitle, documentTemplateUrlTitle, StringComparison.OrdinalIgnoreCase))
            {
                string listId = createdList["id"]?.ToString();
                var updateBody = new JObject { ["displayName"] = listTitle };
                using var updateResponse = await SendGraphRequestAsync(
                    httpClient,
                    HttpMethod.Patch,
                    $"lists/{listId}",
                    updateBody);
                await ReadSuccessResponseAsync(updateResponse, HttpMethod.Patch, $"lists/{listId}");
            }

            _driveIdCache.Remove(listTitle);
            _driveIdCache.Remove(documentTemplateUrlTitle);

            return body;
        }

        public async Task<object> UpdateDocumentLibrary(string listTitle)
        {
            return await CreateDocumentLibrary(listTitle);
        }

        public async Task<bool> DeleteFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            var driveId = await GetDriveIdAsync(listTitle);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(folderName)}";

            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Delete, requestPath);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                return true;
            }

            await ThrowGraphExceptionAsync(response, HttpMethod.Delete, requestPath);
            return false;
        }

        public async Task<bool> FolderExists(string listTitle, string folderName)
        {
            return await GetFolder(listTitle, folderName) != null;
        }

        public async Task<bool> DocumentLibraryExists(string listTitle)
        {
            return await GetDocumentLibrary(listTitle) != null;
        }

        public async Task<object> GetFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            var driveId = await GetDriveIdAsync(listTitle);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(folderName)}";

            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, requestPath);

            if (response.StatusCode == HttpStatusCode.NotFound)
            {
                return null;
            }

            if (response.IsSuccessStatusCode)
            {
                string jsonString = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject(jsonString);
            }

            return null;
        }

        public async Task<object> GetDocumentLibrary(string listTitle)
        {
            var list = await GetListAsync(listTitle);
            return list == null ? null : list.ToObject<object>();
        }

        public async Task<string> AddFile(string folderName, string fileName, Stream fileData, string contentType)
        {
            return await AddFile(DefaultDocumentListTitle, folderName, fileName, fileData, contentType);
        }

        public async Task<string> AddFile(string documentLibrary, string folderName, string fileName, Stream fileData, string contentType)
        {
            folderName = FixFoldername(folderName);
            if (!await FolderExists(documentLibrary, folderName))
            {
                await CreateFolder(documentLibrary, folderName);
            }

            return await UploadFile(fileName, documentLibrary, folderName, fileData, contentType);
        }

        public async Task<string> AddFile(string folderName, string fileName, byte[] fileData, string contentType)
        {
            return await AddFile(DefaultDocumentListTitle, folderName, fileName, fileData, contentType);
        }

        public async Task<string> AddFile(string documentLibrary, string folderName, string fileName, byte[] fileData, string contentType)
        {
            folderName = FixFoldername(folderName);
            if (!await FolderExists(documentLibrary, folderName))
            {
                await CreateFolder(documentLibrary, folderName);
            }

            return await UploadFile(fileName, documentLibrary, folderName, fileData, contentType);
        }

        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, Stream fileData, string contentType)
        {
            using var ms = new MemoryStream();
            fileData.CopyTo(ms);
            return await UploadFile(fileName, listTitle, folderName, ms.ToArray(), contentType);
        }

        public string GetTruncatedFileName(string fileName, string listTitle, string folderName)
        {
            int maxLength = 128;
            fileName = FixFilename(fileName, maxLength);
            folderName = FixFoldername(folderName);

            string itemPath = BuildItemPath(folderName, fileName);
            string requestPath = $"drives/{{driveId}}/{EncodeGraphPath(itemPath)}/content";

            if (requestPath.Length > MaxUrlLength)
            {
                int delta = requestPath.Length - MaxUrlLength;
                maxLength -= delta;
                fileName = FixFilename(fileName, maxLength);
            }

            return fileName;
        }

        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, byte[] data, string contentType)
        {
            folderName = FixFoldername(folderName);
            fileName = GetTruncatedFileName(fileName, listTitle, folderName);

            var driveId = await GetDriveIdAsync(listTitle);
            string itemPath = BuildItemPath(folderName, fileName);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(itemPath)}/content";

            using var httpClient = await GetHttpClientAsync();
            using var request = await CreateGraphRequestAsync(httpClient, HttpMethod.Put, requestPath);

            var byteArrayContent = new ByteArrayContent(data);
            byteArrayContent.Headers.ContentType = new MediaTypeHeaderValue(
                string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType);
            request.Content = byteArrayContent;

            using var response = await httpClient.SendAsync(request);
            if (response.StatusCode is HttpStatusCode.OK or HttpStatusCode.Created)
            {
                return fileName;
            }

            await ThrowGraphExceptionAsync(response, HttpMethod.Put, requestPath);
            return fileName;
        }

        public async Task<string> UpdateListItemFields(AddFileResponse itemData, string listTitle, string contentType, string fileName)
        {
            var list = await GetListAsync(listTitle);
            if (list == null)
            {
                throw new SharePointRestException($"Document library '{listTitle}' was not found.");
            }

            string listId = list["id"]?.ToString();
            string itemId = itemData?.ListItemAllFields?.ID.ToString(CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(itemId))
            {
                throw new SharePointRestException("List item id was not provided.");
            }

            var body = new JObject
            {
                ["Title"] = contentType,
                ["FileLeafRef"] = fileName,
            };

            string requestPath = $"lists/{listId}/items/{itemId}/fields";
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Patch, requestPath, body);
            return await ReadSuccessResponseAsync(response, HttpMethod.Patch, requestPath);
        }

        public async Task<byte[]> DownloadFile(string url)
        {
            var (libraryName, itemPath) = ParseServerRelativeUrl(url);
            var driveId = await GetDriveIdAsync(libraryName);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(itemPath)}/content";

            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, requestPath);
            await EnsureSuccessAsync(response, HttpMethod.Get, requestPath);

            using var ms = new MemoryStream();
            await response.Content.CopyToAsync(ms);
            return ms.ToArray();
        }

        public async Task<bool> DeleteFile(string listTitle, string folderName, string fileName)
        {
            folderName = FixFoldername(folderName);
            return await DeleteFile(BuildServerRelativeUrl(listTitle, folderName, fileName));
        }

        public async Task<bool> DeleteFile(string serverRelativeUrl)
        {
            var (libraryName, itemPath) = ParseServerRelativeUrl(serverRelativeUrl);
            var driveId = await GetDriveIdAsync(libraryName);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(itemPath)}";

            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Delete, requestPath);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                return true;
            }

            await ThrowGraphExceptionAsync(response, HttpMethod.Delete, requestPath);
            return false;
        }

        public async Task<bool> RenameFile(string oldServerRelativeUrl, string newServerRelativeUrl)
        {
            var (oldLibrary, oldItemPath) = ParseServerRelativeUrl(oldServerRelativeUrl);
            var (newLibrary, newItemPath) = ParseServerRelativeUrl(newServerRelativeUrl);

            if (!string.Equals(oldLibrary, newLibrary, StringComparison.OrdinalIgnoreCase))
            {
                throw new SharePointRestException("Renaming files across document libraries is not supported.");
            }

            string newFileName = Path.GetFileName(newItemPath.Replace('/', Path.DirectorySeparatorChar));
            var driveId = await GetDriveIdAsync(oldLibrary);
            string requestPath = $"drives/{driveId}/{EncodeGraphPath(oldItemPath)}";

            using var httpClient = await GetHttpClientAsync();
            using var getResponse = await SendGraphRequestAsync(httpClient, HttpMethod.Get, requestPath);
            string itemJson = await ReadSuccessResponseAsync(getResponse, HttpMethod.Get, requestPath);
            string itemId = JObject.Parse(itemJson)["id"]?.ToString();

            var body = new JObject { ["name"] = newFileName };
            using var patchResponse = await SendGraphRequestAsync(
                httpClient,
                HttpMethod.Patch,
                $"drives/{driveId}/items/{itemId}",
                body);

            if (patchResponse.IsSuccessStatusCode)
            {
                return true;
            }

            await ThrowGraphExceptionAsync(patchResponse, HttpMethod.Patch, $"drives/{driveId}/items/{itemId}");
            return false;
        }

        private async Task<HttpClient> GetHttpClientAsync()
        {
#pragma warning disable CA2000 // Dispose objects before losing scope
            HttpMessageHandler handler = new SocketsHttpHandler
            {
                AllowAutoRedirect = false,
                MaxConnectionsPerServer = 25,
            };
#pragma warning restore CA2000

            var httpClient = new HttpClient(handler)
            {
                BaseAddress = _configuration.GraphBaseUri,
            };
            httpClient.DefaultRequestHeaders.Accept.Add(MediaTypeWithQualityHeaderValue.Parse("application/json"));

            string accessToken = await _authenticator.GetAccessTokenAsync();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            return httpClient;
        }

        private async Task<string> GetSiteIdAsync()
        {
            if (!string.IsNullOrEmpty(_siteId))
            {
                return _siteId;
            }

            var resourceUri = _configuration.Resource;
            string sitePath = resourceUri.AbsolutePath.TrimEnd('/');
            if (string.IsNullOrEmpty(sitePath))
            {
                sitePath = "/";
            }

            string requestPath = $"sites/{resourceUri.Host}:{sitePath}";
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, requestPath);
            string responseContent = await ReadSuccessResponseAsync(response, HttpMethod.Get, requestPath);
            _siteId = JObject.Parse(responseContent)["id"]?.ToString();

            if (string.IsNullOrEmpty(_siteId))
            {
                throw new SharePointRestException("Unable to resolve SharePoint site id from Microsoft Graph.");
            }

            return _siteId;
        }

        private async Task<JToken> GetListAsync(string listTitle)
        {
            string siteId = await GetSiteIdAsync();
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, "lists?$select=id,name,displayName,list");
            string responseContent = await ReadSuccessResponseAsync(response, HttpMethod.Get, "lists");

            foreach (var list in ParseGraphCollection(responseContent))
            {
                if (list["list"]?["template"]?.ToString() != "documentLibrary")
                {
                    continue;
                }

                if (MatchesListTitle(list, listTitle))
                {
                    return list;
                }
            }

            return null;
        }

        private async Task<string> GetDriveIdAsync(string listTitle)
        {
            if (_driveIdCache.TryGetValue(listTitle, out string cachedDriveId))
            {
                return cachedDriveId;
            }

            var list = await GetListAsync(listTitle);
            if (list == null)
            {
                throw new SharePointRestException($"Document library '{listTitle}' was not found.");
            }

            string listId = list["id"]?.ToString();
            using var httpClient = await GetHttpClientAsync();
            using var response = await SendGraphRequestAsync(httpClient, HttpMethod.Get, $"lists/{listId}/drive");
            string responseContent = await ReadSuccessResponseAsync(response, HttpMethod.Get, $"lists/{listId}/drive");
            string driveId = JObject.Parse(responseContent)["id"]?.ToString();

            if (string.IsNullOrEmpty(driveId))
            {
                throw new SharePointRestException($"Unable to resolve drive id for document library '{listTitle}'.");
            }

            _driveIdCache[listTitle] = driveId;
            return driveId;
        }

        private static bool MatchesListTitle(JToken list, string listTitle)
        {
            return string.Equals(list["name"]?.ToString(), listTitle, StringComparison.OrdinalIgnoreCase)
                || string.Equals(list["displayName"]?.ToString(), listTitle, StringComparison.OrdinalIgnoreCase);
        }

        private async Task<HttpRequestMessage> CreateGraphRequestAsync(HttpClient httpClient, HttpMethod method, string relativePath)
        {
            string siteId = await GetSiteIdAsync();
            var request = new HttpRequestMessage(method, $"sites/{siteId}/{relativePath}");
            return request;
        }

        private async Task<HttpResponseMessage> SendGraphRequestAsync(
            HttpClient httpClient,
            HttpMethod method,
            string relativePath,
            JObject body = null)
        {
            var request = await CreateGraphRequestAsync(httpClient, method, relativePath);
            if (body != null)
            {
                request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
            }

            return await httpClient.SendAsync(request);
        }

        private static List<JToken> ParseGraphCollection(string responseContent)
        {
            var responseObject = JObject.Parse(responseContent);
            if (responseObject["value"] is JArray values)
            {
                return values.Children().ToList();
            }

            return new List<JToken>();
        }

        private async Task<string> ReadSuccessResponseAsync(
            HttpResponseMessage response,
            HttpMethod method,
            string relativePath,
            HttpStatusCode? expectedStatusCode = null)
        {
            await EnsureSuccessAsync(response, method, relativePath, expectedStatusCode);
            return response.Content == null ? string.Empty : await response.Content.ReadAsStringAsync();
        }

        private async Task EnsureSuccessAsync(
            HttpResponseMessage response,
            HttpMethod method,
            string relativePath,
            HttpStatusCode? expectedStatusCode = null)
        {
            if (expectedStatusCode.HasValue && response.StatusCode == expectedStatusCode.Value)
            {
                return;
            }

            if (response.IsSuccessStatusCode)
            {
                return;
            }

            await ThrowGraphExceptionAsync(response, method, relativePath);
        }

        private async Task ThrowGraphExceptionAsync(HttpResponseMessage response, HttpMethod method, string relativePath)
        {
            string responseContent = response.Content == null
                ? string.Empty
                : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            var ex = new SharePointRestException(
                $"Microsoft Graph operation returned an invalid status code '{response.StatusCode}'");
            ex.Request = new HttpRequestMessageWrapper(new HttpRequestMessage(method, relativePath), null);
            ex.Response = new HttpResponseMessageWrapper(response, responseContent);
            throw ex;
        }

        private string GetSiteServerRelativePath()
        {
            return _configuration.Resource.AbsolutePath.TrimEnd('/');
        }

        private string BuildServerRelativeUrl(string listTitle, string folderName, string fileName)
        {
            string path = $"{GetSiteServerRelativePath()}/{listTitle}";
            if (!string.IsNullOrEmpty(folderName))
            {
                path += $"/{folderName}";
            }

            if (!string.IsNullOrEmpty(fileName))
            {
                path += $"/{fileName}";
            }

            return path;
        }

        private (string LibraryName, string ItemPath) ParseServerRelativeUrl(string serverRelativeUrl)
        {
            string normalizedUrl = serverRelativeUrl.Trim().TrimStart('/');
            string sitePath = GetSiteServerRelativePath().TrimStart('/');

            if (normalizedUrl.StartsWith(sitePath, StringComparison.OrdinalIgnoreCase))
            {
                normalizedUrl = normalizedUrl.Substring(sitePath.Length).TrimStart('/');
            }

            int separatorIndex = normalizedUrl.IndexOf('/');
            if (separatorIndex < 0)
            {
                return (normalizedUrl, string.Empty);
            }

            string libraryName = normalizedUrl.Substring(0, separatorIndex);
            string itemPath = normalizedUrl.Substring(separatorIndex + 1);
            return (libraryName, itemPath);
        }

        private static string BuildItemPath(string folderName, string fileName)
        {
            if (string.IsNullOrEmpty(folderName))
            {
                return fileName;
            }

            return $"{folderName}/{fileName}";
        }

        private static string EncodeGraphPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return "root";
            }

            var encodedSegments = path
                .Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(Uri.EscapeDataString);
            return $"root:/{string.Join("/", encodedSegments)}:";
        }

        private string GetInvalidCharacters(string osInvalidChars)
        {
            osInvalidChars += "~#%&*()[]{}:;+@^<>|.?!/";
            string invalidChars = Regex.Escape(osInvalidChars);
            return string.Format(CultureInfo.InvariantCulture, @"([{0}]*\.+$)|([{0}]+)", invalidChars);
        }

        private string RemoveInvalidCharacters(string filename)
        {
            var osInvalidChars = new string(Path.GetInvalidFileNameChars());
            return Regex.Replace(filename, GetInvalidCharacters(osInvalidChars), "_");
        }

        private string FixFoldername(string foldername)
        {
            var osInvalidChars = new string(Path.GetInvalidPathChars());
            return Regex.Replace(foldername, GetInvalidCharacters(osInvalidChars), "_");
        }

        private string FixFilename(string filename, int maxLength = 128)
        {
            string result = RemoveInvalidCharacters(filename);
            if (result.Length >= maxLength)
            {
                string extension = Path.GetExtension(result);
                result = Path.GetFileNameWithoutExtension(result).Substring(0, maxLength - extension.Length);
                result += extension;
            }

            return result;
        }
    }
}
