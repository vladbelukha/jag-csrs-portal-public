using Microsoft.Extensions.Logging;
using Microsoft.Rest;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Csrs.Interfaces
{
    /// <summary>
    /// SharePoint file manager implementation for SharePoint Online (SPO).
    /// Uses OAuth2 client credentials (bearer token) for authentication
    /// rather than the SAML / FedAuth cookie flow used by the on-premises implementation.
    /// </summary>
    public class SharePointOnlineFileManager : ISharePointFileManager
    {
        private const string DefaultDocumentListTitle = "Account";
        private const string SharePointSpaceCharacter = "_x0020_";
        private const int MaxUrlLength = 256;

        private readonly SharePointOnlineConfiguration _configuration;
        private readonly ISharePointOnlineAuthenticator _authenticator;
        private readonly ILogger<SharePointOnlineFileManager> _logger;

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

            string serverRelativeUrl = Uri.EscapeUriString(listTitle);
            if (!string.IsNullOrEmpty(folderName))
            {
                serverRelativeUrl += "/" + Uri.EscapeUriString(folderName);
            }
            serverRelativeUrl = EscapeApostrophe(serverRelativeUrl);

            string responseContent = null;
            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/getFolderByServerRelativeUrl('/{serverRelativeUrl}')/files"),
                Headers = { { "Accept", "application/json" } }
            };

            using var httpClient = await GetHttpClientAsync();
            HttpResponseMessage response = null;
            List<FileDetailsList> fileDetailsList = new List<FileDetailsList>();

            try
            {
                response = await httpClient.SendAsync(request);
            }
            catch (ArgumentNullException)
            {
                var ex = new SharePointRestException("The response is null.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The response is null.");
                throw ex;
            }
            catch (InvalidOperationException)
            {
                var ex = new SharePointRestException("The request message was already sent by the HttpClient instance.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The request message was already sent by the HttpClient instance.");
                throw ex;
            }
            catch (HttpRequestException)
            {
                var ex = new SharePointRestException("The request failed due to an underlying issue such as network connectivity, DNS failure, server certificate validation or timeout.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The request failed due to an underlying issue such as network connectivity, DNS failure, server certificate validation or timeout.");
                throw ex;
            }
            catch (TaskCanceledException)
            {
                var ex = new SharePointRestException("The request failed due to timeout.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The request failed due to timeout.");
                throw ex;
            }

            HttpStatusCode statusCode = response.StatusCode;

            if ((int)statusCode == 404)
            {
                _logger.LogInformation("The folder '{ServerRelativeUrl}' is not found.", serverRelativeUrl);
                return fileDetailsList;
            }

            if ((int)statusCode != 200)
            {
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);
                throw ex;
            }
            else
            {
                responseContent = await response.Content.ReadAsStringAsync();
            }

            JObject responseObject = null;
            try
            {
                responseObject = JObject.Parse(responseContent);
            }
            catch (JsonReaderException jre)
            {
                _logger.LogError("Error parsing response: {Content}", responseContent);
                throw jre;
            }

            List<JToken> responseResults = responseObject["value"].Children().ToList();
            foreach (JToken responseResult in responseResults)
            {
                FileDetailsList searchResult = responseResult.ToObject<FileDetailsList>();
                string fileDoctype = Path.GetExtension(searchResult.Name).ToUpper();
                searchResult.DocumentType = fileDoctype ?? string.Empty;
                fileDetailsList.Add(searchResult);
            }

            return fileDetailsList;
        }

        public async Task CreateFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            string relativeUrl = EscapeApostrophe($"/{listTitle}/{folderName}");

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/folders/add('{relativeUrl}')"),
                Headers = { { "Accept", "application/json" } }
            };

            StringContent strContent = new StringContent("", Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            request.Content = strContent;

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);
                throw ex;
            }
        }

        public async Task<object> CreateDocumentLibrary(string listTitle, string documentTemplateUrlTitle = null)
        {
            using var listsRequest = new HttpRequestMessage(HttpMethod.Post, new Uri(_configuration.Resource, "_api/web/Lists"));

            if (string.IsNullOrEmpty(documentTemplateUrlTitle))
            {
                documentTemplateUrlTitle = listTitle;
            }

            var library = CreateNewDocumentLibraryRequest(documentTemplateUrlTitle);
            string jsonString = JsonConvert.SerializeObject(library);
            StringContent strContent = new StringContent(jsonString, Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            listsRequest.Content = strContent;
            listsRequest.Headers.Add("odata-version", "3.0");

            using var httpClient = await GetHttpClientAsync();
            using var listsResponse = await httpClient.SendAsync(listsRequest);
            HttpStatusCode statusCode = listsResponse.StatusCode;

            if (statusCode != HttpStatusCode.Created || !listsResponse.IsSuccessStatusCode)
            {
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                string responseContent = await listsResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(listsRequest, null);
                ex.Response = new HttpResponseMessageWrapper(listsResponse, responseContent);
                throw ex;
            }
            else
            {
                jsonString = await listsResponse.Content.ReadAsStringAsync();
                var ob = JsonConvert.DeserializeObject<DocumentLibraryResponse>(jsonString);

                if (listTitle != documentTemplateUrlTitle)
                {
                    using var titleRequest = new HttpRequestMessage(HttpMethod.Post, new Uri(_configuration.Resource, $"_api/web/lists(guid'{ob.d.Id}')"));
                    var updateRequest = new
                    {
                        __metadata = new { type = "SP.List" },
                        Title = listTitle
                    };
                    jsonString = JsonConvert.SerializeObject(updateRequest);
                    strContent = new StringContent(jsonString, Encoding.UTF8);
                    strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                    titleRequest.Headers.Add("IF-MATCH", "*");
                    titleRequest.Headers.Add("X-HTTP-Method", "MERGE");
                    titleRequest.Content = strContent;

                    using var titleResponse = await httpClient.SendAsync(titleRequest);
                    titleResponse.EnsureSuccessStatusCode();
                }
            }

            return library;
        }

        public async Task<object> UpdateDocumentLibrary(string listTitle)
        {
            using var request = new HttpRequestMessage(HttpMethod.Put, new Uri(_configuration.Resource, "_api/web/Lists"));

            var library = CreateNewDocumentLibraryRequest(listTitle);
            string jsonString = JsonConvert.SerializeObject(library);
            StringContent strContent = new StringContent(jsonString, Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            request.Content = strContent;

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            HttpStatusCode statusCode = response.StatusCode;

            if (statusCode != HttpStatusCode.Created || !response.IsSuccessStatusCode)
            {
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);
                throw ex;
            }

            return library;
        }

        public async Task<bool> DeleteFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            string serverRelativeUrl = $"{listTitle}/{folderName}";

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Delete,
                RequestUri = new Uri(_configuration.Resource, "_api/web/getFolderByServerRelativeUrl('/" + EscapeApostrophe(serverRelativeUrl) + "')"),
                Headers = { { "Accept", "application/json" } }
            };
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("X-HTTP-Method", "DELETE");

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                return true;
            }

            var ex2 = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
            string content = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            ex2.Request = new HttpRequestMessageWrapper(request, null);
            ex2.Response = new HttpResponseMessageWrapper(response, content);
            throw ex2;
        }

        public async Task<bool> FolderExists(string listTitle, string folderName)
        {
            var folder = await GetFolder(listTitle, folderName);
            return folder != null;
        }

        public async Task<bool> DocumentLibraryExists(string listTitle)
        {
            var library = await GetDocumentLibrary(listTitle);
            return library != null;
        }

        public async Task<object> GetFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            string serverRelativeUrl = $"{listTitle}/{folderName}";

            using var endpointRequest = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, "_api/web/getFolderByServerRelativeUrl('/" + EscapeApostrophe(serverRelativeUrl) + "')"),
                Headers = { { "Accept", "application/json" } }
            };

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(endpointRequest);
            string jsonString = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK || response.IsSuccessStatusCode)
            {
                return JsonConvert.DeserializeObject(jsonString);
            }

            return null;
        }

        public async Task<object> GetDocumentLibrary(string listTitle)
        {
            string title = Uri.EscapeUriString(listTitle);

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/lists/GetByTitle('{title}')"),
                Headers = { { "Accept", "application/json" } }
            };

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            string jsonString = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK)
            {
                return JsonConvert.DeserializeObject(jsonString);
            }

            return null;
        }

        public async Task<string> AddFile(string folderName, string fileName, Stream fileData, string contentType)
        {
            return await AddFile(DefaultDocumentListTitle, folderName, fileName, fileData, contentType);
        }

        public async Task<string> AddFile(string documentLibrary, string folderName, string fileName, Stream fileData, string contentType)
        {
            folderName = FixFoldername(folderName);
            bool folderExists = await FolderExists(documentLibrary, folderName);
            if (!folderExists)
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
            bool folderExists = await FolderExists(documentLibrary, folderName);
            if (!folderExists)
            {
                await CreateFolder(documentLibrary, folderName);
            }
            return await UploadFile(fileName, documentLibrary, folderName, fileData, contentType);
        }

        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, Stream fileData, string contentType)
        {
            using var ms = new MemoryStream();
            fileData.CopyTo(ms);
            byte[] data = ms.ToArray();
            return await UploadFile(fileName, listTitle, folderName, data, contentType);
        }

        public string GetTruncatedFileName(string fileName, string listTitle, string folderName)
        {
            int maxLength = 128;
            fileName = FixFilename(fileName, maxLength);
            folderName = FixFoldername(folderName);

            string serverRelativeUrl = GetServerRelativeURL(listTitle, folderName);
            string requestUriString = GenerateUploadRequestUriString(serverRelativeUrl, fileName);

            if (requestUriString.Length > MaxUrlLength)
            {
                int delta = requestUriString.Length - MaxUrlLength;
                maxLength -= delta;
                fileName = FixFilename(fileName, maxLength);
            }

            return fileName;
        }

        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, byte[] data, string contentType)
        {
            folderName = FixFoldername(folderName);
            fileName = GetTruncatedFileName(fileName, listTitle, folderName);

            string serverRelativeUrl = GetServerRelativeURL(listTitle, folderName);
            string requestUriString = GenerateUploadRequestUriString(serverRelativeUrl, fileName);

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(requestUriString),
                Headers = { { "Accept", "application/json" } }
            };

            ByteArrayContent byteArrayContent = new ByteArrayContent(data);
            byteArrayContent.Headers.Add("content-length", data.Length.ToString());
            request.Content = byteArrayContent;

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                return fileName;
            }

            var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            ex.Request = new HttpRequestMessageWrapper(request, null);
            ex.Response = new HttpResponseMessageWrapper(response, responseContent);
            throw ex;
        }

        public async Task<string> UpdateListItemFields(AddFileResponse itemData, string listTitle, string contentType, string fileName)
        {
            string requestUriString = GenerateUpdateListItemUriString(listTitle, itemData.ListItemAllFields.ID.ToString());

            using var request = new HttpRequestMessage(HttpMethod.Post, requestUriString);

            var listItem = CreateUpdateListItemRequestRequest(contentType, fileName, listTitle);
            string jsonString = JsonConvert.SerializeObject(listItem);
            StringContent strContent = new StringContent(jsonString, Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

            request.Headers.Add("Accept", "application/json;odata=verbose");
            request.Headers.Add("X-Http-Method", "MERGE");
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("odata-version", "3.0");
            request.Content = strContent;

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            string streamData = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.NoContent)
            {
                return streamData;
            }

            var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            ex.Request = new HttpRequestMessageWrapper(request, null);
            ex.Response = new HttpResponseMessageWrapper(response, responseContent);
            throw ex;
        }

        public async Task<byte[]> DownloadFile(string url)
        {
            url = EscapeApostrophe(url);

            using var endpointRequest = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('{url}')/$value"),
            };

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(endpointRequest);

            using var ms = new MemoryStream();
            await response.Content.CopyToAsync(ms);
            return ms.ToArray();
        }

        public async Task<bool> DeleteFile(string listTitle, string folderName, string fileName)
        {
            folderName = FixFoldername(folderName);
            string serverRelativeUrl = $"/{listTitle}/{folderName}/{fileName}";
            return await DeleteFile(serverRelativeUrl);
        }

        public async Task<bool> DeleteFile(string serverRelativeUrl)
        {
            serverRelativeUrl = EscapeApostrophe(serverRelativeUrl);

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Delete,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('/{serverRelativeUrl}')"),
                Headers = { { "Accept", "application/json" } }
            };
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("X-HTTP-Method", "DELETE");

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                return true;
            }

            var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            ex.Request = new HttpRequestMessageWrapper(request, null);
            ex.Response = new HttpResponseMessageWrapper(response, responseContent);
            throw ex;
        }

        public async Task<bool> RenameFile(string oldServerRelativeUrl, string newServerRelativeUrl)
        {
            oldServerRelativeUrl = EscapeApostrophe(oldServerRelativeUrl);
            newServerRelativeUrl = EscapeApostrophe(newServerRelativeUrl);

            Uri url = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('{oldServerRelativeUrl}')/moveto(newurl='{newServerRelativeUrl}', flags=1)");
            using var request = new HttpRequestMessage(HttpMethod.Post, url);

            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                return true;
            }

            var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
            string responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            ex.Request = new HttpRequestMessageWrapper(request, null);
            ex.Response = new HttpResponseMessageWrapper(response, responseContent);
            throw ex;
        }

        private async Task<HttpClient> GetHttpClientAsync()
        {
#pragma warning disable CA2000 // Dispose objects before losing scope
            HttpMessageHandler handler = new SocketsHttpHandler
            {
                AllowAutoRedirect = false,
                MaxConnectionsPerServer = 25
            };
#pragma warning restore CA2000

            HttpClient httpClient = new HttpClient(handler);
            httpClient.BaseAddress = _configuration.Resource;
            httpClient.DefaultRequestHeaders.Accept.Add(MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose"));

            string accessToken = await _authenticator.GetAccessTokenAsync();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            string digest = await GetDigest(httpClient);
            if (digest is not null)
            {
                httpClient.DefaultRequestHeaders.Add("X-RequestDigest", digest);
            }

            httpClient.DefaultRequestHeaders.Add("Cache-Control", "no-cache");
            httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");

            return httpClient;
        }

        private async Task<string> GetDigest(HttpClient client)
        {
            string result = null;

            using var endpointRequest = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_configuration.Resource, "_api/contextinfo"),
                Headers = { { "Accept", "application/json;odata=verbose" } }
            };

            var response = await client.SendAsync(endpointRequest);
            string jsonString = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK && jsonString.Length > 1)
            {
                if (jsonString[0] == '{')
                {
                    JToken t = JToken.Parse(jsonString);
                    result = t["d"]["GetContextWebInformation"]["FormDigestValue"].ToString();
                }
                else
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(jsonString);
                    var digests = doc.GetElementsByTagName("d:FormDigestValue");
                    if (digests.Count > 0)
                    {
                        result = digests[0].InnerText;
                    }
                }
            }

            return result;
        }

        private static string EscapeApostrophe(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                return value.Replace("'", "''");
            }
            return value;
        }

        private string GetInvalidCharacters(string osInvalidChars)
        {
            osInvalidChars += "~#%&*()[]{}:;+@^<>|.?!/";
            string invalidChars = Regex.Escape(osInvalidChars);
            return string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);
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

        private string GetServerRelativeURL(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            return Uri.EscapeUriString(listTitle) + "/" + Uri.EscapeUriString(folderName);
        }

        private string GenerateUploadRequestUriString(string folderServerRelativeUrl, string fileName)
        {
            folderServerRelativeUrl = EscapeApostrophe(folderServerRelativeUrl);
            fileName = EscapeApostrophe(fileName);
            Uri requestUri = new Uri(_configuration.Resource, $"_api/web/getFolderByServerRelativeUrl('/{folderServerRelativeUrl}')/Files/add(url='{fileName}',overwrite=true)?$expand=ListItemAllFields");
            return requestUri.ToString();
        }

        private string GenerateUpdateListItemUriString(string listTitle, string id)
        {
            listTitle = EscapeApostrophe(listTitle);
            id = EscapeApostrophe(id);
            Uri requestUri = new Uri(_configuration.Resource, $"_api/web/lists/getbytitle('{listTitle}')/items({id})");
            return requestUri.ToString();
        }

        private static object CreateNewDocumentLibraryRequest(string listName)
        {
            return new
            {
                __metadata = new { type = "SP.List" },
                BaseTemplate = 101,
                Title = listName
            };
        }

        private static object CreateUpdateListItemRequestRequest(string title, string fileName, string listTitle)
        {
            string formattedTitle = listTitle.Replace(" ", SharePointSpaceCharacter);
            string itemType = $"SP.Data.{formattedTitle}Item";
            return new
            {
                __metadata = new { type = itemType },
                FileLeafRef = fileName,
                Title = title
            };
        }
    }
}
