using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
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
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Csrs.Interfaces
{

    public class SharePointFileManager : ISharePointFileManager
    {
        public const string DefaultDocumentUrlTitle = "account";
        public const string ApplicationDocumentListTitle = "Application";
        public const string ApplicationDocumentUrlTitle = "adoxio_application";
        public const string ContactDocumentListTitle = "contact";
        public const string WorkerDocumentListTitle = "Worker Qualification";
        public const string WorkerDocumentUrlTitle = "adoxio_worker";
        public const string EventDocumentListTitle = "adoxio_event";
        public const string FederalReportListTitle = "adoxio_federalreportexport";
        public const string LicenceDocumentUrlTitle = "adoxio_licences";
        public const string LicenceDocumentListTitle = "Licence";

        private const string DefaultDocumentListTitle = "Account";
        private const string SharePointSpaceCharacter = "_x0020_";

        private const int MaxUrlLength = 256; // default maximum URL length.
        private readonly SharePointFileManagerConfiguration _configuration;
        private readonly ISamlAuthenticator _samlAuthenticator;
        private readonly ILogger<SharePointFileManager> _logger;

        public SharePointFileManager(SharePointFileManagerConfiguration configuration, ISamlAuthenticator samlAuthenticator, ILogger<SharePointFileManager> logger)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _samlAuthenticator = samlAuthenticator ?? throw new ArgumentNullException(nameof(samlAuthenticator));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// Escape the apostrophe character.  Since we use it to enclose the filename it must be escaped.
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>Filename, with apropstophes escaped.</returns>
        private string EscapeApostrophe(string filename)
        {
            string result = null;
            if (!string.IsNullOrEmpty(filename))
            {
                result = filename.Replace("'", "''");
            }
            return result;
        }

        /// <summary>
        /// Get file details list from SharePoint filtered by folder name and document type
        /// </summary>
        /// <param name="listTitle"></param>
        /// <param name="folderName"></param>
        /// <param name="documentType"></param>
        /// <returns></returns>
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
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            //using var response = await httpClient.SendAsync(request);
            HttpResponseMessage response = null;

            // create file details list to add from response
            List<FileDetailsList> fileDetailsList = new List<FileDetailsList>();

            try
            {
                response = await httpClient.SendAsync(request);
            }
            catch (ArgumentNullException e)
            {
                var ex = new SharePointRestException("The response is null.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The response is null.");
                throw ex;
            }
            catch (InvalidOperationException e)
            {
                var ex = new SharePointRestException("The request message was already sent by the HttpClient instance.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The request message was already sent by the HttpClient instance.");
                throw ex;
            }
            catch (HttpRequestException e)
            {
                var ex = new SharePointRestException("The request failed due to an underlying issue such as network connectivity, DNS failure, server certificate validation or timeout.");
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, null);
                _logger.LogError("The request failed due to an underlying issue such as network connectivity, DNS failure, server certificate validation or timeout.");
                throw ex;
            }
            catch (TaskCanceledException e)
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
                _logger.LogInformation($"The folder '{serverRelativeUrl}' is not found.");
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

            // parse the response
            JObject responseObject = null;
            try
            {
                responseObject = JObject.Parse(responseContent);
            }
            catch (JsonReaderException jre)
            {
                _logger.LogError("Error in parse the response", responseContent);
                throw jre;
            }
            // get JSON response objects into a list
            List<JToken> responseResults = responseObject["value"].Children().ToList();

            // create .NET objects
            foreach (JToken responseResult in responseResults)
            {
                // JToken.ToObject is a helper method that uses JsonSerializer internally
                FileDetailsList searchResult = responseResult.ToObject<FileDetailsList>();
                //filter by parameter documentType
                string fileDoctype = System.IO.Path.GetExtension(searchResult.Name).ToUpper();
                searchResult.DocumentType = fileDoctype;
                fileDetailsList.Add(searchResult);
                if (searchResult.DocumentType == null)
                {
                    searchResult.DocumentType = string.Empty;
                }
            }
            return fileDetailsList;
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

            // Get the validated file name string
            string result = Regex.Replace(filename, GetInvalidCharacters(osInvalidChars), "_");

            return result;
        }

        private string FixFoldername(string foldername)
        {
            var osInvalidChars = new string(Path.GetInvalidPathChars());

            // Get the validated folder name string
            string result = Regex.Replace(foldername, GetInvalidCharacters(osInvalidChars), "_");

            return result;
        }

        private string FixFilename(string filename, int maxLength = 128)
        {
            string result = RemoveInvalidCharacters(filename);

            // SharePoint requires that the filename is less than 128 characters.

            if (result.Length >= maxLength)
            {
                string extension = Path.GetExtension(result);
                result = Path.GetFileNameWithoutExtension(result).Substring(0, maxLength - extension.Length);
                result += extension;
            }

            return result;
        }

        /// <summary>
        /// Create Folder
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public async Task CreateFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);

            string relativeUrl = EscapeApostrophe($"/{listTitle}/{folderName}");

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/folders/add('{relativeUrl}')"),
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            //string jsonString = "{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '" + relativeUrl + "'}";

            StringContent strContent = new StringContent("", Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

            request.Content = strContent;

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            HttpStatusCode statusCode = response.StatusCode;

            // check to see if the folder creation worked.
            if (!response.IsSuccessStatusCode)
            {
                string responseContent;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                throw ex;
            }
            else
            {
                string jsonString = await response.Content.ReadAsStringAsync();
            }


        }
        /// <summary>
        /// Create Folder
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
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
            // fix for bad request
            listsRequest.Headers.Add("odata-version", "3.0");

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var listsResponse = await httpClient.SendAsync(listsRequest);

            HttpStatusCode statusCode = listsResponse.StatusCode;

            if (statusCode != HttpStatusCode.Created || !listsResponse.IsSuccessStatusCode)
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                responseContent = await listsResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(listsRequest, null);
                ex.Response = new HttpResponseMessageWrapper(listsResponse, responseContent);

                throw ex;
            }
            else
            {
                jsonString = await listsResponse.Content.ReadAsStringAsync();
                var ob = Newtonsoft.Json.JsonConvert.DeserializeObject<DocumentLibraryResponse>(jsonString);

                if (listTitle != documentTemplateUrlTitle)
                {
                    // update list title
                    using var titleRequest = new HttpRequestMessage(HttpMethod.Post, new Uri(_configuration.Resource, "_api/web/lists(guid'{ob.d.Id}')"));
                    var type = new { type = "SP.List" };
                    var request = new
                    {
                        __metadata = type,
                        Title = listTitle
                    };
                    jsonString = JsonConvert.SerializeObject(request);
                    strContent = new StringContent(jsonString, Encoding.UTF8);
                    strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
                    titleRequest.Headers.Add("IF-MATCH", "*");
                    titleRequest.Headers.Add("X-HTTP-Method", "MERGE");
                    titleRequest.Content = strContent;

                    using var titleResponse = await httpClient.SendAsync(titleRequest);
                    jsonString = await titleResponse.Content.ReadAsStringAsync();
                    titleResponse.EnsureSuccessStatusCode();
                }
            }

            return library;
        }

        public async Task<Object> UpdateDocumentLibrary(string listTitle)
        {
            using var request = new HttpRequestMessage(HttpMethod.Put, new Uri(_configuration.Resource, "_api/web/Lists"));

            var library = CreateNewDocumentLibraryRequest(listTitle);

            string jsonString = JsonConvert.SerializeObject(library);
            StringContent strContent = new StringContent(jsonString, Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");
            request.Content = strContent;

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            HttpStatusCode statusCode = response.StatusCode;

            if (statusCode != HttpStatusCode.Created || !response.IsSuccessStatusCode)
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{statusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                throw ex;
            }
            else
            {
                jsonString = await response.Content.ReadAsStringAsync();
            }

            return library;
        }

        private object CreateNewDocumentLibraryRequest(string listName)
        {
            var type = new { type = "SP.List" };
            var request = new
            {
                __metadata = type,
                BaseTemplate = 101,
                Title = listName
            };
            return request;
        }

        private object CreateUpdateListItemRequestRequest(string title, string fileName, string listTitle)
        {
            string formattedTitle = listTitle.Replace(" ", SharePointSpaceCharacter);
            string itemType = $"SP.Data.{formattedTitle}Item";
            var type = new { type = itemType };
            var request = new
            {
                __metadata = type,
                FileLeafRef = fileName,
                Title = title
            };
            return request;
        }

        public async Task<bool> DeleteFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);

            bool result = false;
            // Delete is very similar to a GET.
            string serverRelativeUrl = $"{listTitle}/{folderName}";

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Delete,
                RequestUri = new Uri(_configuration.Resource, "_api/web/getFolderByServerRelativeUrl('/" + EscapeApostrophe(serverRelativeUrl) + "')"),
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            // We want to delete this folder.
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("X-HTTP-Method", "DELETE");

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                result = true;
            }
            else
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                throw ex;
            }

            return result;
        }

        public async Task<bool> FolderExists(string listTitle, string folderName)
        {
            Object folder = await GetFolder(listTitle, folderName);

            return (folder != null);
        }

        public async Task<bool> DocumentLibraryExists(string listTitle)
        {
            Object lisbrary = await GetDocumentLibrary(listTitle);

            return (lisbrary != null);
        }

        public async Task<Object> GetFolder(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);

            Object result = null;
            string serverRelativeUrl = $"{listTitle}/{folderName}";

            using var endpointRequest = new HttpRequestMessage()
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, "_api/web/getFolderByServerRelativeUrl('/" + EscapeApostrophe(serverRelativeUrl) + "')"),
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(endpointRequest);
            string jsonString = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK || response.IsSuccessStatusCode)
            {
                result = JsonConvert.DeserializeObject(jsonString);
            }

            return result;
        }

        public async Task<Object> GetDocumentLibrary(string listTitle)
        {
            Object result = null;
            string title = Uri.EscapeUriString(listTitle);

            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/lists/GetByTitle('{title}')"),
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            string jsonString = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK)
            {
                result = JsonConvert.DeserializeObject(jsonString);
            }

            return result;
        }

        public async Task<string> AddFile(String folderName, String fileName, Stream fileData, string contentType)
        {
            return await this.AddFile(DefaultDocumentListTitle, folderName, fileName, fileData, contentType);
        }

        public async Task<string> AddFile(String documentLibrary, String folderName, String fileName, Stream fileData, string contentType)
        {
            folderName = FixFoldername(folderName);
            bool folderExists = await this.FolderExists(documentLibrary, folderName);
            if (!folderExists)
            {
                await this.CreateFolder(documentLibrary, folderName);
            }

            // now add the file to the folder.

            fileName = await this.UploadFile(fileName, documentLibrary, folderName, fileData, contentType);

            return fileName;

        }

        public async Task<string> AddFile(String folderName, String fileName, byte[] fileData, string contentType)
        {
            return await this.AddFile(DefaultDocumentListTitle, folderName, fileName, fileData, contentType);
        }

        public async Task<string> AddFile(String documentLibrary, String folderName, String fileName, byte[] fileData, string contentType)
        {
            folderName = FixFoldername(folderName);
            bool folderExists = await this.FolderExists(documentLibrary, folderName);
            if (!folderExists)
            {
                await this.CreateFolder(documentLibrary, folderName);
            }

            // now add the file to the folder.

            fileName = await this.UploadFile(fileName, documentLibrary, folderName, fileData, contentType);

            return fileName;

        }

        private string GetServerRelativeURL(string listTitle, string folderName)
        {
            folderName = FixFoldername(folderName);
            string serverRelativeUrl = Uri.EscapeUriString(listTitle) + "/" + Uri.EscapeUriString(folderName);
            return serverRelativeUrl;
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

        /// <summary>
        /// Upload a file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="listTitle"></param>
        /// <param name="folderName"></param>
        /// <param name="fileData"></param>
        /// <param name="contentType"></param>
        /// <returns>Uploaded Filename, or Null if not successful.</returns>
        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, Stream fileData, string contentType)
        {
            // convert the stream into a byte array.
            MemoryStream ms = new MemoryStream();
            fileData.CopyTo(ms);
            byte[] data = ms.ToArray();
            string result = await UploadFile(fileName, listTitle, folderName, data, contentType);
            return result;
        }

        /// <summary>
        /// SharePoint is very particular about the file name length and the total characters in the URL to access a file.
        /// This method returns the input file name or a truncated version of the file name if it is over the max number of characters.
        /// </summary>
        /// <param name="fileName">The file name to check; e.g. "abcdefg1111222233334444.pdf"</param>
        /// <param name="listTitle">The list title</param>
        /// <param name="folderName">The folder name where the file would be uploaded</param>
        /// <returns>The (potentially truncated) file name; e.g. "abcd.pdf"</returns>
        public string GetTruncatedFileName(string fileName, string listTitle, string folderName)
        {
            // SharePoint requires that filenames are less than 128 characters.
            int maxLength = 128;
            fileName = FixFilename(fileName, maxLength);

            folderName = FixFoldername(folderName);

            // SharePoint also imposes a limit on the whole URL
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

        /// <summary>
        /// Upload a file
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="listTitle"></param>
        /// <param name="folderName"></param>
        /// <param name="fileData"></param>
        /// <param name="contentType"></param>
        /// <returns>Uploaded Filename, or Null if not successful.</returns>
        public async Task<string> UploadFile(string fileName, string listTitle, string folderName, byte[] data, string contentType)
        {
            string result = null;
            folderName = FixFoldername(folderName);
            fileName = GetTruncatedFileName(fileName, listTitle, folderName);

            string serverRelativeUrl = GetServerRelativeURL(listTitle, folderName);
            string requestUriString = GenerateUploadRequestUriString(serverRelativeUrl, fileName);

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(requestUriString),
                Headers = {
                        { "Accept", "application/json" }
                    }
            };

            ByteArrayContent byteArrayContent = new ByteArrayContent(data);
            byteArrayContent.Headers.Add(@"content-length", data.Length.ToString());
            request.Content = byteArrayContent;

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            var streamData = response.Content.ReadAsStringAsync().Result;

            var listItemData = JsonConvert.DeserializeObject<AddFileResponse>(streamData);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                return fileName;
            }
            else
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                request.Dispose();
                throw ex;
            }
        }

        /// <summary>
        /// Upload a file
        /// </summary>
        /// <param name="itemData"></param>
        /// <param name="listTitle"></param>
        /// <param name="contentType"></param>
        /// <param name="fileName"></param>
        /// <returns>Uploaded Filename, or Null if not successful.</returns>
        public async Task<string> UpdateListItemFields(AddFileResponse itemData, string listTitle, string contentType, string fileName)
        {
            string result = null;
            string requestUriString = GenerateUpdateListItemUriString(listTitle, itemData.ListItemAllFields.ID.ToString());

            using var request = new HttpRequestMessage(HttpMethod.Post, requestUriString);

            var listItem = CreateUpdateListItemRequestRequest(contentType, fileName, listTitle);

            string jsonString = JsonConvert.SerializeObject(listItem);
            StringContent strContent = new StringContent(jsonString, Encoding.UTF8);
            strContent.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=verbose");

            request.Headers.Add("Accept", "application/json;odata=verbose");
            //endpointRequest.Headers.Add("Accept", "*/*");
            request.Headers.Add("X-Http-Method", "MERGE");
            //endpointRequest.Headers.Add("X-Requested-With", "XMLHttpRequest");
            request.Headers.Add("IF-MATCH", "*");
            //endpointRequest.Headers.Add("IF-MATCH", itemData.ETag);
            request.Headers.Add("odata-version", "3.0");
            request.Content = strContent;

            //ByteArrayContent byteArrayContent = new ByteArrayContent(data);
            //byteArrayContent.Headers.Add(@"content-length", data.Length.ToString());
            //endpointRequest.Content = byteArrayContent;

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);
            var streamData = await response.Content.ReadAsStringAsync();

            if (response.StatusCode == HttpStatusCode.OK || response.StatusCode == HttpStatusCode.NoContent)
            {
                result = streamData;
            }
            else
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                throw ex;
            }

            return result;
        }

        /// <summary>
        /// Download a file
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<byte[]> DownloadFile(string url)
        {
            byte[] result = null;

            url = EscapeApostrophe(url);

            HttpRequestMessage endpointRequest = new HttpRequestMessage
            {
                //The URL is expected to begin with a slash.
                Method = HttpMethod.Get,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('{url}')/$value"),
            };

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(endpointRequest);

            using MemoryStream ms = new MemoryStream();
            await response.Content.CopyToAsync(ms);
            result = ms.ToArray();

            return result;
        }

        private async Task<string> GetDigest(HttpClient client)
        {
            string result = null;

            HttpRequestMessage endpointRequest = new HttpRequestMessage()
            {
                Method = HttpMethod.Post,
                RequestUri = new Uri(_configuration.Resource, "_api/contextinfo"),
                Headers = {
                    { "Accept", "application/json;odata=verbose" }
                }
            };

            // make the request.
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

        /// <summary>
        /// Delete a file
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<bool> DeleteFile(string listTitle, string folderName, string fileName)
        {
            // Delete is very similar to a GET.
            folderName = FixFoldername(folderName);
            string serverRelativeUrl = $"/{listTitle}/{folderName}/{fileName}";

            bool result = await DeleteFile(serverRelativeUrl);
            return result;
        }

        public async Task<bool> DeleteFile(string serverRelativeUrl)
        {
            bool result = false;
            // Delete is very similar to a GET.

            serverRelativeUrl = EscapeApostrophe(serverRelativeUrl);

            using var request = new HttpRequestMessage
            {
                Method = HttpMethod.Delete,
                RequestUri = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('/{serverRelativeUrl}')"),
                Headers = {
                    { "Accept", "application/json" }
                }
            };

            // We want to delete this file.
            request.Headers.Add("IF-MATCH", "*");
            request.Headers.Add("X-HTTP-Method", "DELETE");

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                result = true;
            }
            else
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);

                throw ex;
            }

            return result;
        }

        /// <summary>
        /// Rename a file.  Note that this only works for files with relatively short names due to the max URL length.  It may be possible to allow that to work by using @variables in the URL.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<bool> RenameFile(string oldServerRelativeUrl, string newServerRelativeUrl)
        {
            bool result = false;

            oldServerRelativeUrl = EscapeApostrophe(oldServerRelativeUrl);
            newServerRelativeUrl = EscapeApostrophe(newServerRelativeUrl);

            Uri url = new Uri(_configuration.Resource, $"_api/web/GetFileByServerRelativeUrl('{oldServerRelativeUrl}')/moveto(newurl='{newServerRelativeUrl}', flags=1)");
            using var request = new HttpRequestMessage(HttpMethod.Post, url);

            // make the request.
            using var httpClient = await GetHttpClientAsync();
            using var response = await httpClient.SendAsync(request);

            if (response.StatusCode == HttpStatusCode.OK)
            {
                result = true;
            }
            else
            {
                string responseContent = null;
                var ex = new SharePointRestException($"Operation returned an invalid status code '{response.StatusCode}'");
                responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                ex.Request = new HttpRequestMessageWrapper(request, null);
                ex.Response = new HttpResponseMessageWrapper(response, responseContent);
                throw ex;
            }

            return result;
        }

        private async Task<HttpClient> GetHttpClientAsync()
        {
            var cookieContainer = new CookieContainer();
#pragma warning disable CA2000 // Dispose objects before losing scope
            HttpMessageHandler handler = new SocketsHttpHandler
            {
                UseCookies = true,
                AllowAutoRedirect = false,
                CookieContainer = cookieContainer,
                MaxConnectionsPerServer = 25
            };

            var apiGatewayHost = _configuration.ApiGatewayHost;
            var apiGatewayPolicy = _configuration.ApiGatewayPolicy;
            var resource = _configuration.Resource;
            var authorizationUri = _configuration.AuthorizationUri;
            var relyingPartyIdentifier = _configuration.RelyingPartyIdentifier;
            var username = _configuration.Username;
            var password = _configuration.Password;

            if (!string.IsNullOrEmpty(apiGatewayHost) && !string.IsNullOrEmpty(apiGatewayPolicy))
            {
                // since this is executed on every access to sharepoint, only log at debug level
                _logger.LogDebug("Using {@ApiGateway} for {Resource}",
                    new { Host = apiGatewayHost, Policy = apiGatewayPolicy }, resource);

                handler = new ApiGatewayHandler(handler, apiGatewayHost, apiGatewayPolicy);
            }
#pragma warning restore CA2000 // Dispose objects before losing scope

            HttpClient httpClient = new HttpClient(handler);
            httpClient.BaseAddress = resource;
            httpClient.DefaultRequestHeaders.Accept.Add(MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose"));

            // simplify the parameters
            var authorizationUrl = authorizationUri.ToString();

            string samlToken = await _samlAuthenticator.GetStsSamlTokenAsync(relyingPartyIdentifier, username, password, authorizationUrl);

            await _samlAuthenticator.GetSharepointFedAuthCookieAsync(resource, samlToken, httpClient, cookieContainer, apiGatewayHost, apiGatewayPolicy);

            var digest = await GetDigest(httpClient);
            if (digest is not null)
            {
                httpClient.DefaultRequestHeaders.Add("X-RequestDigest", digest);
            }

            // Standard headers for API access
            httpClient.DefaultRequestHeaders.Add("Cache-Control", "no-cache");
            httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");

            return httpClient;
        }
    }

    class DocumentLibraryResponse
    {
        public DocumentLibraryResponseContent d { get; set; }
    }

    class DocumentLibraryResponseContent
    {
        public string Id { get; set; }
    }

    public class AddFileResponse
    {
        public string ETag { get; set; }
        public ListItemAllFieldsObj ListItemAllFields { get; set; }
    }

    public class ListItemAllFieldsObj
    {
        public int ID { get; set; }
    }

    public class ContextInfoResponse
    {
        public string FormDigestValue { get; set; }
    }

    public class FileSystemItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Documenttype { get; set; }
        public int Size { get; set; }
        public string Serverrelativeurl { get; set; }
        public DateTime Timecreated { get; set; }
        public DateTime Timelastmodified { get; set; }
    }


    public class FileDetailsList
    {
        public string Name { get; set; }
        public string TimeLastModified { get; set; }
        public string TimeCreated { get; set; }
        public string Length { get; set; }
        public string DocumentType { get; set; }
        public string ServerRelativeUrl { get; set; }
    }

}
