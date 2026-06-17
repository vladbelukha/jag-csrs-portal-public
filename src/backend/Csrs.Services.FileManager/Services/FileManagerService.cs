using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Google.Protobuf;
using Google.Protobuf.WellKnownTypes;
using Csrs.Interfaces;
using Grpc.Core;
using Microsoft.Extensions.Logging;

namespace Csrs.Services.FileManager
{
    public class FileManagerService : FileManager.FileManagerBase
    {
        private readonly ISharePointFileManager _sharePointFileManager;
        private readonly ILogger<FileManagerService> _logger;

        public FileManagerService(ISharePointFileManager sharePointFileManager, ILogger<FileManagerService> logger)
        {
            _sharePointFileManager = sharePointFileManager ?? throw new ArgumentNullException(nameof(sharePointFileManager));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public override async Task<CreateFolderReply> CreateFolder(CreateFolderRequest request, ServerCallContext context)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(context);

            _logger.LogDebug("Create Folder");

            var logFolder = WordSanitizer.Sanitize(request.FolderName);

            var listTitle = GetDocumentListTitle(request.EntityName);
            var documentTemplateUrl = GetDocumentTemplateUrlPart(request.EntityName);

            await CreateDocumentLibraryIfMissing(listTitle, documentTemplateUrl);

            var folderExists = false;
            try
            {
                var folder = await _sharePointFileManager.GetFolder(listTitle, request.FolderName);
                if (folder != null) folderExists = true;
            }
            catch (SharePointRestException ex)
            {
                _logger.LogError(ex, "SharePointRestException creating sharepoint folder (status code: {StatusCode})", ex.Response.StatusCode);
                folderExists = false;
            }
            catch (Exception e)
            {
                _logger.LogError(e, "Generic Exception creating sharepoint folder");
                folderExists = false;
            }

            var result = new CreateFolderReply();

            if (folderExists)
            {
                result.ResultStatus = ResultStatus.Success;
            }
            else
            {
                try
                {
                    await _sharePointFileManager.CreateFolder(GetDocumentListTitle(request.EntityName), request.FolderName);
                    var folder = await _sharePointFileManager.GetFolder(listTitle, request.FolderName);
                    if (folder != null) result.ResultStatus = ResultStatus.Success;
                }
                catch (SharePointRestException ex)
                {
                    result.ResultStatus = ResultStatus.Fail;
                    result.ErrorDetail = $"ERROR in creating folder {logFolder}";
                    _logger.LogError(ex, "ERROR in creating {Folder}", logFolder);
                }
                catch (Exception e)
                {
                    result.ResultStatus = ResultStatus.Fail;
                    result.ErrorDetail = $"ERROR in creating folder {logFolder}";
                    _logger.LogError(e, "ERROR in creating {Folder}", logFolder);
                }
            }

            return result;
        }

        public override async Task<FileExistsReply> FileExists(FileExistsRequest request, ServerCallContext context)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(context);

            _logger.LogDebug("File Exists");

            var result = new FileExistsReply();

            List<FileDetailsList> fileDetailsList = null;
            try
            {
                var documentTemplateUrl = GetDocumentTemplateUrlPart(request.EntityName);
                fileDetailsList = await _sharePointFileManager.GetFileDetailsListInFolder(documentTemplateUrl, request.FolderName, request.DocumentType);
                if (fileDetailsList != null)
                {
                    var hasFile = fileDetailsList.Any(f => f.ServerRelativeUrl == request.ServerRelativeUrl);
                    result.ResultStatus = hasFile ? FileExistStatus.Exist : FileExistStatus.NotExist;
                }
            }
            catch (SharePointRestException spre)
            {
                _logger.LogError(spre, "Error determining if file exists");
                result.ResultStatus = result.ResultStatus = FileExistStatus.Error;
                result.ErrorDetail = "Error determining if file exists";
            }
            catch (Exception e)
            {
                result.ResultStatus = FileExistStatus.Error;
                result.ErrorDetail = "Error determining if file exists";
                _logger.LogError(e, "Error determining if file exists");
            }

            return result;
        }

        public override async Task<DeleteFileReply> DeleteFile(DeleteFileRequest request, ServerCallContext context)
        {
            var result = new DeleteFileReply();

            var logUrl = WordSanitizer.Sanitize(request.ServerRelativeUrl);

            try
            {
                var success = await _sharePointFileManager.DeleteFile(request.ServerRelativeUrl);
                result.ResultStatus = success ? ResultStatus.Success : ResultStatus.Fail;
            }
            catch (SharePointRestException ex)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in deleting file {logUrl}";
                _logger.LogError(ex, result.ErrorDetail);
            }
            catch (Exception e)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in deleting file {logUrl}";
                _logger.LogError(e, result.ErrorDetail);
            }

            return result;
        }

        public override async Task<DownloadFileReply> DownloadFile(DownloadFileRequest request, ServerCallContext context)
        {
            _logger.LogDebug("Download file");

            var result = new DownloadFileReply();
            var logUrl = WordSanitizer.Sanitize(request.ServerRelativeUrl);

            try
            {
                var data = await _sharePointFileManager.DownloadFile(request.ServerRelativeUrl);

                if (data != null)
                {
                    result.ResultStatus = ResultStatus.Success;
                    result.Data = ByteString.CopyFrom(data);
                }
                else
                {
                    result.ResultStatus = ResultStatus.Fail;
                }
            }
            catch (SharePointRestException ex)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in downloading file {logUrl}";
                _logger.LogError(ex, result.ErrorDetail);
            }
            catch (Exception e)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in downloading file {logUrl}";
                _logger.LogError(e, result.ErrorDetail);
            }

            return result;
        }

        public override async Task<UploadFileReply> UploadFile(UploadFileRequest request, ServerCallContext context)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(context);

            _logger.LogDebug("Upload file");

            var result = new UploadFileReply();
            var logFileName = WordSanitizer.Sanitize(request.FileName);
            var logFolderName = WordSanitizer.Sanitize(request.FolderName);

            try
            {
                var listTitle = GetDocumentListTitle(request.EntityName);
                var documentTemplateUrlPart = GetDocumentTemplateUrlPart(request.EntityName);

                await CreateDocumentLibraryIfMissing(listTitle, documentTemplateUrlPart);

                var fileName = await _sharePointFileManager.AddFile(
                    GetDocumentTemplateUrlPart(request.EntityName),
                    request.FolderName,
                    request.FileName,
                    request.Data.ToByteArray(),
                    request.ContentType);

                result.FileName = fileName;
                result.ResultStatus = ResultStatus.Success;
            }
            catch (SharePointRestException ex)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in uploading file {logFileName} to folder {logFolderName}";
                _logger.LogError(ex, result.ErrorDetail);
            }
            catch (Exception e)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in uploading file {logFileName} to folder {logFolderName}";
                _logger.LogError(e, result.ErrorDetail);
            }

            return result;
        }

        public override async Task<FolderFilesReply> FolderFiles(FolderFilesRequest request, ServerCallContext context)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(context);

            _logger.LogDebug("Get Folder Files");

            var result = new FolderFilesReply();

            // Get the file details list in folder
            List<FileDetailsList> fileDetailsList = null;

            try
            {
                var d = GetDocumentTemplateUrlPart(request.EntityName);

                fileDetailsList = await _sharePointFileManager.GetFileDetailsListInFolder(d, request.FolderName, request.DocumentType);
                if (fileDetailsList != null)

                {
                    // gRPC ensures that the collection has space to accept new data; no need to call a constructor
                    foreach (var item in fileDetailsList)
                    {
                        // Sharepoint API responds with dates in UTC format
                        var utcFormat = DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal;
                        DateTime parsedCreateDate, parsedLastModified;
                        DateTime.TryParse(item.TimeCreated, CultureInfo.InvariantCulture, utcFormat,
                            out parsedCreateDate);
                        DateTime.TryParse(item.TimeLastModified, CultureInfo.InvariantCulture, utcFormat,
                            out parsedLastModified);

                        var newItem = new FileSystemItem
                        {
                            DocumentType = item.DocumentType,
                            Name = item.Name,
                            ServerRelativeUrl = item.ServerRelativeUrl,
                            Size = int.Parse(item.Length),
                            TimeCreated = Timestamp.FromDateTime(parsedCreateDate),
                            TimeLastModified = Timestamp.FromDateTime(parsedLastModified)
                        };

                        result.Files.Add(newItem);
                    }

                    result.ResultStatus = ResultStatus.Success;
                }
            }
            catch (SharePointRestException spre)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = "Error getting SharePoint File List";
                _logger.LogError(spre, "Error getting SharePoint File List");
            }

            return result;
        }

        public override Task<TokenReply> GetToken(TokenRequest request, ServerCallContext context)
        {
            var result = new TokenReply
            {
                ResultStatus = ResultStatus.Fail,
                ErrorDetail = "Authentication not required"
            };

            return Task.FromResult(result);
        }

        public override Task<TruncatedFilenameReply> GetTruncatedFilename(
            TruncatedFilenameRequest request,
            ServerCallContext context)
        {
            var result = new TruncatedFilenameReply();
            var logFileName = WordSanitizer.Sanitize(request.FileName);
            var logFolderName = WordSanitizer.Sanitize(request.FolderName);

            try
            {
                // Ask SharePoint whether this filename would be truncated upon upload
                var listTitle = GetDocumentListTitle(request.EntityName);
                var maybeTruncated = _sharePointFileManager.GetTruncatedFileName(request.FileName, listTitle, request.FolderName);
                result.FileName = maybeTruncated;
                result.ResultStatus = ResultStatus.Success;
            }
            catch (SharePointRestException ex)
            {
                result.ResultStatus = ResultStatus.Fail;
                result.ErrorDetail = $"ERROR in getting truncated filename {logFileName} for folder {logFolderName}";
                _logger.LogError(ex, result.ErrorDetail);
            }

            return Task.FromResult(result);
        }

        private async Task CreateDocumentLibraryIfMissing(string listTitle, string documentTemplateUrl = null)
        {
            var exists = await _sharePointFileManager.DocumentLibraryExists(listTitle);
            if (!exists)
            {
                await _sharePointFileManager.CreateDocumentLibrary(listTitle, documentTemplateUrl);
            }
        }

        private static string GetDocumentListTitle(string entityName)
        {
            string listTitle;
            switch (entityName.ToLower())
            {
                case "ssg_csrscommunicationmessage":
                    listTitle = "CSRS Communication Message";
                    break;
                case "ssg_csrsfile":
                    listTitle = "CSRS File";
                    break;
                case "application":
                    listTitle = SharePointFileManager.ApplicationDocumentListTitle;
                    break;
                case "contact":
                    listTitle = SharePointFileManager.ContactDocumentListTitle;
                    break;
                case "worker":
                    listTitle = SharePointFileManager.WorkerDocumentListTitle;
                    break;
                case "event":
                    listTitle = SharePointFileManager.EventDocumentListTitle;
                    break;
                case "federal_report":
                    listTitle = SharePointFileManager.FederalReportListTitle;
                    break;
                case "licence":
                    listTitle = SharePointFileManager.LicenceDocumentListTitle;
                    break;
                default:
                    listTitle = entityName;
                    break;
            }

            return listTitle;
        }

        private static string GetDocumentTemplateUrlPart(string entityName)
        {
            var listTitle = "";
            switch (entityName.ToLower())
            {
                case "ssg_csrscommunicationmessage":
                    listTitle = "ssg_csrscommunicationmessage";
                    break;
                case "ssg_csrsfile":
                    listTitle = "ssg_csrsfile";
                    break;
                case "application":
                    listTitle = "adoxio_application";
                    break;
                case "contact":
                    listTitle = SharePointFileManager.ContactDocumentListTitle;
                    break;
                case "worker":
                    listTitle = "adoxio_worker";
                    break;
                case "event":
                    listTitle = SharePointFileManager.EventDocumentListTitle;
                    break;
                case "federal_report":
                    listTitle = SharePointFileManager.FederalReportListTitle;
                    break;
                case "licence":
                    listTitle = SharePointFileManager.LicenceDocumentUrlTitle;
                    break;
                default:
                    listTitle = entityName;
                    break;
            }

            return listTitle;
        }

    }
}