using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Csrs.Interfaces
{
    public interface ISharePointFileManager
    {
        Task<List<FileDetailsList>> GetFileDetailsListInFolder(string listTitle, string folderName, string documentType);

        Task CreateFolder(string listTitle, string folderName);

        Task<object> CreateDocumentLibrary(string listTitle, string documentTemplateUrlTitle = null);

        Task<object> UpdateDocumentLibrary(string listTitle);

        Task<bool> DeleteFolder(string listTitle, string folderName);

        Task<bool> FolderExists(string listTitle, string folderName);

        Task<bool> DocumentLibraryExists(string listTitle);

        Task<object> GetFolder(string listTitle, string folderName);

        Task<object> GetDocumentLibrary(string listTitle);

        Task<string> AddFile(string folderName, string fileName, Stream fileData, string contentType);

        Task<string> AddFile(string documentLibrary, string folderName, string fileName, Stream fileData, string contentType);

        Task<string> AddFile(string folderName, string fileName, byte[] fileData, string contentType);

        Task<string> AddFile(string documentLibrary, string folderName, string fileName, byte[] fileData, string contentType);

        Task<string> UploadFile(string fileName, string listTitle, string folderName, Stream fileData, string contentType);

        string GetTruncatedFileName(string fileName, string listTitle, string folderName);

        Task<string> UploadFile(string fileName, string listTitle, string folderName, byte[] data, string contentType);

        Task<string> UpdateListItemFields(AddFileResponse itemData, string listTitle, string contentType, string fileName);

        Task<byte[]> DownloadFile(string url);

        Task<bool> DeleteFile(string listTitle, string folderName, string fileName);

        Task<bool> DeleteFile(string serverRelativeUrl);

        Task<bool> RenameFile(string oldServerRelativeUrl, string newServerRelativeUrl);
    }
}
