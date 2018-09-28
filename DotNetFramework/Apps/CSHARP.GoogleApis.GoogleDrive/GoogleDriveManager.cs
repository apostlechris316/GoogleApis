/********************************************************************************
 * CSHARP Google Drive Library - This library is meant to be an example to help 
 * you when working with Google Drive. 
 * 
 * LICENSE: 
 *          Free to use provided details on fixes and/or extensions emailed to 
 *          chris.williams@readwatchcreate.com
 *          
 * 
 * DEPENDENCIES: This library depends on the CSHARP.GoogleApis.Common 
 *          
 ********************************************************************************/

using CSHARP.GoogleApis.Common;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;

using System;
using System.Collections.Generic;
using System.IO;

namespace CSHARP.GoogleApis.GoogleDrive
{
    /// <summary>
    /// Manages Accessing Google Drive
    /// </summary>
    public class GoogleDriveManager : BaseGoogleApiManager
    {
        #region Other Useful Mime Type

        public const string CsvTextMimeType = "text/csv";

        #endregion

        #region Sample Google API Mime Types 

        public const string GoogleDriveFolderMimeType = "application/vnd.google-apps.folder";
        public const string GoogleSheetMimeType = "application/vnd.google-apps.spreadsheet";

        #endregion

        #region Sample Scopes for Google Drive

        /// <summary>
        /// Allows access to Drive Service in Read-Only Mode
        /// </summary>
        public string[] ReadOnlyScope = { DriveService.Scope.DriveReadonly };

        /// <summary>
        /// Allows access to Drive Service with full access
        /// </summary>
        public string[] FullAccessScope = { DriveService.Scope.Drive };

        #endregion

        #region Sample Queries When Searching for Files and Folders

        // Google Drive Canned Search Queries for folders
        public const string SearchQueryFolders = "mimeType = '" + GoogleDriveFolderMimeType + "'";
        public const string SearchQueryFolderWithName = SearchQueryWithName + " and " + SearchQueryFolders;
        public const string SearchQueryChildFolders = SearchQueryFolders + " and " + SearchQueryChildren;
        public const string SearchQueryChildFolderWithName = SearchQueryChildFolders + " and '{1}' = name";

        // Google Drive Canned Search Queries for sheets
        public const string SearchQuerySheets = "mimeType = '" + GoogleSheetMimeType + "'";
        public const string SearchQuerySheetsWithName = SearchQueryWithName + " and " + SearchQuerySheets;
        public const string SearchQueryChildSheets = SearchQuerySheets + " and " + SearchQueryChildren;
        public const string SearchQueryChildSheetWithName = SearchQueryChildSheets + " and '{1}' = name";

        public const string SearchQueryChildren = "'{0}' in parents";
        public const string SearchQueryWithName = "'{0}' = name";

        #endregion

        /// <summary>
        /// Creates instance of the Google Drive Service
        /// </summary>
        /// <param name="userCredential">User Credential created by calling GetUserCredentials on this class</param>
        /// <param name="applicationName">The name of the appplication Google is expecting to call</param>
        /// <returns>Returns the DriveService instance upons success.</returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 2
        /// </remarks>
        public DriveService CreateDriveServiceInstance(GoogleCredential credential, string applicationName)
        {
            // if no credentials passed in don't even try to create the drive service
            if (credential == null) throw new ArgumentNullException("credentials are required in order to create the Drive Service");

            // if no application name is supplied return null
            if (string.IsNullOrEmpty(applicationName)) throw new ArgumentNullException("application name is required in order to create the Drive Service");

            // Create Drive API service.
            return new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }

        #region Create Folder Related

        /// <summary>
        /// Creates a Root folder and returns the folder id
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderName">Name of folder to create</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 4
        /// </remarks>
        public string CreateFolder(DriveService driveService, string folderName)
        {
            // if drive service not passed in we cannot create the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to create folder");

            // if no folder name is supplied we cannot create the folder
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException("folder name is required in order to create the Drive Service");

            var folderFile = new Google.Apis.Drive.v3.Data.File
            {
                Name = folderName,
                MimeType = GoogleDriveFolderMimeType
            };
            var result = driveService.Files.Create(folderFile).Execute();
            return result.Id;
        }

        /// <summary>
        /// Creates a folder under the parent folder(s) provided
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderName"></param>
        /// <param name="parents"></param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 5
        /// </remarks>
        public string CreateFolder(DriveService driveService, string folderName, List<string> parents)
        {
            // if drive service not passed in we cannot create the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to create folder");

            // if no folder name is supplied we cannot create the folder
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException("folder name is required in order to create the Drive Service");

            var folderFile = new Google.Apis.Drive.v3.Data.File
            {
                Name = folderName,
                MimeType = GoogleDriveFolderMimeType,
                Parents = parents
            };
            var result = driveService.Files.Create(folderFile).Execute();
            return result.Id;
        }

        #endregion

        #region Delete Related

        /// <summary>
        /// Empties the trash for the Google Drive
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public bool EmptyTrash(DriveService driveService)
        {
            // if drive service not passed in we cannot create the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to create folder");

            try
            {
                driveService.Files.EmptyTrash().Execute();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// Permanently deletes a file.
        /// WARNING: Doe NOT put in the trash.
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="fileId">Id of the File to delete</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 4
        /// </remarks>
        public bool DeleteFile(DriveService driveService, String fileId)
        {
            // if drive service not passed in we cannot delete the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to delete file");

            // if no file id is supplied we cannot delete the folder
            if (string.IsNullOrEmpty(fileId)) throw new ArgumentNullException("file id is required in order to delete file");

            try
            {
                driveService.Files.Delete(fileId).Execute();
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        #endregion

        #region Move File/Folder Related

        /// <summary>
        /// Moves a file from one folder on the drive to another
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="fileId">Id of the file to move</param>
        /// <param name="folderId">Id of folder to move file to</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File MoveFile(DriveService driveService, string fileId, string folderId)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to move file");

            // if no file id is supplied we cannot move the file
            if (string.IsNullOrEmpty(fileId)) throw new ArgumentNullException("file id is required in order to move the file");

            // if no folder id is supplied we cannot move the file
            if (string.IsNullOrEmpty(folderId)) throw new ArgumentNullException("folder id is required in order to move the file");

            // Retrieve the existing parents to remove
            var getRequest = driveService.Files.Get(fileId);
            getRequest.Fields = "parents";
            var file = getRequest.Execute();
            var previousParents = String.Join(",", file.Parents);

            // Move the file to the new folder
            var updateRequest = driveService.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = folderId;
            updateRequest.RemoveParents = previousParents;
            return updateRequest.Execute();
        }

        #endregion

        #region Permissions

        /// <summary>
        /// Creates a user permission object to be used to grant on files
        /// </summary>
        /// <param name="role">Type of access a person has eg. reader, writer</param>
        /// <param name="emailAddress">Email address of user to be granted permission</param>
        /// <returns></returns>
        public Permission CreateUserPermission(string role, string emailAddress)
        {
            // if no email addres is supplied we cannot create permission object
            if (string.IsNullOrEmpty(emailAddress)) throw new ArgumentNullException("email address is required in order to create permission object");

            // if no role is supplied we cannot create permission object
            if (string.IsNullOrEmpty(role)) throw new ArgumentNullException("role is required in order to create permission object");

            return new Permission() { Type = "user", Role = role, EmailAddress = emailAddress };
        }

        /// <summary>
        /// Creates a domain permission object to be used to grant on files.
        /// </summary>
        /// <param name="role">Type of access a person has eg. reader, writer</param>
        /// <param name="domain">Domain portaion of email address to grant permissions to. Eg. sitecore.net, readwatchcreate.com</param>
        /// <returns></returns>
        public Permission CreateDomainPermission(string role, string domain)
        {
            // if no domain is supplied we cannot create permission object
            if (string.IsNullOrEmpty(domain)) throw new ArgumentNullException("domain is required in order to create permission object");

            // if no role is supplied we cannot create permission object
            if (string.IsNullOrEmpty(role)) throw new ArgumentNullException("role is required in order to create permission object");

            return new Permission() { Type = "domain", Role = role, Domain = domain };
        }

        /// <summary>
        /// Assigns a permission to a file
        /// </summary>
        /// <param name="driveService"></param>
        /// <param name="fileId"></param>
        /// <param name="permission"></param>
        /// <returns></returns>
        public bool AssignPermissionToFile(DriveService driveService, string fileId, Permission permission)
        {
            // if drive service not passed in we cannot create the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to set permission");

            if (permission == null) throw new ArgumentNullException("Permisssion is required in order to set the permission");

            // if no file Id is supplied we cannot set permission
            if (string.IsNullOrEmpty(fileId)) throw new ArgumentNullException("fileId is required in order to set permission");

            var request = driveService.Permissions.Create(permission, fileId);
            request.Fields = "id";
            if (request.Execute() == null) return false;
            return true;
        }

        #endregion

        #region Get Folder(s) Related 

        /// <summary>
        /// Gets the folder given the id
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderId">Unique Id representing the folder</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetFolderById(DriveService driveService, string folderId)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get folder");

            // if no folder id is supplied we cannot return the folder
            if (string.IsNullOrEmpty(folderId)) throw new ArgumentNullException("folder id is required in order to get child folder");

            return driveService.Files.Get(folderId).Execute();
        }

        /// <summary>
        /// Gets the folder underneath the root folder given its name
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderName">Name of folder to get</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 3
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetFolderByName(DriveService driveService, string folderName)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child folder");

            // if no folder name is supplied we cannot move the file
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException("folder name is required in order to get child folder");

            var fileList = GetFileListViaSearchQuery(driveService, string.Format(SearchQueryFolderWithName, folderName), false);

            // if file list is null it means something is not right with name query so resort to old school loop.
            if (fileList == null)
            {
                fileList = GetFileList(driveService, false);
                foreach (var file in fileList)
                {
                    // if the foldername matches and the file is of type folder then return
                    if (file.Name == folderName) if(file.MimeType == GoogleDriveFolderMimeType) return file; 
                }

                // file was really not found
                return null;
            }

            // if no file return null
            if (fileList.Count == 0 || fileList.Count > 1) return null;

            return fileList[0];
        }

        /// <summary>
        /// Gets the folder underneath a parent folder given its name
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderName">Name of folder to get</param>
        /// <param name="parentFolderId">Folder to look in</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 5
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetFolderByName(DriveService driveService, string folderName, string parentFolderId)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child folder");

            // if no folder name is supplied we cannot move the file
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException("folder name is required in order to get child folder");

            var fileList = GetFileListViaSearchQuery(driveService, string.Format(SearchQueryChildFolderWithName, parentFolderId, folderName), false);

            // if file list is null it means something is not right with name query so resort to old school loop.
            if (fileList == null)
            {
                fileList = GetFileList(driveService, false);
                foreach (var file in fileList)
                {
                    // if the foldername matches and the file is of type folder then return
                    if (file.Name == folderName) if (file.MimeType == GoogleDriveFolderMimeType) return file;
                }

                // file was really not found
                return null;
            }

            if (fileList.Count == 0 || fileList.Count > 1) return null;

            return fileList[0];
        }

        /// <summary>
        /// Get List of Root Folders from Google Drive. If you are looking for files then use GetFileList.
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="firstPageOnly">If true, will only get the first page of files (default is 10 files)</param>
        /// <returns></returns>
        /// <remarks>
        /// (IMPORTANT: If you get rid of Google.Apis.Drive.v3.Data then it will default to System.IO.File which is not what you want)
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 2
        /// </remarks>
        public IList<Google.Apis.Drive.v3.Data.File> GetFolderList(DriveService driveService, bool firstPageOnly)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get folder list");

            return GetFileListViaSearchQuery(driveService, SearchQueryFolders, false);
        }

        #endregion

        #region File List Related 

        /// <summary>
        /// Gets a file underneath the root folder given its name
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="fileName">Name of file to get</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 3
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetFileByName(DriveService driveService, string fileName)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child item");

            // if no file name is supplied we cannot move the file
            if (string.IsNullOrEmpty(fileName)) throw new ArgumentNullException("item name is required in order to get child file");

            var fileList = GetFileListViaSearchQuery(driveService, string.Format(SearchQueryWithName, fileName), false);

            // if file list is null it means something is not right with name query so resort to old school loop.
            if (fileList == null)
            {
                fileList = GetFileList(driveService, false);
                foreach(var file in fileList)
                {
                    if(file.Name == fileName)
                    {
                        return file;
                    }
                }
            }

            // if no file return null
            if (fileList.Count == 0 || fileList.Count > 1) return null;

            return fileList[0];
        }

        /// <summary>
        /// Get List of Files in the root of the Google Drive
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="firstPageOnly">If true, will only get the first page of files (default is 10 files)</param>
        /// <returns></returns>
        /// <remarks>
        /// (IMPORTANT: If you get rid of Google.Apis.Drive.v3.Data then it will default to System.IO.File which is not what you want)
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public IList<Google.Apis.Drive.v3.Data.File> GetFileList(DriveService driveService, bool firstPageOnly)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child folder");

            return GetFileListViaSearchQuery(driveService, string.Empty, firstPageOnly);
        }

        /// <summary>
        /// Get the list of files/folders below a given folder
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="folderId">Id of folder to get files under</param>
        /// <param name="firstPageOnly">If true, will only get the first page of files (default is 10 files)</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public IList<Google.Apis.Drive.v3.Data.File> GetFileList(DriveService driveService, string folderId, bool firstPageOnly)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child folder");

            return GetFileListViaSearchQuery(driveService, string.Format(SearchQueryChildren, folderId), firstPageOnly);
        }


        /// <summary>
        /// Get the list of files/folders matching a specific search query
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="query">Search query to execute</param>
        /// <param name="firstPageOnly">If true, will only get the first page of files (default is 10 files)</param>
        /// <returns>List of files matching criteria or empty list. If exception thrown will return null list.</returns>
        /// <remarks>
        /// UNIT TESTING: Indirectly Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public IList<Google.Apis.Drive.v3.Data.File> GetFileListViaSearchQuery(DriveService driveService, string query, bool firstPageOnly)
        {
            // if drive service not passed in we cannot move the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child folder");

            var fileList = new List<Google.Apis.Drive.v3.Data.File>();

            // Define parameters of request.
            FilesResource.ListRequest request = driveService.Files.List();
            request.PageSize = 10;
            request.Q = query;

            do
            {
                try
                {
                    var result = request.Execute();
                    fileList.AddRange(result.Files);

                    // If we are only getting the first page then don't get anymore otherwise get full file list.
                    request.PageToken = (firstPageOnly) ? null : result.NextPageToken;
                }
                catch (Exception e)
                {
                    request.PageToken = null;
                    fileList = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return fileList;
        }


        #endregion

        #region GoogleSheet Related

        /// <summary>
        /// Creates a Google Sheet in Google Drive in the given parent folder and returns the id
        /// </summary>
        /// <param name="driveService"></param>
        /// <param name="sheetName">Name of the Google Sheet</param>
        /// <param name="sheetDescription">Brief description of the Google Sheet</param>
        /// <param name="parents">List of parent ids (folders) containing the sheet</param>
        /// <returns></returns>
        public string CreateGoogleSheet(DriveService driveService, string sheetName, string sheetDescription, List<string> parents)
        {
            // if drive service not passed in we cannot create the folder
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to create sheet");

            // if no sheet name is supplied we cannot create the sheet
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException("sheet name is required in order to create sheet");

            var sheetFile = new Google.Apis.Drive.v3.Data.File
            {
                Name = sheetName,
                Description = sheetDescription,
                MimeType = GoogleSheetMimeType,
                Parents = parents
            };
            var result = driveService.Files.Create(sheetFile).Execute();
            return result.Id;
        }

        /// <summary>
        /// Gets the sheet given the id
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// TO DO: We should check if file returned is actually a Sheet
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetSheetById(DriveService driveService, string sheetId)
        {
            // if drive service not passed in we cannot get the sheet
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to create sheet");

            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheetId is required in order to look up the sheet.");

            return driveService.Files.Get(sheetId).Execute();
        }

        /// <summary>
        /// Gets the file of type Google Sheet underneath a parent folder given its name
        /// </summary>
        /// <param name="driveService">The drive service created using CreateDriveServiceInstance</param>
        /// <param name="sheetName">Name of sheet to get</param>
        /// <param name="parentFolderId">Folder to look in</param>
        /// <returns></returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleSheetViaConsole - STEP 1
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File GetSheetFileInChildFolderByName(DriveService driveService, string sheetName, string parentFolderId)
        {
            // if drive service not passed in we cannot get the file
            if (driveService == null) throw new ArgumentNullException("Drive service instance is required in order to get child sheet");

            // if no sheet name is supplied we cannot get the file
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException("sheet name is required in order to get child sheet");

            var fileList = GetFileListViaSearchQuery(driveService, string.Format(SearchQueryChildSheetWithName, parentFolderId, sheetName), false);

            // if file list is null it means something is not right with name query so resort to old school loop.
            if (fileList == null)
            {
                fileList = GetFileList(driveService, false);
                foreach (var file in fileList)
                {
                    // if the foldername matches and the file is of type sheet then return
                    if (file.Name == sheetName) if (file.MimeType == GoogleSheetMimeType) return file;
                }

                // file was really not found
                return null;
            }

            if (fileList.Count == 0 || fileList.Count > 1) return null;

            return fileList[0];
        }

        #endregion

        #region Upload File Related

        /// <summary>
        /// Upload CSV Data as GoogleSheet
        /// </summary>
        /// <param name="service"></param>
        /// <param name="sheetName">Name of sheet to create on Google Drive</param>
        /// <param name="sheetDescription"></param>
        /// <param name="contentByteArray">Byte Array containing the CSV data</param>
        /// <param name="parents"></param>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File UploadCsvAsGoogleSheet(DriveService driveService, string sheetName, string sheetDescription, byte [] contentByteArray, List<string> parents)
        {
            return UploadFile(driveService, sheetName, sheetDescription, GoogleSheetMimeType, CsvTextMimeType, parents);
        }

        /// <summary>
        /// Uploads a file
        /// Adapted from: https://developers.google.com/drive/v2/reference/files/insert
        /// </summary>
        /// <param name="service">a Valid authenticated DriveService</param>
        /// <param name="fileName">path to the file to upload</param>
        /// <param name="parents">Collection of parent folders which contain this file. 
        ///                       Setting this field will put the file in all of the provided folders. root folder.</param>
        /// <returns>If upload succeeded returns the File resource of the uploaded file 
        ///          If the upload fails returns null</returns>
        /// <remarks>
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File UploadFile(DriveService driveService, string fileName, string fileDescription, string mimeType, string contentType, List<string> parents)
        {
            if(System.IO.File.Exists(fileName))
            { 
                // File's content.
                byte[] contentByteArray = System.IO.File.ReadAllBytes(fileName);
                return UploadFile(driveService, fileName, fileDescription, mimeType, contentType, contentByteArray, parents);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Uploads a file
        /// Adapted from: https://developers.google.com/drive/v2/reference/files/insert
        /// </summary>
        /// <param name="service"></param>
        /// <param name="fileName"></param>
        /// <param name="fileDescription"></param>
        /// <param name="driveFileMimeType">The type of file that will exist on Google Drive (eg. "application/vnd.google-apps.spreadsheet" for a GoogleSheet</param>
        /// <param name="contentByteArray"></param>
        /// <param name="parents"></param>
        /// <returns></returns>
        /// <remarks>Look at UploadCsvAsGoogleSheet To see an example on when mimeType and contentType differ. In most cases it will be the same.
        /// UNIT TESTING: NOT Tested via UnitTestGoogleDriveViaConsole
        /// </remarks>
        public Google.Apis.Drive.v3.Data.File UploadFile(DriveService service, string fileName, string fileDescription, string mimeType, string contentType, byte[] contentByteArray, List<string> parents)
        {
            var body = new Google.Apis.Drive.v3.Data.File
            {
                Name = fileName,
                Description = fileDescription,
                MimeType = mimeType,
                Parents = parents
            };

            MemoryStream stream = new MemoryStream(contentByteArray);
            try
            {
                FilesResource.CreateMediaUpload request = service.Files.Create(body, stream, mimeType);
                request.Upload();
                return request.ResponseBody;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        #endregion
    }
}