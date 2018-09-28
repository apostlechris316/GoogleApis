using CSHARP.GoogleApis.GoogleDrive;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestGoogleDriveViaConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var googleDriveManager = new GoogleDriveManager();
            GoogleCredential credentialReadOnly = null;
            GoogleCredential credentialFullAccess = null;
            DriveService driveServiceReadOnly = null;
            DriveService driveServiceFullAccess = null;

            // STEP 0: Pre-test validation
            if (System.IO.File.Exists("client-secret.json") == false)
            {
                Console.WriteLine("******* ERROR: Please ensure you have downloaded your Google API .JSON file and copied into same folder as the .EXE *******");
                Console.ReadKey();
                return;
            }

            #region STEP 1: Get the User Credentials

            #region STEP 1.1: Negative Testing

            // STEP 1.1.1 - Test no credentials file and no scope
            try { var expectedFailure = googleDriveManager.GetUserCredential(string.Empty, null); }
            catch(Exception exception)
            {
                Console.WriteLine("******* STEP 1.1.1 - PASSED : EXPECTED FAILURE with no credentials file: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 1.1.2 - Test credentials file supplied but no scope
            try { var expectedFailure = googleDriveManager.GetUserCredential("client-secret.json", null); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 1.1.2 - PASSED : EXPECTED FAILURE with no scope: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 1.2: Positive Testing

            // STEP 1.2.1 - Test getting credential with ReadOnly Scope
            try
            {
                credentialReadOnly = googleDriveManager.GetUserCredential("client-secret.json", googleDriveManager.ReadOnlyScope);
                if (credentialReadOnly == null)
                {
                    Console.WriteLine("******* STEP 1.2.1 - FAILED - Credentials returned null when getting Read Only Scope permissions");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch(Exception exception)
            {
                Console.WriteLine("******* STEP 1.2.1 - FAILED - Exception thrown getting credentials with Read Only Scope");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // STEP 1.2.1 - Test getting credential with Full Access Scope
            try
            {
                credentialFullAccess = googleDriveManager.GetUserCredential("client-secret.json", googleDriveManager.FullAccessScope);
                if (credentialFullAccess == null)
                {
                    Console.WriteLine("******* STEP 1.2.2 - FAILED - Credentials returned null when getting Full Access Scope permissions");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 1.2.2 - FAILED - Exception thrown getting credentials with Full Access permissions");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("******* STEP 1.2 - PASSED - got both Read Only and Full Access Credentials");
            Console.WriteLine("*******");
            Console.ReadKey();

            #endregion

            #endregion

            #region STEP 2: Get Drive Service Instance

            #region STEP 2.1: Negative Testing

            // STEP 2.1.1 - Test no credentials passed in and no application name
            try { var expectedFailure = googleDriveManager.CreateDriveServiceInstance(null, string.Empty); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 2.1.1 - PASSED : EXPECTED FAILURE with no credentials passed: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 2.1.2 - Test credentials passed in but no application name
            try { var expectedFailure = googleDriveManager.CreateDriveServiceInstance(credentialReadOnly, string.Empty); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 2.1.2 - PASSED : EXPECTED FAILURE with no application name passed: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 2.2: Positive Testing

            // STEP 2.2.1 - Test getting drive service with ReadOnly Scope
            try
            {
                driveServiceReadOnly = googleDriveManager.CreateDriveServiceInstance(credentialReadOnly, "UnitTest Google Drive Library");
                if (driveServiceReadOnly == null)
                {
                    Console.WriteLine("******* STEP 2.2.1 - FAILED - drive service returned null getting it with Read Only credentials. ");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 2.2.1 - FAILED - Exception thrown getting drive service using Read Only credentials.");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // STEP 2.2.2 - Test getting drive service with Full Access Scope
            try
            {
                driveServiceFullAccess = googleDriveManager.CreateDriveServiceInstance(credentialFullAccess, "UnitTest Google Drive Library");
                if (driveServiceFullAccess == null)
                {
                    Console.WriteLine("******* STEP 2.2.2 - FAILED - drive service returned null getting it with Full Access credentials. ");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 2.2.2 - FAILED - Exception thrown getting drive service using Full Access credentials.");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("******* STEP 2.2 - PASSED");
            Console.WriteLine("*******");
            Console.ReadKey();

            #endregion

            #endregion

            #region STEP 3: Testing Drive access at the Root Folder level

            IList<Google.Apis.Drive.v3.Data.File> folders = null;

            #region PRE-STEP: Ensure there are at least 11 folders in the root

            // check if T1 folder exists and if not create it
            var t1Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T1");
            if (t1Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T1")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T1 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // Confirm the create worked
            t1Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T1");
            if (t1Folder == null)
            {
                Console.WriteLine("******* ERROR: Cannot getting test folder T1 in the Root Folder *******");
                Console.ReadKey();
                return;
            }

            // check if T2 folder exists and if not create it
            var t2Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T2");
            if (t2Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T2")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T2 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T3 folder exists and if not create it
            var t3Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T3");
            if (t3Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T3")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T3 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T4 folder exists and if not create it
            var t4Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T4");
            if (t4Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T4")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T4 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T5 folder exists and if not create it
            var t5Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T5");
            if (t5Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T5")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T5 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T6 folder exists and if not create it
            var t6Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T6");
            if (t6Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T6")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T6 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T7 folder exists and if not create it
            var t7Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T7");
            if (t7Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T7")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T7 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T8 folder exists and if not create it
            var t8Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T8");
            if (t8Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T8")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T8 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T9 folder exists and if not create it
            var t9Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T9");
            if (t9Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T9")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T9 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T7 folder exists and if not create it
            var t10Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T10");
            if (t10Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T10")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T10 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T11 folder exists and if not create it
            var t11Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T11");
            if (t11Folder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T11")))
                {
                    Console.WriteLine("******* ERROR: Cannot create test folder T11 in the Root Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            #endregion

            #region STEP 3.1: Negative Testing

            // STEP 3.1.1 - Test no drive service, first page only
            try { var expectedFailure = googleDriveManager.GetFolderList(null, true); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.1 - PASSED : EXPECTED FAILURE with no drive service passed while asking for first page only: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 3.1.2 - Test no drive service, all pages
            try { var expectedFailure = googleDriveManager.GetFolderList(null, false); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.2 - PASSED : EXPECTED FAILURE with no drive service passed while asking for all pages: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 3.2: Positive Testing

            // STEP 3.2.1 - Test drive service, first page only
            folders = googleDriveManager.GetFolderList(driveServiceReadOnly, true);
            if (folders == null)
            {
                Console.WriteLine("******* ERROR: Cannot get the Root Folders (first page only) *******");
                Console.ReadKey();
                return;
            }
            foreach (var folder in folders) { Console.WriteLine(folder.Name); }

            // STEP 3.2.2 - Test drive service, all pages
            folders = googleDriveManager.GetFolderList(driveServiceReadOnly, false);
            if (folders == null)
            {
                Console.WriteLine("******* ERROR: Cannot get the Root Folders (all pages) *******");
                Console.ReadKey();
                return;
            }
            foreach (var folder in folders) { Console.WriteLine(folder.Name); }

            #endregion

            #endregion

            #region STEP 4: Test Delete and/or Create Child Folder

            // Check if "test" folder exists in the root if it exists delete and re-create it
            var testFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "Test");
            if (testFolder != null)
            {
                // Delete it and then we can recreate it for our test.
                // WARNING: When using DeleteFile it totally deletes it from the drive. It does NOT put it in trash.
                if (googleDriveManager.DeleteFile(driveServiceFullAccess, testFolder.Id) == false)
                {
                    Console.WriteLine("******* ERROR: Cannot delete the test folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // Create the test folder in the root
            var testFolderId = googleDriveManager.CreateFolder(driveServiceFullAccess, "Test");
            if (string.IsNullOrEmpty(testFolderId) == true)
            {
                Console.WriteLine("******* ERROR: Cannot create test folder *******");
                Console.ReadKey();
                return;
            }

            // Assign the test folder permissions for "readwatchcreate.com" domain
            var permission = googleDriveManager.CreateUserPermission("reader", "fastcpu88@gmail.com");
            if (permission == null)
            {
                Console.WriteLine("******* ERROR: Could not create reader permission for 'readwatchcreate.com' domain *******");
                Console.ReadKey();
                return;
            }

            if(googleDriveManager.AssignPermissionToFile(driveServiceFullAccess, testFolderId, permission) == false)
            {
                Console.WriteLine("******* ERROR: Could not set reader permission for 'readwatchcreate.com' domain on test folder *******");
                Console.ReadKey();
                return;
            }

            #endregion

            #region STEP 5: Testing Drive access at the Child Folder level

            #region PRE-STEP: Ensure there are at least 11 child folders in the T1 folder

            // check if T1Child folder exists and if not create it
            var t1ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T1Child", testFolderId);
            if (t1ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T1Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T1Child in the Test Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T2Child folder exists and if not create it
            var t2ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T2Child", testFolderId);
            if (t2ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T2Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T2Child folder in the Test Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T3Child folder exists and if not create it
            var t3ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T3Child", testFolderId);
            if (t3ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T3Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T3Child folder in the T3 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T4Child folder exists and if not create it
            var t4ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T4Child", testFolderId);
            if (t4ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T4Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T4Child folder in the T4 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T5Child folder exists and if not create it
            var t5ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T5Child", testFolderId);
            if (t5ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T5Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T5Child folder in the T5 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T6Child folder exists and if not create it
            var t6ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T6Child", testFolderId);
            if (t6ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T6Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T6Child folder in the T6 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T7Child folder exists and if not create it
            var t7ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T7Child", testFolderId);
            if (t7ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T7Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T7Child folder in the T7 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T8Child folder exists and if not create it
            var t8ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T8Child", testFolderId);
            if (t8ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T8Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T8Child folder in the T8 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T9Child folder exists and if not create it
            var t9ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T9Child", testFolderId);
            if (t9ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T9Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T9Child folder in the T9 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T10Child folder exists and if not create it
            var t10ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T10Child", testFolderId);
            if (t10ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T10Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T10Child folder in the T10 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            // check if T11Child folder exists and if not create it
            var t11ChildFolder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T11Child", testFolderId);
            if (t11ChildFolder == null)
            {
                if (string.IsNullOrEmpty(googleDriveManager.CreateFolder(driveServiceFullAccess, "T11Child", new List<string> { testFolderId })))
                {
                    Console.WriteLine("******* ERROR: Cannot create T11Child folder in the T11 Folder *******");
                    Console.ReadKey();
                    return;
                }
            }

            #endregion

            #region STEP 3.1: Negative Testing

            // STEP 3.1.1 - Test no drive service, first page only
            try { var expectedFailure = googleDriveManager.GetFileList(null, true); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.1 - PASSED : EXPECTED FAILURE with no drive service passed while asking for first page only: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 3.1.2 - Test no drive service, all pages
            try { var expectedFailure = googleDriveManager.GetFolderList(null, false); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.2 - PASSED : EXPECTED FAILURE with no drive service passed while asking for all pages: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 3.2: Positive Testing

            // STEP 3.2.1 - Test drive service, first page only
            folders = googleDriveManager.GetFolderList(driveServiceReadOnly, true);
            if (folders == null)
            {
                Console.WriteLine("******* ERROR: Cannot get the Root Folders (first page only) *******");
                Console.ReadKey();
                return;
            }
            foreach (var folder in folders) { Console.WriteLine(folder.Name); }

            // STEP 3.2.2 - Test drive service, all pages
            folders = googleDriveManager.GetFolderList(driveServiceReadOnly, false);
            if (folders == null)
            {
                Console.WriteLine("******* ERROR: Cannot get the Root Folders (all pages) *******");
                Console.ReadKey();
                return;
            }
            foreach (var folder in folders) { Console.WriteLine(folder.Name); }

            #endregion

            #endregion
        }
    }
}
