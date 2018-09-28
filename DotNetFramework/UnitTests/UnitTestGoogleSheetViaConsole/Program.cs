namespace UnitTestGoogleSheetViaConsole
{
    using CSHARP.GoogleApis.GoogleDrive;
    using CSHARP.GoogleApis.GoogleSheet;

    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Drive.v3;
    using Google.Apis.Sheets.v4;

    using System;
    using System.Collections.Generic;

    class Program
    {
        static void Main(string[] args)
        {
            // ************************************************************
            // ASSUMPTION: Unit test has been run on Google Drive first
            // ************************************************************

            var googleDriveManager = new GoogleDriveManager();
            var googleSheetManager = new GoogleSheetManager();

            GoogleCredential credentialReadOnly = null;
            GoogleCredential credentialFullAccess = null;

            DriveService driveServiceReadOnly = null;
            DriveService driveServiceFullAccess = null;

            SheetsService sheetsServiceReadOnly = null;
            SheetsService sheetsServiceFullAccess = null;

            #region STEP 0: Pre-test validation

            if (System.IO.File.Exists("client-secret.json") == false)
            {
                Console.WriteLine("******* ERROR: Please ensure you have downloaded your Google API .JSON file and copied into same folder as the .EXE *******");
                Console.ReadKey();
                return;
            }

            // Get the readonly drive credentials. (If it fails here run UnitTestGoogleDriveViaConsole to troubleshoot)
            credentialReadOnly = googleDriveManager.GetUserCredential("client-secret.json", googleDriveManager.ReadOnlyScope);
            if (credentialReadOnly == null)
            {
                Console.WriteLine("******* PRE-TEST VALIDATION FAILED - Credentials returned null when getting Read Only Scope permissions");
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // Get the full access drive credentials. (If it fails here run UnitTestGoogleDriveViaConsole to troubleshoot)
            credentialFullAccess = googleDriveManager.GetUserCredential("client-secret.json", googleDriveManager.FullAccessScope);
            if (credentialFullAccess == null)
            {
                Console.WriteLine("******* PRE-TEST VALIDATION FAILED - Credentials returned null when getting Full Access Scope permissions");
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // Get Drive Service Instance as Read-Only
            driveServiceReadOnly = googleDriveManager.CreateDriveServiceInstance(credentialReadOnly, "UnitTest Google Sheet Library");
            if (driveServiceReadOnly == null)
            {
                Console.WriteLine("******* STEP 2.2.1 - FAILED - drive service returned null getting it with Read Only credentials. ");
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // Get Drive Service Instance as Full Access
            driveServiceFullAccess = googleDriveManager.CreateDriveServiceInstance(credentialFullAccess, "UnitTest Google Sheet Library");
            if (driveServiceFullAccess == null)
            {
                Console.WriteLine("******* STEP 2.2.2 - FAILED - drive service returned null getting it with Full Access credentials. ");
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            var t1Folder = googleDriveManager.GetFolderByName(driveServiceReadOnly, "T1");
            if (t1Folder == null)
            {
                Console.WriteLine("******* ERROR: T1 folder should already exist. Please ensure you ran UnitTestGoogleDriveViaConsole first. *******");
                Console.ReadKey();
                return;
            }

            // Check if sheet exists and so delete it
            var sheetFile = googleDriveManager.GetSheetFileInChildFolderByName(driveServiceReadOnly, "TestSheet", t1Folder.Id);
            if (sheetFile != null)
            {
                googleDriveManager.DeleteFile(driveServiceFullAccess, sheetFile.Id);
            }

            #endregion

            #region STEP 1: Create the test sheet

            #region STEP 1.1: Negative Testing

            // STEP 1.1.1 - Test no drive service
            try { var expectedFailure = googleDriveManager.CreateGoogleSheet(null, string.Empty, string.Empty, null); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 1.1.1 - PASSED : EXPECTED FAILURE with no drive service passed: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 3.1.2 - Test drive service, no sheet name
            try { var expectedFailure = googleDriveManager.CreateGoogleSheet(driveServiceFullAccess, string.Empty, string.Empty, new List<string> { t1Folder.Id }); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.2 - PASSED : EXPECTED FAILURE with drive service but no sheet name: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 1.2 Positive Testing

            var sheetFileId = googleDriveManager.CreateGoogleSheet(driveServiceFullAccess, "TestSheet", "Sheet used to Unit test CSHARP Sheet Library.", new List<string> { t1Folder.Id });
            if(string.IsNullOrEmpty(sheetFileId))
            {
                Console.WriteLine("******* ERROR: Failed to create test sheet *******");
                Console.ReadKey();
                return;
            }

            #endregion

            #endregion

            #region STEP 2: Assign Read Permissions to fastcpu88@gmail.com

            // Get the permission object so we can use it to assign permission to the sheet
            var permission = googleDriveManager.CreateUserPermission("reader", "fastcpu88@gmail.com");
            if (permission == null)
            {
                Console.WriteLine("******* ERROR: Failed to create user permission object for fastcpu88@gmail.com *******");
                Console.ReadKey();
                return;
            }

            // Assign the permission to the sheet
            if (googleDriveManager.AssignPermissionToFile(driveServiceFullAccess, sheetFileId, permission) == false)
            {
                Console.WriteLine("******* ERROR: Failed to set reader permission for user fastcpu88@gmail.com for spreadsheet *******");
                Console.ReadKey();
                return;
            }

            // Get the permission object so we can use it to assign permission to the sheet
            permission = googleDriveManager.CreateUserPermission("writer", "fastcpu88@gmail.com");
            if (permission == null)
            {
                Console.WriteLine("******* ERROR: Failed to create user permission object for fastcpu88@gmail.com *******");
                Console.ReadKey();
                return;
            }

            // Assign the permission to the sheet
            if (googleDriveManager.AssignPermissionToFile(driveServiceFullAccess, sheetFileId, permission) == false)
            {
                Console.WriteLine("******* ERROR: Failed to set write permission for user fastcpu88@gmail.com for spreadsheet *******");
                Console.ReadKey();
                return;
            }

            #endregion

            #region STEP 3: Get Sheets Service Instance

            #region STEP 3.1: Negative Testing

            // STEP 3.1.1 - Test no credentials passed in and no application name
            try { var expectedFailure = googleSheetManager.CreateSheetServiceInstance(null, string.Empty); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.1 - PASSED : EXPECTED FAILURE with no credentials passed: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            // STEP 2.1.2 - Test credentials passed in but no application name
            try { var expectedFailure = googleSheetManager.CreateSheetServiceInstance(credentialReadOnly, string.Empty); }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.1.2 - PASSED : EXPECTED FAILURE with no application name passed: ");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
            }

            #endregion

            #region STEP 3.2: Positive Testing

            // STEP 3.2.1 - Test getting sheet service with ReadOnly Scope
            try
            {
                sheetsServiceReadOnly = googleSheetManager.CreateSheetServiceInstance(credentialReadOnly, "UnitTest Google Sheet Library");
                if (sheetsServiceReadOnly == null)
                {
                    Console.WriteLine("******* STEP 3.2.1 - FAILED - sheets service returned null getting it with Read Only credentials. ");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.2.1 - FAILED - Exception thrown getting sheets service using Read Only credentials.");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            // STEP 3.2.2 - Test getting sheet service with Full Access Scope
            try
            {
                sheetsServiceFullAccess = googleSheetManager.CreateSheetServiceInstance(credentialFullAccess, "UnitTest Google Sheet Library");
                if (sheetsServiceFullAccess == null)
                {
                    Console.WriteLine("******* STEP 3.2.2 - FAILED - sheets service returned null getting it with Full Access credentials. ");
                    Console.WriteLine("*******");
                    Console.ReadKey();
                    return;
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine("******* STEP 3.2.2 - FAILED - Exception thrown getting sheets service using Full Access credentials.");
                Console.WriteLine(exception.ToString());
                Console.WriteLine("*******");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("******* STEP 3.2 - PASSED");
            Console.WriteLine("*******");
            Console.ReadKey();

            #endregion

            #endregion

            #region STEP 4: Read and Write Values from test sheet

            #region STEP 4.1 - Read range (should be blank)

            var testRange = "Sheet1!A2:E2";
            var valueResults = googleSheetManager.GetSheetDataRange(sheetsServiceReadOnly, sheetFileId, testRange);

            if(valueResults == null)
            {
                Console.WriteLine("******* ERROR: Failed to get range in sheet *******");
                Console.ReadKey();
                return;
            }

            // Review the cells are empty. The Values are stored in an array of arrays
            IList<IList<Object>> values = valueResults.Values;
            if (values == null || values.Count == 0)
            {
                Console.WriteLine("******* PASSED: Expected range in sheet to be empty as we just created it *******");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("******* ERROR: Failed as sheet range should have no data *******");
                Console.ReadKey();
                return;
            }

            #endregion

            #region STEP 4.2 - Write a single value and confirm it got written

            var singleCellTestRange = "Sheet1!A2";
            googleSheetManager.UpdateSheetCellValue(sheetsServiceFullAccess, sheetFileId, singleCellTestRange, "New Text In A2");
            var singleCellValueResult = googleSheetManager.GetSheetDataRange(sheetsServiceReadOnly, sheetFileId, singleCellTestRange);
            if (singleCellValueResult == null || singleCellValueResult.Values.Count == 0)
            {
                Console.WriteLine("******* ERROR: Failed to read back single cell value *******");
                Console.ReadKey();
                return;
            }

            if (singleCellValueResult.Values[0].Count == 0)
            {
                Console.WriteLine("******* ERROR: Failed to read back single cell value *******");
                Console.ReadKey();
                return;
            }

            if (singleCellValueResult.Values[0][0].ToString() != "New Text In A2")
            {
                Console.WriteLine("******* ERROR: Failed as single text value was not updated in the sheet *******");
                Console.ReadKey();
                return;
            }
            #endregion

            #region STEP 5: Bulk Write data to test sheet

            // Getting the range again as we did update a cell in the range.
            valueResults = googleSheetManager.GetSheetDataRange(sheetsServiceReadOnly, sheetFileId, testRange);

            // Fill in values in the results
            for(int rowIndex = 0; rowIndex < valueResults.Values.Count; rowIndex++)
            {
                for (int colIndex = 0; colIndex < valueResults.Values[rowIndex].Count; colIndex++)
                {
                    valueResults.Values[rowIndex][colIndex] = rowIndex.ToString() + ":" + colIndex.ToString();
                }
            }

            // Write the values
            googleSheetManager.UpdateSheetDataRange(sheetsServiceFullAccess, sheetFileId, testRange, valueResults);

            // Getting the range again to confirm updates.
            valueResults = googleSheetManager.GetSheetDataRange(sheetsServiceReadOnly, sheetFileId, testRange);

            // Check the expected values in the results
            for (int rowIndex = 0; rowIndex < valueResults.Values.Count; rowIndex++)
            {
                for (int colIndex = 0; colIndex < valueResults.Values[rowIndex].Count; colIndex++)
                {
                    if(valueResults.Values[rowIndex][colIndex].ToString() != rowIndex.ToString() + ":" + colIndex.ToString())
                    {
                        Console.WriteLine("******* ERROR: Failed as the value we expected to be updated was not: " + rowIndex.ToString() + ":" + colIndex.ToString() + " * ******");
                        Console.ReadKey();
                        return;
                    }
                }
            }

            #endregion

            #endregion
        }
    }
}
