/********************************************************************************
 * CSHARP Google Sheet Library - This library is meant to be an example to help 
 * you when working with Google Sheet. 
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
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;

namespace CSHARP.GoogleApis.GoogleSheet
{
    /// <summary>
    /// Manages Accessing and manipulating Google Sheets
    /// </summary>
    public class GoogleSheetManager : BaseGoogleApiManager
    {
        /// <summary>
        /// Allows access to Sheet Service in Read-Only Mode
        /// </summary>
        public string[] ReadOnlyScope = { SheetsService.Scope.SpreadsheetsReadonly };

        /// <summary>
        /// Allows access to Sheet Service with Full Access
        /// </summary>
        public string[] FullAccessScope = { SheetsService.Scope.Spreadsheets };

        #region Canned Sheet Ranges

        // (FIXED) Syntax error in getting the first row. Needs to be a range not single digit.
        public string RangeFirstRow = "1:1";          // Returns all the elements in the first row

        public string RangeFirstTwoRows = "1:2";    // Returns all the elements in the first and second rows
        public string RangeFirstColumn = "A:A";     // Returns all the elements in the first column

        #endregion

        public GoogleSheetManager() 
        {

        }

        /// <summary>
        /// Constructor that allows override of JSON Credentials file
        /// </summary>
        /// <param name="overrideCredentialsJsonFilePath">Full path to Credentials JSON file</param>
        public GoogleSheetManager(string overrideCredentialsJsonFilePath) : base(overrideCredentialsJsonFilePath)
        {

        }

        /// <summary>
        /// Creates instance of the Google Sheet Service
        /// </summary>
        /// <param name="userCredential">User Credential created by calling GetUserCredentials on this class</param>
        /// <param name="applicationName">The name of the appplication Google is expecting to call</param>
        /// <returns>Returns the SheetService instance upons success.</returns>
        /// <remarks>
        /// UNIT TESTING: Tested via UnitTestGoogleDriveViaConsole - STEP 2
        /// </remarks>
        public SheetsService CreateSheetServiceInstance(GoogleCredential credential, string applicationName)
        {
            // if no credentials passed in don't even try to create the drive service
            if (credential == null) throw new ArgumentNullException("credentials are required in order to create the Sheet Service");

            // if no application name is supplied return null
            if (string.IsNullOrEmpty(applicationName)) throw new ArgumentNullException("application name is required in order to create the Sheet Service");

            // Create Drive API service.
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }

        #region Column Mapping Related

        /// <summary>
        /// Returns the column letter for the column index (Useful when working with ValueRange)
        /// </summary>
        /// <param name="columnIndex">Index of the column in the sheet</param>
        /// <returns>The letter matching the column index</returns>
        /// <remarks>Adapted from https://stackoverflow.com/questions/43097563/google-sheet-api-v4-java-get-column-letter-from-column-number
        /// </remarks>
        public string ColumnIndexToLetter(int columnIndex)
        {
            int Base = 26;
            String chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            int TempNumber = columnIndex;
            string outputColumnName = string.Empty;
            while (TempNumber > 0)
            {
                int position = TempNumber % Base;
                outputColumnName = (position == 0 ? "Z" : chars.Substring(position > 0 ? position - 1 : 0, 1)) + outputColumnName;
                TempNumber = (TempNumber - 1) / Base;
            }
            return outputColumnName;
        }

        /// <summary>
        /// Returns the column index for the column letter passed in
        /// </summary>
        /// <remarks>Adapted from https://stackoverflow.com/questions/43097563/google-sheet-api-v4-java-get-column-letter-from-column-number
        /// </remarks>
        public int ColumnLetterToIndex(string columnLetter)
        {
            int outputColumnNumber = 0;

            if (columnLetter == null || columnLetter.Length == 0)
            {
                throw new ArgumentNullException("columnLetter is required");
            }

            int i = columnLetter.Length - 1;
            int t = 0;
            while (i >= 0)
            {
                char curr = columnLetter.Substring(i, 1)[0];
                outputColumnNumber = outputColumnNumber + (int)Math.Pow(26, t) * (curr - 'A' + 1);
                t++;
                i--;
            }
            return outputColumnNumber;
        }

        #endregion

        /// <summary>
        /// Gets the first row of a sheet given the sheet id and name
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="sheetName">(Optional) Name of the sheet (tabs at the bottom)</param>
        /// <returns></returns>
        public ValueRange GetFirstRowOfSheet(SheetsService sheetsService, string sheetId, string sheetName)
        {
            // if sheets service not passed in we cannot get sheet data
            if (sheetsService == null) throw new ArgumentNullException("Sheets service instance is required in order to get sheet data");

            // if no sheet id is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheet id is required in order to get sheet data");

            return GetSheetDataRange(sheetsService, sheetId, (string.IsNullOrEmpty(sheetName) == false ? "'" + sheetName + "'!" : "") + RangeFirstRow);
        }

        /// <summary>
        /// Gets all the elements in the first column
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="sheetName">(Optional) Name of the sheet (tabs at the bottom)</param>
        /// <returns></returns>
        public ValueRange GetFirstColumn(SheetsService sheetsService, string sheetId, string sheetName)
        {
            // if sheets service not passed in we cannot get sheet data
            if (sheetsService == null) throw new ArgumentNullException("Sheets service instance is required in order to get sheet data");

            // if no sheet id is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheet id is required in order to get sheet data");

            return GetColumn(sheetsService, sheetId, sheetName, "A");
        }

        /// <summary>
        /// Gets the data from the given row
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="sheetName">(Optional) Name of the sheet (tabs at the bottom)</param>
        /// <param name="row">index of the row to read</param>
        /// <returns></returns>
        public ValueRange GetRow(SheetsService sheetsService, string sheetId, string sheetName, int row)
        {
            // if sheets service not passed in we cannot get sheet data
            if (sheetsService == null) throw new ArgumentNullException("Sheets service instance is required in order to get sheet data");

            // if no sheet id is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheet id is required in order to get sheet data");

            // if no row is supplied we cannot get sheet data
            if (row < 0) throw new ArgumentNullException("a row zero or greater is required in order to get sheet data");

            return GetSheetDataRange(sheetsService, sheetId, (string.IsNullOrEmpty(sheetName) == false ? "'" + sheetName + "'!" : "") + row);
        }

        /// <summary>
        /// Gets all the elements in the given column
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="sheetName">(Optional) Name of the sheet (tabs at the bottom)</param>
        /// <param name="column">The column to retrieve data for</param>
        /// <returns></returns>
        public ValueRange GetColumn(SheetsService sheetsService, string sheetId, string sheetName, string column)
        {
            // if sheets service not passed in we cannot get sheet data
            if (sheetsService == null) throw new ArgumentNullException("Sheets service instance is required in order to get sheet data");

            // if no sheet id is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheet id is required in order to get sheet data");

            // if no column is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(column)) throw new ArgumentNullException("column is required in order to get sheet data");

            return GetSheetDataRange(sheetsService, sheetId, (string.IsNullOrEmpty(sheetName) == false ? "'" + sheetName + "'!" : "") + column + ":" + column);
        }

        /// <summary>
        /// Gets data for a range of rows and cells in a Google Sheet
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="range">Range of data to get eg. Class Data!A2:E</param>
        public ValueRange GetSheetDataRange(SheetsService sheetsService, string sheetId, string range)
        {
            // if sheets service not passed in we cannot get sheet data
            if (sheetsService == null) throw new ArgumentNullException("Sheets service instance is required in order to get sheet data");

            // if no sheet id is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(sheetId)) throw new ArgumentNullException("sheet id is required in order to get sheet data");

            // if no range is supplied we cannot get sheet data
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException("range is required in order to get sheet data");

            SpreadsheetsResource.ValuesResource.GetRequest request =
                    sheetsService.Spreadsheets.Values.Get(sheetId, range);

            return request.Execute();
        }

        /// <summary>
        /// Updates a single cell. Note if you are updating a lot of data it is better to update a range for performance.
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="range">Range of data to get eg. Class Data!A2:E</param>
        /// <param name="newCellValue"></param>
        /// <returns></returns>
        public UpdateValuesResponse UpdateSheetCellValue(SheetsService sheetsService, string sheetId, string range, string newCellValue)
        {
            // A Value Range is an array of arrays.
            ValueRange valueRange = new ValueRange
            {
                MajorDimension = "COLUMNS"//"ROWS";//COLUMNS
            };
            var oblist = new List<object>() { newCellValue };
            valueRange.Values = new List<IList<object>> { oblist };

            return UpdateSheetDataRange(sheetsService, sheetId, range, valueRange);
        }

        /// <summary>
        /// Appends a row to the sheet.
        /// </summary>
        /// <param name="sheetsService">The sheets service created using CreateSheetsServiceInstance</param>
        /// <param name="sheetId">Unique Id representing the sheet</param>
        /// <param name="sheetName">SheetName, if supplied is used to generate the range</param>
        /// <param name="nextRow"></param>
        /// <param name="valueRange"></param>
        /// <returns></returns>
        public AppendValuesResponse AppendSheetData(SheetsService sheetsService, string sheetId, string sheetName, long nextRow, ValueRange valueRange)
        {
            // (FIX) Row Range is {SheetName}!{row}:{row}
            SpreadsheetsResource.ValuesResource.AppendRequest request = sheetsService.Spreadsheets.Values.Append(valueRange, sheetId,
                (string.IsNullOrEmpty(sheetName) == false ? "'" + sheetName + "'!" : "") + nextRow.ToString() + ":" + nextRow.ToString());
            request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            return request.Execute();
        }

        /// <summary>
        /// Updates a range of data in a sheet. 
        /// </summary>
        /// <param name="sheetsService"></param>
        /// <param name="sheetId"></param>
        /// <param name="range"></param>
        /// <param name="valueRange"></param>
        /// <returns></returns>
        public UpdateValuesResponse UpdateSheetDataRange(SheetsService sheetsService, string sheetId, string range, ValueRange valueRange)
        {
            // Update the cell
            SpreadsheetsResource.ValuesResource.UpdateRequest update = sheetsService.Spreadsheets.Values.Update(valueRange, sheetId, range);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            return update.Execute();
        }
    }
}
