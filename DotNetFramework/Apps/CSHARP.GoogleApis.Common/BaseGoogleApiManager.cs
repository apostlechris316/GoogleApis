/********************************************************************************
 * CSHARP Google APIs Library - Example code to help you get started connecting 
 * to Google APIs from a Console app or Powershell. 
 *  
 * LICENSE: Free to use provided details on fixes and/or extensions emailed to 
 *          chris.williams@readwatchcreate.com
 *          
 * This blog article provides more information on how this works:
 *  
 * http://aspnetguild.blogspot.ca/2017/12/reading-and-writing-files-to-google.html
 * 
 * ********************************************************************************/

using System.IO;
using Google.Apis.Auth.OAuth2;
using System;
using System.Configuration;

namespace CSHARP.GoogleApis.Common
{
    /// <summary>
    /// Base class for all the Google API Manager eg. Google Drive, Google Sheets, etc
    /// </summary>
    public class BaseGoogleApiManager
    {
        public BaseGoogleApiManager()
        {

        }

        /// <summary>
        /// Name of the application (used when calling Google Sheet API)
        /// </summary>
        public string ApplicationName
        {
            get
            {
                // if overridden then return that
                if (string.IsNullOrEmpty(_overrideCredentialsJsonFilePath) == false) return _overrideCredentialsJsonFilePath;

                // otherwise get default from config file
                return (ConfigurationManager.AppSettings["CrmContactApplicationName"] != null) ? ConfigurationManager.AppSettings["CrmContactApplicationName"] : "Google APIs";
            }
        }
        private string _overrideApplicationName = "Google APIs";

        /// <summary>
        /// Base constructor for Google API Managers
        /// </summary>
        /// <param name="overrideCredentialsJsonFilePath">Full path to Credentials JSON file</param>
        public BaseGoogleApiManager(string overrideCredentialsJsonFilePath)
        {
            _overrideCredentialsJsonFilePath = overrideCredentialsJsonFilePath;
        }

        /// <summary>
        /// Credentials JSON File Path (If empty assumes client-secret.json in the same folder as application.
        /// </summary>
        public string CredentialsJsonFilePath
        {
            get
            {
                // if overridden then return that
                if (string.IsNullOrEmpty(_overrideCredentialsJsonFilePath) == false) return _overrideCredentialsJsonFilePath;

                // otherwise get default from config file
                return (ConfigurationManager.AppSettings["GoogleCredentialsJsonFilePathContact"] != null) ? ConfigurationManager.AppSettings["GoogleCredentialsJsonFilePathContact"] : "client-secret.json";
            }
        }

        private string _overrideCredentialsJsonFilePath = string.Empty;

        /// <summary>
        /// Gets the user credentials using default Credentials JSON
        /// </summary>
        /// <param name="scopes">Determines how the drive will be accessed. You can use one of the scope constants like ReadOnlyScope</param>
        /// <returns></returns>
        public GoogleCredential GetUserCredentials(string [] scopes)
        {
            return GetUserCredential(CredentialsJsonFilePath, scopes);
        }

        /// <summary>
        /// Gets the user credentials used to access Google APIs 
        /// </summary>
        /// <param name="credentialsFileName">Name of credentials .json file downloaded from Google including extension. This method expects the credentials file to be in the application execution path unless you specify the full path</param>
        /// <param name="scopes">Determines how the drive will be accessed. You can use one of the scope constants like ReadOnlyScope</param>
        /// <remarks>
        /// TESTED via UnitTestGoogleDriveViaConsole Step 1
        /// </remarks>
        public GoogleCredential GetUserCredential(string credentialsFileName, string[] scopes)
        {
            // if no credentials filename is supplied we cannot look up user credentials
            if (string.IsNullOrEmpty(credentialsFileName)) throw new ArgumentNullException("credentials JSON filename is required in order to get user credentials.");

            // Scopes determine permissions so if we do not have a scope then not worth asking for access to nothing.
            if (scopes == null) throw new ArgumentNullException("At least one scope is required to get user credentials.");
            if (scopes.Length < 1) throw new ArgumentNullException("At least one scope is required to get user credentials.");

            // Read the credentials from the file in preparation for service creations.
            GoogleCredential credential = null;
            using (var stream = new FileStream(credentialsFileName, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream);
                credential = credential.CreateScoped(scopes);
            }

            return credential;
        }
    }
}
