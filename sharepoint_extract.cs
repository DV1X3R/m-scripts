#region Namespaces
using System;
using System.Net;
using System.Linq;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
#endregion

namespace ST_
{
    static class SpExtensions
    {
        // Extension method to handle random 503 error
        public static void ExecuteQueryR(this ClientContext clientContext, int retryAttempts, int delayBetweenRetriesMs)
        {
            while (true)
            {
                try
                {
                    clientContext.ExecuteQuery();
                    return;
                }
                catch (WebException e)
                {
                    if (((e.Response as HttpWebResponse).StatusCode == (HttpStatusCode)503) && (retryAttempts > 0))
                    {
                        retryAttempts--;
                        System.Threading.Thread.Sleep(delayBetweenRetriesMs);
                    }
                    else throw;
                }
            }
        }
    }

    static class SpUtils
    {
        private static Uri GetPortalUri(string sourceUrl)
        {
            return new Uri(sourceUrl.Substring(0, sourceUrl.IndexOf("sharepoint.com/") + 15));
        }

        private static string GetSpFileRelativePath(string sourceUrl)
        {
            int portalLength = sourceUrl.IndexOf("sharepoint.com/") + 15;
            return "/" + sourceUrl.Substring(portalLength, sourceUrl.Length - portalLength);
        }

        private static ClientContext GetSpContext(Uri uri, string userName, string password)
        {
            var securePassword = new System.Security.SecureString();
            foreach (var ch in password) securePassword.AppendChar(ch);
            return new ClientContext(uri) { Credentials = new SharePointOnlineCredentials(userName, securePassword) };
        }

        public static void DownloadSpFile(string userName, string password, string fileUrl, string targetFilePath, int retryAttempts, int delayBetweenRetriesMs)
        {
            using (var clientContext = GetSpContext(GetPortalUri(fileUrl), userName, password))
            {
                clientContext.ExecuteQueryR(retryAttempts, delayBetweenRetriesMs);
                using (var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, GetSpFileRelativePath(fileUrl)))
                {
                    System.IO.Directory.CreateDirectory(targetFilePath.Substring(0, targetFilePath.LastIndexOf('\\')));
                    using (var fileStream = System.IO.File.Create(targetFilePath))
                    {
                        fileInfo.Stream.CopyTo(fileStream);
                    }
                }
            }
        }

        public static void DownloadLatestSpFileFromFolder(string userName, string password, string folderUrl, string targetFilePath, int retryAttempts, int delayBetweenRetriesMs)
        {
            var filesInFolder = new List<string>();

            // populate 'filesInFolder' list
            using (var clientContext = GetSpContext(new Uri(folderUrl), userName, password))
            {
                var documents = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(documents.RootFolder.Folders);
                clientContext.ExecuteQueryR(retryAttempts, delayBetweenRetriesMs);

                foreach (var documentsFolder in documents.RootFolder.Folders)
                {
                    if (documentsFolder.Name == "General")
                    {
                        clientContext.Load(documentsFolder.Files);
                        clientContext.ExecuteQueryR(retryAttempts, delayBetweenRetriesMs);

                        foreach (var file in documentsFolder.Files)
                        {
                            filesInFolder.Add(file.Name);
                        }
                    }
                }
            }

            // download the latest file
            var fileUrl = folderUrl + "/Shared Documents/General/" + filesInFolder.OrderBy(x => x).Last();
            DownloadSpFile(userName, password, fileUrl, targetFilePath, retryAttempts, delayBetweenRetriesMs);
        }
    }

    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        #region Help:  Using Integration Services variables and parameters in a script
        /* To use a variable in this script, first ensure that the variable has been added to 
         * either the list contained in the ReadOnlyVariables property or the list contained in 
         * the ReadWriteVariables property of this script task, according to whether or not your
         * code needs to write to the variable.  To add the variable, save this script, close this instance of
         * Visual Studio, and update the ReadOnlyVariables and 
         * ReadWriteVariables properties in the Script Transformation Editor window.
         * To use a parameter in this script, follow the same steps. Parameters are always read-only.
         * 
         * Example of reading from a variable:
         *  DateTime startTime = (DateTime) Dts.Variables["System::StartTime"].Value;
         * 
         * Example of writing to a variable:
         *  Dts.Variables["User::myStringVariable"].Value = "new value";
         * 
         * Example of reading from a package parameter:
         *  int batchId = (int) Dts.Variables["$Package::batchId"].Value;
         *  
         * Example of reading from a project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].Value;
         * 
         * Example of reading from a sensitive project parameter:
         *  int batchId = (int) Dts.Variables["$Project::batchId"].GetSensitiveValue();
         * */

        #endregion

        #region Help:  Firing Integration Services events from a script
        /* This script task can fire events for logging purposes.
         * 
         * Example of firing an error event:
         *  Dts.Events.FireError(18, "Process Values", "Bad value", "", 0);
         * 
         * Example of firing an information event:
         *  Dts.Events.FireInformation(3, "Process Values", "Processing has started", "", 0, ref fireAgain)
         * 
         * Example of firing a warning event:
         *  Dts.Events.FireWarning(14, "Process Values", "No values received for input", "", 0);
         * */
        #endregion

        #region Help:  Using Integration Services connection managers in a script
        /* Some types of connection managers can be used in this script task.  See the topic 
         * "Working with Connection Managers Programatically" for details.
         * 
         * Example of using an ADO.Net connection manager:
         *  object rawConnection = Dts.Connections["Sales DB"].AcquireConnection(Dts.Transaction);
         *  SqlConnection myADONETConnection = (SqlConnection)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Sales DB"].ReleaseConnection(rawConnection);
         *
         * Example of using a File connection manager
         *  object rawConnection = Dts.Connections["Prices.zip"].AcquireConnection(Dts.Transaction);
         *  string filePath = (string)rawConnection;
         *  //Use the connection in some code here, then release the connection
         *  Dts.Connections["Prices.zip"].ReleaseConnection(rawConnection);
         * */
        #endregion

        /// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
        public void Main()
        {
            // TODO: Add your code here
            string userName = Dts.Variables["$Project::MAIL_AO365"].Value.ToString();
            string password = Dts.Variables["$Project::PASSWORD_AO365"].Value.ToString();
            string sourceUrl = Dts.Variables["User::Sharepoint_Source_Link"].Value.ToString();
            string targetFilePath = Dts.Variables["User::Sharepoint_Destination_File"].Value.ToString();

            try
            {
                SpUtils.DownloadSpFile(userName, password, sourceUrl, targetFilePath, 20, 2000);
                Dts.Variables["User::Sharepoint_Success"].Value = true;
            }
            catch (Exception e)
            {
                Dts.Variables["User::Sharepoint_Error_Message"].Value = e.Message;
                Dts.Variables["User::Sharepoint_Success"].Value = false;
            }

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        #region ScriptResults declaration
        /// <summary>
        /// This enum provides a convenient shorthand within the scope of this class for setting the
        /// result of the script.
        /// 
        /// This code was generated automatically.
        /// </summary>
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

    }
}
