#region Help:  Introduction to the script task
/* The Script Task allows you to perform virtually any operation that can be accomplished in
 * a .Net application within the context of an Integration Services control flow. 
 * 
 * Expand the other regions which have "Help" prefixes for examples of specific ways to use
 * Integration Services features within this script task. */
#endregion


#region Namespaces
using System;
using Microsoft.SharePoint.Client;
#endregion

namespace ST_ca3ecad38994454dbf5c83efe2dfa1d8
{
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

        private static ClientContext GetSpContext(Uri portalUrl, string userName, string password)
        {
            var securePassword = new System.Security.SecureString();
            foreach (var ch in password) securePassword.AppendChar(ch);
            return new ClientContext(portalUrl) { Credentials = new SharePointOnlineCredentials(userName, securePassword) };
        }

        private static void DownloadSpFile(Web web, string fileUrl, string targetFile)
        {
            var ctx = (ClientContext)web.Context;
            ctx.ExecuteQuery();

            using (var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(ctx, fileUrl))
            {
                //Directory.CreateDirectory(targetPath);
                //var fileName = Path.Combine(targetPath, Path.GetFileName(fileUrl));
                using (var fileStream = System.IO.File.Create(targetFile))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                }
            }
        }

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
            string sourceUrl = Dts.Variables["$Project::Sharepoint_Link"].Value.ToString();
            string targetFile = Dts.Variables["$Project::FlatFile_Path"].Value.ToString();

            int portalIndex = sourceUrl.IndexOf("sharepoint.com/") + 15;
            string portalUrl = sourceUrl.Substring(0, portalIndex);
            string fileUrl = "/" + sourceUrl.Substring(portalIndex, sourceUrl.Length - portalUrl.Length);

            using (var ctx = GetSpContext(new Uri(portalUrl), userName, password))
            {
                var web = ctx.Web;
                try
                {
                    DownloadSpFile(web, fileUrl, targetFile);
                    Dts.Variables["User::Sharepoint_Success"].Value = true;
                }
                catch (Exception e)
                {
                    //Console.WriteLine(e.Message); Console.ReadLine();
                    Dts.Variables["User::Sharepoint_Error_Message"].Value = e.Message;
                    Dts.Variables["User::Sharepoint_Success"].Value = false;
                }
                //catch (IdcrlException e) { } // Credentials error
                //catch (System.Net.WebException e) { } // 404 file not found
                //catch (UnauthorizedAccessException e) { } // File write error
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
