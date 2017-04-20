// -----------------------------------------------------------------------
// <copyright file="VoiceAttackPlugin.cs" company="Insequence Corporation">
//      Copyright (c) Insequence Corporation. All Rights Reserved
// </copyright>
// -----------------------------------------------------------------------
namespace ReadGoogleSpreadsheetVAPlugin
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Threading;
    using Google.Apis.Auth.OAuth2;
    using Google.Apis.Services;
    using Google.Apis.Sheets.v4;
    using Google.Apis.Sheets.v4.Data;
    using Google.Apis.Util.Store;

    /// <summary>
    /// Voice Attack Plugin Class
    /// </summary>
    public class VoiceAttackPlugin
    {
        /// <summary>
        /// The scopes to use in the Google API Request
        /// </summary>
        private static string[] scopes = { SheetsService.Scope.SpreadsheetsReadonly };

        /// <summary>
        /// The Application Name for the Google API
        /// </summary>
        private static string applicationName = "Read Google Spreadsheets Voice Attack Plugin";

        /// <summary>
        /// The Google API <see cref="UserCredential"/>
        /// </summary>
        private static UserCredential credentials = null;

        /// <summary>
        /// A handle to the Google Spreadsheet API
        /// </summary>
        private static SheetsService service = null;

        /// <summary>
        /// A <see cref="CancellationTokenSource"/> that can be used to cancel a request in progress.
        /// </summary>
        private static CancellationTokenSource cts = null;

        /// <summary>
        /// A list of variables that have been sent to voice attack, used for clearing values on subsequent calls.
        /// </summary>
        private static List<string> voiceAttackVariables = new List<string>();

        /// <summary>
        /// Returns the name of the Plugin for VoiceAttack
        /// </summary>
        /// <returns>Name of the Plugin</returns>
        public static string VA_DisplayName()
        {
            return "Read Google Spreadsheet Plugin - v1.0.0";
        }

        /// <summary>
        /// Returns information about this plugin.
        /// </summary>
        /// <returns>About information for the plugin.</returns>
        public static string VA_DisplayInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Read Google Spreadsheet Plugin");
            sb.AppendLine("Author: Jeffrey Geer");
            sb.Append("Reads data from a Google Doc's Spreadsheet");
            return sb.ToString();
        }

        /// <summary>
        /// Gets an identifying value for VoiceAttack
        /// </summary>
        /// <returns><see cref="Guid"/> value to identify this plugin.</returns>
        public static Guid VA_Id()
        {
            // {CD39E757-B3B4-4C00-9AA2-8B861FAC391D}
            return new Guid("{CD39E757-B3B4-4C00-9AA2-8B861FAC391D}");
        }

        /// <summary>
        /// This is called when a Stop Command is issued by Voice Attack, not much I can do though as these are HTTP calls.
        /// </summary>
        public static void VA_StopCommand()
        {
            cts?.Cancel();
        }

        /// <summary>
        /// This is called when VoiceAttack initializes the plugin.
        /// </summary>
        /// <param name="vaProxy">A connection to VoiceAttack, used for passing back values.</param>
        public static void VA_Init1(dynamic vaProxy)
        {
            try
            {
                string clientSecretPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "client_secret.json");
                using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
                {
                    string credPath = Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/ReadGoogleSheetsVAPlugin.json");

                    cts = new CancellationTokenSource();
                    var credTask = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        scopes,
                        "user",
                        cts.Token,
                        new FileDataStore(credPath, true));
                    credTask.Wait();
                    cts = null;
                    if (credTask.IsCanceled)
                    {
                        credentials = null;
                        service = null;
                        return;
                    }
                    else
                    {
                        credentials = credTask.Result;
                    }
                }
                
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credentials,
                    ApplicationName = applicationName,
                });
            }
            catch (Exception ex)
            {
                var sb = new StringBuilder();
                sb.AppendLine("ReadGoogleSpreadsheetVAPlugin Error");
                sb.AppendLine("====================================");
                sb.AppendFormat("Message: {0}", ex.Message).AppendLine();
                sb.AppendLine("Stacktrace:");
                sb.AppendLine(ex.StackTrace);
                System.Windows.Forms.MessageBox.Show(sb.ToString(), "ReadGoogleSpreadsheetVAPlugin Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Called when VoiceAttack Exits.
        /// </summary>
        /// <param name="vaProxy">A connection to VoiceAttack, used for passing back values.</param>
        public static void VA_Exit1(dynamic vaProxy)
        {
            cts?.Cancel();
        }

        /// <summary>
        /// Called when the Plugin is called by VoiceAttack.
        /// </summary>
        /// <param name="vaProxy">A connection to VoiceAttack, used for passing back values.</param>
        public static void VA_Invoke1(dynamic vaProxy)
        {
            // Set in progress Variable (for Async Calls)
            vaProxy.SetBoolean("ReadingGoogleSheet", true);
            
            // Clear all previous variables.
            foreach (var variable in voiceAttackVariables)
            {
                vaProxy.SetText(variable, null);
            }
            voiceAttackVariables.Clear();

            // Get the Google Sheet Id, which should be in the Context, Exit if it is Empty or Null
            string sheetId = vaProxy.Context;
            if (string.IsNullOrEmpty(sheetId) || sheetId == "ClearData")
            {
                vaProxy.SetBoolean("ReadingGoogleSheet", false);
                return;
            }

            // They have made an authorization request, so call the authorization and exit
            if (sheetId == "Authorize")
            {
                VA_Init1(vaProxy);
                vaProxy.SetBoolean("ReadingGoogleSheet", false);
                return;
            }

            // A variable named "RequestSheetRange" specifies what data to get.
            string range = vaProxy.GetText("RequestSheetRange");
            if (range != null)
            {
                if (service == null)
                {
                    vaProxy.SetText("SheetRange", "Google API Connection is not Initialized!");
                    voiceAttackVariables.Add("SheetRange"); // Include this variable into the list for clearing
                }

                // Get data from Spreadsheet
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(sheetId, range);
                cts = new CancellationTokenSource();
                var requestTask = request.ExecuteAsync(cts.Token);
                requestTask.Wait();
                cts = null;
                if (requestTask.IsCanceled)
                {
                    vaProxy.SetBoolean("ReadingGoogleSheet", false);
                    return;
                }

                ValueRange response = requestTask.Result;
                var values = response.Values;
                if (values != null && values.Count > 0)
                {
                    // We got some data. Data will be passed back to Voice Attack in two ways.
                    // 1) A variable named "SheetRange" will contain all values in the specified range Comma Delimted, with each row on its own line
                    // 2) Each individual value will be in a variable named like "SheetRange[<R>][<C>]" where "<R>" is the Row Number (0 based) and "<C>" is the Column Number (0 based).
                    var sb = new StringBuilder();
                    for (int r = 0; r < values.Count; r++)
                    {
                        var newLine = true;
                        for (int c = 0; c < values[r].Count; c++)
                        {
                            if (!newLine)
                            {
                                sb.Append(", ");
                            }

                            newLine = false;
                            sb.Append(values[r][c]);

                            // Alright, store the data in a variable named "SheetRange[<R>][<C>]" where "<R>" is the Row Number (0 based) and "<C>" is the Column Number (0 based).
                            var cellName = string.Format("SheetRangeRow{0}Col{1}", r, c);
                            vaProxy.SetText(cellName, values[r][c].ToString());
                            voiceAttackVariables.Add(cellName); // Keep a list of variables so that we can clear them on next call.
                        }

                        sb.AppendLine();
                    }

                    // Put all data into "SheetRange"
                    vaProxy.SetText("SheetRange", sb.ToString());
                    voiceAttackVariables.Add("SheetRange"); // Include this variable into the list for clearing
                }
            }

            vaProxy.SetBoolean("ReadingGoogleSheet", false);
        }
    }
}
