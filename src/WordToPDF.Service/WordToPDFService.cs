using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Net.Http;
using System.Reflection;
using Serilog;
using WordToPDF.Library;


namespace WordToPDF.Service
{
    public class WordToPDFService
    {
        protected System.Timers.Timer _processTimer;
        protected bool _processLock;
        protected List<IDocumentQueue> _documentQueues;
        protected string _heartbeatEndpoint;

        public WordToPDFService()
        {
            Log.Debug("Construction of WordToPDFService");
            _documentQueues = new List<IDocumentQueue>();
        }

        public void Start()
        {
            Log.Warning($"Starting {Assembly.GetEntryAssembly().GetName()} on {Environment.MachineName}");

            // enable folder watching if it is configured
            if (ConfigurationManager.AppSettings["FolderQueue:Enabled"] == "true")
            {
                _documentQueues.Add(new FolderDocumentQueue(ConfigurationManager.AppSettings["FolderQueue:WatchPath"], ConfigurationManager.AppSettings["FolderQueue:SourceName"]));
            }

            // enable database table watching if it is configured
            if (ConfigurationManager.AppSettings["DatabaseQueue:Enabled"] == "true")
            {
                ConnectionStringSettings connectionString = ConfigurationManager.ConnectionStrings[ConfigurationManager.AppSettings["DatabaseQueue:ConnectionStringName"]];
                IDbConnection connection = DbProviderFactories.GetFactory(connectionString.ProviderName).CreateConnection();
                connection.ConnectionString = connectionString.ConnectionString;
                DatabaseDocumentQueue databaseDocumentQueue = new DatabaseDocumentQueue(connection, ConfigurationManager.AppSettings["DatabaseQueue:TableName"]);
                _documentQueues.Add(databaseDocumentQueue);
            }

            // configure the UptimeRobot Heartbeat client and endpoint
            _heartbeatEndpoint = ConfigurationManager.AppSettings["UptimeRobot:HeartbeatEndpoint"];

            _processLock = false;
            int processTimerDelay = int.Parse(ConfigurationManager.AppSettings["ProcessTimer:DelayMS"]);
            _processTimer = new System.Timers.Timer(processTimerDelay) { AutoReset = true };
            _processTimer.Elapsed += (sender, eventArgs) => ProcessDocuments();
            _processTimer.Start();
        }

        public void Stop()
        {
            Log.Warning($"Stopping {Assembly.GetEntryAssembly().GetName()} on {Environment.MachineName}");
            _processTimer.Stop();
            _processTimer = null;
        }

        public void ProcessDocuments()
        {
            ConvertService convertService = null;

            if(_processLock)
            {
                Log.Debug("Re-entry into service process, still busy.");
                return;
            }

            _processLock = true;

            if (!String.IsNullOrEmpty(_heartbeatEndpoint))
            {
                // synchronous use of async. reading repsonse string as async synchronizes the call
                using (var client = new HttpClient())
                {
                    var response = client.GetAsync(_heartbeatEndpoint).Result;
                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = response.Content;
                        responseContent.ReadAsStringAsync();
                        Log.Debug("UptimeRobot heartbeat link touched");
                    }
                }
            }

            try
            {
                Log.Debug("Processing pending documents.");
                foreach (IDocumentQueue documentQueue in _documentQueues)
                {
                    while (documentQueue.Count() > 0)
                    {
                        if (convertService == null)
                        {
                            convertService = new ConvertService();
                            convertService.Initialize();
                        }
                        DocumentTarget documentTarget = documentQueue.NextDocument();
                        Log.Information($"Generating PDF for {documentQueue.SourceName()}: {documentTarget.Id}: {documentTarget.InputFile}");
                        try
                        {
                            string outFile = documentTarget.OutputFile;
                            documentTarget.ResultCode = convertService.Convert(documentTarget.InputFile, ref outFile);
                            documentTarget.OutputFile = outFile;
                        }
                        catch (Exception e)
                        {
                            Log.Warning("Conversion exception {0}", e.Message);
                            documentTarget.ResultCode = (int)ExitCode.InternalError;
                        }
                        Log.Information($"Generated PDF for {documentTarget.Id} as {documentTarget.OutputFile} with Result {documentTarget.ResultCode}");
                        documentQueue.CompleteDocument(documentTarget);
                    }
                }
            }
            catch (Exception e)
            {
                Log.Error("Unhandled exception in ProcessDocuments {0}", e);
            }
            _processLock = false;
        }
    }
}
