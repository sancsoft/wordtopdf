using System;
using System.Collections.Generic;
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

        public WordToPDFService()
        {
            Log.Debug("Construction of WordToPDFService");
            _documentQueues = new List<IDocumentQueue>();
        }

        public void Start()
        {
            Log.Warning($"Starting {Assembly.GetEntryAssembly().GetName()} on {Environment.MachineName}");

            _documentQueues.Add(new FolderDocumentQueue(@"C:\Users\mterry\git\wordtopdf\src\TestFolder", "Test Folder"));

            _processLock = false;
            _processTimer = new System.Timers.Timer(10000) { AutoReset = true };
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
            int resultCode = -1;

            if(_processLock)
            {
                Log.Debug("Re-entry into service process, still busy.");
                return;
            }

            _processLock = true;
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
                            resultCode = convertService.Convert(documentTarget.InputFile, documentTarget.OutputFile);
                        }
                        catch (Exception e)
                        {
                            Log.Warning("Conversion exception {0}", e.Message);
                            documentTarget.ResultCode = -1;
                        }
                        Log.Information($"Generated PDF for {documentTarget.Id} as {documentTarget.OutputFile} with Result {resultCode}");
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
