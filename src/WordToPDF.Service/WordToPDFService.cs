using System;
using System.Reflection;
using Serilog;

namespace WordToPDF.Service
{
    public class WordToPDFService
    {
        protected System.Timers.Timer _processTimer;
        protected bool _processLock;

        public WordToPDFService()
        {
            Log.Debug("Construction of WordToPDFService");
        }

        public void Start()
        {
            Log.Warning($"Starting {Assembly.GetEntryAssembly().GetName()} on {Environment.MachineName}");
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
            if(_processLock)
            {
                Log.Debug("Re-entry into service process, still busy.");
                return;
            }

            _processLock = true;
            try
            {

            }
            catch (Exception e)
            {
                Log.Error("Unhandled exception in ProcessDocuments {0}", e);
            }
            _processLock = false;
        }
    }
}
