using System.Reflection;
using Serilog;
using Topshelf;

namespace WordToPDF.Service
{
    class Program
    {
        static void Main(string[] args)
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Verbose()
                .Enrich.WithProperty("appname", Assembly.GetEntryAssembly().GetName().Name)
                .CreateLogger();

            AssemblyName assemblyName = Assembly.GetEntryAssembly().GetName();
            Log.Information(assemblyName.Name + " Version " + assemblyName.Version.ToString());
            Log.Information("Copyright (c) 2020 - Sanctuary Software Studio, Inc.");

            HostFactory.Run(x =>
            {
                x.Service<WordToPDFService>(s =>
                {
                    s.ConstructUsing(name => new WordToPDFService());
                    s.WhenStarted(tc => tc.Start());
                    s.WhenStopped(tc => tc.Stop());
                });
                x.RunAsLocalSystem();

                x.SetDescription("WordToPDF Service");
                x.SetDisplayName("WordToPDF");
                x.SetServiceName("WordToPDF");
                x.UseSerilog();
            });

        }
    }
}
