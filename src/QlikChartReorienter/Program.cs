using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Qlik.Engine;
using Qlik.Engine.Communication;
using Qlik.Sense.Client.Visualizations;
using Qlik.Sense.Client.Visualizations.Components;

namespace QlikChartReorienter
{
	class Program
    {
        private static StreamWriter _logStreamWriter;

        private static void Write(string message)
        {
            Console.Write(message);
            _logStreamWriter?.Write(message);
        }

        private static void WriteLine(string message)
        {
            Write(message + Environment.NewLine);
        }
		
        static void Main(string[] args)
        {
            var now = DateTime.Now;
            var logFileName = "qlik-chart-reorient-log-" + now.ToString("yyyy-MM-ddTHH-mm-ss") + ".log";
            Console.WriteLine("Writing to log file: " + logFileName);
            _logStreamWriter = new StreamWriter(logFileName,false);
            var commitChanges = false;
            if (args.Length > 0)
                commitChanges = args.First() == "commit";

            if (!commitChanges)
                WriteLine("Dry run only. No changes will be committed.");

            var location = Location.FromUri(@"https:\\localhost");
            var certs = CertificateManager.LoadCertificateFromStore();
			location.AsDirectConnection("INTERNAL", "sa_api", certs, false);

            var appIds = location.GetAppIdentifiers().ToArray();
            WriteLine($"Scanning {appIds.Length} apps...");

            var reoriented = new List<IAppIdentifier>();
            var n = 0;
            foreach (var appIdentifier in appIds)
            {
                n++;
                WriteLine($"({n}/{appIds.Length}) - Scanning \"{appIdentifier.AppName}\" ({appIdentifier.AppId})");
                if (ReorientForApp(location, appIdentifier, commitChanges))
                    reoriented.Add(appIdentifier);
            }

            WriteLine($"Reoriented charts in the following {reoriented.Count} apps:");
            foreach (var appIdentifier in reoriented)
            {
                WriteLine($"\t{appIdentifier.AppName} ({appIdentifier.AppId})");
            }

            if (commitChanges)
            {
                WriteLine("*** Changes applied! ***");
            }
            else
            {
                WriteLine("Dry run only. No changes applied. Use argument \"commit\" to apply changes.");
            }
        }

        private static bool ReorientForApp(ILocation location, IAppIdentifier appIdentifier, bool commitChanges)
        {
            using (var app = location.App(appIdentifier, session:Session.Random, noData:true))
            {
                var infos = app.GetAllInfos();
                var combocharts = infos.Where(i => i.Type == "combochart").ToArray();
                WriteLine($"\t  |- Found {combocharts.Length} combo charts");
                var modified = 0;
                foreach (var combochart in combocharts)
                {
                    var o = app.GetGenericObject(combochart.Id);
                    var extendsId = o.Properties.ExtendsId;
                    if (!string.IsNullOrEmpty(extendsId))
                        o = app.GetGenericObject(extendsId);

                    var p = o.GetProperties().As<CombochartProperties>();
                    if (p.Orientation == Orientation.Horizontal)
                    {
                        WriteLine($"\t  |- Changing {(string.IsNullOrEmpty(extendsId) ? "chart" : "master object")} to vertical: " + p.Info.Id);
                        p.Orientation = Orientation.Vertical;
                        if (commitChanges)
                        {
                            try
                            {
                                o.SetProperties(p);
                                modified++;
                            }
                            catch
                            {
                                WriteLine("\t  |- Failed to set properties.");
                            }
                        }
                    }
                }
                WriteLine($"\t  \\- Modified {modified} combo charts.");

                // app.DoSave();
                return modified > 0;
            }
        }
    }
}
