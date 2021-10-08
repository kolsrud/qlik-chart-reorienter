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
        private static bool _commitChanges = false;
        private static bool _scanAllApps = false;
        private static string _scanSpecificAppId = null;

        private static void Write(string message)
        {
            Console.Write(message);
            _logStreamWriter?.Write(message);
        }

        private static void WriteLine(string message)
        {
            Write(message + Environment.NewLine);
        }


        private static void ParseArgs(string[] args)
        {
            if (args.Length == 0)
                PrintUsage();

            _commitChanges = args.Contains("--commit");
            _scanAllApps = args.Contains("--all");
            _scanSpecificAppId = GetArg(args, "--app");
            if (!_scanAllApps && (_scanSpecificAppId == null))
            {
                Error("Exactly one of the flags --all and --app must be defined.");
            }
            if (_scanAllApps && (_scanSpecificAppId != null))
            {
                Error("Illegal combination of flags. The flags --all and --app are mutually exclusive.");
            }
        }

        private static void Error(string msg)
        {
            Console.WriteLine("--------------------------------------------------------------------------");
            Console.WriteLine("--- Error: " + msg);
            Console.WriteLine("--------------------------------------------------------------------------");
            PrintUsage();
        }

        private static string GetArg(string[] args, string arg)
        {
            var i = Array.IndexOf(args, arg);
            if (i == -1)
                return null;

            if (i == arg.Length - 1)
            {
                Error("No argument found for app flag.");
            }

            var result = args[i + 1];
            if (!Guid.TryParse(result, out _))
            {
                Error($"Unable to parse argument \"{args[i+1]}\" as Guid.");
            }

            return result;
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage:   " + System.AppDomain.CurrentDomain.FriendlyName + " (--all | --app <appId>) [--commit]");
            Console.WriteLine("Example: " + System.AppDomain.CurrentDomain.FriendlyName + " --all --commit");
            Console.WriteLine("         " + System.AppDomain.CurrentDomain.FriendlyName + " --app 9183fa90-69f6-4864-862e-d6ff75865fb0 --commit");
            Environment.Exit(0);
        }

        static void Main(string[] args)
        {
            ParseArgs(args);

            var now = DateTime.Now;
            var logFileName = "qlik-chart-reorient-log-" + now.ToString("yyyy-MM-ddTHH-mm-ss") + ".log";
            Console.WriteLine("Writing to log file: " + logFileName);
            _logStreamWriter = new StreamWriter(logFileName,false);

            if (!_commitChanges)
                WriteLine("Dry run only. No changes will be committed.");

            var location = Location.FromUri(@"https:\\localhost");
            var certs = CertificateManager.LoadCertificateFromStore();
			location.AsDirectConnection("INTERNAL", "sa_api", certs, false);

            AppIdentifier[] appIds;
            appIds = _scanAllApps ? location.GetAppIdentifiers().OfType<AppIdentifier>().ToArray() : new []{new AppIdentifier{AppId=_scanSpecificAppId}};

            WriteLine($"Scanning {appIds.Length} apps...");

            var reoriented = new List<IAppIdentifier>();
            var n = 0;
            foreach (var appIdentifier in appIds)
            {
                n++;
                WriteLine($"({n}/{appIds.Length}) - Scanning \"{appIdentifier.AppName ?? "<App name unknown>"}\" ({appIdentifier.AppId})");
                if (ReorientForApp(location, appIdentifier, _commitChanges))
                    reoriented.Add(appIdentifier);
            }

            WriteLine($"Reoriented charts in the following {reoriented.Count} apps:");
            foreach (var appIdentifier in reoriented)
            {
                WriteLine($"\t{appIdentifier.AppName} ({appIdentifier.AppId})");
            }

            if (_commitChanges)
            {
                WriteLine("*** Changes applied! ***");
            }
            else
            {
                WriteLine("Dry run only. No changes applied. Use argument \"commit\" to apply changes.");
            }
        }

        private static bool ReorientForApp(ILocation location, AppIdentifier appIdentifier, bool commitChanges)
        {
            try
            {
                using (var app = location.App(appIdentifier, session: Session.Random, noData: true))
                {
                    if (appIdentifier.AppName == null)
                    {
                        appIdentifier.AppName = app.GetAppProperties().Title;
                    }

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
                            Write(
                                $"\t  |- Changing {(string.IsNullOrEmpty(extendsId) ? "chart" : "master object")} to vertical: " +
                                p.Info.Id);
                            p.Orientation = Orientation.Vertical;
                            try
                            {
                                if (commitChanges)
                                {
                                    WriteLine("");
                                    o.SetProperties(p);
                                }
                                else
                                {  
                                    WriteLine($" (Dry run only. No changes applied.)");
                                }

                                modified++;
                            }
                            catch
                            {
                                WriteLine("");
                                WriteLine("\t  |- Failed to set properties.");
                            }
                        }
                    }

                    WriteLine($"\t  \\- Modified {modified} combo charts.");

                    // app.DoSave();
                    return modified > 0;
                }
            }
            catch (Exception e)
            {
                WriteLine("Error encountered: " + e);
            }

            return false;
        }
    }
}
