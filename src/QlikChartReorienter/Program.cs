using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using Qlik.Engine;
using Qlik.Engine.Communication;
using Qlik.Sense.Client;
using Qlik.Sense.Client.Visualizations;
using Qlik.Sense.Client.Visualizations.Components;

namespace QlikChartReorienter
{
	class Program
    {
        private enum Mode
        {
            ComboChartOrientation,
            ComboChartBarAxis
        }

        private static StreamWriter _logStreamWriter;
        private static string _url = null;
        private static string _certs = null;
        private static SecureString _certPwd = null;
        private static bool _commitChanges = false;
        private static bool _scanAllApps = false;
        private static string _scanSpecificAppId = null;
        private static Mode _mode = Mode.ComboChartOrientation;

        private static void Write(string message)
        {
            Console.Write(message);
            _logStreamWriter?.Write(message);
            _logStreamWriter?.Flush();
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
            _certs = GetArg(args, "--certs");
            var certPwd = GetArg(args, "--certPwd");
            if (certPwd != null)
            {
                _certPwd = MakeSecureString(certPwd);
            }

            _url = GetArg(args, "--url");
            if (_url == null)
            {
                _url = "localhost";
            }
            else if (_certs == null)
            {
                _certs = ".";
            }

            _scanSpecificAppId = GetArgGuid(args, "--app");
            var modeStr = GetArg(args, "--mode");
            switch (modeStr?.ToLower() ?? Mode.ComboChartOrientation.ToString().ToLower())
            {
                case "combochartorientation":
                    _mode = Mode.ComboChartOrientation;
                    break;
                case "combochartbaraxis":
                    _mode = Mode.ComboChartBarAxis;
                    break;
                default:
                    Error($"Unknown operations mode {modeStr}");
                    break;
            }

            if (!_scanAllApps && (_scanSpecificAppId == null))
            {
                Error("Exactly one of the flags --all and --app must be defined.");
            }
            if (_scanAllApps && (_scanSpecificAppId != null))
            {
                Error("Illegal combination of flags. The flags --all and --app are mutually exclusive.");
            }
        }

        private static SecureString MakeSecureString(string certPwd)
        {
            var result = new SecureString();
            foreach (var c in certPwd)
            {
                result.AppendChar(c);
            }
            return result;
        }

        private static void Error(string msg)
        {
            Console.WriteLine("--------------------------------------------------------------------------");
            Console.WriteLine("--- Error: " + msg);
            Console.WriteLine("--------------------------------------------------------------------------");
            PrintUsage();
        }

        private static string GetArgGuid(string[] args, string argName)
        {
            var result = GetArg(args, argName);
            if (result == null)
                return null;

            if (!Guid.TryParse(result, out _))
            {
                Error($"Unable to parse argument \"{result}\" as Guid.");
            }

            return result;
        }

        private static string GetArg(string[] args, string argName)
        {
            var i = Array.IndexOf(args, argName);
            if (i == -1)
                return null;

            if (i == args.Length - 1)
            {
                Error($"No argument found for flag flag {argName}.");
            }
            return args[i + 1];
        }

        static void PrintUsage()
        {
            var exeName = System.AppDomain.CurrentDomain.FriendlyName;
            Console.WriteLine($"Usage:    {exeName} (--all | --app <appId>) [--mode <mode>] [--commit]");
            Console.WriteLine("                                  [--url <url>] [--certs <path>] [--certPwd <pwd>]");
            Console.WriteLine($"          <mode> ::= ComboChartOrientation | ComboChartBarAxis");
            Console.WriteLine($"Modes:    ComboChartOrientation - Force combo chart orientations to vertical.");
            Console.WriteLine($"          ComboChartBarAxis     - Force combo chart bar axis to primary.");
            Console.WriteLine($"Examples: {exeName} --all --commit");
            Console.WriteLine($"          {exeName} --url my.server.com --certs C:\\path\\to\\certs --certPwd qwerty --all --commit");
            Console.WriteLine($"          {exeName} --app 9183fa90-69f6-4864-862e-d6ff75865fb0 --commit");
            Console.WriteLine($"          {exeName} --all --mode ComboChartBarAxis --commit");
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

            var location = Location.FromUri(@"https:\\" + _url);
            var certs = _certs == null ? CertificateManager.LoadCertificateFromStore() :
                                               _certPwd == null ? CertificateManager.LoadCertificateFromDirectory(_certs) :
                                               CertificateManager.LoadCertificateFromDirectory(_certs, _certPwd);
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

                    var masterObjectList = app.GetMasterObjectList();
                    var masterComboCharts = masterObjectList.Layout.AppObjectList.Items.Where(i => i.Data.Visualization == "combochart").ToArray();
                    WriteLine($"\t  |- Found {masterComboCharts.Length} master item combo charts.");
                    var modified = ReorientCharts(app, masterComboCharts.Select(chart => chart.Info), _mode, commitChanges);
                    app.DestroyGenericSessionObject(masterObjectList.Id);

                    var infos = app.GetAllInfos();
                    var combocharts = infos.Where(i => i.Type == "combochart").ToArray();
                    WriteLine($"\t  |- Found {combocharts.Length} combo charts");
                    modified += ReorientCharts(app, combocharts, _mode, commitChanges);

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

        private static int ReorientCharts(IApp app, IEnumerable<NxInfo> infos, Mode mode, bool commitChanges)
        {
            var modified = 0;
            foreach (var combochart in infos)
            {
                var o = app.GetGenericObject(combochart.Id);
                var extendsId = o.Properties.ExtendsId;
                if (!string.IsNullOrEmpty(extendsId))
                    o = app.GetGenericObject(extendsId);

                modified += PerformModification(o, mode, commitChanges);
            }

            return modified;
        }

        private static int PerformModification(GenericObject o, Mode mode, bool commitChanges)
        {
            switch (mode)
            {
                case Mode.ComboChartOrientation: return ForceComboChartOrientation(o, commitChanges);
                case Mode.ComboChartBarAxis: return ForceComboChartBarAxis(o, commitChanges);
                default:
                    Error($"No modification method defined for mode {mode}");
                    return 0;
            }
        }

        private static int ForceComboChartOrientation(IGenericObject o, bool commitChanges)
        {
            var p = o.GetProperties().As<CombochartProperties>();
            if (p.Orientation == Orientation.Horizontal)
            {
                Write($"\t  |- Changing {(string.IsNullOrEmpty(p.ExtendsId) ? "chart" : "master object")} to vertical: " +  p.Info.Id);
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

                    return 1;
                }
                catch
                {
                    WriteLine("");
                    WriteLine("\t  |- Failed to set properties.");
                }
            }

            return 0;
        }

        private static int ForceComboChartBarAxis(IGenericObject o, bool commitChanges)
        {
            var p = o.GetProperties().As<CombochartProperties>();
            var defs = p.HyperCubeDef.Measures.Select(m => m.Def);
            foreach (var def in defs)
            {
                if (def.Series.Type == CombochartSeriesType.Bar && def.Series.Axis != 0)
                {
                    Write($"\t  |- Changing {(string.IsNullOrEmpty(p.ExtendsId) ? "chart" : "master object")} bar axis to primary: " + p.Info.Id);
                    def.Series.Axis = 0;
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

                        return 1;
                    }
                    catch
                    {
                        WriteLine("");
                        WriteLine("\t  |- Failed to set properties.");
                    }
                }
            }

            return 0;
        }
    }
}
