using System;
using System.ComponentModel.Design;
using System.Linq;
using Qlik.Engine;
using Qlik.Sense.Client.Visualizations;
using Qlik.Sense.Client.Visualizations.Components;

namespace QlikChartReorienter
{
	class Program
	{
		static void Main(string[] args)
        {
            var location = Location.FromUri(@"http:\\localhost");
			location.AsDirectConnectionToPersonalEdition();

            var appIds = location.AppsWithNameOrDefault("ComboChartReorienter");
            foreach (var appIdentifier in appIds)
            {
                using (var app = location.App(appIdentifier))
                {
                    var infos = app.GetAllInfos();
                    var combocharts = infos.Where(i => i.Type == "combochart").ToArray();
                    Console.WriteLine($"Found {combocharts.Length} combo charts");
                    foreach (var combochart in combocharts)
                    {
                        var o = app.GetGenericObject(combochart.Id);
                        var extendsId = o.Properties.ExtendsId;
                        if (!string.IsNullOrEmpty(extendsId))
                            o = app.GetGenericObject(extendsId);

                        var p = o.GetProperties().As<CombochartProperties>();
                        if (p.Orientation == Orientation.Horizontal)
                        {
                            Console.WriteLine($"Changing {(string.IsNullOrEmpty(extendsId) ? "chart" : "master object")} to vertical: " + p.Info.Id);
                            p.Orientation = Orientation.Vertical;
                            // o.SetProperties(p);
                        }
                    }

                    app.DoSave();
                }
            }
        }
	}
}
