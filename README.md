# qlik-chart-reorienter

Tool for setting orientation of combo charts to vertical. The tool will connect to an engine on localhost using certificates which means that the tool must be run on a node where the Qlik Sense Engine is running. The tool can either scan one particular app or all apps in the system.

```
Usage:    QlikChartReorienter.exe (--all | --app <appId>) [--mode <mode>] [--commit]
          <mode> ::= ComboChartOrientation | ComboChartBarAxis
Modes:    ComboChartOrientation - Force combo chart orientations to vertical.
          ComboChartBarAxis     - Force combo chart bar axis to primary.
Examples: QlikChartReorienter.exe --all --commit
          QlikChartReorienter.exe --app 9183fa90-69f6-4864-862e-d6ff75865fb0 --commit
          QlikChartReorienter.exe --all --mode ComboChartBarAxis --commit
```
