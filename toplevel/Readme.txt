#################################################
#  SQA Automation Software guide for developers #
#################################################

1) Creating New Project
   - Right-click Solution 'toplevel'
   - Add -> New Project
   - Select "Class Library"
   - Name the project (Naming convention : Capital letter for start letter, eg : ExampleTest)
   - Rename Class1.cs to proper class name (Eg : Example.cs)
   - In new project, add references as follows
     - References -> Add Reference
	 - Tick PluginContracts, SAAL, System.Windows.Forms.Ribbon from Solution
	 - Tick System.Windows.Form from Assemblies
   - Copy from ExampleTest project (Example.cs) inside the namespace and change the class name accordingly
   - Add button in mainUI form and add following line to call the plugin and populate ribbon tab
     - m_plugin = _Plugins["Example Test"];
	 - this.ribbon1.Tabs.Add(m_plugin.EquipmentSetting);

 
2) Emgu related files
   - Make sure to copy this directory to run directory (tessdata, X86 and X64)

#################################################
#  TroubleShooting                              #
#################################################

1) Getting Error "The OLE DB provider "Microsoft.ACE.OLEDB.12.0" has not been registered". when importing excel file.
   - Install 2007 Office System Driver from https://www.microsoft.com/en-us/download/confirmation.aspx?id=23734