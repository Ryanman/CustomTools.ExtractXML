# CustomTools.ExtractXML
Extracts XML from a .xls file, includes a subfolder feature.


I initially created this tool as part of a client request that got scrapped. It can't be totally useless, right?

##How to Use
* Open in VS. This project uses .NET 4.5 but there's no reason really, that's just how the cookie crumbles. 
* Compile (should take less than a second) and then hit run.
* The application will first ask for the folder with the XML files in it. Do the full directory.
* Then it will ask for the location of the schema file. 
  * You can only validate against one schema at a time. Include the schema extension, probably .xsd
* You can chooose the name of the validation report. Make sure it follows filename rules.
  * If you don't choose one and just press enter, the tool will create "ValidationReport.txt"
* The tool will show you a (hideous) progress bar, and give you a list of files that have failed to validate against the schema
* The tool will also output the validation report, which contains more details about specific errors
