# VBA Project: emptycDataSet
This cross reference list for repo (emptycDataSet) was automatically created on 4/18/2017 10:33:04 AM by VBAGit.For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")
You can see [library and dependency information here](dependencies.md)

###Below is a cross reference showing which modules and procedures reference which others
*module*|*proc*|*referenced by module*|*proc*
---|---|---|---
cBrowser||googleSheets|getSchemaNewSheets
cBrowser||googleSheets|getDataNewSheets
cBrowser||googleSheets|googleWireExample
cCell||cDataRow|create
cDataColumn||cDataSet|create
cDataRow||cDataSet|create
cDataSet||googleSheets|googleWireExample
cHeadingRow||cDataSet|Class_Initialize
cJobject||googleSheets|getJobjectFromWire
cJobject||googleSheets|importGoogleWorkbookNewSheets
cJobject||googleSheets|getSheetsInSchemaNewSheets
cOauth2||oauthExamples|getGoogled
cregXLib||regXLib|rxMakeRxLib
cRest||restLibrary|restQuery
cStringChunker||cJobject|recurseSerialize
cStringChunker||cJobject|unSplitToString
cStringChunker||cJobject|serialize
cUAMeasure||UAMeasure|registerUA
regXLib|rxReplace|usefulcJobject|cleanGoogleWire
regXLib|rxReplace|usefulcJobject|hackJSObjectToJSON
regXLib|rxReplace|usefulcJobject|hackJSONPObjectToJSON
UAMeasure|registerUA|restLibrary|restQuery
usefulcJobject|cleanGoogleWire|googleSheets|getJobjectFromWire
usefulcJobject|JSONParse|googleSheets|getSchemaNewSheets
usefulSheetStuff|sheetExists|googleSheets|getDataNewSheets
usefulSheetStuff|sheetExists|googleSheets|deleteSomeOfTheSheets
usefulStuff|isSomething|googleSheets|getSheetsInSchemaNewSheets
usefulStuff|isSomething|googleSheets|getSchemaNewSheets
usefulStuff|makeKey|googleSheets|importGoogleWorkbookNewSheets
