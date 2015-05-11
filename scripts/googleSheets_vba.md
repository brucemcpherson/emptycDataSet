# VBA Project: **emptycDataSet**
## VBA Module: **[googleSheets](/scripts/googleSheets.vba "source is here")**
### Type: StdModule  

This procedure list for repo (emptycDataSet) was automatically created on 5/11/2015 12:43:47 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in googleSheets

---
VBA Procedure: **testWorkBookImportPublicSheets**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub testWorkBookImportPublicSheets()*  

**no arguments required for this procedure**


---
VBA Procedure: **testWorkBookImportNewSheets**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub testWorkBookImportNewSheets()*  

**no arguments required for this procedure**


---
VBA Procedure: **testSelectedImportNewSheets**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub testSelectedImportNewSheets()*  

**no arguments required for this procedure**


---
VBA Procedure: **importGoogleWorkbookNewSheets**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function importGoogleWorkbookNewSheets(key As String, Optional deleteAllSheetsFirst As Boolean = False, Optional replaceConflictingSheets = True, Optional oauthNeeded As Boolean = False, Optional headers As Boolean = False, Optional listOfSheets = vbNullString) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||
deleteAllSheetsFirst|Boolean|True| False|
replaceConflictingSheets|Variant|True||
oauthNeeded|Boolean|True| False|
headers|Boolean|True| False|
listOfSheets|Variant|True||


---
VBA Procedure: **getSchemaNewSheets**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getSchemaNewSheets(url As String, authHeader As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||
authHeader|String|False||


---
VBA Procedure: **getDataNewSheets**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub getDataNewSheets(sheetJob As cJobject, authHeader As String, accessToken As String, Optional headers As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
authHeader|String|False||
accessToken|String|False||
headers|Boolean|True| True|


---
VBA Procedure: **getSheetsInSchemaNewSheets**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getSheetsInSchemaNewSheets(schemaJob As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
schemaJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **getJobjectFromWire**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getJobjectFromWire(jsonData As String, Optional t As eDeserializeType = eDeserializeNormal, Optional url As String = vbNullString) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jsonData|String|False||
t|eDeserializeType|True| eDeserializeNormal|
url|String|True| vbNullString|


---
VBA Procedure: **deleteAllTheSheets**  
Type: **Function**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function deleteAllTheSheets() As Worksheet*  

**no arguments required for this procedure**


---
VBA Procedure: **deleteSomeOfTheSheets**  
Type: **Function**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function deleteSomeOfTheSheets(job As cJobject) As Worksheet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **hackTheLastSheet**  
Type: **Function**  
Returns: **Worksheet**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function hackTheLastSheet() As Worksheet*  

**no arguments required for this procedure**


---
VBA Procedure: **googleWireExample**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub googleWireExample()*  

**no arguments required for this procedure**
