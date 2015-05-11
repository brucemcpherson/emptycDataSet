# VBA Project: **emptycDataSet**
## VBA Module: **[googleWireExample](/scripts/googleWireExample.vba "source is here")**
### Type: StdModule  

This procedure list for repo (emptycDataSet) was automatically created on 5/11/2015 1:03:16 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in googleWireExample

---
VBA Procedure: **testWorkBookImport**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub testWorkBookImport()*  

**no arguments required for this procedure**


---
VBA Procedure: **testPublicWorkBookImport**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub testPublicWorkBookImport()*  

**no arguments required for this procedure**


---
VBA Procedure: **importGoogleWorkbook**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function importGoogleWorkbook(key As String, Optional deleteAllSheetsFirst As Boolean = False, Optional replaceConflictingSheets = True, Optional oauthNeeded As Boolean = False, Optional headers As Boolean = True) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||
deleteAllSheetsFirst|Boolean|True| False|
replaceConflictingSheets|Variant|True||
oauthNeeded|Boolean|True| False|
headers|Boolean|True| True|


---
VBA Procedure: **getData**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub getData(sheetJob As cJobject, authHeader As String, accessToken As String, Optional headers As Boolean = True)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
authHeader|String|False||
accessToken|String|False||
headers|Boolean|True| True|


---
VBA Procedure: **getSchema**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getSchema(url As String, authHeader As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||
authHeader|String|False||


---
VBA Procedure: **getSheetsInSchema**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getSheetsInSchema(schemaJob As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
schemaJob|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **googleWireExample**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub googleWireExample()*  

**no arguments required for this procedure**
