# VBA Project: **emptycDataSet**
## VBA Module: **[restLibrary](/scripts/restLibrary.vba "source is here")**
### Type: StdModule  

This procedure list for repo (emptycDataSet) was automatically created on 5/11/2015 12:42:11 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in restLibrary

---
VBA Procedure: **generalQuery**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function generalQuery(sheetName As String, libEntry As String, queryString As String, Optional breport As Boolean = True, Optional queryCanBeBlank As Boolean = False, Optional appendQuery As String = vbNullString) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetName|String|False||
libEntry|String|False||
queryString|String|False||
breport|Boolean|True| True|
queryCanBeBlank|Boolean|True| False|
appendQuery|String|True| vbNullString|


---
VBA Procedure: **generalDataSetQuery**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function generalDataSetQuery(sheetName As String, libEntry As String, colName As String, Optional breport As Boolean = True, Optional queryCanBeBlank As Boolean = False, Optional appendQuery As String = vbNullString, Optional collectionNeeded As Boolean = True) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetName|String|False||
libEntry|String|False||
colName|String|False||
breport|Boolean|True| True|
queryCanBeBlank|Boolean|True| False|
appendQuery|String|True| vbNullString|
collectionNeeded|Boolean|True| True|


---
VBA Procedure: **generalReport**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function generalReport(cr As cRest, breport As Boolean) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cr|[cRest](/libraries/cRest_cls.md "cRest")|False||
breport|Boolean|False||


---
VBA Procedure: **getRestLibrary**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getRestLibrary() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **createRestLibrary**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function createRestLibrary() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **restQuery**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function restQuery(Optional sheetName As String = vbNullString, Optional sEntry As String = vbNullString, Optional sQuery As String = vbNullString, Optional sQueryColumn As String = vbNullString, Optional sRestUrl As String = vbNullString, Optional sResponseResults As String = vbNullString, Optional bTreeSearch As Boolean = True, Optional bPopulate As Boolean = True, Optional bClearMissing As Boolean = True, Optional complain As Boolean = True, Optional queryCanBeBlank As Boolean = False, Optional sFix As String = vbNullString, Optional user As String = vbNullString, Optional pass As String = vbNullString, Optional append As Boolean = False, Optional stampQuery As String = vbNullString, Optional appendQuery As String = vbNullString, Optional collectionNeeded As Boolean = True, Optional postData As String = vbNullString, Optional resultsFormat As erResultsFormat = erUnknown) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sheetName|String|True| vbNullString|
sEntry|String|True| vbNullString|
sQuery|String|True| vbNullString|
sQueryColumn|String|True| vbNullString|
sRestUrl|String|True| vbNullString|
sResponseResults|String|True| vbNullString|
bTreeSearch|Boolean|True| True|
bPopulate|Boolean|True| True|
bClearMissing|Boolean|True| True|
complain|Boolean|True| True|
queryCanBeBlank|Boolean|True| False|
sFix|String|True| vbNullString|
user|String|True| vbNullString|
pass|String|True| vbNullString|
append|Boolean|True| False|
stampQuery|String|True| vbNullString|
appendQuery|String|True| vbNullString|
collectionNeeded|Boolean|True| True|
postData|String|True| vbNullString|
resultsFormat|erResultsFormat|True| erUnknown|


---
VBA Procedure: **abandonType**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function abandonType(sEntry, qType As erRestType, targetType As erRestType) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sEntry|Variant|False||
qType|erRestType|False||
targetType|erRestType|False||


---
VBA Procedure: **whichType**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function whichType(t As erRestType) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
t|erRestType|False||


---
VBA Procedure: **createHeadingsFromKeys**  
Type: **Function**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function createHeadingsFromKeys(job As cJobject, ds As cDataSet) As cDataSet*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|False||


---
VBA Procedure: **getAndMakeJobjectFromXML**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getAndMakeJobjectFromXML(url As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||


---
VBA Procedure: **makeJobjectFromXML**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function makeJobjectFromXML(theXml As String, Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
theXml|String|False||
complain|Boolean|True| True|


---
VBA Procedure: **getAndMakeJobjectAuto**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function getAndMakeJobjectAuto(url As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
url|String|False||
