# VBA Project: **emptycDataSet**
## VBA Module: **[cRest](/libraries/cRest.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (emptycDataSet) was automatically created on 5/11/2015 12:42:58 PM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cRest

---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**


---
VBA Procedure: **jObjects**  
Type: **Get**  
Returns: **Collection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get jObjects() As Collection*  

**no arguments required for this procedure**


---
VBA Procedure: **datajObject**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get datajObject() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **jObject**  
Type: **Get**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get jObject(Optional complain As Boolean = True) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
complain|Boolean|True| True|


---
VBA Procedure: **erType**  
Type: **Get**  
Returns: **erRestType**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get erType() As erRestType*  

**no arguments required for this procedure**


---
VBA Procedure: **response**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get response() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **encodedUri**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get encodedUri() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **queryhCell**  
Type: **Get**  
Returns: **[cCell](/libraries/cCell_cls.md "cCell")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get queryhCell() As cCell*  

**no arguments required for this procedure**


---
VBA Procedure: **queryString**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let queryString(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **restUrlStem**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let restUrlStem(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **queryString**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get queryString() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **dset**  
Type: **Get**  
Returns: **[cDataSet](/libraries/cDataSet_cls.md "cDataSet")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get dset() As cDataSet*  

**no arguments required for this procedure**


---
VBA Procedure: **respRootJob**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function respRootJob(job As cJobject) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **stripDots**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function stripDots(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **dotsTail**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function dotsTail(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **isDots**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function isDots(s As String) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **childOrFindJob**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function childOrFindJob(job As cJobject, s As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
s|String|False||


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function init(Optional rData As String = "responsedata.results", Optional et As erRestType = erQueryPerRow, Optional hc As cCell = Nothing, Optional rq As String = vbNullString, Optional ds As cDataSet = Nothing, Optional pop As Boolean = True, Optional pUrl As String = vbNullString, Optional clearmissing As Boolean = True, Optional treesearch As Boolean = False, Optional complain As Boolean = True, Optional sIgnore As String = vbNullString, Optional user As String = vbNullString, Optional pass As String = vbNullString, Optional append As Boolean = False, Optional stampQuery As cCell = Nothing, Optional appendQuery As String = vbNullString, Optional libAccept As String = vbNullString, Optional bWire As Boolean = False, Optional collectionNeeded As Boolean = True, Optional bAlwaysEncode As Boolean = False, Optional timeout As Long = 0, Optional postData As String = vbNullString, Optional resultsFormat As erResultsFormat = erJSON, Optional oa As cOauth2 = Nothing) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rData|String|True| "responsedata.results"|
et|erRestType|True| erQueryPerRow|
hc|[cCell](/libraries/cCell_cls.md "cCell")|True| Nothing|
rq|String|True| vbNullString|
ds|[cDataSet](/libraries/cDataSet_cls.md "cDataSet")|True| Nothing|
pop|Boolean|True| True|
pUrl|String|True| vbNullString|
clearmissing|Boolean|True| True|
treesearch|Boolean|True| False|
complain|Boolean|True| True|
sIgnore|String|True| vbNullString|
user|String|True| vbNullString|
pass|String|True| vbNullString|
append|Boolean|True| False|
stampQuery|[cCell](/libraries/cCell_cls.md "cCell")|True| Nothing|
appendQuery|String|True| vbNullString|
libAccept|String|True| vbNullString|
bWire|Boolean|True| False|
collectionNeeded|Boolean|True| True|
bAlwaysEncode|Boolean|True| False|
timeout|Long|True| 0|
postData|String|True| vbNullString|
resultsFormat|erResultsFormat|True| erJSON|
oa|[cOauth2](/libraries/cOauth2_cls.md "cOauth2")|True| Nothing|


---
VBA Procedure: **executeSingle**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function executeSingle(Optional rurl As String = vbNullString, Optional qry As String = vbNullString, Optional complain As Boolean = True, Optional sFix As String = vbNullString ) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
rurl|String|True| vbNullString|
qry|String|True| vbNullString|
complain|Boolean|True| True|
sFix|String|True| vbNullString|


---
VBA Procedure: **execute**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function execute(Optional qry As String = vbNullString, Optional sFix As String = vbNullString, Optional complain As Boolean = True) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
qry|String|True| vbNullString|
sFix|String|True| vbNullString|
complain|Boolean|True| True|


---
VBA Procedure: **populateOneRow**  
Type: **Function**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function populateOneRow(job As cJobject, dr As cDataRow) As cDataRow*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
dr|[cDataRow](/libraries/cDataRow_cls.md "cDataRow")|False||


---
VBA Procedure: **populateRows**  
Type: **Function**  
Returns: **[cRest](/libraries/cRest_cls.md "cRest")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function populateRows(job As cJobject, Optional complain As Boolean = True) As cRest*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
job|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
complain|Boolean|True| True|


---
VBA Procedure: **getValueFromJo**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Function getValueFromJo(jo As cJobject, originalKey As String) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
jo|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
originalKey|String|False||


---
VBA Procedure: **browser**  
Type: **Get**  
Returns: **[cBrowser](/libraries/cBrowser_cls.md "cBrowser")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get browser() As cBrowser*  

**no arguments required for this procedure**


---
VBA Procedure: **Class_Initialize**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Initialize()*  

**no arguments required for this procedure**
