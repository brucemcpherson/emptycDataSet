# VBA Project: **emptycDataSet**
## VBA Module: **[cBrowser](/libraries/cBrowser.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (emptycDataSet) was automatically created on 4/18/2017 10:33:01 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cBrowser

---
VBA Procedure: **browser**  
Type: **Get**  
Returns: **InternetExplorer**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get browser() As InternetExplorer*  

**no arguments required for this procedure**


---
VBA Procedure: **isOk**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get isOk() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **status**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get status() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **responseHeaders**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get responseHeaders() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **optionURL**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get optionURL() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **successCode**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get successCode() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **deniedCode**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get deniedCode() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **Text**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get Text() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **url**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get url() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **init**  
Type: **Function**  
Returns: **[cBrowser](/libraries/cBrowser_cls.md "cBrowser")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function init() As cBrowser*  

**no arguments required for this procedure**


---
VBA Procedure: **Navigate**  
Type: **Function**  
Returns: **[cBrowser](/libraries/cBrowser_cls.md "cBrowser")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Navigate(fn As String, Optional lockDown As Boolean = False, Optional visible As Boolean = True) As cBrowser*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fn|String|False||
lockDown|Boolean|True| False|
visible|Boolean|True| True|


---
VBA Procedure: **httpPost**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function httpPost(fn As String, Optional data As String = vbNullString, Optional isjson As Boolean = False, Optional authHeader As String = vbNullString, Optional additionalHeaders As Object = Nothing, Optional method As String = "POST") As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fn|String|False||
data|String|True| vbNullString|
isjson|Boolean|True| False|
authHeader|String|True| vbNullString|
additionalHeaders|Object|True| Nothing|
method|String|True| "POST"|


---
VBA Procedure: **extraHeaders**  
Type: **Function**  
Returns: **Object**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function extraHeaders(ohttp As Object, additionalHeaders As cJobject) As Object*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ohttp|Object|False||
additionalHeaders|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||


---
VBA Procedure: **httpGET**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function httpGET(fn As String, Optional authUser As String = vbNullString, Optional authPass As String = vbNullString, Optional accept As String = vbNullString, Optional timeout As Long = 0, Optional authHeader As String = vbNullString, Optional additionalHeaders As Object = Nothing, Optional method As String = "GET") As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
fn|String|False||
authUser|String|True| vbNullString|
authPass|String|True| vbNullString|
accept|String|True| vbNullString|
timeout|Long|True| 0|
authHeader|String|True| vbNullString|
additionalHeaders|Object|True| Nothing|
method|String|True| "GET"|


---
VBA Procedure: **storeStuff**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub storeStuff(o As Object)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
o|Object|False||


---
VBA Procedure: **Element**  
Type: **Function**  
Returns: **IHTMLElement**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Element(eID As String) As IHTMLElement*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
eID|String|False||


---
VBA Procedure: **elementTags**  
Type: **Function**  
Returns: **IHTMLElementCollection**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function elementTags(tag As String) As IHTMLElementCollection*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
tag|String|False||


---
VBA Procedure: **ElementText**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get ElementText(eID As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
eID|String|False||


---
VBA Procedure: **kill**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub kill()*  

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


---
VBA Procedure: **Class_Terminate**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub Class_Terminate()*  

**no arguments required for this procedure**


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
VBA Procedure: **pIeOB_OnQuit**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub pIeOB_OnQuit()*  

**no arguments required for this procedure**


---
VBA Procedure: **pIeOB_TitleChange**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub pIeOB_TitleChange(ByVal Text As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
ByVal|String|False||
