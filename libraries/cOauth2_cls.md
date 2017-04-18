# VBA Project: **emptycDataSet**
## VBA Module: **[cOauth2](/libraries/cOauth2.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (emptycDataSet) was automatically created on 4/18/2017 10:33:02 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cOauth2

---
VBA Procedure: **googleAuth**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function googleAuth(scopeEntry As String, Optional replacementConsole As cJobject = Nothing, Optional clientID As String = vbNullString, Optional clientSecret As String = vbNullString, Optional complain As Boolean = True, Optional cloneFromScopeEntry As String = vbNullString) As cOauth2*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
scopeEntry|String|False||
replacementConsole|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|
clientID|String|True| vbNullString|
clientSecret|String|True| vbNullString|
complain|Boolean|True| True|
cloneFromScopeEntry|String|True| vbNullString|


---
VBA Procedure: **hasToken**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get hasToken() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **authHeader**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get authHeader() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **token**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get token() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **denied**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get denied() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **revoke**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function revoke() As cOauth2*  

**no arguments required for this procedure**


---
VBA Procedure: **getUserConsent**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getUserConsent() As cOauth2*  

**no arguments required for this procedure**


---
VBA Procedure: **getToken**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getToken(Optional phase As String = "authorization_code") As cOauth2*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
phase|String|True| "authorization_code"|


---
VBA Procedure: **addSeconds**  
Type: **Function**  
Returns: **Date**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function addSeconds(d As Date, n As Long) As Date*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
d|Date|False||
n|Long|False||


---
VBA Procedure: **isAuthenticated**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get isAuthenticated() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **tokenType**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get tokenType() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **expiresIn**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get expiresIn() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **expires**  
Type: **Get**  
Returns: **Date**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get expires() As Date*  

**no arguments required for this procedure**


---
VBA Procedure: **code**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get code() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **hasRefreshToken**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get hasRefreshToken() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **isExpired**  
Type: **Get**  
Returns: **Boolean**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get isExpired() As Boolean*  

**no arguments required for this procedure**


---
VBA Procedure: **refreshToken**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Property Get refreshToken() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **getItemValue**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getItemValue(key As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
key|String|False||


---
VBA Procedure: **createUrl**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function createUrl(parameterType As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
parameterType|String|False||


---
VBA Procedure: **generatePhaseParameters**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function generatePhaseParameters(whichPhase As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
whichPhase|String|False||


---
VBA Procedure: **tearDown**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function tearDown() As cOauth2*  

**no arguments required for this procedure**


---
VBA Procedure: **salt**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let salt(p As String)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|String|False||


---
VBA Procedure: **encrypt**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function encrypt(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **decrypt**  
Type: **Function**  
Returns: **String**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function decrypt(s As String) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **getRegistryPackage**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function getRegistryPackage(authFlavor As String, scopeEntry As String) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
authFlavor|String|False||
scopeEntry|String|False||


---
VBA Procedure: **setRegistryPackage**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function setRegistryPackage() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **describeDialog**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function describeDialog() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **googlePackage**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function googlePackage(Optional consolePackage As cJobject = Nothing) As cJobject*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
consolePackage|[cJobject](/libraries/cJobject_cls.md "cJobject")|True| Nothing|


---
VBA Procedure: **addFromOther**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Sub addFromOther(c As cJobject, p As cJobject, k As String, Optional ok As String = vbNullString)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
c|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
p|[cJobject](/libraries/cJobject_cls.md "cJobject")|False||
k|String|False||
ok|String|True| vbNullString|


---
VBA Procedure: **addGoogleScope**  
Type: **Function**  
Returns: **[cOauth2](/libraries/cOauth2_cls.md "cOauth2")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function addGoogleScope(s As String) As cOauth2*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
s|String|False||


---
VBA Procedure: **makeBasicGoogleConsole**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Private**  
Description: ****  

*Private Function makeBasicGoogleConsole() As cJobject*  

**no arguments required for this procedure**


---
VBA Procedure: **skeletonPackage**  
Type: **Function**  
Returns: **[cJobject](/libraries/cJobject_cls.md "cJobject")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function skeletonPackage() As cJobject*  

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
