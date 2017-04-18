# VBA Project: **emptycDataSet**
## VBA Module: **[cCell](/libraries/cCell.cls "source is here")**
### Type: ClassModule  

This procedure list for repo (emptycDataSet) was automatically created on 4/18/2017 10:33:03 AM by VBAGit.
For more information see the [desktop liberation site](http://ramblings.mcpher.com/Home/excelquirks/drivesdk/gettinggithubready "desktop liberation")

Below is a section for each procedure in cCell

---
VBA Procedure: **row**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get row() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **column**  
Type: **Get**  
Returns: **Long**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get column() As Long*  

**no arguments required for this procedure**


---
VBA Procedure: **parent**  
Type: **Get**  
Returns: **[cDataRow](/libraries/cDataRow_cls.md "cDataRow")**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get parent() As cDataRow*  

**no arguments required for this procedure**


---
VBA Procedure: **myKey**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get myKey() As String*  

**no arguments required for this procedure**


---
VBA Procedure: **where**  
Type: **Get**  
Returns: **Range**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get where() As Range*  

**no arguments required for this procedure**


---
VBA Procedure: **refresh**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get refresh() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **toString**  
Type: **Get**  
Returns: **String**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get toString(Optional sFormat As String = vbNullString, Optional followFormat As Boolean = False, Optional deLocalize As Boolean = False) As String*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
sFormat|String|True| vbNullString|
followFormat|Boolean|True| False|
deLocalize|Boolean|True| False|


---
VBA Procedure: **value**  
Type: **Get**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Get value() As Variant*  

**no arguments required for this procedure**


---
VBA Procedure: **value**  
Type: **Let**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Property Let value(p As Variant)*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|False||


---
VBA Procedure: **needSwap**  
Type: **Function**  
Returns: **Boolean**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function needSwap(cc As ccell, e As eSort) As Boolean*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
cc|ccell|False||
e|eSort|False||


---
VBA Procedure: **Commit**  
Type: **Function**  
Returns: **Variant**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function Commit(Optional p As Variant) As Variant*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
p|Variant|True||


---
VBA Procedure: **create**  
Type: **Function**  
Returns: **ccell**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Function create(par As cDataRow, colNum As Long, rCell As Range, Optional v As Variant) As ccell*  

*name*|*type*|*optional*|*default*|*description*
---|---|---|---|---
par|[cDataRow](/libraries/cDataRow_cls.md "cDataRow")|False||
colNum|Long|False||
rCell|Range|False||
v|Variant|True||


---
VBA Procedure: **tearDown**  
Type: **Sub**  
Returns: **void**  
Return description: ****  
Scope: **Public**  
Description: ****  

*Public Sub tearDown()*  

**no arguments required for this procedure**
