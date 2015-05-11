'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:52 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/3423885/raw/cRest.cls
' This is a generalized class for dealing with rest queries
Option Explicit
' v2.13
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Public Enum erRestType
    erQueryPerRow   ' one query for each row
    erSingleQuery   ' one single query fills multiple rows
End Enum
Public Enum erResultsFormat
    erJSON   ' expected json
    erAUTO   ' detect
    erXML    ' expect xml
    erUnknown ' not yet decided
End Enum
Public Enum erAuthType
    erOAUTH2  ' google oauth2
End Enum
Private pResultsFormat As erResultsFormat
Private pCollectionNeeded As Boolean
Private perType As erRestType
Private pQueryhCell As cCell
Private pQueryString As String
Private pDset As cDataSet
Private pPopulate As Boolean
Private pRestUrlStem As String
Private pResponse As String
Private pResponseData As String
Private pClearMissingCells As Boolean
Private pJobjects As Collection
Private pTreeSearch As Boolean
Private pIgnore As String
Private pPass As String
Private pUser As String
Private pAccept As String
Private pDatajObject As cJobject
Private pAppend As Boolean
Private pStamp As cCell
Private pAppendQuery As String
Private pWire As Boolean
Private pAlwaysEncode As Boolean
Private pTimeout As Long
Private pPost As String
Private pBrowser As cBrowser
Private poAuth2 As cOauth2


Public Sub tearDown()
    Dim cj As cJobject
    If Not pDset Is Nothing Then pDset.tearDown
    Set pDatajObject = Nothing
    If Not pJobjects Is Nothing And pCollectionNeeded Then
        For Each cj In pJobjects
            cj.tearDown
        Next cj
        Set pJobjects = Nothing
    End If
    If Not pBrowser Is Nothing Then
        pBrowser.tearDown
        Set pBrowser = Nothing
    End If
    If Not poAuth2 Is Nothing Then
        poAuth2.tearDown
        Set poAuth2 = Nothing
    End If
End Sub
Public Property Get jObjects() As Collection
    Set jObjects = pJobjects
End Property
Public Property Get datajObject() As cJobject
    Set datajObject = pDatajObject
End Property
Public Property Get jObject(Optional complain As Boolean = True) As cJobject

    If pResultsFormat = erXML Then
        Set jObject = xmlStringToJobject(pResponse)
        
    ElseIf pResultsFormat = erAUTO Then
        Set jObject = xmlStringToJobject(pResponse, False)
        If (jObject Is Nothing) Then
           Set jObject = JSONParse(pResponse, , complain)
        End If
    
    Else
        Set jObject = JSONParse(pResponse, , complain)
    End If
    
End Property
Public Property Get erType() As erRestType
    erType = perType
End Property
Public Property Get response() As String
    response = pResponse
End Property
Public Property Get encodedUri() As String
    Dim p As Long, s As String, sq As String
    s = pRestUrlStem
    sq = pQueryString & pAppendQuery
' sometimes a query can contain stuff not to be encoded.. we identify this by the presence of = in the query string
' and encode after that. if there is no = then no encoding is needed
    If (Not pAlwaysEncode) Then
        p = InStrRev(sq, "=")
        If (p > 0) Then
            s = s & Left(sq, p)
            If (p < Len(sq)) Then s = s & URLEncode(Mid(sq, p + 1))
        Else
            s = s & sq
        End If
    Else
        s = s & URLEncode(sq)
    End If
    encodedUri = s
End Property

Public Property Get queryhCell() As cCell
    Set queryhCell = pQueryhCell
End Property
Public Property Let queryString(p As String)
    pQueryString = p
End Property
Public Property Let restUrlStem(p As String)
    pRestUrlStem = p
End Property
Public Property Get queryString() As String
    queryString = pQueryString
End Property
Public Property Get dset() As cDataSet
    Set dset = pDset
End Property
Private Function respRootJob(job As cJobject) As cJobject

    If pResponseData = vbNullString Then
        Set respRootJob = job
    Else
        Set respRootJob = childOrFindJob(job, pResponseData)
    End If
    
End Function
Private Function stripDots(s As String) As String
    ' this is to deal with sytax like object..field in columns headers and returns the object part
    stripDots = rxReplace("\.{2}.*", s, "")
End Function
Private Function dotsTail(s As String) As String
    ' this returns the field of object..field
    dotsTail = rxGroup("(.*?)\.{2}(.*)", s, 2)
End Function
Private Function isDots(s As String) As Boolean
    ' detects if there are dots like object..field
    isDots = stripDots(s) <> s
End Function

Public Function childOrFindJob(job As cJobject, s As String) As cJobject
    Dim t As String
    t = stripDots(s)
    If pTreeSearch Then
        Set childOrFindJob = job.find(t)
    Else
        Set childOrFindJob = job.child(t)
    End If
    
End Function
Public Function init(Optional rData As String = "responsedata.results", _
                     Optional et As erRestType = erQueryPerRow, _
                     Optional hc As cCell = Nothing, _
                     Optional rq As String = vbNullString, _
                     Optional ds As cDataSet = Nothing, _
                     Optional pop As Boolean = True, _
                     Optional pUrl As String = vbNullString, _
                     Optional clearmissing As Boolean = True, _
                     Optional treesearch As Boolean = False, _
                     Optional complain As Boolean = True, _
                     Optional sIgnore As String = vbNullString, _
                     Optional user As String = vbNullString, _
                     Optional pass As String = vbNullString, _
                     Optional append As Boolean = False, _
                     Optional stampQuery As cCell = Nothing, _
                     Optional appendQuery As String = vbNullString, _
                     Optional libAccept As String = vbNullString, _
                     Optional bWire As Boolean = False, _
                     Optional collectionNeeded As Boolean = True, _
                     Optional bAlwaysEncode As Boolean = False, _
                     Optional timeout As Long = 0, _
                     Optional postData As String = vbNullString, _
                     Optional resultsFormat As erResultsFormat = erJSON, _
                     Optional oa As cOauth2 = Nothing) As cRest
    ' query control
    pAlwaysEncode = bAlwaysEncode
    pAppendQuery = appendQuery
    pResponseData = rData
    perType = et
    pTreeSearch = treesearch
    pClearMissingCells = clearmissing
    pIgnore = sIgnore
    pPass = pass
    pUser = user
    pAppend = append
    pAccept = libAccept
    pWire = bWire
    pCollectionNeeded = collectionNeeded
    pTimeout = timeout
    pPost = postData
    Set pStamp = stampQuery
    pResultsFormat = resultsFormat
    Set poAuth2 = oa
    ' every jobject for every row
    Set pJobjects = New Collection
    Set pQueryhCell = hc
    If perType = erQueryPerRow And pQueryhCell Is Nothing Then
        If complain Then
            MsgBox ("You need to specify a column for rowwise queries")
        End If
        Exit Function
    End If
    pQueryString = rq
    pRestUrlStem = pUrl
    
    ' data set control
    Set pDset = ds
    If pDset Is Nothing Then
        If Not hc Is Nothing Then
            Set pDset = hc.parent.parent
        End If
    End If
    
    ' population control
    pPopulate = pop
    If pDset Is Nothing And pPopulate Then
        If complain Then
            MsgBox ("You need to specify a dataset")
        End If
        Exit Function
    End If
    

    Set init = Me
End Function
Private Function executeSingle(Optional rurl As String = vbNullString, _
                        Optional qry As String = vbNullString, _
                        Optional complain As Boolean = True, _
                        Optional sFix As String = vbNullString _
                        ) As cJobject
   Dim job As cJobject, jot As cJobject, authHeader As String

    ' we can supply this at run time
    If rurl <> vbNullString Then pRestUrlStem = rurl
    If qry <> vbNullString Then pQueryString = qry

    If (sFix = vbNullString) Then
        Debug.Print encodedUri
        ' we can use post if any post data is specified
        authHeader = vbNullString
        If Not poAuth2 Is Nothing Then
            authHeader = poAuth2.authHeader
        End If
        If (pPost = vbNullString) Then
            pResponse = pBrowser.httpGET(encodedUri, pUser, pPass, pAccept, pTimeout, authHeader)
        Else
            pResponse = pBrowser.httpPost(encodedUri, pPost, , authHeader)
        End If
        If pIgnore <> vbNullString And pResultsFormat = erJSON Then
            If Left(pResponse, Len(pIgnore)) = pIgnore Then
                pResponse = "{" & "'crest'" & ":" _
                    & Mid(pResponse, Len(pIgnore) + 1) & "}"
            End If
        End If
        ' now another tweak no quotes round the keys
        If pWire Then pResponse = rxReplace("(\w+)(:)", pResponse, "'$1':")
    Else
        pResponse = sFix
    End If
    
    ' deserialize the result
    Set job = jObject(complain)
    If (job Is Nothing) Then Exit Function
        
    If pCollectionNeeded Then
        pJobjects.add job
    Else
        pJobjects.add job.key
    End If
    
    If Not job.isValid Then
        If complain Then
            MsgBox ("Badly formed jSon returned for query" & _
               pQueryString & "-" & pResponse)
        End If
    Else
        If respRootJob(job) Is Nothing Then
            If complain Then
                MsgBox ("No results for query " _
                        & pQueryString _
                        & "-" & job.serialize(True))
            End If
        Else
            Set executeSingle = job
        End If
    End If
    
End Function
Public Function execute(Optional qry As String = vbNullString, _
                        Optional sFix As String = vbNullString, _
                        Optional complain As Boolean = True) As cRest
    Dim cc As cCell, dr As cDataRow, job As cJobject, n As Long
    Select Case perType
        Case erSingleQuery
        'clear current data
            If Not pDset Is Nothing Then
                If Not pDset.where Is Nothing And Not pAppend Then
                    pDset.where.ClearContents
                End If
            End If
            ' do a single query that populates multiple rows
            Set job = executeSingle(, qry, complain, sFix)
            If Not job Is Nothing Then
                If (Not populateRows(job, complain) Is Nothing) Then
                    'update the dataset with the new values
                    Set pDset = pDset.rePopulate
                    Set execute = Me
                End If
            End If
            
        Case erQueryPerRow
        ' do a query for each row
            With pDset
                n = 0
                For Each dr In .rows
                    Set job = executeSingle(, URLEncode(dr.cell(pQueryhCell.toString).toString), , sFix)
                    If Not job Is Nothing Then
                        If (Not populateOneRow(job, dr) Is Nothing) Then n = n + 1
                    End If
                Next dr
                ' write it all back to the sheet
                If (n > 0) Then
                    If (pClearMissingCells) Then
                        .bigCommit
                    Else
                        .flushDirtyColumns
                    End If

                End If
            End With
            
        Case Default
            Debug.Assert False
            Exit Function
            
    End Select
    Set execute = Me

End Function
Private Function populateOneRow(job As cJobject, dr As cDataRow) As cDataRow
    Dim jo As cJobject, dc As cCell, jof As cJobject
    'populate cells with response - this populate cells in this row

    If Not pPopulate Then Exit Function
    ' this is a real update data call
    
    Set jo = respRootJob(job)
    Set pDatajObject = jo
    For Each dc In pDset.headings
        ' leave the query column intact
        If (dc.column <> pQueryhCell.column) Then
            ' update with new value
            Set jof = childOrFindJob(jo, dc.toString)
            If Not jof Is Nothing Then
                dr.cell(dc.column).value = getValueFromJo(jof, dc.toString)
                ' if its not a reponse item, then clear it if required
            ElseIf pClearMissingCells Then
                dr.cell(dc.column).value = Empty
            End If
        End If
    Next dc
    ' for convenience
    Set populateOneRow = dr

End Function
Private Function populateRows(job As cJobject, Optional complain As Boolean = True) As cRest
    Dim jo As cJobject, dc As cCell, iAppend As Long, dotless As String
    'populate cells with response - this populate cells in this row
    Set pDatajObject = respRootJob(job)
    If Not pPopulate Then Exit Function
    ' this is a real update data call
    iAppend = 0
    If pAppend Then iAppend = pDset.rows.count
    With pDatajObject
        If .hasChildren Then
            For Each jo In .children
                ' now match whatever column headings there are
                For Each dc In pDset.headings
                    dotless = stripDots(dc.toString)
                    If Not jo.child(dotless) Is Nothing Then
                        dc.where.Offset(jo.childIndex + iAppend).value = _
                                getValueFromJo(jo.child(dotless), dc.toString)
                    ' this is to deal with when a query which is supposed to create and array doesnt
                    ElseIf Not jo.parent.child(dotless) Is Nothing Then
                        dc.where.Offset(jo.childIndex + iAppend).value = _
                                getValueFromJo(jo.parent.child(dotless), dc.toString)
                    End If
                Next dc
                If Not pStamp Is Nothing Then
                    pStamp.where.Offset(jo.childIndex + iAppend).value = queryString
                End If
                ' this is where it wasnt actually an array, but should have been
                If Not .isArrayRoot Then Exit For
            Next jo
            Set populateRows = Me
        Else
            If complain Then
                MsgBox ("Could find no data for query " _
                        & pQueryString _
                        & "-" & job.serialize)
            End If
        End If
    End With
End Function
Function getValueFromJo(jo As cJobject, originalKey As String) As Variant
    Dim s As String, jom As cJobject, searchKey As String, _
        needDots As Boolean, t As String, jot As cJobject
    
    ' now deal with syntax like obj..field
    needDots = isDots(originalKey)
    If needDots Then searchKey = dotsTail(originalKey)
    
    If jo.isArrayRoot Then
        s = vbNullString
        If jo.hasChildren Then
            ' this is aboutmaking a comma separated array
            For Each jom In jo.children
                t = vbNullString
                With jom
                    If (needDots) Then
                        Set jot = .find(searchKey)
                        If Not jot Is Nothing Then
                            t = jot.toString
                        End If
                    Else
                        t = .toString
                    End If
                    
                    If t <> vbNullString Then
                        If s <> vbNullString Then s = s & ","
                        s = s & t
                    End If
                End With
            Next jom
        End If
        getValueFromJo = s
    Else
        getValueFromJo = jo.value
    End If
    
End Function
Public Property Get browser() As cBrowser
    Set browser = pBrowser
End Property
Private Sub Class_Initialize()
    Set pBrowser = New cBrowser
End Sub

