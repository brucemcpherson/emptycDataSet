'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:53 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/6974763/raw/googleSheets.vba
Option Explicit
' the new google sheets are bit different than the old.
' these are the new versions.
' note that you can also use dbAbstraction for updating, querying and reading sheets.
' however this is a port of the old to the new sheets layout
' public new sheets need to be published, private not.
' v1.0
' note that things have changed between google sheets old and new - public sheets must be published, private sheets need oauth2
' import google workbooks

Public Sub testWorkBookImportPublicSheets()
    Dim key As String
    
    ' this example imports all the worksheets in a workbook with this Key from a public published sheet
    ' you need to have enabled oauth2
    key = "12pTwh5Wzg0W4ZnGBiUI3yZY8QFoNI8NNx_oCPynjGYY"
    
    If Not importGoogleWorkbookNewSheets(key) Then
        MsgBox ("failed to import workbook at " & key)
    
    End If
End Sub
Public Sub testWorkBookImportNewSheets()
    Dim key As String
    
    ' this example imports all the worksheets in a workbook with this Key from a  private sheet
    ' you need to have enabled oauth2
    key = "12pTwh5Wzg0W4ZnGBiUI3yZY8QFoNI8NNx_oCPynjGYY"
    
    If Not importGoogleWorkbookNewSheets(key, , , True) Then
        MsgBox ("failed to import workbook at " & key)
    
    End If
End Sub
Public Sub testSelectedImportNewSheets()
    Dim key As String
    
    ' this example imports the given list of private sheets
    Dim listOfSheets As String
    listOfSheets = "carriers,exceldemo"
    key = "12pTwh5Wzg0W4ZnGBiUI3yZY8QFoNI8NNx_oCPynjGYY"
    
    If Not importGoogleWorkbookNewSheets(key, , , True, , listOfSheets) Then
        MsgBox ("failed to import workbook at " & key)
    
    End If
End Sub
Public Function importGoogleWorkbookNewSheets(key As String, _
            Optional deleteAllSheetsFirst As Boolean = False, _
            Optional replaceConflictingSheets = True, _
            Optional oauthNeeded As Boolean = False, _
            Optional headers As Boolean = False, _
            Optional listOfSheets = vbNullString) As Boolean
    
    ' this will import all the sheets in a given Google SpreadSheet

    Dim authHeader As String, url As String, schema As cJobject, _
            sheetHack As Worksheet, sheetJob As cJobject, accessToken As String
    
    ' if not oauth then sheet needs to have been published
    If (oauthNeeded) Then
        With getGoogled("viz")
            If (.hasToken) Then
                authHeader = .authHeader
                accessToken = .token
            Else
                Exit Function
            End If
            .tearDown
        End With
    Else
        authHeader = vbNullString
    End If
    
    ' this is the endpoint for google VIZ imports
    url = "https://spreadsheets.google.com/feeds/worksheets/" + key
    If (oauthNeeded) Then
        url = url & "/private"
    Else
        url = url & "/public"
    End If
    
    url = url & "/full"
    
    Set schema = getSchemaNewSheets(url, authHeader)
    importGoogleWorkbookNewSheets = False

    ' now we really should have the schema
    If Not schema Is Nothing Then
        ' get the sheets in this workbook and their urls
        Set sheetJob = getSheetsInSchemaNewSheets(schema)
        
        If Not sheetJob Is Nothing Then
            ' check we have all the sheets needed
            Dim doJob As cJobject, job As cJobject, joc As cJobject
            If (listOfSheets = vbNullString) Then
                Set doJob = sheetJob
            Else
                Dim a As Variant, i As Long
                a = Split(listOfSheets, ",")
                Set doJob = New cJobject
                With doJob.init(Nothing).addArray
                    For i = LBound(a) To UBound(a)
                        For Each job In sheetJob.children
                            If (makeKey(a(i)) = makeKey(job.toString("name"))) Then
                                With .add
                                    For Each joc In job.children
                                        .add joc.key, joc.value
                                    Next joc
                                End With
                            End If
                        Next job
                    Next i
                End With
                If (doJob.children.count <> arrayLength(a)) Then
                    MsgBox ("did not find all required sheets " & doJob.stringify)
                End If
            End If
            Application.Calculation = xlCalculationManual
            
            ' clean up existing sheets id necessary
            If deleteAllSheetsFirst Then
                Set sheetHack = deleteAllTheSheets
            
            ElseIf replaceConflictingSheets Then
                Set sheetHack = deleteSomeOfTheSheets(doJob)
            
            End If
            ' get the new data
            getDataNewSheets doJob, authHeader, accessToken, headers
            sheetJob.tearDown
            doJob.tearDown
            ' final clean up of sheet that should have been deleted
            If Not sheetHack Is Nothing Then
                Application.DisplayAlerts = False
                sheetHack.Delete
                Application.DisplayAlerts = True
            End If
            importGoogleWorkbookNewSheets = True
            Application.Calculation = xlCalculationAutomatic
        End If
        
        schema.tearDown
    Else
        MsgBox ("unable to get schema from " & url)

    End If
    
End Function

Private Function getSchemaNewSheets(url As String, authHeader As String) As cJobject
    Dim cb As cBrowser, jsonData As String, jObject As cJobject
    Set cb = New cBrowser
    With cb
        jsonData = .httpGET(url, , , , , authHeader)
        If .status = 200 And jsonData <> vbNullString Then
            ' first try for aJSON result
            Set jObject = JSONParse(jsonData, , False)
            If (jObject Is Nothing Or Not jObject.isValid) Then
                Set jObject = makeJobjectFromXML(jsonData, False)
            End If
            If (isSomething(jObject) And jObject.isValid) Then
                Set getSchemaNewSheets = getJobjectFromWire(jObject.stringify, eDeserializeNormal, url)
                jObject.tearDown
            End If
        Else
            Set getSchemaNewSheets = Nothing
        End If
        
        .tearDown
    End With
End Function

Private Sub getDataNewSheets(sheetJob As cJobject, authHeader As String, accessToken As String, Optional headers As Boolean = True)
    
    Dim cj As cJobject, job As cJobject, w As Worksheet, _
        r As Range, jr As cJobject, jc As cJobject, c As cJobject, url As String, _
        cb As cBrowser, jsonData As String, jor As cJobject, s As String, joc As cJobject, start As Long
    
    Set cb = New cBrowser
    For Each job In sheetJob.children
        ' we'll refresh after each sheet
        Application.ScreenUpdating = False
        ' create a new sheet
        Set w = sheetExists(job.toString("name"), False)
        If (w Is Nothing) Then
            Set w = Worksheets.add
            w.name = job.toString("name")
            Set r = w.Cells.Resize(1, 1)
                
            ' now get data for this sheet - the vizualisation api needs the access token on the url
            url = job.toString("url")
            If (accessToken <> vbNullString) Then
                url = url & "&access_token=" & accessToken
            End If
            If (headers) Then
                url = url & "&headers=-1"
            Else
                url = url & "&headers=0"
            End If
            
            jsonData = cb.httpGET(url, , , , , authHeader)
            If (cb.status <> 200) Then
                MsgBox ("failed to get data: error " & cb.status & " for " & w.name)
            Else
                ' get as data
                Set c = getJobjectFromWire(jsonData, eDeserializeGoogleWire)
                If Not c Is Nothing Then
                    Set jc = c.find("cols")
                    Set jr = c.find("rows")
                    
                    ' column headings
                    start = 0
                    If Not jc Is Nothing And headers Then
                        For Each jor In jc.children
                            With jor
                                s = .child("label").value
                                If s = vbNullString Then
                                    s = .child("id").value
                                End If
                                r.Offset(start, .childIndex - 1).value = s
                            End With
                        Next jor
                        start = 1
                    End If
        
                    ' and these are the rows
                    If Not jr Is Nothing Then
                        For Each jor In jr.children
                            With jor
                                For Each joc In .child("c").children
                                    r.Offset(.childIndex + start - 1, joc.childIndex - 1).value = joc.child("v").value
                                Next joc
                            End With
                        Next jor
                    End If
                    c.tearDown
                Else
                    MsgBox ("no data for " & w.name)
                End If
            End If
            
        Else
            MsgBox ("could not add " & job.toString("name") & " sheet already exists")
        End If
        
        Application.ScreenUpdating = True
        
    Next job
    cb.tearDown

End Sub

Public Function getSheetsInSchemaNewSheets(schemaJob As cJobject) As cJobject
    ' discovers the URL of all sheets in the schema
    Dim cj As cJobject, job As cJobject, s As String, _
        sheetJob As cJobject, url As String, joc As cJobject

    ' find the child with the feed directory
    Set cj = schemaJob.find("feed.entry")
    If cj Is Nothing Then
        MsgBox ("could not find feed urls in schema entries")
    Else
    
        ' these will be the urls for each sheet to get
        url = vbNullString

        For Each job In cj.children
            Dim jobLink As cJobject, relLink As cJobject, jItem As cJobject
            
            ' results will go here
            If sheetJob Is Nothing Then
                Set sheetJob = New cJobject
                sheetJob.init(Nothing).addArray
            End If
            
            ' get the name of worksheet
            With sheetJob.add
                .add "name", job.child("title.text").value
                Set jobLink = job.child("link")
                For Each joc In jobLink.children
                    If (InStr(1, joc.toString("rel"), "visualizationApi") > 0) Then
                        .add "url", joc.toString("href")
                        Exit For
                    End If
                Next joc
                Debug.Assert isSomething(.childExists("url"))
            End With
            
        Next job
        Set getSheetsInSchemaNewSheets = sheetJob
    End If
    sheetJob.stringify
End Function


' decode data returned from httpGet
Public Function getJobjectFromWire(jsonData As String, _
        Optional t As eDeserializeType = eDeserializeNormal, _
        Optional url As String = vbNullString) As cJobject
    

    Dim c As cJobject, s As String
    Set c = New cJobject
    If t = eDeserializeGoogleWire Then
        s = cleanGoogleWire(jsonData)
    Else
        s = CStr(jsonData)
    End If
    
    With c.init(Nothing)
        .add "url", CStr(url)
        .add("data").append JSONParse(s, t)
    End With
    
    Set getJobjectFromWire = c
End Function
Public Function deleteAllTheSheets() As Worksheet
    ' dont want complaints
    Application.DisplayAlerts = False
    
    ' delete all but one sheet ( you cant delete the last sheet in a book )
    While Sheets.count > 1
      Sheets(1).Delete
    Wend
    
    Application.DisplayAlerts = True
    Set deleteAllTheSheets = hackTheLastSheet
End Function
Public Function deleteSomeOfTheSheets(job As cJobject) As Worksheet
    ' dont want complaints
    Dim cj As cJobject, sh As Worksheet
    Application.DisplayAlerts = False
    
    For Each cj In job.children
        Set sh = sheetExists(cj.toString("name"), False)
        If Not sh Is Nothing Then
            If Sheets.count > 1 Then
                sh.Delete
            Else
                Set deleteSomeOfTheSheets = hackTheLastSheet
                Debug.Assert cj.childIndex = job.children.count
            End If
        End If
    Next cj

End Function

Private Function hackTheLastSheet() As Worksheet
    ' rename it to something random
    Randomize
    Sheets(1).name = "canbedeleted" & CStr(Rnd())
    Set hackTheLastSheet = Sheets(1)

End Function




'--- below - will only work for non-oauth2/ single sheet - use the previous examples for oauth2/entire workbooks
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse

Public Sub googleWireExample()
'NOTE - THIS METHOD HAS LARGELY BEEN SUPERCEDED BY THE EXAMPLE AT THE TOP OF THIS MODULE
'testPublicWorkBookImport
'------------------------
    Dim dset As cDataSet, dsClone As cDataSet, jo As cJobject, cb As cBrowser, url As String
    Dim sWire As String


    url = "https://docs.google.com/a/mcpher.com/spreadsheet/tq?key=0AodxbO8eOvBZdE93VnNiaVNRdjdxMXJNMWJlNVRMWGc#gid=0"
    ' get the google wire string
    ' to test, delete the contents of the worksheet sheet 'clone' - this will fill it up
    Set cb = New cBrowser
    sWire = cb.httpGET(url)
    ' load to a dataset
    Set dset = New cDataSet
    With dset
        .populateGoogleWire sWire, Range("Clone!$a$1")
        
        If .where Is Nothing Then
            MsgBox ("No data to process")
        Else
            ' it worked
            
        End If
    End With
    dset.tearDown
End Sub
