'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:55 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/6974763/raw/googleWireExample.vba
Option Explicit
'NOTE THIS MODULE HAS BEEN DEPRECATED FOR NEW SHEETS, it will still work on old sheets
'For new sheets USE THE GOOGLESHEETS module
'import google workbooks
'v1.4
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 15/10/2013 10:52:07 : from manifest:5055578 gist https://gist.github.com/brucemcpherson/6974763/raw/googleWireExample.vba

' import google workbooks
Public Sub testWorkBookImport()
    Dim key As String
    key = "0At2ExLh4POiZdE43aGo4TENEWlVOeFBkRlVPcEhIbnc"
    If Not importGoogleWorkbook(key, , , True) Then
        MsgBox ("failed to import workbook at " & key)
    
    End If
End Sub
Public Sub testPublicWorkBookImport()
    Dim key As String

    key = "0AodxbO8eOvBZdE93VnNiaVNRdjdxMXJNMWJlNVRMWGc"
    
    If Not importGoogleWorkbook(key) Then
        MsgBox ("failed to import workbook at " & key)
    
    End If
End Sub
Public Function importGoogleWorkbook(key As String, _
            Optional deleteAllSheetsFirst As Boolean = False, _
            Optional replaceConflictingSheets = True, _
            Optional oauthNeeded As Boolean = False, _
            Optional headers As Boolean = True) As Boolean
    
    ' this will import all the sheets in a given Google SpreadSheet

    Dim authHeader As String, url As String, schema As cJobject, _
            sheetHack As Worksheet, sheetJob As cJobject, accessToken As String
    
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
    
    url = url & "/basic?alt=json"
    Set schema = getSchema(url, authHeader)
    importGoogleWorkbook = False
    
    
    ' now we really should have the schema
    If Not schema Is Nothing Then
        ' get the sheets in this workbook and their urls
        Set sheetJob = getSheetsInSchema(schema)
        
        If Not sheetJob Is Nothing Then
            Application.Calculation = xlCalculationManual
            
            ' clean up existing sheets id necessary
            If deleteAllSheetsFirst Then
                Set sheetHack = deleteAllTheSheets
            
            ElseIf replaceConflictingSheets Then
                Set sheetHack = deleteSomeOfTheSheets(sheetJob)
            
            End If
            ' get the new data
            getData sheetJob, authHeader, accessToken, headers
            sheetJob.tearDown
            
            ' final clean up of sheet that should have been deleted
            If Not sheetHack Is Nothing Then
                Application.DisplayAlerts = False
                sheetHack.Delete
                Application.DisplayAlerts = True
            End If
            importGoogleWorkbook = True
            Application.Calculation = xlCalculationAutomatic
        End If
        
        schema.tearDown
    Else
        MsgBox ("unable to get schema from " & url)

    End If
    
End Function
Private Sub getData(sheetJob As cJobject, authHeader As String, accessToken As String, Optional headers As Boolean = True)
    
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
' get the schema for the given workbook
Private Function getSchema(url As String, authHeader As String) As cJobject
    Dim cb As cBrowser, jsonData As String
    Set cb = New cBrowser
    With cb
        jsonData = .httpGET(url, , , , , authHeader)
        If .status = 200 And jsonData <> vbNullString Then
            Set getSchema = getJobjectFromWire(jsonData, eDeserializeNormal, url)
        Else
            Set getSchema = Nothing
        End If
        .tearDown
    End With
End Function
Public Function getSheetsInSchema(schemaJob As cJobject) As cJobject
    ' discovers the URL of all sheets in the schema
    Dim cj As cJobject, job As cJobject, s As String, _
        sheetJob As cJobject, url As String, joc As cJobject

    ' find the child with the feed directory
    Set cj = schemaJob.find("feed.entry")
    If cj Is Nothing Then
        MsgBox ("could not find feed urls in schema entries")
    Else
    
        ' these will be the urls for each sheet to get
        For Each job In cj.children
            url = vbNullString
            
            ' we're going to use the link for viz api
            For Each joc In job.child("link").children
                s = joc.toString("rel")
                If InStr(1, s, "visualizationApi") Then
                    url = joc.toString("href")
                    Exit For
                End If
            Next joc
            
            ' need link + title
            If (url = vbNullString Or job.find("title.$t") Is Nothing) Then
                MsgBox ("couldnt find the vizapi link or title in the schema")
                Exit For
            End If
            
            ' initialize if first time
            If sheetJob Is Nothing Then
                Set sheetJob = New cJobject
                sheetJob.init(Nothing).addArray
            End If
            
            ' add an array item describing the entry
            With sheetJob.add
                .add "url", url
                .add "name", job.find("title.$t").toString
            End With
    
        Next job
        Set getSheetsInSchema = sheetJob
    End If
    
End Function
'


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
