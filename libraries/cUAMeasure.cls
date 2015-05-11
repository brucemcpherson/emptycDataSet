'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:58 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/7453152/raw/cUAMeasure.cls
Option Explicit
'v2.2
Private pUACode As String
Private pBrowser As cBrowser
Private pID As String
Private pPostData As String
Private pUrl As String
Private pVersion As String
Private pOptOut As Boolean

Public Property Get UACode() As String
    UACode = pUACode
End Property

Public Property Get id() As String
    id = pID
End Property
Public Property Get browser() As cBrowser
    Set browser = pBrowser
End Property
Public Function postAppView(page As String) As cUAMeasure
    
    Set postAppView = Me
    On Error GoTo failed
    If Not pOptOut Then
        pPostData = "v=1&tid=" & pUACode & "&cid=" _
        & pID & "&t=appview&an=EXCEL&av=" & pVersion & "&cd=" & page

        pBrowser.httpPost pUrl, pPostData & "&sc=start"
    End If


    Exit Function
failed:
    ' if it ever fails then we assume tis being blocked by firewall or something and pass
    pOptOut = True
End Function
Public Function postAppKill() As cUAMeasure
    If Not pOptOut Then
        pBrowser.httpPost pUrl, pPostData & "&sc=end"
    End If
    Set postAppKill = Me
End Function
Private Sub Class_Initialize()
    Set pBrowser = New cBrowser

    ' change this to true to opt out of analytic reporting
    pOptOut = False
    
    If Not pOptOut Then
        pUrl = "http://www.google-analytics.com/collect"
        pID = getUserHash()
        pUACode = getUACode()
        pVersion = getVersion()
    End If
    
End Sub
Public Sub tearDown()
    If Not pBrowser Is Nothing Then
        pBrowser.tearDown
        Set pBrowser = Nothing
    End If
End Sub




