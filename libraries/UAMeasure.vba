'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:58 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/7453196/raw/UAMeasure.vba
'v2.1
Option Explicit
Public Sub testua()
    With registerUA("developing_testua")
        ' do something
        sleep 5
        ' kill session
        .postAppKill
        If Not .browser.isOk Then
            Debug.Print .browser.status
        End If
        .tearDown
    End With
End Sub
Public Function registerUA(page As String) As cUAMeasure
    Dim c As cUAMeasure
    Set c = New cUAMeasure
    With c
        .postAppView (page)
        ' silent fail
        If Not .browser.isOk Then
            Debug.Print .browser.status
        End If
    End With
    Set registerUA = c
End Function
Public Function getUACode() As String
    getUACode = "UA-45711027-1"
End Function
Public Function getVersion() As String
    getVersion = "cDataSet.v.3.012"
End Function
Public Function getUserHash() As String

' more an more systems dont have this, so just abandonding
    getUserHash = getSalt()
    
    'getUserHash = encryptSha1(getSalt(), _
     '   Application.ThisWorkbook.FullName & Application.UserName)
End Function
Public Function getSalt() As String
    getSalt = "vNIXE0xscrmjlyV-12Nj_BvUPaw="
End Function



