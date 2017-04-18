'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:56 PM : from manifest:5055578 gist https://gist.github.com/brucemcpherson/6937450/raw/oAuthExamples.vba
Option Explicit
' oauth examples
' v1.2
' convienience function for auth..
Public Function getGoogled(scope As String, _
                                Optional replacementpackage As cJobject = Nothing, _
                                Optional clientID As String = vbNullString, _
                                Optional clientSecret As String = vbNullString, _
                                Optional complain As Boolean = True, _
                                Optional cloneFromeScope As String = vbNullString) As cOauth2
    Dim o2 As cOauth2
    Set o2 = New cOauth2
    With o2.googleAuth(scope, replacementpackage, clientID, clientSecret, complain, cloneFromeScope)
        If Not .hasToken And complain Then
            MsgBox ("Failed to authorize to google for scope " & scope & ":denied code " & o2.denied)
        End If
    End With
    
    Set getGoogled = o2
End Function
Private Sub testOauth2()
    Dim myConsole As cJobject
    ' if you are calling for the first time ever you can either provide your
    ' clientid/clientsecret - or pass the the jsonparse retrieved from the google app console
    ' normally all this stuff comes from encrpted registry store
    
    ' first ever
    'Set myConsole = makeMyGoogleConsole
    'With getGoogled("analytics", myConsole)
    '    Debug.Print .authHeader
   '     .tearDown
   ' End With

    'or you can do first ever like this
    With getGoogled("viz", , "1092408392628-lga1vk2kr2bipvbo0esru331gg9fnp8r.apps.googleusercontent.com", "2ct4Skci3lDrPtQ-Y3dRD50f")
        Debug.Print .authHeader
        .tearDown
    End With
    
    With getGoogled("drive", , , , , "viz")
        Debug.Print .authHeader
        .tearDown
    End With
    ' all other times this is what is needed
    With getGoogled("drive")
        Debug.Print .authHeader

        .tearDown
    End With
    ' lets auth and have a look at the contents
    'Debug.Print objectStringify(getGoogled("drive"))
    
    ' all other times this is what is needed
    With getGoogled("analytics", , , , , "drive")
        Debug.Print .authHeader
        .tearDown
    End With
    
    ' here's an example of cloning credentials from a different scope for 1st time in
    With getGoogled("urlshortener", , , , , "drive")
        Debug.Print .authHeader
        .tearDown
    End With
    
    With getGoogled("urlshortener")
        Debug.Print .authHeader
        .tearDown
    End With
    
    ' if you made one, clean it up
    If Not myConsole Is Nothing Then
        myConsole.tearDown
    End If
End Sub

Private Function makeMyGoogleConsole() As cJobject
    Dim consoleJSON As String
 
     consoleJSON = _
    "{'installed':{'auth_uri':'https://accounts.google.com/o/oauth2/auth'," & _
      "'client_secret':'xxxxxxxx'," & _
      "'token_uri':'https://accounts.google.com/o/oauth2/token'," & _
      "'client_email':'','redirect_uris':['urn:ietf:wg:oauth:2.0:oob','oob']," & _
      "'client_x509_cert_url':'','client_id':'xxxxxxx.apps.googleusercontent.com'," & _
      "'auth_provider_x509_cert_url':'https://www.googleapis.com/oauth2/v1/certs'}}"
      
      Set makeMyGoogleConsole = JSONParse(consoleJSON)

End Function
   