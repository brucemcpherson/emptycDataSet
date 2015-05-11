'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:57 PM : from manifest:7471153 gist https://gist.github.com/brucemcpherson/6937529/raw/usefulEncrypt.vba
Option Explicit
' this stuff needs the capicom library and will only work for 32bit excel
' v1.5
' for 64 bit - no encryption is being done yet
'

Public Function encryptMessage(ByVal TobeEncrypted As String, ByVal salt As String) As String

' adapted from http://msdn.microsoft.com/en-us/library/windows/desktop/aa382018(v=vs.85).aspx
' needs a reference to capicom

#If VBA7 And Win64 Then
'TODO -  64bit remains unencrypted
    MsgBox ("warning- 64 bit excel encryption not yet implemented - will return plain text")
    encryptMessage = TobeEncrypted
#Else

    On Error GoTo ErrorHandler
    Const CAPICOM_ENCRYPTION_ALGORITHM_DES = 2
    Dim message As Object
    Set message = CreateObject("CAPICOM.EncryptedData")
    message.content = TobeEncrypted
    message.SetSecret (salt)

    message.Algorithm.name = CAPICOM_ENCRYPTION_ALGORITHM_DES
    encryptMessage = message.encrypt
    Set message = Nothing
    Exit Function

ErrorHandler:
    If Err.number > 0 Then
        MsgBox ("Visual Basic error found:" & Err.description)
    Else
        MsgBox ("CAPICOM error found : " & Err.number)
    End If

#End If
End Function
Public Function decryptMessage(ByVal encrypted As String, ByVal salt As String) As String
    
#If VBA7 And Win64 Then
'TODO -  64bit remains unencrypted
    decryptMessage = encrypted
#Else
    On Error GoTo ErrorHandler
    Dim message As Object
    Set message = CreateObject("CAPICOM.EncryptedData")
    message.SetSecret salt
    message.decrypt encrypted
    decryptMessage = message.content
    Set message = Nothing
    Exit Function

ErrorHandler:
    If Err.number > 0 Then
        MsgBox "Visual Basic error found:" & Err.description
    Else
    '    Check for a bad password error.
        If Err.number = -2146893819 Then
            MsgBox "Error. The password may not be correct."
        Else
            MsgBox "CAPICOM error found : " & Err.number
        End If
    End If
#End If
End Function

Public Function encryptSha1(ByVal keyString As String, ByVal str As String) As String

    Dim encode As Object, encrypt As Object, s As String, _
        t() As Byte, b() As Byte, privateKeyBytes() As Byte
        
    Set encode = CreateObject("System.Text.asciiEncoding")
    Set encrypt = CreateObject("System.Security.Cryptography.HMACSHA1")
    s = Replace(keyString, "-", "+")
    s = Replace(s, "_", "/")
    privateKeyBytes = decodeBase64(s)

    encrypt.key = privateKeyBytes
    t = encode.Getbytes_4(str)
    b = encrypt.ComputeHash_2((t))
    
    s = tob64(b)
    s = Replace(s, "+", "-")
    encryptSha1 = Replace(s, "/", "_")
    
    Set encode = Nothing
    Set encrypt = Nothing

End Function

Public Function tob64(ByRef arrData() As Byte) As String

    Dim objXML As Object, objNode
    'Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = CreateObject("MSXML2.DOMDocument")

    ' byte array to base64
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    tob64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

Public Function decodeBase64(ByVal strData As String) As Byte()
    Dim objXML As Object, objNode As Object
    'Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    decodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function





