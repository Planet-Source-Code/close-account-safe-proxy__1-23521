Attribute VB_Name = "modMain"
'Safe Proxy by David Fiala  - Released to planet-source-code on: May 29, 2001
'djf1010@aol.com
'This is one of my older programs I built and I probally wouldn't release it to the public the way it is
'but I don't plan on fixing it any time soon. Enjoy...

Option Explicit

Public Function GetUrlServer(strHeaders As String) As String
    Dim intPos As Integer 'declare a variable
    Dim strOut As String 'declare a variable
    intPos = "11" 'set the value of a variable
    Do While Right(strOut, 1) <> "/"
        intPos = intPos + 1 'set the value of a variable
        strOut = strOut & Mid(strHeaders, intPos, 1) 'set the value of a variable
    Loop 'Do While Right(strOut, 1) <> "/"
    strOut = Mid(strOut, 1, Len(strOut) - 1) 'set the value of a variable
    GetUrlServer = strOut 'return a value
End Function

Public Function GetUrlFile(strHeaders As String) As String
    Dim strOut As String 'declare a variable
    Dim intSPOS As Integer 'declare a variable
    Dim intEPOS As Integer 'declare a variable
    Dim intTPOS As Integer 'declare a variable
    intTPOS = 1 'set the value of a variable
    intSPOS = Len(GetUrlServer(strHeaders)) + 12 'set the value of a variable
    Do While Mid(strHeaders, intTPOS, 8) <> " HTTP/1."
        intTPOS = intTPOS + 1 'set the value of a variable
    Loop 'Do While Mid(strHeaders, intTPOS, 8) <> " HTTP/1."
    GetUrlFile = Mid(strHeaders, intSPOS, intTPOS - intSPOS) 'return a value
End Function

Public Function NewGet(strHeaders As String) As String
    Dim strNew As String 'declare a variable
    strNew = strHeaders 'set the value of a variable
    strNew = Replace(strNew, "http://" & GetUrlServer(strHeaders), "") 'set the value of a variable
    NewGet = strNew 'return a value
End Function

Public Function Blocks(strHeaders As String) As String
    Dim strNew As String 'declare a variable
    strNew = strHeaders 'set the value of a variable
    strNew = BlockUpload(strNew)
'    strNew = Replace(strNew, "type=file", "type=hidden value=><b>Sorry, file uploads have been restricted.</b><rem ") 'block out the uploads
'    strNew = Replace(strNew, "type=""file""", "type=hidden value=><b>Sorry, file uploads have been restricted.</b><rem ") 'block out the uploads
    Blocks = strNew 'return a value
End Function

Public Function NewPost(strHeaders As String) As String
    Dim strNew As String 'declare a variable
    strNew = strHeaders 'set the value of a variable
    strNew = Replace(strNew, "http://" & GetUrlServerPost(strHeaders), "") 'set the value of a variable
    NewPost = strNew 'return a value
End Function

Public Function GetUrlServerPost(strHeaders As String) As String
    Dim intPos As Integer 'declare a variable
    Dim strOut As String 'declare a variable
    intPos = "12" 'set the value of a variable
    Do While Right(strOut, 1) <> "/"
        intPos = intPos + 1 'set the value of a variable
        strOut = strOut & Mid(strHeaders, intPos, 1) 'set the value of a variable
    Loop 'Do While Right(strOut, 1) <> "/"
    strOut = Mid(strOut, 1, Len(strOut) - 1) 'set the value of a variable
    GetUrlServerPost = strOut 'return a value
End Function

Public Function BlockUpload(strHeaders As String)
    Dim strLCase As String
    Dim strReturn As String
    strReturn = strHeaders
    strLCase = LCase(strReturn)
Search:
    If InStr(1, strLCase, "type=""file""", vbTextCompare) = 0 Then
        'no type="file" now we search for type=file
        GoTo Search2
    Else
        'do search for: type="file"
        strReturn = Mid(strReturn, 1, InStr(1, strLCase, "type=""file""", vbTextCompare) - 1) & " type=hidden value=><b>Sorry, you are not allowed to upload.</b><rem " & Mid(strReturn, InStr(1, strLCase, "type=""file""", vbTextCompare) + 9)
        strLCase = LCase(strReturn)
    End If 'f InStr(1, strLCase, "type=""file""", vbTextCompare) = 0
Search2:
    If InStr(1, strLCase, "type=file", vbTextCompare) = 0 Then
        'no type=file now we are done
        GoTo Done
    Else
        'search for: type=file
        strReturn = Mid(strReturn, 1, InStr(1, strLCase, "type=file", vbTextCompare) - 1) & " type=hidden value=><b>Sorry, you are not allowed to upload.</b><rem " & Mid(strReturn, InStr(1, strLCase, "type=file", vbTextCompare) + 9)
        strLCase = LCase(strReturn)
        GoTo Search2
    End If 'If InStr(1, strLCase, "type=file", vbTextCompare) = 0
Done:
    BlockUpload = strReturn
End Function
