Dim myURL As String, sFilename As String
    myURL = "http://somelink.com/fileofchoice.pdf"
    sFilename = Environ("SystemDrive") & Environ("HomePath") & _
            Application.PathSeparator & "Desktop" & Application.PathSeparator & _
            "file.pdf"
   
    Dim WinHttpReq As Object, oStream As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False ', "username", "password"
   WinHttpReq.Send
   
    myURL = WinHttpReq.ResponseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.ResponseBody
        oStream.SaveToFile sFilename, 2  ' 1 = no overwrite, 2 = overwrite
       oStream.Close
    End If
