'Do action on workbook open
Private Sub Workbook_Open()

    Dim objReq As WinHttp.WinHttpRequest
    Set objReq = New WinHttp.WinHttpRequest 'Used to support cookies
    Dim knackAuth As String
    Dim argumentString As String
    Dim ws As Worksheet
    Dim strFile As String
    Let argumentString = "{""email"":""example@email.com"",""password"":""password"", ""remember"":false}" 'Remember flag leaves session valid for 48 hours
    
    'Get cookie authentication'
    knackAuth = "https://us-api.knack.com/v1/accounts/session/" 'POST to for login
    objReq.Open "POST", knackAuth, False
    objReq.setRequestHeader "Content-Type", "application/json"
    objReq.send argumentString

    'Get cookie repsonse header, holds cookie auth token
    Let cookies = objReq.getResponseHeader("Set-Cookie")
    Let searchTarget = ";"
    Let Length = InStr(1, cookies, searchTarget, 1)
    Let auth = Mid(cookies, 1, Length - 1) 'connect.sid s%3Ankzcs0ItRFGeDsC03HzqHVasrz_qGRHv.L8NXm8ehXSc66HclBRr2hw6fDSn8YYIKLGS5C9QO1RA
    
    'Get file from knack API
    objReq.Option(WinHttpRequestOption_EnableRedirects) = True
    objReq.Open "GET", "https://api.knack.com/v1/objects/object_1/records/export/applications/APP-ID?&type=csv", False
    objReq.setRequestHeader "Cookie", auth 'Setup cookie with authentication token
    objReq.send
    
    'File Path
    strFile = "C:\Users\" & Environ("username") & "\file.csv" 'Must be a valid path. IE: subfolders must exist or this will fail
    
    'Create file stream, save csv to local disk
    If objReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write objReq.responseBody
        oStream.SaveToFile strFile, 2 '1 = no overwrite, 2 = overwrite
        oStream.Close
    End If
    
    Set ws = ActiveWorkbook.Sheets("sheet") 'Set to current worksheet name
    
    'Insert data into worksheet starting at range specified
    With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True 'Specified options must match source data, otherwise this will fail
         .Refresh
    End With
    
    'Cleanup connections
    Dim Conn As WorkbookConnection
    For Each Conn In ThisWorkbook.Connections
        If Conn.Name <> ".sheet" And Conn.Name <> "sheet1" And Conn.Name <> "sheet2" Then
            Conn.Delete
        End If
    Next Conn
End Sub
