Sub DownloadAndMoveReport(URLID As String, ReportName As String, Optional username As String = "", Optional password As String = "", Optional Searchterm As String = "")
    If Dir("H:\STOP.txt") <> "" Then Exit Sub
    
    Dim driver As Object
    Dim downloadFolderPath As String
    Dim destFilePath As String
    Dim SharepointFilePath As String

    
    downloadFolderPath = Environ("USERPROFILE") & "\Downloads\"
    
    If Searchterm <> "" Then
    
        destFilePath = "G:\Server Processes\Shared data\" & ReportName & " " & Searchterm & ".csv"
'        SharepointFilePath = "\\eastwestltd.sharepoint.com@SSL\DavWWWRoot\sites\ProcessOptimisationProjectManagement\Shared Documents\Shared data\" & ReportName & " " & Searchterm & ".csv"
        
        Else
        
        destFilePath = "G:\Server Processes\Shared data\" & ReportName & ".csv"
'        SharepointFilePath = "\\eastwestltd.sharepoint.com@SSL\DavWWWRoot\sites\ProcessOptimisationProjectManagement\Shared Documents\Shared data\" & ReportName & ".csv"
        
    End If
    
    ' Delete any existing file in the Downloads folder
    On Error Resume Next
    Kill downloadFolderPath & ReportName & ".csv"
    On Error GoTo 0

    ' Default to helpdesk account if no username and password are provided
    If username = "" Or password = "" Then
        username = ""
        password = ""
    End If

'--------------------------------------
    ' Start Chrome silently with retry (3 attempts)
    Set driver = StartChromeSilent(3)

    ' If it still failed, exit gracefully (no MsgBox)
    If driver Is Nothing Then
        ' Optionally log to Immediate window
        Debug.Print "Chrome failed to start (", Now, ")"
    End If

    ' Safe maximize (waits for readiness to avoid hangs)
    SafeMaximize driver
'--------------------------------------

    ' Set up Selenium WebDriver
'    Set driver = CreateObject("Selenium.WebDriver")

    driver.Timeouts.PageLoad = 20000000
    driver.Timeouts.ImplicitWait = 20000000

    driver.SetProfile Environ("USERPROFILE") & "\AppData\Local\Google\Chrome\User Data May25Problem"
    driver.AddArgument "profile-directory=Profile 2"

'    driver.Start "chrome"
'    driver.Wait 1000
'    driver.Window.Maximize

    ' Log in and handle potential logoff
    Do Until Dir(downloadFolderPath & ReportName & ".csv") <> ""
        driver.Get "https://go.joblogic.com/uk/Account/LogOn"
        
        Dim keys As New Selenium.keys
        driver.SendKeys keys.Escape
         
        driver.Wait 10000

        If InStr(driver.Url, "Login") = 0 Then
            driver.FindElementById("accountMenu").Click
            driver.Wait 1000
            driver.FindElementById("logOffMenu").Click
            driver.Wait 1000
        End If

        If InStr(driver.Url, "AlreadyLogin") > 0 Then
            On Error Resume Next
            driver.FindElementById("accountMenu").Click
            driver.Wait 1000
            driver.FindElementById("logOffMenu").Click
            driver.Wait 1000
        End If


        ' Clear and enter username
        With driver.FindElementByName("UserName")
            driver.Wait 500
            .SendKeys keys.Control & "a"
            driver.Wait 500
            .SendKeys keys.Delete
            driver.Wait 500
            .SendKeys username
        End With
        
        driver.Wait 1000
        
        ' Clear and enter password
        With driver.FindElementByName("Password")
            driver.Wait 500
            .SendKeys keys.Control & "a"
            driver.Wait 500
            .SendKeys keys.Delete
            driver.Wait 500
            .SendKeys password
        End With
'        driver.FindElementByName("UserName").SendKeys username
'        driver.Wait 1000
'        driver.FindElementByName("Password").SendKeys password
'        driver.Wait 1000
        driver.FindElementByName("Password").SendKeys vbCrLf ' Simulate Enter key press
        driver.Wait 5000
        driver.Get "https://go.joblogic.com/account/processlogin"
        driver.Wait 2000
        driver.SendKeys keys.Escape
        ' Wait for the URL to change from the login page
        Dim timeout As Single
        timeout = Timer + 30 ' Wait for up to 30 seconds
        Do While InStr(driver.Url, "Login") > 0
            If Timer > timeout Then
                driver.Get "https://go.joblogic.com/account/processlogin"
                driver.Wait 3000
            End If
            driver.Wait 1000
        Loop

        ' Ensure the post-login page has loaded
        If Not ElementExists(driver, "OutstandingVsComplete_Last30Days") Then
            driver.Get "https://go.joblogic.com/account/processlogin"
            driver.Wait 5000
        End If

        Dim stopLoop As Boolean
        Dim RunLoop As Boolean 'New code added 31/03/2025
        
        stopLoop = False
        RunLoop = True
        
        Do
            On Error Resume Next
            driver.Get "https://go.joblogic.com/uk/Report/Preview?id=" & URLID
            On Error GoTo 0
            
            '--------------New code added 31/03/2025 -------------------------
            driver.Wait 10000
            If InStr(driver.Url, "AlreadyLogin") > 0 Then
                stopLoop = True
                RunLoop = False
            End If
            
            '-------------------------------
            
            'New Code for search term

            
            ' Example condition to stop the loop early
            Debug.Print "Checking if 'report-data-confirm' exists..."
                        
            If ElementExists(driver, "report-data-confirm") Then
                    If ElementExists(driver, "searchTerm") Then
                        driver.FindElementById("searchTerm").Click
                        driver.SendKeys Searchterm
                        driver.Wait 500
                        driver.SendKeys keys.Enter
                        driver.Wait 500
                    End If
            
                Debug.Print "'report-data-confirm' exists, clicking it."
                            
                driver.Wait 2000
                On Error Resume Next
                driver.FindElementById("report-data-confirm").Click
                On Error GoTo 0
                stopLoop = True
            Else
                Debug.Print "'report-data-confirm' does not exist. Retrying..."
            End If

            DoEvents ' Allow other processes to run
        Loop While Not stopLoop ' Adjust the number of retries as needed

        If RunLoop = True Then 'New code added 31/03/2025
            Dim downloadStartTime As Single
            downloadStartTime = Timer
            Do Until Dir(downloadFolderPath & ReportName & ".csv") <> ""
                    If ElementExists(driver, "report-data-confirm") Then 'NEW CODE added 17/04/2025
                        DoEvents
                    Else
                        Exit Do
                    End If
            Loop
        End If 'New code added 31/03/2025
    Loop

    ' Log out and close the browser
    driver.FindElementById("accountMenu").Click
    driver.Wait 1000
    driver.FindElementById("logOffMenu").Click
    driver.Wait 1000
    driver.Quit

    If Dir(destFilePath) <> "" Then
        ' Remove read-only attribute if set
        SetAttr destFilePath, vbNormal
        Kill destFilePath
    End If

    'move the file to the Sharepoint folder
'
'    FileCopy downloadFolderPath & ReportName & ".csv", _
'            SharepointFilePath


    
    ' Move the file to the destination path
    Name downloadFolderPath & ReportName & ".csv" As destFilePath

    ' set read-only attribute
    SetAttr destFilePath, vbReadOnly


End Sub

Function ElementExists(driver As Object, elementId As String) As Boolean
    On Error Resume Next
    Dim elements As Object
    Set elements = driver.FindElementsById(elementId) ' Plural version to avoid errors
    
    If elements.Count = 0 Then
        Debug.Print "Element '" & elementId & "' NOT found."
        ElementExists = False
    Else
        Debug.Print "Element '" & elementId & "' found."
        ElementExists = True
    End If
    
    Set elements = Nothing
    On Error GoTo 0
End Function

' Starts Chrome silently with retry. Returns Nothing if all attempts fail.
Function StartChromeSilent(Optional retries As Integer = 3) As WebDriver
    Dim d As WebDriver
    Dim attempt As Integer

    On Error GoTo TryAgain

    For attempt = 1 To retries
        Err.Clear
        Set d = New WebDriver
        d.Start "chrome"

        ' Success
        Set StartChromeSilent = d
        Exit Function

TryAgain:
        ' Clean up partial session and retry quietly
        On Error Resume Next
        If Not d Is Nothing Then d.Quit
        On Error GoTo 0
        On Error GoTo TryAgain

        Err.Clear
        ' brief pause before retry
        Application.Wait Now + TimeValue("0:00:01")
    Next attempt

    ' All attempts failed: return Nothing silently
    Set StartChromeSilent = Nothing
End Function


' Waits for the driver to actually have a browser window (readiness).
' Returns True if ready within timeoutSecs, otherwise False. Silent.
Function WaitForWindowReady(ByVal d As WebDriver, Optional timeoutSecs As Double = 5) As Boolean
    On Error GoTo SafeExit
    Dim t As Single: t = Timer

    Do
        DoEvents
        If Not d Is Nothing Then
            ' Accessing .Window can throw while Chrome is still initializing
            On Error Resume Next
            Dim hasWin As Boolean
            hasWin = Not (d.Window Is Nothing)
            On Error GoTo 0

            If hasWin Then
                WaitForWindowReady = True
                Exit Function
            End If
        End If

        If Timer - t >= timeoutSecs Then Exit Do
    Loop

SafeExit:
    ' If we reach here without setting True, it's False
End Function


' Safe maximize that won’t hang: waits for readiness first, then tries to maximize.
' Silent: no MsgBox, no Exit Sub.
Sub SafeMaximize(ByVal d As WebDriver)
    If d Is Nothing Then Exit Sub

    If WaitForWindowReady(d, 5) Then
        On Error Resume Next
        d.Window.Maximize
        On Error GoTo 0
    End If
End Sub
