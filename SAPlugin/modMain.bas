Attribute VB_Name = "modMain"
Option Explicit


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_DWORD = 4

Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Public Const ERROR_SUCCESS = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259
Dim wsh As New WshShell


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal _
    lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Const S_OK = 0
Private Const SYNCHRONIZE = &H100000
Const PROCESS_QUERY_INFORMATION = &H400
Const STILL_ACTIVE = &H103

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


'these are our general options
Private msTempPath As String        'the SA temp path
Private msSAPath As String          'path to the SA executable
Private msSARulesPath As String     'path to the SA rules
Private mnQuarantine As Integer     'the quarantine options
Private msQuarantinePath As String  'the path to the quarantine directory
Private mbKeepOriginal As Boolean   'whether to use the original message (i.e. when spam isn't detected, not alter it)
Private mlMaxMessageSize As Long    'any message that is over this size is not processed
Private m_sArgs As String
Private Log As Logger
Private mSpamcPath As String
Dim oFSO As New FileSystemObject


Sub HandleSpam(ByVal sMsgCommandFile, ByVal sGUID As String, ByVal sMsgFile As String, Optional Testing As Boolean = False)

                    Select Case mnQuarantine
                    Case 0          'pass the SA altered one through to user
                        Log.Log "if Quarantine == 0, passing the modified message to the user"
                        Log.Log "So copying " & msTempPath & sGUID & ".MAI" & " to " & sMsgFile
                        oFSO.CopyFile msTempPath & sGUID & ".MAI", sMsgFile
                        ' Ensure this goes to the Junk folder in ME:
                        InsertHeader sMsgFile
                
                    Case 1          'quarantine
                        'this is a spam email, so we quarantine it
                        'quaranting involves moving the message, and the SA result (so we can check why it was rejected)
                        'we actually keep original message so we can pass it through later unaltered
                        'the extension for the SA output will be changed to .SA
                        'the copied message file also has a new filename
                        Log.Log "Handling SPAM case where Quarantine == 1"
                        Debug.Assert oFSO.FileExists(sMsgCommandFile)
                        oFSO.CopyFile sMsgCommandFile, msQuarantinePath & sGUID & ".MAI"
                        oFSO.CopyFile sMsgFile, msQuarantinePath & "Messages\" & sGUID & ".MAI"
                        oFSO.CopyFile msTempPath & sGUID & ".MAI", msQuarantinePath & "Messages\" & sGUID & ".SA"
                    
                        'since we have quarantined it, we want to delete from the queues directory so that it doesn't get picked up
                        If Testing Then
                            Log.Log "NOTE: WE ARE IN TEST MODE, WHERE QUARANTINE == 1, SO NOT DELETING ANYTHING"
                            Log.Log "IF YOU SEE THIS IN A REAL EXECUTABLE, THE DEVELOPER MADE AN ERROR (DIDN'T REMOVE THE TESTING FLAG) WHEN HE BUILT THE EXE)"
                        Else
                            Log.Log "Deleting file: " & sMsgCommandFile & "..."
                            oFSO.DeleteFile sMsgCommandFile
                            Log.Log "Deleting file: " & sMsgFile & " ... "
                            oFSO.DeleteFile sMsgFile
                        End If
                        
                    Case 2
                        ' Delete the email because it is a spam and options tell us to
                        Log.Log "Handling SPAM case where Quarantine == 2"
                        Log.Log "Deleting mail because it is spam, and the options tell us to do so."
                        Log.Log "Deleting file: " & sMsgCommandFile & "..."
                        If Not Testing Then
                            oFSO.DeleteFile sMsgCommandFile
                            Log.Log "Deleting file: " & sMsgFile & " ... "
                            oFSO.DeleteFile sMsgFile
                        Else
                            Log.Log "NOTE: WE ARE IN TEST MODE, WHERE QUARANTINE == 1, SO NOT DELETING ANYTHING"
                            Log.Log "IF YOU SEE THIS IN A REAL EXECUTABLE, THE DEVELOPER MADE AN ERROR (DIDN'T REMOVE THE TESTING FLAG) WHEN HE BUILT THE EXE)"
                        End If
                    End Select
                    
                    

End Sub

Function GetFileContents(filePath As String) As String
    Dim Handle As Integer
    On Local Error GoTo Fail
    Handle = FreeFile
    Dim fopen As Boolean
    
    Open filePath For Binary Access Read As #Handle
        fopen = True
        GetFileContents = Space$(LOF(Handle))
        Get #Handle, , GetFileContents
        Close #Handle
        Handle = 0
    Exit Function
    
Fail:
    If fopen Then
        Close #Handle
    End If
    
    
End Function

Function WriteFileContents(filePath As String, Parts() As String) As Boolean
    Dim Handle As Integer
    On Local Error GoTo Fail
    Handle = FreeFile
    Dim Backup As String: Backup = filePath & ".bak"
    Dim fopen As Boolean
    
    Name filePath As Backup

    Open filePath For Binary Access Write As #Handle
        fopen = True
        Dim i As Integer
        For i = 0 To UBound(Parts)
            Put #Handle, , Parts(i)
        Next i
    
    Close #Handle
     
    Kill Backup
    WriteFileContents = True
    Exit Function
Fail:
On Local Error Resume Next
    If fopen Then Close #Handle
    Name Backup As filePath
End Function


Sub InsertHeader(filePath As String, Optional hdr As String = "X-ME-Content: Deliver-To=Junk")
    
    Const SPAMSTATUS As String = "X-Spam-Flag: YES"
    
    Dim sAll As String
    sAll = GetFileContents(filePath)
    Dim loc As Long
    loc = InStr(sAll, SPAMSTATUS)
    Dim headerPart As String
    Dim msgPart As String
    
    If (loc > 0) Then
        ' find the first newline after
        Dim x As Integer
        
        While x <> 10
            loc = loc + 1
            x = CInt(Asc(Mid$(sAll, loc, 1)))
        Wend
        Dim Parts(2) As String
        Parts(0) = Mid$(sAll, 1, loc)
        Parts(1) = hdr
        Parts(2) = Mid$(sAll, loc)
        
        Dim Inserted As Boolean: Inserted = WriteFileContents(filePath, Parts)
        If Inserted Then
            Log.Log "Added header: " & hdr & ", so the email arrives in the junk folder"
            'Dim tmp As String: tmp = GetFileContents(filePath)
            'Debug.Assert (InStr(tmp, hdr) > 0)
           
        Else
            Log.Log "UNEXPECTED: did not insert header: " & hdr
        End If
        
    Else

    End If
    
    
    
End Sub

Sub Main()

    
    On Error GoTo Errhandler
    Set Log = New Logger
    Dim myTimer As New cTimer
    Set myTimer.Log = Log
    myTimer.ID = "Checking this email for spam took: "
    
    Dim sCommand As String
    sCommand = Command$
    ' Passed Params from the MTA: MessageID ConnectorCode
    Dim args() As String
    args() = Split(sCommand, " ")
    m_sArgs = sCommand
  
    
    Log.Log "Command Arguments: " & sCommand
    
    If LenB(m_sArgs) = 0 Then
        Log.Log "Expected two arguments: something like 59769876.MAI <space> SMTP from MailEnable"
        Log.Log "Actually got empty arguments"
        End
    End If
    
    If (InStr(m_sArgs, " ") <= 0) Then
        Log.Log "Expected two arguments: something like 59769876.MAI <space> SMTP from MailEnable"
        Log.Log "Actually got " & m_sArgs
        End
    End If
       
    Dim Splut() As String
    Splut = Split(m_sArgs, " ")
    If UBound(Splut) <> 1 Then
        Log.Log "ERROR: Expected exactly two arguments: something like 59769876.MAI <space> SMTP from MailEnable"
        If (UBound(Splut) = 3) Then
            Log.Log "ERROR: This Plugin is for an MTA delivery event, not a MAILBOX delivery event"
        End If
        Log.Log "Actually got " & m_sArgs
        End
    End If

    If Not LoadOptions Then
        Exit Sub
    End If

    Dim sMsgCommandFile As String
    Dim sMsgFile As String
    Dim sTempPath As String

    Dim oStream As TextStream
    Dim sTemp As String
    Dim lResult As Double
    Dim sGUID As String
    Dim bSpam As Boolean
    
    ' set this to false for the actual release code
    Const Testing As Boolean = False
    Const TestingKnownSpam = False
    Const TestingKnownNotSpam = False
    
    ' You will need to copy these files from the sa folder to the correct Data Directory when testing.
    If (TestingKnownSpam) Then
        args(0) = "sample-spam.MAI"
    ElseIf (TestingKnownNotSpam) Then
        args(0) = "sample-nonspam.MAI"
    End If
    
    sMsgFile = GetRegistryString("SOFTWARE\Mail Enable\Mail Enable", "Data Directory") & "\QUEUES\" & args(1) & "\Inbound\Messages\" & args(0)
    sMsgCommandFile = GetRegistryString("SOFTWARE\Mail Enable\Mail Enable", "Data Directory") & "\QUEUES\" & args(1) & "\Inbound\" & args(0)
    
    Log.Log "'sMsgFile' = " & sMsgFile
    Log.Log "'sMsgCommandFile' = " & sMsgCommandFile
    
    
    
    If TestingKnownSpam Or TestingKnownNotSpam Then
        Debug.Assert oFSO.FileExists(App.Path & "\" & args(0))
        oFSO.CopyFile App.Path & "\" & args(0), sMsgFile, True
    Else
        oFSO.CopyFile sMsgCommandFile, App.Path & "\" & "SampleCommandFile.txt"
    End If

    
    
    'Debug.Assert (oFSO.FileExists(sMsgCommandFile))
    Debug.Assert (oFSO.FileExists(sMsgFile))

    If oFSO.FileExists(sMsgFile) Then
        If (mlMaxMessageSize > 0) And (oFSO.GetFile(sMsgFile).Size > mlMaxMessageSize) Then
            Set oFSO = Nothing
            Log.Log "Message file: " & sMsgFile & " is too big"
            Exit Sub
        End If
    Else
        Set oFSO = Nothing
        Log.Log "Message file: " & sMsgFile & " does not exist"
        Exit Sub
    End If
    
    sGUID = GetGUID
    Dim input_file_in_quotes As String
    Dim output_file_in_quotes As String
    input_file_in_quotes = Chr$(34) & sMsgFile & Chr$(34)
    output_file_in_quotes = Chr$(34) & msTempPath & sGUID & ".MAI" & Chr$(34)
    

        ' spamc -d hostname < sample-spam.txt > output.txt
    sTemp = mSpamcPath & " -E -d 127.0.0.1 < " & input_file_in_quotes & " > " & output_file_in_quotes
    Set oStream = oFSO.OpenTextFile(msTempPath & sGUID & ".BAT", ForAppending, True)
    oStream.WriteLine sTemp
    oStream.Close
    Set oStream = Nothing
    
    
    'then we process the message through SA outputting to our temp dir
    Dim ret As Long
    ret = ShellSynchronous(sGUID)
    
    bSpam = (ret = 1)
    
    If TestingKnownSpam Then
        Debug.Assert bSpam
        Debug.Assert (ret = 0 Or ret = 1)
    Else
        If TestingKnownNotSpam Then
            Debug.Assert Not bSpam
            Debug.Assert (ret = 0 Or ret = 1)
        End If
    End If
    
   If (ret = 0 Or ret = 1) Then
        If bSpam Then
            HandleSpam sMsgCommandFile, sGUID, sMsgFile, Testing
        Else
            Log.Log "This message seemed to be HAM (OK, not junk)"
        End If
    
     
    Else
            ' SA didn't process it... delete our bat file, log it and exit
            Log.Log "A SPAMC ERROR OCCURRED: " & modSAErrors.ReturnCodeToString(CInt(ret))
            oFSO.DeleteFile msTempPath & sGUID & ".BAT"
            Log.Log "SpamAssassin didn't return an email. Configuration problem. Return code was: "
            Set oFSO = Nothing
            Exit Sub
    End If


    If Not bSpam Then
        If mbKeepOriginal Then
            'don't do anything!
        Else
            'if we aren't quarantining, we are replacing the message file with the SA returned one
            oFSO.CopyFile msTempPath & sGUID & ".MAI", sMsgFile, True
        End If
    End If
    
    ' delete the files we have created
    oFSO.DeleteFile msTempPath & sGUID & ".MAI"
    oFSO.DeleteFile msTempPath & sGUID & ".BAT"
    Set oFSO = Nothing

    Exit Sub

Errhandler:
    Log.Log "Could not process pickup event for Connector: " & Err.Description
    
End Sub


Private Function LoadOptions() As Boolean

    Dim oFSO As New FileSystemObject
    Dim sTemp As String * 255
    Dim lResult As Long
    Dim configFilePath As String
    
    configFilePath = App.Path & "\spamassassinconfig.ini"
    
    If oFSO.FileExists(App.Path & "\spamassassinconfig.ini") Then
    
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "SPAMCPATH", vbNullString, sTemp, 255, configFilePath)
        If lResult > 0 Then
            mSpamcPath = Trim$(Left$(sTemp, lResult))
            If Not oFSO.FileExists(mSpamcPath) Then
                Log.Log "Spamc path " & mSpamcPath & " does not exist, falling back to using the less efficient spamassassin executable"
                Log.Log "Note: the path should look something like " & "V:\spamd\sa\spamc.exe" & " (make sure you include the .exe extension)"
                mSpamcPath = vbNullString
            End If
        End If
        
        If LenB(mSpamcPath) = 0 Then
            Log.Log "Spamc path is incorrect or not set in the ini file, exiting"
            LoadOptions = False
            Exit Function
        End If
        
        
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "QUARANTINE", "", sTemp, 255, configFilePath)
        If lResult > 0 Then
            mnQuarantine = Trim$(Left$(sTemp, lResult))
        End If
        
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "MAXMESSAGESIZE", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
        If lResult > 0 Then
            mlMaxMessageSize = Trim$(Left$(sTemp, lResult))
        End If
       
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "KEEPORIGINAL", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
        If lResult > 0 Then
            If Trim$(Left$(sTemp, lResult)) = "1" Then
                mbKeepOriginal = True
            Else
                mbKeepOriginal = False
            End If
        End If
        
        msQuarantinePath = GetRegistryString("SOFTWARE\Mail Enable\Mail Enable", "Quarantine Directory")
        If (Not oFSO.FolderExists(msQuarantinePath)) Then
            lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "QUARANTINEPATH", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
            If lResult > 0 Then
                msQuarantinePath = Trim$(Left$(sTemp, lResult))
            End If

        End If
        
        If LenB(msQuarantinePath) > 0 Then
            If Right$(msQuarantinePath, 1) <> "\" Then
                msQuarantinePath = msQuarantinePath & "\"
            End If
            
                            ' check to see if there is a Messages subdir, if not create it
                If Not oFSO.FolderExists(oFSO.BuildPath(msQuarantinePath, "Messages")) Then
                    oFSO.CreateFolder oFSO.BuildPath(msQuarantinePath, "Messages")
                End If
                                
                ' make sure it is valid -- note: EMPTY is ok (which is why we checked above)
                If Not oFSO.FolderExists(msQuarantinePath & "Messages") Then
                    Log.Log "Quarantine path is not a valid directory."
                    Err.Raise 20001, "LoadOptions", "Quarantine path is not a valid directory."
                End If
        End If


        
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "SPAMASSASSINEXE", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
        If lResult > 0 Then
            msSAPath = Trim$(Left$(sTemp, lResult))
            If (oFSO.FileExists(msSAPath) = False) Then
                    Log.Log "Spamassassin executable file: " & msSAPath & " not found. Please set it in the " & App.Path & "\spamassassinconfig.ini" & " file."
                    Err.Raise 20001, "LoadOptions", "SpamAssassin exe path does not exist."
            End If
        End If
        
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "SPAMASSASSINRULESPATH", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
        If lResult > 0 Then
            msSARulesPath = Trim$(Left$(sTemp, lResult))
            If Not oFSO.FolderExists(msSARulesPath) Then
                    Log.Log "Quantine path is not a valid directory."
                    Err.Raise 20001, "LoadOptions", "Spamassassin rules path is not a valid directory. Do you need to run sa-update?"
            End If
        End If
        
        lResult = GetPrivateProfileString("SPAMASSASSIN Plugin Config", "TEMPPATH", "", sTemp, 255, App.Path & "\spamassassinconfig.ini")
        If lResult > 0 Then
            msTempPath = Trim$(Left$(sTemp, lResult))
            
           ' Relative or absolute path?
            If Mid$(msTempPath, 2, 1) = ":" Then
            
            Else
                ' make the path absolute (helps when debugging)
                msTempPath = App.Path & "\" & msTempPath
            End If

            If Right$(msTempPath, 1) <> "\" Then
                msTempPath = msTempPath & "\"
            End If
            
            If Not oFSO.FolderExists(msTempPath) Then
                MakePath (msTempPath)
            End If
            
            ' check we made it:
            If Not oFSO.FolderExists(msTempPath) Then
                Log.Log "Unable to find or create TEMPPATH: " & msTempPath
                Log.Log "Check the ini file"
                Exit Function
            End If

        End If
        
        If LenB(msTempPath) = 0 Then
            Log.Log "Cannot create temp directory if it is empty. Check the ini file!"
            LoadOptions = False
            Exit Function
        End If
        
        If Not oFSO.FolderExists(msTempPath) Then
            Log.Log "Cannot create temp directory " & msTempPath & " Check the ini file!"
            LoadOptions = False
            Exit Function
        End If
                  
        If Len(msQuarantinePath) = 0 Then
            Log.Log "Turning off quarantining for SpamAssassin because no path set in config file"
            mnQuarantine = 0
        End If
        
        Set oFSO = Nothing
        LoadOptions = True
    Else
        Log.Log "Can't run SpamAssassin because no plugin config file at path: " & App.Path
        LoadOptions = False
    End If
    
    
End Function

Private Function ShellAndWait(sCmd As String) As Long


Dim cmd As WshExec
Dim myTimer As cTimer
Dim Slept As Long
Const TIMEOUT As Long = 10000

Set myTimer = New cTimer

Call wsh.Run(sCmd, 1, 1)

Do While cmd.Status = WshRunning
    Sleep 10
    Slept = Slept + 10
    DoEvents
    If (Slept > TIMEOUT) Then
        ShellAndWait = -1
        Exit Function
    End If
    DoEvents
Loop

ShellAndWait = cmd.ExitCode
Dim sErr As String
sErr = cmd.StdErr.ReadAll
If LenB(sErr) Then
    If cmd.ExitCode >= 0 Then
        ShellAndWait = -1
    End If
    Log.Log "ERROR: shelled command returned " & sErr
    Debug.Assert 0
End If
' return code of 1 means it's SPAM when running SPAMC
Log.Log "Shelled Program exited with code: " & cmd.ExitCode
Log.Log "Running " & sCmd & " took " & myTimer.Elapsed & " milliseconds."
'Log.Log "STDOUT from shelled program: " & cmd.StdOut.ReadAll
'Log.Log "STDERR from shelled program: " & cmd.StdErr.ReadAll

End Function

Private Function ShellSynchronous(ByVal sGUID As String)
        Dim hProcess As Long
        Dim RetVal As Long
        Dim tmr As New cTimer
        tmr.ID = "Running SA Process"
        Set tmr.Log = Log
        
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(Environ("comspec") & " /c """ & msTempPath & sGUID & ".BAT""", vbHide))
        If (hProcess = 0) Then
            Debug.Assert False
            Log.Log "OpenProcess Failed: " & Err.LastDllError
            ShellSynchronous = -100
            Exit Function
        End If
        
        Do
            GetExitCodeProcess hProcess, RetVal
            DoEvents: Sleep 100
        Loop While RetVal = STILL_ACTIVE
        
        'close handle
        CloseHandle hProcess
        ShellSynchronous = RetVal
        
End Function


Function GetRegistryString(ByVal sKeyName As String, ByVal szValueName As String) As String
    
    On Error Resume Next
    
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lRetVal As Long
    Dim hKey As Long
    Dim sValue As String
    Dim lValue As Long
    
    lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lrc = RegQueryValueExNULL(hKey, szValueName, 0&, lType, 0&, cch)
    sValue = String(cch, 0)
    
    Select Case lType
        Case REG_DWORD
            lrc = RegQueryValueExLong(hKey, szValueName, 0&, lType, lValue, cch)
            GetRegistryString = lValue
        Case Else
            lrc = RegQueryValueExString(hKey, szValueName, 0&, lType, sValue, cch)
            GetRegistryString = Left$(sValue, cch - 1)
    End Select
    
    RegCloseKey (hKey)
    
End Function


Public Function GetGUID() As String
    
    Dim lResult As Long
    Dim lguid As GUID
    Dim MyguidString As String
    Dim MyGuidString1 As String
    Dim MyGuidString2 As String
    Dim MyGuidString3 As String
    Dim DataLen As Integer
    Dim StringLen As Integer
    Dim i%
    
    On Error GoTo Errhandler
    
    lResult = CoCreateGuid(lguid)
    If lResult = S_OK Then
        MyGuidString1 = Hex$(lguid.Data1)
        StringLen = Len(MyGuidString1)
        DataLen = Len(lguid.Data1)
        MyGuidString1 = LeadingZeros(2 * DataLen, StringLen) & MyGuidString1
        MyGuidString2 = Hex$(lguid.Data2)
        StringLen = Len(MyGuidString2)
        DataLen = Len(lguid.Data2)
        MyGuidString2 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString2)
        MyGuidString3 = Hex$(lguid.Data3)
        StringLen = Len(MyGuidString3)
        DataLen = Len(lguid.Data3)
        MyGuidString3 = LeadingZeros(2 * DataLen, StringLen) & Trim$(MyGuidString3)
        GetGUID = MyGuidString1 & MyGuidString2 & MyGuidString3
        For i% = 0 To 7
            MyguidString = MyguidString & Format$(Hex$(lguid.Data4(i%)), "00")
        Next i%
        GetGUID = GetGUID & MyguidString
    Else
        GetGUID = "00000000"
    End If
    
    Exit Function

Errhandler:
    GetGUID = "00000000"
End Function

Function LeadingZeros(ExpectedLen As Integer, ActualLen As Integer) As String
    LeadingZeros = String$(ExpectedLen - ActualLen, "0")
End Function

Public Sub MakePath(ByVal Folder As String)

    Dim arTemp() As String
    Dim i As Long
    Dim FSO As Scripting.FileSystemObject
    Dim cFolder As String

    Set FSO = New Scripting.FileSystemObject

    arTemp = Split(Folder, "\")
    For i = LBound(arTemp) To UBound(arTemp)
        cFolder = cFolder & arTemp(i) & "\"
        If Not FSO.FolderExists(cFolder) Then
            Call FSO.CreateFolder(cFolder)
        End If
    Next

End Sub
