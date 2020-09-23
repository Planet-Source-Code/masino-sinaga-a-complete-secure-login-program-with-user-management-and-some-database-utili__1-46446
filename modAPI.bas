Attribute VB_Name = "modAPI"
'Utk menonaktifkan Ctrl-Alt-Del
Declare Function SystemParametersInfo Lib "user32" Alias _
"SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'---------- Compact Database --------------------
Public Declare Function GetTempPath Lib "kernel32" Alias _
    "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer _
    As String) As Long
'---------- End of Compact database -------------

'To show program from TaskList;
'so, program can not be-"End-Task"!!!
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID _
As Long, ByVal dwType As Long) As Long

'This for SHUTDOWN Komputer from program
Public Const EWX_LOGOFF = 0  'Log off
Public Const EWX_SHUTDOWN = 1 'Shut down system
Public Const EWX_REBOOT = 2 ' Reboot.
Public Const EWX_FORCE = 4 'Force all app closed.
Public Const EWX_POWEROFF = 8 'Shut down (Casing ATX)
Public Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'This for setting program in INI File
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
ByVal lpFileName As String) As Long

Public INIFileName As String

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName _
As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'---------- INI File ----------


'--------- "Browse for Folder"
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Public Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList _
As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Public Type BrowseInfo
  hWndOwner As Long
  pIDLRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type
'---------- End of browse for folder ------------


'--- Start TopMostOfAll...
Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, _
ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'--- End of TopMostOfAll...

'--- Start run at startup
Public Type SECURITY_ATTRIBUTES
nLength As Long
lpSecurityDescriptor As Long
bInheritHandle As Long
End Type

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Enum T_KeyClasses
HKEY_CLASSES_ROOT = &H80000000
HKEY_CURRENT_CONFIG = &H80000005
HKEY_CURRENT_USER = &H80000001
HKEY_LOCAL_MACHINE = &H80000002
HKEY_USERS = &H80000003
End Enum

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY _
Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) _
And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const REG_BINARY = 3
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4
Private Const REG_NOTIFY_CHANGE_NAME = &H1
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_OPTION_BACKUP_RESTORE = 4
Private Const REG_OPTION_CREATE_LINK = 2
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_RESERVED = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME _
Or REG_NOTIFY_CHANGE_ATTRIBUTES Or _
REG_NOTIFY_CHANGE_LAST_SET Or _
REG_NOTIFY_CHANGE_SECURITY)
Private Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or _
REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or _
REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

Public Sub DeleteValue(rClass As T_KeyClasses, Path As String, sKey As String)
Dim hKey As Long
Dim res As Long
res = RegOpenKeyEx(rClass, Path, 0, KEY_ALL_ACCESS, hKey)
res = RegDeleteValue(hKey, sKey)
RegCloseKey hKey
End Sub

Public Function SetRegValue(KeyRoot As T_KeyClasses, Path As String, sKey As _
String, NewValue As String) As Boolean
Dim hKey As Long
Dim KeyValType As Long
Dim KeyValSize As Long
Dim KeyVal As String
Dim tmpVal As String
Dim res As Long
Dim i As Integer
Dim X As Long
res = RegOpenKeyEx(KeyRoot, Path, 0, KEY_ALL_ACCESS, hKey)
If res <> 0 Then GoTo Errore
tmpVal = String(1024, 0)
KeyValSize = 1024
res = RegQueryValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
Select Case res
Case 2
KeyValType = REG_SZ
Case Is <> 0
GoTo Errore
End Select
Select Case KeyValType
Case REG_SZ
tmpVal = NewValue
Case REG_DWORD
X = Val(NewValue)
tmpVal = ""
For i = 0 To 3
tmpVal = tmpVal & Chr(X Mod 256)
X = X \ 256
Next
End Select
KeyValSize = Len(tmpVal)
res = RegSetValueEx(hKey, sKey, 0, KeyValType, tmpVal, KeyValSize)
If res <> 0 Then GoTo Errore
SetRegValue = True
RegCloseKey hKey
Exit Function
Errore:
SetRegValue = False
RegCloseKey hKey
End Function
'End setting of run at startup

Public Function ReadFromINIToControls(Objek, MyAppName As String)
Dim Contrl As Control
Dim TempControlName As String * 101, TempControlValue As String * 101
On Error Resume Next
For Each Contrl In Objek
If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Or (TypeOf _
Contrl Is OptionButton) Or (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is CheckBox) Then
TempControlName = Contrl.Name
If (TypeOf Contrl Is TextBox) Or (TypeOf Contrl Is ComboBox) Then 'Or _
   '(TypeOf Contrl Is MaskEdBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "", _
   TempControlValue, Len(TempControlValue), INIFileName)
Else 'If (TypeOf Contrl Is CheckBox) Then
   Result = GetPrivateProfileString(MyAppName, TempControlName, "0", _
   TempControlValue, Len(TempControlValue), INIFileName)
End If
  
If (TypeOf Contrl Is OptionButton) Then
   If Contrl.Index = Val(TempControlValue) Then Contrl = True
Else
    Contrl = TempControlValue
   If (TypeOf Contrl Is ComboBox) Then
      If Len(Contrl.Text) = 0 Then Contrl.ListIndex = 0
      End If
   End If
End If
Next
End Function

Public Function SaveFromControlsToINI(Objek, MyAppName As String)
Dim Contrl As Control
Dim TempControlName As String, TempControlValue As String
On Error Resume Next
For Each Contrl In Objek
  If (TypeOf Contrl Is CheckBox) Or (TypeOf Contrl Is ComboBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Value
    If (TypeOf Contrl Is ComboBox) Then
      TempControlValue = Contrl.Text
      If TempControlValue = "" Then TempControlValue = 1
    End If
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If
  If (TypeOf Contrl Is TextBox) Then
    TempControlName = Contrl.Name
    TempControlValue = Contrl.Text
    Result = WritePrivateProfileString(MyAppName, TempControlName, _
    TempControlValue, INIFileName)
  End If
  If (TypeOf Contrl Is OptionButton) Then
    TempControlValue = Contrl.Value
    If TempControlValue = True Then
      TempControlName = Contrl.Name
      TempControlValue = Contrl.Index
      Result = WritePrivateProfileString(MyAppName, TempControlName, _
      TempControlValue, INIFileName)
    End If
  End If
Next
End Function

'Compact database Access
Public Sub CompactJetDatabase(Location As String, _
           Optional BackupOriginal As Boolean = True)
On Error GoTo CompactErr
 Dim strBackupFile As String
 Dim strTempFile As String
 'Check whether database exists
 If Len(Dir(Location)) Then
    'If backup is neccessary, then do it.
    If BackupOriginal = True Then
        strBackupFile = GetTemporaryPath & "CompactData.mdb"
        If Len(Dir(strBackupFile)) Then Kill strBackupFile
        FileCopy Location, strBackupFile
    End If
    DoEvents
    'Make temporary file..
    strTempFile = GetTemporaryPath & "TempData.mdb"
    If Len(Dir(strTempFile)) <> 0 Then Kill strTempFile
    DoEvents
    'See.. this is how we compact Access database
    'with password protected.
    'DBEngine.CompactDatabase strOldDb, strNewDb,,,";pwd=yourpassword;"
    DBEngine.CompactDatabase Location, strTempFile, , , ";pwd=masino2002;"
    DoEvents
    'Delete a real database
    Kill Location
    DoEvents
    'Copy temporary compacted file to real one
    FileCopy strTempFile, Location
    DoEvents
    'Delete database temporary file
    Kill strTempFile
    DoEvents
    FinishWaiting
    'Show message
    MsgBox "Database was successfully compacted!", _
           vbInformation, "Compact Database"
 End If
 Exit Sub
CompactErr:
  Screen.MousePointer = vbDefault
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to compact database.")
  Call Message("Compact Database failed. Please try another time.")
  Select Case Err.Number
         Case 53
             MsgBox "File not found!", _
                    vbExclamation, "File Not Found"
         Case 70  'Database in in use by application
             MsgBox "Database is in use by application!" & vbCrLf & _
                    "" & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try " & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Database"
         Case 75  'Path/file not found
             MsgBox "Path/file not found." & vbCrLf & _
                    "Please select database!", vbCritical, _
                    "Database"
         Case 3343  'Not Access 97
             MsgBox "Database is not Access 97" & vbCrLf & _
                    "or not Access (mdb) file!", vbCritical, _
                    "Unknown Database"
         Case 3356  '
             MsgBox "Database is in use by application!" & vbCrLf & _
                    "" & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try " & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Database"
         Case Else
             MsgBox Err.Number & " - " & _
                    Err.Description & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try" & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Error"
     Exit Sub
  End Select

End Sub

'Get database temporal path while copying...
Public Function GetTemporaryPath()
  Dim strFolder As String
  Dim lngResult As Long
  strFolder = String(MAX_PATH, 0)
  lngResult = GetTempPath(MAX_PATH, strFolder)
  If lngResult <> 0 Then
    GetTemporaryPath = Left(strFolder, InStr(strFolder, _
                       Chr(0)) - 1)
  Else
    GetTemporaryPath = ""
  End If
End Function

'Repair database Access
Public Sub RepairJetDatabase(Location As String, _
           Optional BackupOriginal As Boolean = True)
On Error GoTo CompactErr
 Dim strBackupFile As String
 Dim strTempFile As String
 'Check whether database exists...
 If Len(Dir(Location)) Then
    'If backup neccessary, then do it!
    If BackupOriginal = True Then
       strBackupFile = GetTemporaryPath & "RepairData.mdb"
       If Len(Dir(strBackupFile)) Then Kill strBackupFile
       FileCopy Location, strBackupFile
    End If
    DoEvents
    'Repair database file now...!
    DBEngine.RepairDatabase Location
    DoEvents
    FinishWaiting
    MsgBox "Database was successfully repaired!", _
           vbInformation, "Repair Database"
 End If
 Exit Sub
CompactErr:
  Screen.MousePointer = vbDefault
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  Call Message("Repair Database failed. Please logout, and try again after login.")
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to repairt database.")
  Select Case Err.Number
         Case 53
             MsgBox "File not found!", _
                    vbExclamation, "File Not Found"
         Case 70  'Database is in use by application
             MsgBox "Database is in use by application!" & vbCrLf & _
                    "" & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try " & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Database"
         Case 75  'Path/file not found
             MsgBox "Path/file not found." & vbCrLf & _
                    "Please select database!", vbCritical, _
                    "Database"
         Case 3343  'Not Access 97
             MsgBox "Database is not Access 97" & vbCrLf & _
                    "or not Access (mdb) file!", vbCritical, _
                    "Unknown Database"
         Case 3356  '
             MsgBox "Database is in use by application!" & vbCrLf & _
                    "" & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try " & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Database"
         Case Else
             MsgBox Err.Number & " - " & _
                    Err.Description & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try" & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Error"
     Exit Sub
  End Select
End Sub


