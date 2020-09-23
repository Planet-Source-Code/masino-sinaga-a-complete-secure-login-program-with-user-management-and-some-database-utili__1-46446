VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000F&
   Caption         =   "Login Program, (c) Masino Sinaga, masino_sinaga@yahoo.com"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   4440
      Top             =   1680
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11456
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "18/06/2003"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "08:48"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuStart 
      Caption         =   "&Start"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "L&ogout"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuOperator 
         Caption         =   "&Operator"
      End
      Begin VB.Menu mnuManager 
         Caption         =   "&Manager"
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "&Administrator"
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuChangePass 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "&Compact Database"
      End
      Begin VB.Menu mnuRepair 
         Caption         =   "&Repair Database"
      End
      Begin VB.Menu mnuBackupDB 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "&Setting..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuTips 
         Caption         =   "&Tips of day"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmMain.frm
'Description : ADO Code database programming with Login Program
'              that shows you how to make a login program
'              with managing user who uses this program.
'              User with level 'Admin' can add, update, edit,
'              delete user account, including managing log
'              activities to Database and transfer them to
'              text file (log file).
'              This is menu for this program and there are some
'              procedure that used by some menu here.
'              Reference:
'              - "Microsoft ActiveX Data Objects 2.0 Library"
'              - "Microsoft Data Binding Collection VB 6.0 (SP4)" <--(added automatically by VB6)
'              - "Microsoft DAO 3.51 Object Library (compatible
'                with Microsoft Access 97)
'              Component:
'              - "Microsoft Data Grid Control 6.0 (SP5) (OLEDB)"
'              - "Microsoft Windows Common Controls 6.0 (SP4)"
'              - "Microsoft Common Dialog Control 6.0 (SP3)"
'              Database: Microsoft Access 97
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 13, 2003
'Location    : Jakarta, INDONESIA
'-----------------------------------------------------------

Dim strTujuan As String
Dim StatusBackup As Boolean
Dim Mnt As String
Dim bQuitFromExit As Boolean

Public Sub CheckSoftware(X As Form)
On Error GoTo Pesan
    Dim SaveTitle$
    If App.PrevInstance Then
        SaveTitle$ = App.Title
        MsgBox "Login program by Masino Sinaga is now running!", _
               vbCritical, "Running"
        App.Title = ""
        X.Caption = ""
        AppActivate SaveTitle$
        SendKeys "%{ENTER}", True
        End
    End If
    Exit Sub
Pesan:
    End
    Exit Sub
End Sub

Private Sub MDIForm_Load()
  Call CheckSoftware(frmMain)
  'Ctrl-Alt-Del is off...
  DisableCtrlAltDelete (True)
  'Hide application from tasklist
  HideApp (True)
  sLuar = False
  INIFileName = App.Path & "\SettingLogin.ini"
  
  If frmTips.DontShowTipsAtStartUp.Value = 0 Then
     Screen.MousePointer = vbDefault
     Call Message("Tips of day.")
     frmMain.Show
     frmTips.Show 1
     frmTips.ZOrder 0
  Else
     Unload frmTips
  End If
  
  DoEvents
  
  AmbilSetting
    
  If gloSet.LindungLayar = 1 Then
     Timer1.Enabled = True
     Mnt = Format(Str(CInt(gloSet.IntervalMenit)), "00")
     gloSet.MenitDelay = "00:" & Mnt & ":00"
  Else
     Timer1.Enabled = False
     Mnt = "00"
     Call Message("Screen saver off")
  End If
  'Inisialisasi semua variabel dan Timer
  Gerak = False
  Aksi = False
  Timer1.Interval = 500
  Timer1.Enabled = True
  Awal = Time
  bQuitFromExit = False
  m_blnLogin = False
  mnuLogout.Enabled = False
  mnuTrans.Enabled = False
  mnuChangePass.Enabled = False
  mnuUtility.Enabled = False
End Sub

Public Sub MDIForm_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, Y As Single)
   Awal = Time
   Aksi = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If bQuitFromExit = False Then
   If db Is Nothing Then OpenConnection
   Call SaveActivityToLogDB("Exit from Program.")
   Dim Jawab1 As Integer, Jawab2 As Integer
     Jawab1 = MsgBox("Are you sure you want to quit from this program?", _
     vbQuestion + vbYesNo, "Keluar")
     If Jawab1 = vbYes Then
        Jawab2 = MsgBox("Your are successfully quit from this program right now." & Chr(13) & _
        "Program suggests you to shutdown your computer " & Chr(13) & _
        "if you do not use it any longer.  " & Chr(13) & _
        "" & vbCrLf & _
        "Do you want program to shut down this computer right now?  ", _
        vbQuestion + vbYesNo + vbDefaultButton2, "Shutdown Komputer")
        If Jawab2 = vbYes Then
           X = ExitWindowsEx(1, 0)
        End If
        Call SaveFromControlsToINI(frmSetting, "Setting")
        Call SaveFromControlsToINI(frmTips, "Tips")
        Call DisableCtrlAltDelete(False)
        CloseAllForms
     ElseIf Jawab1 = vbNo Then
        Cancel = -1
     End If
  db.Close
  Set db = Nothing
 End If
End Sub

Private Sub mnuAbout_Click()
  MsgBox "(c) Masino Sinaga, JAKARTA - INDONESIA" & vbCrLf & _
         "Start at Saturday, June 14, 2003" & vbCrLf & _
         "E-mail: masino_sinaga@yahoo.com" & vbCrLf & _
         "" & vbCrLf & _
         "I have tried this program before I posted" & vbCrLf & _
         "to www.planet-source-code.com. I didn't" & vbCrLf & _
         "find any error so far. For your information," & vbCrLf & _
         "I use OS WindowsME and VB6SP5." & vbCrLf & _
         "If you find error, please let me know." & vbCrLf & _
         "" & vbCrLf & _
         "You are free to use this code as you want" & vbCrLf & _
         "but please don't forget to write down my" & vbCrLf & _
         "name in your about app if you use my code." & vbCrLf & _
         "Thank you very much. Enjoy!" & vbCrLf & _
         "", vbInformation, "About"
End Sub

Private Sub mnuAdmin_Click()
  frmUser.Show
End Sub

Private Sub mnuChangePass_Click()
  frmChangePass.Show
End Sub

Private Sub mnuContents_Click()
  MsgBox "This program shows you how a user login " & _
         "to program and can access menu based on " & vbCrLf & _
         "his/her level. I also made a form to manage " & _
         "user account only by user with 'Admin' level." & vbCrLf & _
         "" & vbCrLf & _
         "Password field was encrypted in database." & _
         "Only user with level 'Admin' can manage " & vbCrLf & _
         "all user account include monitoring user " & _
         "log activity and decrypt the password field " & vbCrLf & _
         "if a user forgot his/her password. " & _
         "Admin can even delete the record on log file." & vbCrLf & _
         "" & vbCrLf & _
         "In this example, please use the folowing " & _
         "UserID and Password to try to login to program." & vbCrLf & _
         " " & vbCrLf & _
         "ADMIN:" & vbCrLf & _
         "- UserID = Masino " & vbCrLf & _
         "- Password = Sinaga" & vbCrLf & "" & vbCrLf & _
         "MANAGER:" & vbCrLf & _
         "- UserID = Manager " & vbCrLf & _
         "- Password = Manager" & vbCrLf & "" & vbCrLf & _
         "OPERATOR:" & vbCrLf & _
         "- UserID = Operator " & vbCrLf & _
         "- Password = Operator" & vbCrLf & _
         "", vbInformation, "Help"
End Sub

Private Sub mnuExit_Click()
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Exit from Program.")
  Call SaveFromControlsToINI(frmSetting, "Setting")
  Call SaveFromControlsToINI(frmTips, "Tips")
  DisableCtrlAltDelete (False)
  bQuitFromExit = True
  CloseAllForms
End Sub

Private Sub mnuLogin_Click()
  Call StartWaiting("Please wait, connecting data to database...")
  DoEvents
  frmLogin.Show 1 ', frmMain
End Sub

'Public, agar bisa diakses dari luar form/menu ini
Public Sub VerifyLevel()
  mnuLogin.Enabled = False
  mnuLogout.Enabled = True
  mnuTrans.Enabled = True
  mnuUtility.Enabled = True
  'mnuCompact.Enabled = False
  'mnuRepair.Enabled = False
  'mnuBackupDB.Enabled = False
  mnuChangePass.Enabled = True
  If m_Level = "Admin" Then
    mnuOperator.Enabled = True
    mnuManager.Enabled = True
    mnuAdmin.Enabled = True
  ElseIf m_Level = "Manager" Then
    mnuAdmin.Enabled = False
    mnuOperator.Enabled = False
    mnuManager.Enabled = True
  ElseIf m_Level = "Operator" Then
    mnuAdmin.Enabled = False
    mnuManager.Enabled = False
    mnuOperator.Enabled = True
  End If
End Sub

Private Sub mnuLogout_Click()
  'First, save this activity
  Call SaveActivityToLogDB("Logout from program.")
  'Then close database connection
  If Not db Is Nothing Then
     db.Close
     Set db = Nothing  'Close connection
  End If
  mnuLogout.Enabled = False
  mnuTrans.Enabled = False
  mnuChangePass.Enabled = False
  
  mnuLogin.Enabled = True
  mnuUtility.Enabled = False
  m_UserID = ""
  StatusBar1.Panels(3).Text = ""
  Call UnloadAllExceptOne("frmMain")
End Sub

Private Sub CompactRepairBackup()
  mnuUtility.Enabled = True
  mnuCompact.Enabled = True
  mnuRepair.Enabled = True
  mnuBackupDB.Enabled = True
  Call UnloadAllExceptOne("frmMain")
End Sub

Private Sub mnuManager_Click()
  frmManager.Show
End Sub

Private Sub mnuOperator_Click()
  frmOperator.Show
End Sub

Private Sub mnuCompact_Click()
Dim intAnswer As Integer
On Error GoTo Message
  Call Message("Compact Database, to compress you database file size.")
  intAnswer = MsgBox("Compact database will cause all menu that still" & vbCrLf & _
                     "opened will be closed immadiately by program." & vbCrLf & _
                     "" & vbCrLf & _
                     "Are you sure you want to compact database now?", _
                     vbQuestion + vbYesNo + vbDefaultButton2, _
                     "Compact Database")
  If intAnswer = vbYes Then
     'This open connection for writing activity to log DB
     Call StartWaiting("Please wait, compacting your database...")
     If db Is Nothing Then OpenConnection
     Call SaveActivityToLogDB("Start compact database.")
     DoEvents
     CompactRepairBackup
     StatusBar1.Panels(3).Text = m_UserID
     If Not db Is Nothing Then
       db.Close
       Set db = Nothing
     End If
     If Dir(App.Path & "\Data.ldb") <> "" Then
        Set db = Nothing
     End If
     DoEvents
     Call CompactJetDatabase(App.Path & "\Data.mdb")
     'This open connection for writing activity to log DB
     DoEvents
     If db Is Nothing Then OpenConnection
     Call SaveActivityToLogDB("Finish compact database.")
     DoEvents
  End If
  Exit Sub
Message:
  Screen.MousePointer = vbDefault
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to compact database.")
  DoEvents
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub mnuRepair_Click()
Dim intAnswer As Integer
On Error GoTo Message
  Call Message("Repair your database file.")
  intAnswer = MsgBox("Repair database will cause all menu that still" & vbCrLf & _
                     "opened will be closed immadiately by program." & vbCrLf & _
                     "" & vbCrLf & _
                     "Are you sure you want to repair your database now?", _
                     vbQuestion + vbYesNo + vbDefaultButton2, _
                     "Repair Database")
  If intAnswer = vbYes Then
     DoEvents
     Call StartWaiting("Please wait, repairing your database file...")
     DoEvents
     CompactRepairBackup
     DoEvents
     Call SaveActivityToLogDB("Start repair database.")
     DoEvents
     StatusBar1.Panels(3).Text = m_UserID
     DoEvents
     If Not db Is Nothing Then
       db.Close
       Set db = Nothing
     End If
     DoEvents
     If Dir(App.Path & "\Data.ldb") <> "" Then
        Set db = Nothing
     End If
     DoEvents
     Call RepairJetDatabase(App.Path & "\Data.mdb")
     DoEvents
     Call SaveActivityToLogDB("Finish repair database.")
  End If
  Exit Sub
Message:
  Screen.MousePointer = vbDefault
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to repair database.")
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  MsgBox Err.Number & " - " & Err.Description
End Sub

'Backup database
Private Sub mnuBackupDB_Click()
Dim Jawab As Integer
On Error GoTo Message
  Call Message("Backup your database file to another location.")
  Jawab = MsgBox("Backup database will cause all menu that still" & vbCrLf & _
                 "opened will be closed immadiately by program." & vbCrLf & _
                 "" & vbCrLf & _
                 "Are you sure you want to backup your database now?", _
                 vbQuestion + vbYesNo + vbDefaultButton2, _
                 "Backup Database")
  If Jawab = vbYes Then
     CompactRepairBackup
     DoEvents
     StatusBar1.Panels(3).Text = m_UserID
     strTujuan = ""
     DoEvents
     OpenBrowseForFolder
     DoEvents
     Call StartWaiting("Please wait, backup your database file...")
     DoEvents
     Call SaveActivityToLogDB("Start backup database to '" & Left(strTujuan, 70) & "'.")
     If Not db Is Nothing Then
       db.Close
       Set db = Nothing
     End If
     If Dir(App.Path & "\Data.ldb") <> "" Then
        Set db = Nothing
     End If
     If strTujuan = "" Then
        Unload frmWait
        MsgBox "Backup was canceled by user!", _
               vbExclamation, "Cancel Backup Database"
        Exit Sub
     End If
     DoEvents
     Call BackupDatabase(strTujuan)
     If StatusBackup = True Then
       Call SaveActivityToLogDB("Finish backup database to '" & Left(strTujuan, 70) & "'.")
       FinishWaiting
       MsgBox "Database file is successfully copied to " & vbCrLf & _
              "new directory/file as shown below: " & vbCrLf & _
              "" & vbCrLf & _
              "'" & strTujuan & "\BackupData.mdb'", _
              vbInformation, "Backup Database"
     End If
  End If
  Exit Sub
Message:
  Screen.MousePointer = vbDefault
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to backup database.")
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub BackupDatabase(strLokasi As String)
On Error GoTo Message
  FileCopy App.Path & "\Data.mdb", strTujuan & "\BackupData.mdb"
  StatusBackup = True
  Exit Sub
Message:
  Screen.MousePointer = vbDefault
  StatusBackup = False
  If FormLoadedByName("frmWait") = True Then
     Unload frmWait
     Set frmWait = Nothing
  End If
  If db Is Nothing Then OpenConnection
  Call SaveActivityToLogDB("Failed to backup database.")
  Call Message("Backup Database failed. Please logout, and try again after login.")
  Select Case Err.Number
         Case 53
             MsgBox "File not found!", _
                    vbExclamation, "File Not Found"
         Case 70  'Database is in use by another
             MsgBox "Database was using by another user!" & vbCrLf & _
                    "" & vbCrLf & _
                    "Close database file or you must" & vbCrLf & _
                    "exit from this program and try " & vbCrLf & _
                    "login again, then repeat process.", _
                    vbExclamation, "Database"
         Case 75  'Path/file not found
             MsgBox "Path/file not found." & vbCrLf & _
                    "Please select database!", vbCritical, _
                    "Database"
         Case Else
             MsgBox Err.Number & " - " & Err.Description
     Exit Sub
  End Select
End Sub

Private Sub OpenBrowseForFolder()
Dim lpIDList As Long
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo
  szTitle = "Choose destination folder/directory..."
  With tBrowseInfo
     .hWndOwner = Me.hWnd
     .lpszTitle = lstrcat(szTitle, "")
     .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
     strTujuan = Space(MAX_PATH)
     SHGetPathFromIDList lpIDList, strTujuan
         strTujuan = Left(strTujuan, InStr(strTujuan, vbNullChar) - 1)
  End If
End Sub

Private Sub mnuSetting_Click()
  ProgramActivation
  frmSetting.Show 1
End Sub

Private Sub mnuTips_Click()
  ProgramActivation
  frmTips.Show 1
End Sub

Private Sub Timer1_Timer()
Dim intDuration As Date
  Aksi = False
  If Aksi = False Then
     Gerak = False
     Timer1.Enabled = True
  Else
     Gerak = True
     Timer1.Enabled = False
  End If
  If LindungLayar = 1 Then
       Timer1.Enabled = True
       Mnt = Format(Str(CInt(gloSet.IntervalMenit)), "00")
       gloSet.MenitDelay = "00:" & Mnt & ":00"
       If Gerak = False Then
         intDuration = Time - Awal
         StatusBar1.Panels(2).Text = "Stand-by: " & Format(intDuration, "hh:mm:ss")
         If Right(StatusBar1.Panels(2).Text, 8) = gloSet.MenitDelay Then
              If StaKonek = True Then
                 frmScreenSaver.Show 1
              End If
         End If
       End If
  Else
       Timer1.Enabled = False
       Mnt = "00"
       frmMain.StatusBar1.Panels(2).Text = "Screen saver off"
  End If
End Sub
