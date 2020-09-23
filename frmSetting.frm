VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   3660
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5565
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Pre&view"
         Height          =   350
         Left            =   3960
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox ScreenSaverPassword 
         Caption         =   "Use user password to protect the screen saver"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.TextBox MinuteStandBy 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "1"
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox ScreenSaver 
         Caption         =   "Display screen saver after program idles:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Value           =   1  'Checked
         Width           =   3145
      End
      Begin VB.CheckBox RunProgramAtStartUp 
         Caption         =   "Run program at startup. Please make exe named 'LOGIN.exe'."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   4905
      End
      Begin VB.CheckBox DontShowTipsAtStartUp 
         Caption         =   "Do not show tips at program startup"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Value           =   1  'Checked
         Width           =   4575
      End
      Begin MSComCtl2.UpDown UpDn 
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "MinuteStandBy"
         BuddyDispid     =   196612
         OrigLeft        =   3840
         OrigTop         =   840
         OrigRight       =   4080
         OrigBottom      =   1125
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "minute(s)"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   350
      HelpContextID   =   13
      Left            =   4200
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  ProgramActivation
  If frmTips.DontShowTipsAtStartUp.Value = 1 Then
     DontShowTipsAtStartUp.Value = 1
  ElseIf frmTips.DontShowTipsAtStartUp.Value = 0 Then
     DontShowTipsAtStartUp.Value = 0
  End If
  Unload Me
End Sub

Private Sub cmdApply_Click()
  ProgramActivation
  SaveDataToSetting
  GetDataFromSetting
  cmdApply.Enabled = False
End Sub

Private Sub cmdOK_Click()
  ProgramActivation
  SaveDataToSetting
  GetDataFromSetting
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  ProgramActivation
  frmScreenSaver.Show 1
End Sub

Private Sub Form_Load()
  ProgramActivation
  GetDataFromSetting
  DoEvents
  If ScreenSaver.Value = 1 Then
     MinuteStandBy.Enabled = True
     MinuteStandBy.BackColor = vbWhite
     UpDn.Enabled = True
     cmdPreview.Enabled = True
  Else
     MinuteStandBy.Enabled = False
     MinuteStandBy.BackColor = &H8000000F
     UpDn.Enabled = False
     cmdPreview.Enabled = False
  End If
  If DontShowTipsAtStartUp.Value = 1 Then
     frmTips.DontShowTipsAtStartUp.Value = 1
  Else
     frmTips.DontShowTipsAtStartUp.Value = 0
  End If
  cmdApply.Enabled = False
  Call Message("This menu for program setting.")
End Sub

Private Sub DontShowTipsAtStartUp_Click()
  ProgramActivation
  If DontShowTipsAtStartUp.Value = 1 Then
     frmTips.DontShowTipsAtStartUp.Value = 1
  Else
     frmTips.DontShowTipsAtStartUp.Value = 0
  End If
  cmdApply.Enabled = True
End Sub

Private Sub DontShowTipsAtStartUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ProgramActivation
   Call Message("This menu for setting program to INI file.")
End Sub

Private Sub MinuteStandBy_Change()
  ProgramActivation
  If MinuteStandBy.Text = "" Or MinuteStandBy.Text = "0" Then
     MinuteStandBy.Text = "1"
     SendKeys "{Home}+{End}"
  End If
  cmdApply.Enabled = True
End Sub

Private Sub MinuteStandBy_KeyPress(KeyAscii As Integer)
  ProgramActivation
  If KeyAscii = 13 Then cmdOK.SetFocus
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
        MsgBox "Must be numerical data!", _
               vbInformation, "Numeric"
        MinuteStandBy.SetFocus
   End If
End Sub

Private Sub ScreenSaverPassword_Click()
  ProgramActivation
  cmdApply.Enabled = True
  If ScreenSaverPassword.Value = 1 Then
     ScreenSaver.Value = 1
  End If
End Sub

Private Sub ScreenSaver_Click()
  ProgramActivation
  If ScreenSaver.Value = 1 Then
     MinuteStandBy.Enabled = True
     MinuteStandBy.BackColor = vbWhite
     UpDn.Enabled = True
     cmdPreview.Enabled = True
  Else
     MinuteStandBy.Enabled = False
     MinuteStandBy.BackColor = &H8000000F
     UpDn.Enabled = False
     cmdPreview.Enabled = False
     ScreenSaverPassword.Value = 0
  End If
  cmdApply.Enabled = True
End Sub

Private Sub RunProgramAtStartUp_Click()
  ProgramActivation
  If RunProgramAtStartUp.Value = 1 Then
     SetRegValue HKEY_LOCAL_MACHINE, _
     "Software\Microsoft\Windows\CurrentVersion\Run", "LOGIN", App.Path & "\LOGIN.exe"
     cmdApply.Enabled = True
     Exit Sub
  End If
  If RunProgramAtStartUp.Value = 0 Then
     DeleteValue HKEY_LOCAL_MACHINE, _
     "Software\Microsoft\Windows\CurrentVersion\Run", "LOGIN"
     cmdApply.Enabled = True
     Exit Sub
  End If
End Sub

Private Sub UpDn_Change()
  ProgramActivation
  cmdApply.Enabled = True
End Sub

Private Sub SaveDataToSetting()
  If ScreenSaver.Value = 1 Then
     LindungLayar = 1
     frmMain.Timer1.Enabled = True
  Else
     LindungLayar = 0
     frmMain.StatusBar1.Panels(2).Text = "Screen saver off"
  End If
  Call SaveFromControlsToINI(frmSetting, "Setting")
End Sub

Private Sub GetDataFromSetting()
  Dim Mnt As String
  Call ReadFromINIToControls(frmSetting, "Setting")
  If ScreenSaver.Value = 1 Then
     LindungLayar = 1
     frmMain.Timer1.Enabled = True
  Else
     LindungLayar = 0
  End If
  AmbilSetting
  Mnt = Format(Str(CInt(MinuteStandBy.Text)), "00")
  gloSet.MenitDelay = "00:" & Mnt & ":00"
End Sub

