VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTips 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN Program Ver 1.0 - Tips of day"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   HelpContextID   =   17
   Icon            =   "frmTips.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgBar1 
      Height          =   225
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox SecondChangeTips 
      Height          =   285
      Left            =   1605
      MaxLength       =   2
      TabIndex        =   10
      ToolTipText     =   "Masukkan nilai detik untuk interval pergantian tips"
      Top             =   3810
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3360
      Left            =   80
      Picture         =   "frmTips.frx":0442
      ScaleHeight     =   3300
      ScaleWidth      =   4740
      TabIndex        =   6
      Top             =   80
      Width           =   4800
      Begin VB.Label lblCounter 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   2415
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.TextBox LastTipsNumber 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox NumberOfTips 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "< &Previous"
      Height          =   350
      HelpContextID   =   17
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Ke tips sebelumnya"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      HelpContextID   =   17
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Simpan setting Tips dan selesai dengan tips"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next >"
      Height          =   350
      HelpContextID   =   17
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ke tips berikutnya"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox DontShowTipsAtStartUp 
      Caption         =   "Don't show tips at program startup"
      Height          =   255
      Left            =   100
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Ceklist jika tidak ingin menampilkan tips ini saat program dijalankan "
      Top             =   3480
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox AutoChangeTips 
      Caption         =   "Change tip after"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Ceklist jika ingin tips otomatis berganti setiap nilai detik di sebelahnya"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblSeconds 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   4560
      TabIndex        =   13
      Top             =   3840
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "second(s)"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   3855
      Width           =   855
   End
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tips As New Collection
Const TIP_FILE = "TipsLogin.txt"
Const conMinimized = 1
Dim CurrentTip As Integer, i As Integer, lR As Long
Dim intDuration As Integer

Private Sub DoNextTip()
    If CurrentTip > i Then
        CurrentTip = 1
    End If
    CurrentTip = CurrentTip + 1
    If CurrentTip > i Then
        CurrentTip = 1
    End If
    lblCounter = "Tips number " & CurrentTip & " of " & i & ""
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    frmTips.DisplayCurrentTip
End Sub
Private Sub DoPreviousTip()
    If CurrentTip < 1 Then
        CurrentTip = i
    End If
    CurrentTip = CurrentTip - 1
    If CurrentTip < 1 Then
        CurrentTip = i
    End If
    lblCounter = "Tips number " & CurrentTip & " of " & i & ""
    frmTips.DisplayCurrentTip
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer  ' Descriptor for file.
    Dim Counter As Integer
    InFile = FreeFile
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    Open sFile For Input As InFile
    i = 0
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
        i = i + 1
    Wend
    Close InFile
    frmTips.DisplayCurrentTip
    lblCounter = "Tips number " & CurrentTip & " of " & i & ""
    LoadTips = True
End Function

Private Sub cmdNextTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Display next tips.")
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Finish with tips of day.")
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Display previous tips.")
End Sub

Private Sub SecondChangeTips_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Enter the value tips will automatically changing.")
End Sub

Private Sub DontShowTipsAtStartUp_Click()
  ProgramActivation
  If DontShowTipsAtStartUp.Value = 1 Then
     frmSetting.DontShowTipsAtStartUp.Value = 1
  Else
     frmSetting.DontShowTipsAtStartUp.Value = 0
  End If
End Sub

Private Sub lblTipText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Display current tips.")
End Sub

Private Sub AutoChangeTips_Click()
  ProgramActivation
  If AutoChangeTips.Value = 1 Then
     lblSeconds.Visible = True
     SecondChangeTips.Enabled = True
     prgBar1.Visible = True
     SecondChangeTips.BackColor = vbWhite
     If SecondChangeTips.Text = "" Then
        SecondChangeTips.Text = "5"
        lblSeconds.Caption = "5"
        SecondChangeTips.SetFocus
     End If
     Timer1.Enabled = True
  Else
     prgBar1.Visible = False
     lblSeconds.Visible = False
     SecondChangeTips.BackColor = &H8000000F
     SecondChangeTips.Enabled = False
     Timer1.Enabled = False
  End If
End Sub

Private Sub cmdNextTip_Click()
  ProgramActivation
    DoNextTip
End Sub
Private Sub cmdPrevious_Click()
  ProgramActivation
    DoPreviousTip
End Sub

Private Sub Form_Load()
  ProgramActivation
  Dim ShowAtStartup As Long
  Dim LeftTips, TopTips As Integer
  TopTips = Screen.Height / 2 - Me.Height / 2
  LeftTips = Screen.width / 2 - Me.width / 2
  CurrentTip = Int(i * Rnd + 1)
  lblCounter = "Tips number " & CurrentTip & " of " & i & ""
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = _
        "File " & TIP_FILE & " not found." & vbCrLf & _
        "There is no tips can be displayed right now!" & vbCrLf & _
        "Please send e-mail to: masino_sinaga@yahoo.com"
        lblCounter.Caption = ""
    End If
    OptionStartData
    If Len(Trim(LastTipsNumber.Text)) = 0 Then
        CurrentTip = 1
    Else
        CurrentTip = Val(LastTipsNumber.Text)
        For Counter = 1 To CurrentTip
            DoNextTip
        Next Counter
    End If
    If SecondChangeTips.Text = "" Then SecondChangeTips.Text = "5"
    lblSeconds.Caption = "0"
    Timer1.Interval = 1000
    If AutoChangeTips.Value = 0 Then
       lblSeconds.Visible = False
       prgBar1.Visible = False
       Timer1.Enabled = False
       SecondChangeTips.BackColor = &H8000000F
       SecondChangeTips.Enabled = False
    Else
       lblSeconds.Visible = True
       prgBar1.Visible = True
       Timer1.Enabled = True
       SecondChangeTips.BackColor = vbWhite
       SecondChangeTips.Enabled = True
       intDuration = CInt(SecondChangeTips.Text)
    End If
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Function OptionStartData()
    Call ReadFromINIToControls(frmTips, "Tips")
End Function

Private Sub cmdOK_Click()
    ProgramActivation
    NumberOfTips.Text = i
    LastTipsNumber.Text = CurrentTip
    Call SaveFromControlsToINI(frmTips, "Tips")
    Unload Me
    CurrentTip = 0
    Call Message("")
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Awal = Time
   Aksi = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   frmMain.Caption = "LOGIN PROGRAM Ver 1.0"
   ProgramActivation
   cmdOK_Click
End Sub

Private Sub AutoChangeTips_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Checklist if you want program to change it automatically.")
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Displaying the current tips.")
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
   If SecondChangeTips.Text = "0" Then
      lblSeconds.Caption = "0"
      lblSeconds.Visible = False
      prgBar1.Visible = False
      Timer1.Enabled = False
      Exit Sub
   Else
      lblSeconds.Visible = True
      prgBar1.Visible = True
   End If
   If CInt(SecondChangeTips.Text) = CInt(lblSeconds.Caption) Then
      DoNextTip
   End If
   If CInt(lblSeconds.Caption) >= 0 And CInt(lblSeconds.Caption) < CInt(SecondChangeTips.Text) Then
      lblSeconds.Caption = Str(CInt(lblSeconds.Caption) + 1)
      prgBar1.Value = (CInt(lblSeconds.Caption) / CInt(SecondChangeTips.Text)) * 100
   Else
      lblSeconds.Caption = "1"
      prgBar1.Value = (CInt(lblSeconds.Caption) / CInt(SecondChangeTips.Text)) * 100
   End If
End Sub

Private Sub SecondChangeTips_Change()
On Error Resume Next
   ProgramActivation
   If SecondChangeTips.Text = "" Or SecondChangeTips.Text = "0" Then
      lblSeconds.Visible = False
      prgBar1.Visible = False
      lblSeconds.Caption = "0"
      SecondChangeTips.Text = "0"
      SendKeys "{Home}+{End}"
      Timer1.Enabled = False
      Exit Sub
   Else
      lblSeconds.Visible = True
      prgBar1.Visible = True
      Timer1.Enabled = True
   End If
   If CInt(SecondChangeTips.Text) < 0 Or _
      CInt(SecondChangeTips.Text) > 60 Or _
      SecondChangeTips.Text = "" Then
      SecondChangeTips.Text = "5"
      Exit Sub
   End If
   intDuration = CInt(SecondChangeTips.Text)
End Sub

Private Sub SecondChangeTips_Click()
  ProgramActivation
  SendKeys "{Home}+{End}"
End Sub

Private Sub SecondChangeTips_GotFocus()
  ProgramActivation
  SendKeys "{Home}+{End}"
End Sub

Private Sub SecondChangeTips_KeyPress(KeyAscii As Integer)
  ProgramActivation
  If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbEnter) Then
        Beep
        KeyAscii = 0
   End If
  strValid = "0123456789"
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
     KeyAscii = 0
  End If
End Sub

Private Sub SecondChangeTips_LostFocus()
ProgramActivation
On Error Resume Next
   If CInt(SecondChangeTips.Text) < 0 Or _
      CInt(SecondChangeTips.Text) > 60 Or _
      SecondChangeTips.Text = "" Then
      SecondChangeTips.Text = "5"
      MsgBox "Must be beetwen 0 and 60.", _
             vbCritical, "Invalid"
      SecondChangeTips.SetFocus
      Exit Sub
   End If
End Sub
