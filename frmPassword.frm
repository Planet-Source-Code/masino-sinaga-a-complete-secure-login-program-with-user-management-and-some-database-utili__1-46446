VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Saver Password"
   ClientHeight    =   1815
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3990
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1072.362
   ScaleMode       =   0  'User
   ScaleWidth      =   3746.394
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblPesan 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "User_ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblPetunjuk 
      BackStyle       =   0  'Transparent
      Caption         =   "Type your User_ID and Password..."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblJam 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intCnt As Integer
Dim Ket As String, Status As String
Dim UserSuccess As Boolean, PassSuccess As Boolean
Dim maks As Integer

Private Sub cmdCancel_Click()
  ProgramActivation
  Call Message("Login was canceled by user...")
  Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Pesan
  ProgramActivation
  Dim UsrID As String, Pasw As String
  Dim i As Integer
  If db Is Nothing Then OpenConnection
  OpenTableUser
  UsrID = Trim(txtUserID.Text)
  If frmMain.StatusBar1.Panels(3).Text <> "" Then
    If Trim(UsrID) <> Trim(m_UserID) Then
       MsgBox "User_ID does not match to current/active User_ID." & vbCrLf & _
              "", vbCritical, "Access Denied"
       txtUserID.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
     End If
  End If
  strPassword = Trim(txtPassword.Text)
  EncryptDecrypt
  txtPassword.Text = Temp$
  ReDim tabUser(NumOfUser)
      If UsrID = "" Then
         Ket = "User ID is empty!"
         Status = "User"
         GoTo intCntKesalahan
         txtUserID.SetFocus
      ElseIf txtPassword.Text = "" Then
         Ket = "Password is empty!"
         Status = "Passwo"
         GoTo intCntKesalahan
         txtPassword.SetFocus
      Else
         rsUser.MoveFirst
         DoEvents
         For i = 1 To NumOfUser
             tabUser(i).UserID = rsUser!User_ID
             If Trim(txtUserID.Text) = Trim(tabUser(i).UserID) Then
                UserSuccess = True
                Pasw = Trim(rsUser!password)
                If UserSuccess = True Then
                   DoEvents
                   If Trim(txtPassword.Text) = Pasw Then
                       PassSuccess = True
                       If PassSuccess = True Then
                          If UserSuccess = True And PassSuccess = True Then
                             res = SetWindowPos(frmPassword.hWnd, _
                                   HWND_NOTOPMOST, 0, 0, 0, 0, _
                                   flags)
                             res = SetWindowPos(frmScreenSaver.hWnd, _
                                   HWND_NOTOPMOST, 0, 0, 0, 0, _
                                   flags)
                             Unload frmPassword
                             Set frmPassword = Nothing
                             Unload frmScreenSaver
                             Set frmScreenSaver = Nothing
                             Exit Sub
                          End If
                       Else
                          PassSuccess = False
                       End If
                   Else
                       PassSuccess = False
                   End If
                Else
                   UserSuccess = False
                End If
             Else
             End If
             rsUser.MoveNext
             Next i
             If UserSuccess = True And PassSuccess = False Then
                UsrID = ""
                Pasw = ""
                Status = "Passwo"
                Ket = "Wrong password!"
                txtPassword.SetFocus
                SendKeys "{Home}_+{End}"
             ElseIf UserSuccess = False And PassSuccess = True Or PassSuccess = False Then
                UsrID = ""
                Pasw = ""
                Status = "User"
                Ket = "User ID is not registered."
                txtUserID.SetFocus
                SendKeys "{Home}+{End}"
             End If
intCntKesalahan:
             intCnt = intCnt + 1
             Call GetLoginFails(intCnt)
             Exit Sub
      End If
      Exit Sub
Pesan:
  Select Case Err.Number
         Case 3704
              MsgBox "Login failed. Please try again!", _
                     vbExclamation, "Fail"
         Case Else
              MsgBox Err.Number & " - " & Err.Description & vbCrLf & _
                     "E-mail this error to masino_sinaga@yahoo.com", _
         vbCritical, "Error"
  End Select
End Sub

Private Function CekFormPassword() As Boolean
  If FormLoadedByName("frmPassword") Then
     CekFormPassword = True
  Else
     CekFormPassword = False
  End If
End Function

Private Sub Form_Load()
Dim Jawab As Integer
Dim CekTanggal As String
   res = SetWindowPos(frmScreenSaver.hWnd, _
                      HWND_NOTOPMOST, 0, 0, 0, 0, _
                      flags)
   res = SetWindowPos(frmPassword.hWnd, _
                      HWND_TOPMOST, 0, 0, 0, 0, _
                      flags)
Ulangi:
  intCnt = 0
  UserSuccess = False
  PassSuccess = False
  If db Is Nothing Then OpenConnection
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
ProgramActivation
Dim strValid As String
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789!@#$%^&*()+=-_ "
  If KeyAscii = 27 Then
     cmdCancel_Click
  ElseIf KeyAscii = 13 Then
     txtPassword.SetFocus
     SendKeys "{Home}+{End}"
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
ProgramActivation
Dim strValid As String
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789!@#$%^&*()+=-_ "
  If KeyAscii = 27 Then
     cmdCancel_Click
  ElseIf KeyAscii = vbKeyBack Then
     Exit Sub
  ElseIf KeyAscii = vbKeyDelete Then
     Exit Sub
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
     KeyAscii = 0
  End If
  cmdOK.Default = True
End Sub

Function GetLoginFails(intCnt)
On Error GoTo JalanPintas
Dim intLastResult As Integer
    intLastResult = 0
    intLastResult = intLastResult + intCnt
    MsgBox Ket & ". Please try again.", _
           vbInformation, "Invalid"
    If Status = "User" Then
       txtUserID.SetFocus
    Else
       txtPassword.SetFocus
    End If
    SendKeys "{Home}+{End}"
    Exit Function
JalanPintas:
End Function
