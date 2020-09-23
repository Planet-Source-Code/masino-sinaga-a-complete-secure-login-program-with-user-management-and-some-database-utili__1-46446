VERSION 5.00
Begin VB.Form frmChangePass 
   Caption         =   "Change Password"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   HelpContextID   =   12
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   350
      HelpContextID   =   12
      Left            =   5280
      TabIndex        =   23
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame fraPassword 
      Height          =   2055
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   11415
      Begin VB.TextBox txtOldPass 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtNewPass1 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtNewPass2 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4920
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter old password:"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter new password:"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-enter new password to confirm:"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1320
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   350
      Left            =   6840
      TabIndex        =   6
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   350
      Left            =   3960
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame fraCari 
      Height          =   3975
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   11415
      Begin VB.TextBox txtUserID 
         Height          =   300
         Left            =   4920
         TabIndex        =   0
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblPet 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your UserID then press Enter on your keyboard..."
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label lblLevel 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   18
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label lblAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   780
         Left            =   4920
         TabIndex        =   17
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label lblPhone 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   16
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label lblOccupation 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   15
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4920
         TabIndex        =   14
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation:"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UserID:"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name  : frmChangePass.frm
'Description: To change user password. Only user
'             who has 'Admin' level can change
'             another user include him or herself.
'             User with 'Operator' or 'Manager' level
'             can only change him or her password.
'Author     : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site   : http://www30.brinkster.com/masinosinaga/
'             http://www.geocities.com/masino_sinaga/
'Date       : Saturday, June 13, 2003
'Location   : Jakarta, INDONESIA
'---------------------------------------------------------------

Dim rsChangePass As ADODB.Recordset
Dim SuccessUser As Boolean, Pasw As String

Private Sub cmdBack_Click()
  ProgramActivation
  Unload Me
  Set frmGantiPassword = Nothing
  frmMain.Show
End Sub

Private Sub cmdBack_GotFocus()
  ProgramActivation
  Call Message("Back to main menu and close this form...")
End Sub

Private Sub cmdBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Back to main menu and close this form...")
End Sub

Private Sub cmdOK_Click()
  ProgramActivation
  Dim UsrID As String
  Dim i As Integer
  UsrID = txtUserID.Text
  If m_Level = "Operator" Or m_Level = "Manager" Then
    If m_UserID <> txtUserID.Text Then
       MsgBox "You can not change another user's password." & vbCrLf & _
              "You can only change your password...", _
              vbExclamation, "Change Password"
       cmdOK.Enabled = False
       txtUserID.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
  End If

  ReDim tabUser((rsChangePass.RecordCount))
  If UsrID = "" Then
     MsgBox "User ID can not be empty!", _
            vbExclamation, "User ID"
     txtUserID.SetFocus
  Else
     rsChangePass.MoveFirst
     For i = 1 To rsChangePass.RecordCount
        tabUser(i).UserID = rsChangePass!User_ID
        If txtUserID.Text = tabUser(i).UserID Then
          SuccessUser = True
          Pasw = rsChangePass!password
          lblPet.Visible = False
          txtUserID.Locked = True
          lblName.Caption = rsChangePass!Name
          lblOccupation.Caption = rsChangePass!occupation
          lblPhone.Caption = rsChangePass!phone
          lblAddress.Caption = rsChangePass!Address
          lblLevel.Caption = rsChangePass!Level
          cmdOK.Enabled = False
          OpenInputNewPassword
          Exit For
        Else
          SuccessUser = False
        End If
        rsChangePass.MoveNext
     Next i
     If SuccessUser = False Then
        MsgBox "This User ID is not registered, yet!" & vbCrLf & _
               "" & vbCrLf & _
               "If you want to use this program," & vbCrLf & _
               "your user account must be registered" & vbCrLf & _
               "first in this program. Contact your" & vbCrLf & _
               "Administrator for further information.", _
               vbExclamation, "Not Registered"
        txtUserID.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
     End If
     Exit Sub
  End If
End Sub

Private Sub OpenInputNewPassword()
  ProgramActivation
  txtOldPass.Enabled = True
  txtNewPass1.Enabled = True
  txtNewPass2.Enabled = True
  txtUserID.BackColor = vbWhite
  txtOldPass.BackColor = vbWhite
  txtNewPass1.BackColor = vbWhite
  txtNewPass2.BackColor = vbWhite
  txtOldPass.SetFocus
End Sub

Private Sub cmdOK_GotFocus()
  ProgramActivation
  Call Message("Validate your User_ID...")
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Validate your User_ID...")
End Sub

Private Sub cmdChange_Click()
  ProgramActivation
  Dim intAnswer As Integer
  Dim KeyAscii As Integer
  On Error GoTo GantiErr
  If Len(Trim(txtOldPass.Text)) = 0 Then
     MsgBox "Old password can not be empty!", _
            vbExclamation, "Old Password"
     txtOldPass.SetFocus
     Exit Sub
  End If
  strPassword = Pasw
  EncryptDecrypt
  If txtOldPass.Text <> Temp$ Then
     MsgBox "Old password is wrong. Please correct it!", _
            vbExclamation, "Old Password"
     txtOldPass.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub
  End If

  If Len(Trim(txtNewPass1.Text)) = 0 Then
     MsgBox "New password can not be empty!", _
     vbExclamation, "New Password"
     txtNewPass1.SetFocus
     Exit Sub
  End If

  If txtNewPass1.Text <> txtNewPass2.Text Then
     MsgBox "The confirmation password does not match!" & vbCrLf & _
            "Please, re-type password confirmation...", _
            vbExclamation, "Password Confirmation"
     txtNewPass2.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub
  End If

  If IsNumeric(txtNewPass1.Text) = True Then
     MsgBox "Password can not contain numeric only!" & vbCrLf & _
            "" & vbCrLf & _
            "Password may contains all character only " & vbCrLf & _
            "or combination of numeric and character.", _
            vbInformation, "Password"
     txtNewPass1.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub
  End If

  'If old password matches and new password is valid, then
  'encrypt this password, but confirm to user first...
  intAnswer = MsgBox("Are you sure you want to change this password?", _
              vbQuestion + vbYesNo, "Change Password")
  If intAnswer = vbYes Then
     strPassword = txtNewPass1.Text
     EncryptDecrypt
     txtNewPass1.Text = Temp$
     strPassword = txtNewPass2.Text
     EncryptDecrypt
     txtNewPass2.Text = Temp$
     db.Execute "UPDATE T_User " & _
                "SET [Password] = '" & txtNewPass2.Text & "' " & _
                "WHERE User_ID LIKE '" & txtUserID.Text & "'"
     Call SaveActivityToLogDB("Change password '" & txtUserID.Text & "'.")
     EmptyText
     txtUserID.Locked = False
     MsgBox "Password has been changed successfully!", _
            vbInformation, "Finished Changed"
     Exit Sub
  ElseIf intAnswer = vbNo Or KeyAscii = 27 Then
     Exit Sub
  End If
GantiErr:
  MsgBox Err.Number & " - " & Err.Description & vbCrLf & _
         "Please e-mail this error to: masino_sinaga@yahoo.com", _
         vbCritical, "Error"
End Sub

Private Sub cmdChange_GotFocus()
  ProgramActivation
  Call Message("Change password now.")
End Sub

Private Sub cmdChange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Change the password now.")
End Sub

Private Sub cmdReset_Click()
  ProgramActivation
  EmptyText
End Sub

Private Sub Form_Activate()
  Dim oFrm As Form
  'This will close all forms (if already open)
  'except this form and main form (frmMain).
  'Use this on Activate event procedure.
  For Each oFrm In Forms
    If oFrm.Name <> Me.Name And Not _
      (TypeOf oFrm Is MDIForm) Then
    Unload oFrm
    Set oFrm = Nothing
  End If
Next
Call SaveActivityToLogDB("Start access Change Password menu.")
End Sub

Private Sub Form_Load()
  ProgramActivation
  If db Is Nothing Then OpenConnection
  Set rsChangePass = New ADODB.Recordset
  rsChangePass.Open "SELECT * FROM T_User", _
                    db, adOpenKeyset, adLockOptimistic
  SuccessUser = False
  cmdChange.Enabled = False
  txtUserID.Locked = False
  cmdOK.Enabled = False
  Call Message("This menu for changing user password.")
  Screen.MousePointer = vbDefault
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("This menu for changing user password.")
End Sub

Private Sub fraCari_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
End Sub

Private Sub fraPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
End Sub

Private Sub txtNewPass1_GotFocus()
  ProgramActivation
  Call Message("Enter new password...")
End Sub

Private Sub txtNewPass1_KeyPress(KeyAscii As Integer)
  ProgramActivation
  Dim strValid As String
  strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then
     cmdBack_Click
  ElseIf KeyAscii = vbKeyBack Then
     Exit Sub
  ElseIf KeyAscii = vbKeyDelete Then
     Exit Sub
  ElseIf KeyAscii = 13 Then
     txtNewPass2.SetFocus
     Exit Sub
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
     KeyAscii = 0
  End If
End Sub

Private Sub txtNewPass1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Enter new password...")
End Sub

Private Sub txtNewPass2_GotFocus()
  ProgramActivation
  Call Message("Re-enter new password to confirm...")
End Sub

Private Sub txtNewPass2_KeyPress(KeyAscii As Integer)
  ProgramActivation
  Dim strValid As String
  strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then
     cmdBack_Click
  ElseIf KeyAscii = vbKeyBack Then
     Exit Sub
  ElseIf KeyAscii = vbKeyDelete Then
     Exit Sub
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 _
     And KeyAscii <> vbKeyBack _
     And KeyAscii <> vbKeyDelete _
     And KeyAscii <> 13 Then
     KeyAscii = 0
  End If
  If Len(Trim(txtNewPass2.Text)) > 0 Then
     cmdChange.Enabled = True
     cmdChange.Default = True
  Else
     cmdChange.Enabled = False
     cmdChange.Default = False
  End If
End Sub

Private Sub txtNewPass2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Re-type your new password...")
End Sub

Private Sub txtOldPass_GotFocus()
  ProgramActivation
  Call Message("Enter your old password...")
End Sub

Private Sub txtOldPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Enter your old password...")
End Sub

Private Sub txtUserID_Change()
  ProgramActivation
  If Len(txtUserID.Text) = 0 Then
     cmdOK.Default = False
     cmdOK.Enabled = False
  Else
     cmdOK.Enabled = True
     cmdOK.Default = True
  End If
End Sub

Private Sub txtUserID_GotFocus()
  ProgramActivation
  Call Message("Enter your User_ID...")
End Sub

Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyEscape Then
     cmdBack_Click
  End If
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
  ProgramActivation
  If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
  Call SaveActivityToLogDB("Finish access Change Password menu.")
  frmMain.Show
End Sub

Private Sub txtOldPass_KeyPress(KeyAscii As Integer)
  ProgramActivation
  Dim strValid As String
  strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then
     cmdBack_Click
  ElseIf KeyAscii = vbKeyBack Then
     Exit Sub
  ElseIf KeyAscii = vbKeyDelete Then
     Exit Sub
  ElseIf KeyAscii = 13 Then
     txtNewPass1.SetFocus
     Exit Sub
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 _
     And KeyAscii <> vbKeyBack _
     And KeyAscii <> vbKeyDelete _
     And KeyAscii <> 13 Then
     KeyAscii = 0
   End If
End Sub

Private Sub EmptyText()
  lblPet.Visible = True
  lblName.Caption = ""
  lblOccupation.Caption = ""
  lblPhone.Caption = ""
  lblAddress.Caption = ""
  lblLevel.Caption = ""
  txtUserID.Text = ""
  txtOldPass.Text = ""
  txtNewPass1.Text = ""
  txtNewPass2.Text = ""
  txtOldPass.Enabled = False
  txtNewPass1.Enabled = False
  txtNewPass2.Enabled = False
  txtOldPass.BackColor = &H8000000F
  txtNewPass1.BackColor = &H8000000F
  txtNewPass2.BackColor = &H8000000F
  cmdChange.Enabled = False
  cmdOK.Enabled = False
  txtUserID.Locked = False
  txtUserID.SetFocus
End Sub

Private Sub txtUserID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Type your User_ID and press Enter...")
End Sub
