VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   HelpContextID   =   1
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      HelpContextID   =   1
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      HelpContextID   =   1
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   350
      HelpContextID   =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   350
      HelpContextID   =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   120
   End
   Begin VB.Label lblCounter 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3100
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblGuide 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your User_ID and Password..."
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "User_ID:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1245
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmLogin.frm
'Description : Validate User ID dan Password from user input.
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 13, 2003
'Location    : Jakarta, INDONESIA
'-----------------------------------------------------------
Option Explicit

Dim intTry As Integer, intCount As Byte
Dim strDesc As String, strStatus As String
Dim blnUserOK As Boolean, blnPassOK As Boolean
Dim strName As String, strSay As String

Private Sub cmdCancel_Click()
  ProgramActivation
  Call Message("Login canceled by user.")
  Unload Me
  frmMain.Show
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Cancel login now.")
End Sub

Private Sub cmdOK_Click()
ProgramActivation
On Error GoTo MessErr
  Dim i As Integer
  Dim strPassEncrypted As String
  Dim strUserID As String
   
  Timer1.Enabled = False
  
  'Change mouse pointer to hourglass because
  'there will be a process for a while...
  Screen.MousePointer = vbHourglass
  
  'If we have not been connected to database, connect it!
  If db Is Nothing Then OpenConnection
  'Open table T_User
  OpenTableUser
  'Get User_ID input from user
  strUserID = txtUserID.Text
  'Get Password input from user
  strPassword = Trim(txtPassword.Text)
  'Decrypt Password now. This procedure is in the .bas module
  EncryptDecrypt
  'Take the result (Password decrypted) to txtPassword
  txtPassword.Text = Temp$
  'Re-declare dynamic array with number of record in T_User
  ReDim tabUser(NumOfUser)
      'If txtUserID still blank
      If strUserID = "" Then
         strDesc = "User ID is empty."
         strStatus = "User"
         GoTo CountFail
         txtUserID.SetFocus
      'If Password still blank
      ElseIf txtPassword.Text = "" Then
         strDesc = "Password is empty."
         strStatus = "Passw"
         GoTo CountFail
         txtPassword.SetFocus
      Else 'Both UserID and Password are not empty string
         rsUser.MoveFirst  'Always from first record
         DoEvents
         'Iteration as many as number of record in T_User
         For i = 1 To NumOfUser
             'Get UserID from recordset
             tabUser(i).UserID = rsUser!User_ID
             'If UserID matches with UserID from user
             If Trim(txtUserID.Text) = Trim(tabUser(i).UserID) Then
                blnUserOK = True  'User ID found here!
                'Get encrypted Password from recordset
                strPassEncrypted = Trim(rsUser!password)
                If blnUserOK = True Then 'If UserID OK
                   'And If encrypted Password from user matches
                   'with encrypted Password from recordset
                   If Trim(txtPassword.Text) = strPassEncrypted Then
                       blnPassOK = True 'Password found here!
                       If blnPassOK = True Then
                          'If both UserID and Password OK
                          If blnUserOK = True And blnPassOK = True Then
                             'Get Name, Level and take it to local var
                             'Get UserID and take it public var
                             strName = rsUser!Name
                             m_Level = rsUser![Level]
                             m_UserID = rsUser!User_ID
                             DoAfterLoginOK
                             Screen.MousePointer = vbDefault
                             Exit Sub
                          End If
                       Else 'Password not found yet
                          blnPassOK = False
                       End If
                   Else 'Password not found yet
                       blnPassOK = False
                   End If
                Else 'UserID not found yet
                   blnUserOK = False
                End If
             Else 'UserID not match or not found
                'blnUserOK = False
             End If
         rsUser.MoveNext 'Move to next record
         Next i
         'Unmatch combination can be happened
         If blnUserOK = True And blnPassOK = False Then
            strUserID = ""
            strPassEncrypted = ""
            strStatus = "Passw"
            strDesc = "Wrong Password."
            txtPassword.SetFocus
            SendKeys "{Home}_+{End}"
         ElseIf blnUserOK = False And blnPassOK = True Or blnPassOK = False Then
            strUserID = ""
            strPassEncrypted = ""
            strStatus = "User"
            strDesc = "UserID does not exist."
            txtUserID.SetFocus
            SendKeys "{Home}+{End}"
         End If
CountFail:
         intTry = intTry + 1  'Update counter if fail
         Call FailValidation(intTry) 'Take it to this function
         Exit Sub
      End If
      Exit Sub 'Dont't forget this... :)
MessErr:
  Screen.MousePointer = vbDefault
  Select Case Err.Number
         Case 3704
              MsgBox "Login failed. Please re-login!", _
                     vbExclamation, "Failed"
              If FormLoadedByName("frmLogin") = True Then
                 Unload frmLogin
              End If
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description, vbCritical, "Error"
  End Select
End Sub

Private Sub DoAfterLoginOK()
  Dim strHour As Byte
  'Get hour from system
  strHour = Hour(Time)
  
  'This is for greeting when user login successful
  If strHour >= 0 And strHour < 11 Then
     strSay = "Good morning"
  ElseIf strHour >= 11 And strHour < 15 Then
     strSay = "Good day"
  ElseIf strHour >= 15 And strHour < 18 Then
     strSay = "Good afternoon"
  ElseIf strHour >= 18 And strHour <= 23 Then
     strSay = "Good evening"
  End If

  'Reference to main form (frmMain)
  With frmMain
     DoEvents
     .StatusBar1.Panels(1).Text = "Status: Login OK."
     .StatusBar1.Panels(3).Text = m_UserID
     MsgBox "" & strSay & ", " & strName & "! " & _
            "You are successfully login." & Chr(13) & _
            "" & Chr(13) & _
            "Welcome to this program and have a nice day...", _
            vbInformation, "Login OK"
     EmptyLoginForm
     Unload frmLogin
     Set frmLogin = Nothing
     DoEvents
     Call SaveActivityToLogDB("Login successful.")
     'Check user level for menu accessing
     m_blnLogin = True
     .VerifyLevel
     .Show
  End With
End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Validate user account.")
End Sub

Private Sub Form_Load()
  ProgramActivation
  intTry = 0
  blnUserOK = False
  blnPassOK = False
  If db Is Nothing Then OpenConnection
  Call Message("Validate user account.")
  FinishWaiting
  Screen.MousePointer = vbDefault
End Sub

Private Sub EmptyLoginForm()
  txtUserID.Text = ""
  txtPassword.Text = ""
End Sub

Private Sub Timer1_Timer()
 If Timer1.Enabled = True Then
  If intCount > 0 And intCount < 6 Then
     intCount = intCount - 1
  Else
     intCount = 5
  End If
  If intCount = 0 Then
     Call SaveActivityToLogDB("'" & txtUserID.Text & "' fails login. Program end.")
     End
  End If
  lblCounter.Caption = intCount
 End If
End Sub

Private Sub txtPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Enter your Password...")
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
ProgramActivation
'Validate every character that user type in txtUserID
Dim strValid As String
'This is the valid string user can type to this textbox
'It's up to you, if you want to add another character...
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then 'If user hit Esc button in keyboard
     cmdCancel_Click    'Exit from login
  ElseIf KeyAscii = 13 Then 'If user hit Enter
     txtPassword.SetFocus   'move to next field (Password)
     SendKeys "{Home}+{End}" 'Highlight Password
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeySpace Then
     KeyAscii = 0  '
  End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
ProgramActivation
'Validate every character that user type in txtUserID
Dim strValid As String
'This is the valid string user can type to this textbox
'It's up to you, if you want to add another character...
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
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

Function FailValidation(intTry)
On Error GoTo JalanPintas
Dim LastResult As Integer
    LastResult = 0
    LastResult = LastResult + intTry
    If strStatus <> "User" Then
       m_UserID = txtUserID.Text
    End If
    If LastResult < 3 Then
      Call SaveActivityToLogDB("" & txtUserID.Text & " Try: #" & LastResult & ": " & strDesc & "")
      MsgBox "This is the chance number " & LastResult & ": " & strDesc & "" & vbCrLf & _
             "" & vbCrLf & _
             "You still got " & 3 - LastResult & " more chances.", _
             vbExclamation, "Trying #" & LastResult
      Screen.MousePointer = vbDefault
    Else
      Call SaveActivityToLogDB("" & txtUserID.Text & " Try: #" & LastResult & ": " & strDesc & "")
      MsgBox "This is the chance number " & LastResult & ": " & strDesc & "" & vbCrLf & _
             "" & vbCrLf & _
             "Sorry, your 3 chances is up. " & _
             "Please try again another time.", _
             vbCritical, "Access Denied"
      Screen.MousePointer = vbDefault
    End If
    If LastResult = 3 Then  'Just for three times trying!
       lblGuide.Visible = False
       lblWarning.Visible = True
       lblWarning.Move lblGuide.Left, lblGuide.Top
       lblCounter.Visible = True
       intCount = 5
       'Display warning...
       lblWarning.Caption = "This program will automaticaly end in    second..."
       lblCounter.Caption = "5"
       DoEvents
       Timer1.Enabled = True
       DoEvents
       LockLogin
    End If
    If strStatus = "User" Then
       txtUserID.SetFocus
    Else
       txtPassword.SetFocus
    End If
    SendKeys "{Home}+{End}"
    Exit Function
JalanPintas:
End Function

Private Sub txtUserID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ProgramActivation
  Call Message("Enter your UserID...")
End Sub

Private Sub LockLogin()
  txtUserID.Enabled = False
  txtUserID.BackColor = &HC0C0C0
  txtPassword.Enabled = False
  txtPassword.BackColor = &HC0C0C0
  cmdOK.Enabled = False
  cmdOK.Default = False
  cmdCancel.Default = False
End Sub
