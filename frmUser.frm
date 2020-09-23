VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUser 
   Caption         =   "User Utility"
   ClientHeight    =   7590
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11055
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Back"
         Height          =   585
         Left            =   9360
         TabIndex        =   38
         Top             =   6000
         Width           =   1080
      End
      Begin VB.CommandButton cmdDataGrid 
         Caption         =   "Adjust Data&Grid's columns width based on the longest field in underlying source"
         Height          =   585
         Left            =   6240
         TabIndex        =   36
         Top             =   6000
         Width           =   3015
      End
      Begin VB.Frame fraEntri 
         Height          =   3615
         Left            =   480
         TabIndex        =   15
         Top             =   360
         Width           =   7695
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2280
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   2835
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Password"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   2280
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   24
            Top             =   2520
            Width           =   2460
         End
         Begin VB.TextBox txtFields 
            DataField       =   "User_ID"
            Height          =   285
            Index           =   5
            Left            =   2280
            MaxLength       =   20
            TabIndex        =   23
            Top             =   2205
            Width           =   2460
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Address"
            Height          =   285
            Index           =   3
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1545
            Width           =   4570
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Occupation"
            Height          =   285
            Index           =   2
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   21
            Top             =   1230
            Width           =   4570
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Phone"
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   20
            Top             =   915
            Width           =   4570
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Name"
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   30
            TabIndex        =   19
            Top             =   600
            Width           =   4570
         End
         Begin VB.ComboBox cboLevel 
            Height          =   315
            ItemData        =   "frmUser.frx":030A
            Left            =   2280
            List            =   "frmUser.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1860
            Width           =   2460
         End
         Begin VB.CommandButton cmdLacakPassword 
            Caption         =   "Decrypt Pass&word"
            Height          =   350
            Left            =   5160
            MouseIcon       =   "frmUser.frx":030E
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CommandButton cmdApplyPassword 
            Caption         =   "Change &Password"
            Height          =   350
            Left            =   5160
            MouseIcon       =   "frmUser.frx":0460
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            DataField       =   "Level"
            Height          =   285
            Index           =   4
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   26
            Top             =   1880
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "*) Must be unique. You can not add new user with User_ID that already exists in database."
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   3240
            Width           =   6735
         End
         Begin VB.Label lblPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-type password:"
            Height          =   255
            Left            =   720
            TabIndex        =   34
            Top             =   2820
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   33
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "User_ID   *)"
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   32
            Top             =   2205
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   31
            Top             =   1875
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   30
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation:"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   29
            Top             =   1245
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   28
            Top             =   915
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   27
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.PictureBox picStatBox 
         Height          =   600
         Left            =   480
         ScaleHeight     =   540
         ScaleWidth      =   5595
         TabIndex        =   9
         Top             =   6000
         Width           =   5655
         Begin VB.CommandButton cmdFirst 
            Caption         =   "First"
            Height          =   350
            Left            =   120
            TabIndex        =   13
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Prev"
            Height          =   350
            Left            =   840
            TabIndex        =   12
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            Height          =   350
            Left            =   4080
            TabIndex        =   11
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   "Last"
            Height          =   350
            Left            =   4800
            TabIndex        =   10
            Top             =   100
            UseMaskColor    =   -1  'True
            Width           =   705
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1410
            TabIndex        =   14
            Top             =   120
            Width           =   2760
         End
      End
      Begin VB.PictureBox picButtons 
         Height          =   3465
         Left            =   8640
         ScaleHeight     =   3405
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   480
         Width           =   1755
         Begin VB.CommandButton cmdViewLogFile 
            Caption         =   "View Log &File"
            Height          =   350
            Left            =   240
            TabIndex        =   39
            Top             =   2760
            Width           =   1200
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   350
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   350
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1200
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   350
            Left            =   240
            TabIndex        =   6
            Top             =   1680
            Width           =   1200
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   350
            Left            =   240
            TabIndex        =   5
            Top             =   2040
            Width           =   1200
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&Update"
            Enabled         =   0   'False
            Height          =   350
            Left            =   240
            TabIndex        =   4
            Top             =   960
            Width           =   1200
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Enabled         =   0   'False
            Height          =   350
            Left            =   240
            TabIndex        =   3
            Top             =   1320
            Width           =   1200
         End
         Begin VB.CommandButton cmdViewLogDB 
            Caption         =   "View &Log DB"
            Height          =   350
            Left            =   240
            TabIndex        =   2
            Top             =   2400
            Width           =   1200
         End
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Height          =   1785
         Left            =   480
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4080
         Width           =   9960
         _ExtentX        =   17568
         _ExtentY        =   3149
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmUser.frm
'Description : Add, update, delete, cancel, refresh,
'              change password, and decrypt user password.
'              This menu just for user with user level 'Admin'.
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 14, 2003
'Location    : Jakarta, INDONESIA
'--------------------------------------------------------------

Option Explicit

Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim bCancelResult As Boolean
Dim Nomor As Integer
Dim NomorCari As Integer

Private Sub cboLevel_KeyPress(KeyAscii As Integer)
  ProgramActivation
  If mbEditFlag = True Then
     If KeyAscii = 13 Then txtFields(6).SetFocus
  Else
     If KeyAscii = 13 Then txtFields(5).SetFocus
  End If
End Sub

Private Sub cmdDataGrid_Click()
  ProgramActivation
  Dim intRecord As Integer
  Dim intField As Integer
  intRecord = adoPrimaryRS.RecordCount
  intField = adoPrimaryRS.Fields.Count - 1
  'call the procedure here...
  Call AdjustDataGridColumns _
  (grdDataGrid, adoPrimaryRS, intRecord, intField, True)
End Sub

Private Sub cmdApplyPassword_Click()
  ProgramActivation
  On Error GoTo GantiErr
  cmdApplyPassword.Enabled = False
  OpenConfirmation
  OpenPassword
  'Get password from textbox, then decrypt it.
  strPassword = txtFields(6).Text
  EncryptDecrypt
  txtFields(6).Text = Temp$
  txtPassword.Text = txtFields(6).Text
  lblStatus.Caption = "Edit data user..."
  mbEditFlag = True
  SetButtons False
  txtFields(6).SetFocus
  grdDataGrid.Enabled = False
  Exit Sub
GantiErr:
  MsgBox Err.Description
End Sub

Private Sub cmdViewLogDB_Click()
  ProgramActivation
  frmLogDB.Show
End Sub

Private Sub cmdViewLogFile_Click()
  ProgramActivation
  Call StartWaiting("Please wait, preparing data to display...")
  frmLogFile.Show
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
Call SaveActivityToLogDB("Start access Administrator menu.")
End Sub

Private Sub Form_Load()
ProgramActivation
On Error GoTo MessErr
  If db Is Nothing Then OpenConnection
  'Fill in the combobox for user level choices
  cboLevel.AddItem "Operator"
  cboLevel.AddItem "Manager"
  cboLevel.AddItem "Admin"
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "SHAPE {select * FROM T_User Order by User_ID} AS ParentCMD APPEND ({SELECT * FROM T_User Order by User_ID } AS ChildCMD RELATE User_ID TO User_ID) AS ChildCMD", db, adOpenStatic, adLockOptimistic
  bCancelResult = False
  Dim oText As TextBox
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  cboLevel.Text = txtFields(4).Text
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  mbDataChanged = False
  mbAddNewFlag = False
  CloseEntryForm
  grdDataGrid.TabStop = False
  cmdDataGrid_Click
  Call Message("This menu is for user who has level: 'Admin' only.")
  Screen.MousePointer = vbDefault  'Normalkan mouse
  Exit Sub
MessErr:  'Jika terjadi kesalahan... laporkan ke saya  :)
  MsgBox Err.Number & vbCrLf & Err.Description
  Unload Me  'dan langsung selesai
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If cmdUpdate.Enabled = True And cmdCancel.Enabled = True Then
     MsgBox "You have to save or cancel the changes " & vbCrLf & _
            "that you have just made before quit!", _
            vbExclamation, "Warning"
     cmdUpdate.SetFocus
     Cancel = -1
     Exit Sub
  End If
  On Error Resume Next
  Call SaveActivityToLogDB("Finish access Administrator menu.")
End Sub


Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Nomor = adoPrimaryRS.AbsolutePosition
  lblStatus.Caption = "Record number " & CStr(Nomor) & " of " & adoPrimaryRS.RecordCount
  CheckUserNavigation
End Sub

Private Sub CheckUserNavigation()
  With adoPrimaryRS
   If (.RecordCount > 1) Then
      If (.BOF) Or _
         (.AbsolutePosition = 1) Then
          cmdFirst.Enabled = False
          cmdPrevious.Enabled = False
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      ElseIf (.EOF) Or _
          (.AbsolutePosition = .RecordCount) Then
          cmdNext.Enabled = False
          cmdLast.Enabled = False
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
      Else
          cmdFirst.Enabled = True
          cmdPrevious.Enabled = True
          cmdNext.Enabled = True
          cmdLast.Enabled = True
      End If
   Else
      cmdFirst.Enabled = False
      cmdPrevious.Enabled = False
      cmdNext.Enabled = False
      cmdLast.Enabled = False
   End If
 End With
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  ProgramActivation
  On Error GoTo AddErr
  OpenEntryForm
  OpenConfirmation
    
  cboLevel.Text = "Operator"
  cmdApplyPassword.Enabled = False
  grdDataGrid.Enabled = False
  txtFields(0).SetFocus
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    mbAddNewFlag = True
    lblStatus.Caption = "Add new user..."
    SetButtons False
  End With
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
ProgramActivation
Dim intAnswer As Integer
Dim adoCekAdmin As New ADODB.Recordset
On Error GoTo DeleteErr
adoCekAdmin.Open "SELECT * FROM T_User WHERE [Level] = 'Admin'", db
If adoCekAdmin.RecordCount = 1 And cboLevel.Text = "Admin" Then
   MsgBox "You can not delete this record because" & Chr(13) & _
          "there must be at least one record in this" & Chr(13) & _
          "table with user level: 'Admin'." & vbCrLf & _
          "" & vbCrLf & _
          "If you want to delete this record, please" & vbCrLf & _
          "add a new record with user level 'Admin'" & vbCrLf & _
          "and then you can delete this record.", _
          vbCritical, "Access Denied"
   Exit Sub
End If
If adoPrimaryRS.EOF Then
   MsgBox "User data is empty!", vbCritical, "Empty"
   Exit Sub
End If
intAnswer = MsgBox("Are you sure you want to delete this record?", _
        vbQuestion + vbDefaultButton2 + vbYesNo, _
        "Delete Record")
If intAnswer = vbYes Then
   With adoPrimaryRS
    Call SaveActivityToLogDB("Delete user '" & txtFields(5).Text & "'")
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
   End With
   Exit Sub
Else
   Exit Sub
End If
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  ProgramActivation
  On Error GoTo RefreshErr
  CloseConfirmation
  If bCancelResult = True Then
     SetButtons True
     bCancelResult = False
  End If
  cmdAdd.Enabled = True
  cmdEdit.Enabled = True
  cmdUpdate.Enabled = True
  cmdDelete.Enabled = True
  cmdCancel.Enabled = True
  cmdRefresh.Enabled = True
  SetButtons True
  
  mbEditFlag = False
  mbAddNewFlag = False
  
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  Exit Sub
RefreshErr:
  cmdAdd.Enabled = False
  cmdEdit.Enabled = False
  cmdUpdate.Enabled = False
  cmdDelete.Enabled = False
  cmdCancel.Enabled = False
  cmdRefresh.Enabled = True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark <> 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  bCancelResult = True
  cmdRefresh_Click
  Exit Sub
End Sub

Private Sub cmdEdit_Click()
  ProgramActivation
  On Error GoTo EditErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
  End With
  OpenConfirmation
  OpenEntryForm
  txtFields(5).Enabled = False
  strPassword = txtFields(6).Text
  EncryptDecrypt
  txtFields(6).Text = Temp$
  txtPassword.Text = txtFields(6).Text
  lblStatus.Caption = "Edit data user..."
  mbEditFlag = True
  SetButtons False
  cmdApplyPassword.Enabled = False
  grdDataGrid.Enabled = False
  txtFields(0).SetFocus
  Exit Sub
EditErr:
  MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
  ProgramActivation
  On Error Resume Next
  CloseConfirmation
  CloseEntryForm
  cmdRefresh_Click
  If bCancelResult = True Then
     SetButtons True
     Exit Sub
  End If
  If mbEditFlag = True Then
     strPassword = txtFields(6).Text
     EncryptDecrypt
     txtFields(6).Text = Temp$
  End If
  SetButtons True
  cmdApplyPassword.Enabled = True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark <> 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  grdDataGrid.Enabled = True
End Sub

Private Sub cmdUpdate_Click()
ProgramActivation
Dim i As Integer
Dim strUsr As String
  On Error GoTo UpdateErr
  strUsr = txtFields(5).Text
  cmdUpdate.Default = False
  txtFields(4).Text = cboLevel.Text
  If Len(txtFields(0).Text) = 0 Then
     MsgBox "Enter user name.", vbCritical, "Name"
     txtFields(0).SetFocus
     Exit Sub
  ElseIf Len(txtFields(1).Text) = 0 Then
     MsgBox "Enter phone number.", vbCritical, "Phone"
     txtFields(1).SetFocus
     Exit Sub
  ElseIf Len(txtFields(2).Text) = 0 Then
     MsgBox "Enter occupation.", vbCritical, "Occupation"
     txtFields(2).SetFocus
     Exit Sub
  ElseIf Len(txtFields(3).Text) = 0 Then
     MsgBox "Enter home address.", vbCritical, "Address"
     txtFields(3).SetFocus
     Exit Sub
  ElseIf Len(cboLevel.Text) = 0 Then
     MsgBox "Choose user level.", vbCritical, "Level"
     cboLevel.SetFocus
     Exit Sub
  ElseIf Len(txtFields(5).Text) = 0 Then
     MsgBox "Enter User_ID.", vbCritical, "User_ID"
     txtFields(5).SetFocus
     Exit Sub
  ElseIf Len(txtFields(6).Text) = 0 Then
     MsgBox "Enter password.", vbCritical, "Password"
     txtFields(6).SetFocus
     Exit Sub
  ElseIf Len(txtPassword.Text) = 0 Then
     MsgBox "Password confirmation does not match!", _
            vbCritical, "Password Confirmation"
     txtPassword.SetFocus
     Exit Sub
  End If
  'Check UserID
  'Check double data in primary key (field)
  Dim cekID As New ADODB.Recordset
  cekID.Open "SELECT * FROM T_User WHERE User_ID=" & _
             "'" & Trim(txtFields(5).Text) & "'", db
  If cekID.RecordCount > 0 And mbAddNewFlag Then
     MsgBox "User_ID '" & txtFields(5).Text & "' already exists. " & vbCrLf & _
            "Please change to another User_ID!", _
            vbExclamation, "Double User_ID"
     txtFields(5).SetFocus: SendKeys "{Home}+{End}"
     Set cekID = Nothing
     Exit Sub
  End If

  If txtFields(6).Text <> txtPassword.Text Then
     MsgBox "Password confirmation does not match.", _
            vbCritical, "Password Confirmation"
     txtPassword.SetFocus
     SendKeys "{Home}+{End}"
     Exit Sub
  End If
  strPassword = txtFields(6).Text
  EncryptDecrypt
  txtFields(6).Text = Temp$
  adoPrimaryRS.UpdateBatch adAffectAll
  adoPrimaryRS.MoveLast
  If mbAddNewFlag = True Then
     Call SaveActivityToLogDB("Add new user '" & txtFields(5).Text & "'")
  End If
  Call SaveActivityToLogDB("Save changes for user '" & strUsr & "'")
  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  grdDataGrid.Enabled = True
  cmdApplyPassword.Enabled = True
  mbDataChanged = False
  CloseConfirmation
  CloseEntryForm
  lblStatus.Caption = "Record number " _
    & CStr(adoPrimaryRS.AbsolutePosition) & " of " _
    & CStr(adoPrimaryRS.RecordCount)
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
ProgramActivation
On Error Resume Next
  SendKeys "{Tab}"
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  ProgramActivation
  On Error GoTo GoFirstError
  adoPrimaryRS.MoveFirst
  mbDataChanged = False
  cboLevel.Text = txtFields(4).Text
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  ProgramActivation
  On Error GoTo GoLastError
  adoPrimaryRS.MoveLast
  mbDataChanged = False
  cboLevel.Text = txtFields(4).Text
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  ProgramActivation
  On Error GoTo GoNextError
  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    adoPrimaryRS.MoveLast
  End If
  mbDataChanged = False
  cboLevel.Text = txtFields(4).Text
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  ProgramActivation
  On Error GoTo GoPrevError
  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  cboLevel.Text = txtFields(4).Text
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Enabled = bVal
  cmdEdit.Enabled = bVal
  cmdUpdate.Enabled = Not bVal
  cmdCancel.Enabled = Not bVal
  cmdDelete.Enabled = bVal
  cmdClose.Enabled = bVal
  cmdRefresh.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cmdViewLogDB.Enabled = bVal
  cmdViewLogFile.Enabled = bVal
End Sub

Private Sub Form_Unload(Cancel As Integer)
  cmdClose_Click
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
  Response = -1
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
  If cmdAdd.Enabled = False And mbAddNewFlag = True Then
     cboLevel.Text = cboLevel.List(0)
  Else
     cboLevel.Text = txtFields(4).Text
  End If
End Sub

Private Sub txtFields_Change(Index As Integer)
ProgramActivation
On Error Resume Next
 Select Case Index
        Case 4
             cboLevel.Text = txtFields(4).Text
        Case Else
 End Select
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
ProgramActivation
Select Case Index
       Case 0, 1, 2, 3, 4, 5, 6
            If KeyCode = vbKeyEscape Then
               cmdCancel_Click
            End If
       Case Else
End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
ProgramActivation
Dim strValid As String
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
Select Case Index
       Case 0
            If KeyAscii = 13 Then
               txtFields(1).SetFocus
               SendKeys "{Home}+{End}"
            End If
       Case 1
            If KeyAscii = 13 Then
               txtFields(2).SetFocus
               SendKeys "{Home}+{End}"
            End If
       Case 2
            If KeyAscii = 13 Then
               txtFields(3).SetFocus
               SendKeys "{Home}+{End}"
            End If
       Case 3
            If KeyAscii = 13 Then
               cboLevel.SetFocus
            End If
       Case 5
            If KeyAscii = 13 Then
               txtFields(6).SetFocus
               SendKeys "{Home}+{End}"
            End If
            'Can not input comma (,) here...
            Dim strValUser As String
            strValUser = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789 "
            If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
               KeyAscii = 0
            End If
       
       Case 6
            If KeyAscii = 27 Then
               cmdCancel_Click
            ElseIf KeyAscii = vbKeyBack Then
               Exit Sub
            ElseIf KeyAscii = vbKeyDelete Then
               Exit Sub
            ElseIf KeyAscii = 13 Then
               txtPassword.SetFocus
               SendKeys "{Home}+{End}"
               Exit Sub
            End If
            If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
               KeyAscii = 0
            End If
       Case Else
End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
ProgramActivation
Dim strValid As String
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then
     cmdCancel_Click
  ElseIf KeyAscii = vbKeyBack Then
     Exit Sub
  ElseIf KeyAscii = vbKeyDelete Then
     Exit Sub
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
     KeyAscii = 0
  End If
  cmdUpdate.Default = True
End Sub

Sub CloseConfirmation()
  lblPassword.Visible = False
  txtPassword.Visible = False
  cmdLacakPassword.Enabled = True
End Sub

Sub OpenConfirmation()
  lblPassword.Visible = True
  txtPassword.Visible = True
  txtPassword.Text = ""
  cmdLacakPassword.Enabled = False
End Sub

Private Sub cmdLacakPassword_Click()
  ProgramActivation
  strPassword = txtFields(6).Text
  EncryptDecrypt
  MsgBox "User_ID = " & txtFields(5).Text & " " & Chr(13) & _
         "Password = " & (Temp$) & "", _
         vbInformation, "Password Confirmation"
End Sub

Sub CloseEntryForm()
Dim i As Integer
   For i = 0 To 6
       txtFields(i).Enabled = False
   Next i
   cboLevel.Enabled = False
   txtPassword.Enabled = False
End Sub

Sub OpenEntryForm()
Dim i As Integer
   For i = 0 To 6
       txtFields(i).Enabled = True
   Next i
   cboLevel.Enabled = True
   txtPassword.Enabled = True
End Sub

Sub OpenPassword()
Dim i As Integer
   For i = 0 To 5
       txtFields(i).Enabled = False
   Next i
   cboLevel.Enabled = False
   txtFields(6).Enabled = True
   txtPassword.Enabled = True
   SendKeys "{Home}+{End}"
End Sub

Sub CheckUserID()
Dim i As Integer
Dim NumOfRec As Integer
ReDim cekUser(adoPrimaryRS.RecordCount)
  NumOfRec = adoPrimaryRS.RecordCount
  adoPrimaryRS.MoveFirst
  For i = 1 To adoPrimaryRS.RecordCount
      cekUser(i).User = adoPrimaryRS.Fields("User_ID")
      adoPrimaryRS.MoveNext
  Next i
  adoPrimaryRS.MoveFirst
End Sub
