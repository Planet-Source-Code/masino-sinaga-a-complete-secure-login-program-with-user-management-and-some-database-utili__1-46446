VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogDB 
   Caption         =   "User Activity Log DB"
   ClientHeight    =   7590
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmLogDB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame fraTransfer 
      Caption         =   "Transfer to Text File"
      Height          =   1095
      Left            =   7920
      TabIndex        =   35
      Top             =   4920
      Width           =   3615
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go!"
         Height          =   350
         Left            =   2520
         TabIndex        =   39
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   37
         Text            =   "MyLogFile-Login"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   ".txt"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblFileName 
         Caption         =   "File name:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraDelete 
      Caption         =   "Delete:"
      Height          =   975
      Left            =   7920
      TabIndex        =   31
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   350
         Left            =   2520
         TabIndex        =   34
         Top             =   360
         Width           =   960
      End
      Begin VB.OptionButton optDelete 
         Caption         =   "&Records in Datagrid"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   550
         Width           =   1935
      End
      Begin VB.OptionButton optDelete 
         Caption         =   "&Selected record"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame fraSort 
      Caption         =   "Sort:"
      Height          =   1095
      Left            =   7920
      TabIndex        =   27
      Top             =   2450
      Width           =   3615
      Begin VB.ComboBox cboField 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Descending"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optSort 
         Caption         =   "Ascending"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "&Sort"
         Height          =   350
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraFilter 
      Caption         =   "Filter:"
      Height          =   735
      Left            =   7920
      TabIndex        =   19
      Top             =   1560
      Width           =   3615
      Begin VB.TextBox txtFilter 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "&Filter"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraFind 
      Caption         =   "Find:"
      Height          =   1095
      Left            =   7920
      TabIndex        =   15
      Top             =   300
      Width           =   3615
      Begin VB.CommandButton cmdFindFirst 
         Caption         =   "&Find First"
         Enabled         =   0   'False
         Height          =   350
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1320
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   350
         Left            =   2400
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDataGrid 
      Caption         =   "Adjus Data&Grid Columns"
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   6480
      Width           =   2280
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Back"
      Height          =   375
      Left            =   10320
      TabIndex        =   30
      Top             =   6480
      Width           =   1200
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   6480
      Width           =   1200
   End
   Begin VB.PictureBox picStatBox 
      Height          =   600
      Left            =   720
      ScaleHeight     =   540
      ScaleWidth      =   5595
      TabIndex        =   9
      Top             =   6360
      Width           =   5655
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   350
         Left            =   4800
         TabIndex        =   13
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next"
         Height          =   350
         Left            =   4080
         TabIndex        =   12
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Prev"
         Height          =   350
         Left            =   840
         TabIndex        =   11
         Top             =   100
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   350
         Left            =   120
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
   Begin VB.TextBox txtFields 
      DataField       =   "Time"
      Height          =   285
      Index           =   3
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date"
      Height          =   285
      Index           =   2
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Activity"
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   740
      Width           =   6015
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User_ID"
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   4695
      Left            =   720
      TabIndex        =   8
      Top             =   1320
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   8281
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   6
      Top             =   420
      Width           =   525
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   420
      Width           =   525
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity:"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   735
      Width           =   675
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User_ID:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   420
      Width           =   675
   End
End
Attribute VB_Name = "frmLogDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmLogDB.frm
'Description : Display records of user activities in program
'              with: find first, find next, filter, sort,
'              and transfer data from database to text file.
'              This menu just for user with user level 'Admin'.'
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 14, 2003
'Location    : Jakarta, INDONESIA
'--------------------------------------------------------------

Option Explicit

Dim WithEvents adoPrimaryRS As ADODB.Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim adoFilter As ADODB.Recordset
Dim adoSort As ADODB.Recordset

Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim Nomor As Long

Dim intCount As Integer
Dim intPosition As Integer
Dim bFound As Boolean
Dim strFind As String, strFindNext As String
Dim strResult As String
Dim bCancel As Boolean
Dim adoField As ADODB.Field


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

Private Sub cmdFilter_Click()
ProgramActivation
On Error Resume Next
  Set adoFilter = New ADODB.Recordset
  Set adoFilter = adoPrimaryRS
  FilterInAllFields
End Sub

Private Sub cmdFindFirst_Click()
  ProgramActivation
  FindFirstInAllFields
End Sub

Private Sub cmdFindNext_Click()
  ProgramActivation
  FindNextInAllFields
End Sub

Private Sub cmdGo_Click()
  ProgramActivation
  TransferLogDBToTextFile
End Sub

Function TransferLogDBToTextFile()
ProgramActivation
Dim cn As ADODB.Connection
Dim rsTextFile As ADODB.Recordset
Dim nmf, nmdir As String
Dim strUserID, strActivity, strDate, strTime, v_Rp As String
Dim FileNumber As Integer, i As Long

On Error GoTo MessDBToTxt
nmdir = App.Path
nmf = nmdir & "\" & txtFileName.Text & ".txt"
FileNumber = FreeFile
  Set cn = New ADODB.Connection
  cn.CursorLocation = adUseClient
  cn.Open "PROVIDER=MSDataShape;" & _
          "Data PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & App.Path & "\Data.mdb;" & _
          "Jet OLEDB:Database Password=masino2002;"""
  Set rsTextFile = New ADODB.Recordset
  rsTextFile.Open "SELECT * FROM T_Log", cn
  Open nmf For Output As #FileNumber
    rsTextFile.MoveFirst
    frmWait.prgBar1.Max = rsTextFile.RecordCount
    i = 0
    Do While Not rsTextFile.EOF
      DoEvents
      strUserID = rsTextFile.Fields(0)
      strActivity = rsTextFile.Fields(1)
      strDate = rsTextFile.Fields(2)
      strTime = rsTextFile.Fields(3)
      DoEvents
      i = i + 1
      DoEvents
      frmWait.prgBar1.Value = i
      DoEvents
      Print #FileNumber, _
          strUserID & "," & strActivity & "," _
          & strDate & "," & strTime
      DoEvents
      rsTextFile.MoveNext
    Loop
    frmWait.prgBar1.Value = 0
    rsTextFile.Close
    MsgBox "Transfer log user-activities from database to log text file successful." & vbCrLf & _
           "" & vbCrLf & _
           "You can see this log file at: " & vbCrLf & _
           App.Path & "\" & txtFileName.Text & ".txt." & vbCrLf & _
           "" & vbCrLf & _
           "Please exit from this program first if you want" & vbCrLf & _
           "to view the log file with Notepad or Wordpad.", _
           vbInformation, "Transfer OK"
    Exit Function
  Close #FileNumber
  If Not cn Is Nothing Then
     cn.Close
     Set cn = Nothing
  End If
  Exit Function
MessDBToTxt:
  If Not cn Is Nothing Then
     cn.Close
     Set cn = Nothing
  End If
  Select Case Err.Number
         Case 55
              MsgBox "Text file already open right now." & vbCrLf & _
                     "Please type another file name!", _
                     vbInformation, "Already Open"
              txtFileName.SetFocus
              SendKeys "{Home}+{End}"
         Case Else
              MsgBox Err.Number & " " & _
                     Err.Description, vbCritical, "Error"
  End Select
  If Not rsTextFile Is Nothing Then Set rsTextFile = Nothing
End Function

Private Sub Form_Activate()
ProgramActivation
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
  cmdDataGrid_Click
End Sub

Private Sub Form_Load()
ProgramActivation
Dim rs As ADODB.Recordset
  If Not db Is Nothing Then OpenConnection
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {select * from T_Log Order by Date, Time} AS ParentCMD APPEND ({select * from T_Log Order by Date, Time} AS ChildCMD RELATE User_ID TO User_ID) AS ChildCMD", db, adOpenStatic, adLockOptimistic
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set grdDataGrid.DataSource = adoPrimaryRS
  Set rs = New ADODB.Recordset
  rs.Open "T_Log", _
          db, adOpenKeyset, adLockOptimistic, adCmdTable
  cboField.Clear
  For Each adoField In rs.Fields
      cboField.AddItem adoField.Name
  Next
  cboField.Text = cboField.List(0)
  rs.Close
  mbDataChanged = False
  cmdLast.Value = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ProgramActivation
  If mbEditFlag Or mbAddNewFlag Then Exit Sub
  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset) '
  ProgramActivation
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
  'This is where you put validation code
  'This event gets called when the following actions occur
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

Private Sub cmdDelete_Click()
ProgramActivation
On Error GoTo DeleteErr
Dim intAnswer As Integer
  If optDelete(0).Value = True Then
    intAnswer = _
        MsgBox("Are you sure you want to delete selected record?", _
        vbQuestion + vbDefaultButton2 + vbYesNo, _
        "Delete Record")
    If intAnswer <> vbYes Then Exit Sub
    With adoPrimaryRS
      .Delete
      .MoveNext
      If .EOF Then .MoveLast
    End With
  Else  'all record in a recordset
    If adoFilter Is Nothing Then
       intAnswer = _
        MsgBox("Are you sure you want to delete all records in DataGrid?", _
        vbQuestion + vbDefaultButton2 + vbYesNo, _
        "Delete All Record in DataGrid")
       If intAnswer <> vbYes Then Exit Sub
       Dim i As Long, num As Long
       num = adoPrimaryRS.RecordCount
       adoPrimaryRS.MoveFirst
       For i = 1 To num
           adoPrimaryRS.Delete
           If adoPrimaryRS.EOF And adoPrimaryRS.BOF Then
           Else
              adoPrimaryRS.MoveNext
           End If
       Next i
       MsgBox "All records in DataGrid has been deleted..." & vbCrLf & _
              "Press 'Refresh' to display the other records.", _
              vbInformation, "Delete Recordset OK"
    Else
       intAnswer = _
        MsgBox("Are you sure you want to delete all records in DataGrid?", _
        vbQuestion + vbDefaultButton2 + vbYesNo, _
        "Delete All Record in DataGrid")
       If intAnswer <> vbYes Then Exit Sub
       Dim j As Long, jlh As Long
       jlh = adoFilter.RecordCount
       adoFilter.MoveFirst
       For j = 1 To jlh
           adoFilter.Delete
           If adoFilter.EOF And adoFilter.BOF Then
           Else
              adoFilter.MoveNext
           End If
       Next j
       MsgBox "All records in DataGrid has been deleted..." & vbCrLf & _
              "Press 'Refresh' to display the other records.", _
              vbInformation, "Delete Recordset OK"
    End If
  End If
  Exit Sub
DeleteErr:
  Select Case Err.Number
         Case 3021  'No record
              cmdDelete.Enabled = False
              MsgBox "There is no record in DataGrid now." & vbCrLf & _
                     "Press 'Refresh' to display record!", vbInformation, "No Record"
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description, vbCritical, "Error"
  End Select
End Sub

Private Sub cmdRefresh_Click()
  ProgramActivation
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set adoSort = Nothing
  Set adoFilter = Nothing
  Set grdDataGrid.DataSource = Nothing
  Set adoPrimaryRS = New ADODB.Recordset
  adoPrimaryRS.Open "SHAPE {select * from T_Log Order by [Date], [Time]} AS ParentCMD APPEND ({select * from T_Log Order by [Date], [Time]} AS ChildCMD RELATE User_ID TO User_ID) AS ChildCMD", db, adOpenStatic, adLockOptimistic
  Set grdDataGrid.DataSource = adoPrimaryRS.DataSource
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  Set grdDataGrid.DataSource = adoPrimaryRS
  cmdDelete.Enabled = True
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  ProgramActivation
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  ProgramActivation
  On Error GoTo GoFirstError
  If adoFilter Is Nothing Then
     adoPrimaryRS.MoveFirst
  Else
     adoFilter.MoveFirst
  End If
  mbDataChanged = False
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  ProgramActivation
  On Error GoTo GoLastError
  If adoFilter Is Nothing Then
     adoPrimaryRS.MoveLast
  Else
     adoFilter.MoveLast
  End If
  mbDataChanged = False
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  ProgramActivation
  On Error GoTo GoNextError
  If adoFilter Is Nothing Then
     If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
     If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveLast
        MsgBox "This is the last record.", _
               vbInformation, "Last Record"
     End If
  Else
     If Not adoFilter.EOF Then adoFilter.MoveNext
     If adoFilter.EOF And adoFilter.RecordCount > 0 Then
        Beep
        adoFilter.MoveLast
        MsgBox "This is the last record.", _
               vbInformation, "Last Record"
     End If
  End If
  mbDataChanged = False
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  ProgramActivation
  On Error GoTo GoPrevError
  If adoFilter Is Nothing Then
     If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
     If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
        Beep
        adoPrimaryRS.MoveFirst
        MsgBox "This is the first record.", _
               vbInformation, "First Record"
     End If
  Else
     If Not adoFilter.BOF Then adoFilter.MovePrevious
     If adoFilter.BOF And adoFilter.RecordCount > 0 Then
        Beep
        adoFilter.MoveFirst
        MsgBox "This is the first record.", _
               vbInformation, "First Record"
     End If
  End If
  mbDataChanged = False
  Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub

'This will search data in all fields for the very first time
Private Sub FindFirstInAllFields()
Dim strstrResult As String, strFound As String
Dim i As Integer, j As Integer, k As Integer
  'Always start from first record
  adoPrimaryRS.MoveFirst
  strFind = txtFind.Text
Ulang:
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
     adoPrimaryRS.MoveLast
     MsgBox "'" & strFind & "' not found.", _
            vbExclamation, "Finished Searching"
     cmdFindNext.Enabled = False
     Exit Sub
  End If
  strstrResult = "":  strFound = ""
  For i = 0 To 3  'This iteration for data in textbox
      strResult = UCase(txtFields(i).Text)
      If InStr(1, UCase(txtFields(i).Text), UCase(strFind)) > 0 Then
         strstrResult = "" & strstrResult & "Found '" & strFind & "' at:" & vbCrLf & _
                      ""
       For j = 0 To 3 'This iteration for data in datagrid
          strResult = UCase(txtFields(j).Text)
          If InStr(1, UCase(txtFields(j).Text), UCase(strFind)) > 0 Then
             strFindNext = strFind
             'If we found it, tell user which position
             'it is...
              strstrResult = strstrResult & "" & vbCrLf & _
                 "  Record number " & CStr(adoPrimaryRS.AbsolutePosition) & "" & vbCrLf & _
                 "  - Field name: " & txtFields(j).DataField & "" & vbCrLf & _
                 "  - Contains: " & txtFields(j).Text & "" & vbCrLf & _
                 "  - Columns number: " & j + 1 & " in DataGrid."
             For k = 0 To adoPrimaryRS.Fields.Count - 1
                If adoPrimaryRS.Fields(k).Name = "ChildCMD" Then
                  Exit For
               End If
               strFound = strFound & vbCrLf & _
                         adoPrimaryRS.Fields(k).Name & ": " & _
                         vbTab & adoPrimaryRS.Fields(k).Value
             Next k
             'Because we found, make cmdFindNext active...
             cmdFindNext.Enabled = True
          Else
          End If
       Next j  'End of iteration in datagrid
       Exit Sub
    Else
    End If
  Next i  'End of iteration in textBox
  'If we don't find in first record, move to next record
  adoPrimaryRS.MoveNext
  GoTo Ulang
End Sub


'This will search data from the record position
'we found in FindFirstInAllFields procedure above.
'
Private Sub FindNextInAllFields()
Dim m As Integer, n As Integer, k As Integer
Dim strstrResult As String, strFound As String
strFindNext = strFind
If Len(Trim(strResult)) = 0 Then
   FindFirstInAllFields
   Exit Sub
End If
'Start from record position we found in FindFirstInAllFields
adoPrimaryRS.MoveNext
strFound = "": strstrResult = ""
Ulang:
  'If we don't find it
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
     adoPrimaryRS.MoveLast
     MsgBox "'" & strFindNext & "' not found.", _
            vbExclamation, "Finished Searching"
     Exit Sub
  End If
  For n = 0 To 3  'This iteration for textbox
    strResult = UCase(txtFind.Text)
    'If we found it, all or similiar to it
    If InStr(1, UCase(txtFields(n).Text), UCase(strFindNext)) > 0 Then
       strstrResult = "" & strstrResult & "Found '" & strFindNext & "' at:" & vbCrLf & _
                      ""
       For m = 0 To 3 'This iteration for datagrid
          strResult = UCase(txtFind.Text)
          If InStr(1, UCase(txtFields(m).Text), UCase(strFindNext)) > 0 Then
             'If we found, tell user which record position
             'it is..
             strstrResult = strstrResult & vbCrLf & _
                 "  Record number " & CStr(adoPrimaryRS.AbsolutePosition) & "" & vbCrLf & _
                 "  - Field name: " & txtFields(m).DataField & "" & vbCrLf & _
                 "  - Contains: " & txtFields(m).Text & "" & vbCrLf & _
                 "  - Column number: " & m + 1 & " in DataGrid."
             For k = 0 To adoPrimaryRS.Fields.Count - 1
                If adoPrimaryRS.Fields(k).Name = "ChildCMD" Then
                  Exit For
               End If
               'Get all data we found in that record
               strFound = strFound & vbCrLf & _
                         adoPrimaryRS.Fields(k).Name & ": " & _
                         vbTab & adoPrimaryRS.Fields(k).Value
             Next k
             Exit Sub
          Else
          End If
       Next m  'End of iteration in DataGrid
       Exit Sub
    Else
    End If
  Next n  'End of iteration in TextBox
  adoPrimaryRS.MoveNext
  GoTo Ulang
End Sub

Private Sub FilterInAllFields()
Dim rs As New ADODB.Recordset
Dim kriteria As String
Dim strCriteria As String, strField As String
Dim intField As Integer, i As Integer, j As Integer
Dim tabField() As String
On Error GoTo Message
  'Always start from the first record
  rs.Open "SELECT * FROM T_Log", db
  rs.MoveFirst
  'To get the criteria and to make SQL Statement
  'shorter, we can use this way...
  strCriteria = ""
  intField = rs.Fields.Count
  ReDim tabField(intField)
  intField = rs.Fields.Count
  ReDim tabTgl(intField)
  
  i = 0
  For Each adoField In rs.Fields
      tabField(i) = adoField.Name
      i = i + 1
  Next
      
  For i = 0 To intField - 1
     If i <> intField - 1 Then
        strField = strField & tabField(i) & ","
        strCriteria = strCriteria & _
           tabField(i) & " LIKE '%" & txtFilter.Text & "%' Or "
     Else
        strField = strField & tabField(i) & " "
        strCriteria = strCriteria & tabField(i) & " LIKE '%" & txtFilter.Text & "%' "
     End If
  Next i
  Set adoFilter = New ADODB.Recordset
     adoFilter.Open _
     "SHAPE " & _
     "{SELECT " & strField & " FROM T_Log " & _
     "WHERE " & strCriteria & " ORDER BY [Date], [Time]} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT " & strField & " FROM T_Log " & _
     "WHERE " & strCriteria & " ORDER BY [Date], [Time]} " & _
     "AS ChildCMD RELATE User_ID TO User_ID) " & _
     "AS ChildCMD", db, adOpenStatic, adLockOptimistic
   
  If adoFilter.RecordCount > 0 Then
     Set grdDataGrid.DataSource = adoFilter.DataSource
     Set adoPrimaryRS = adoFilter
     Dim oTextData As TextBox
     For Each oTextData In Me.txtFields
         Set oTextData.DataSource = adoFilter.DataSource
     Next
     'Go to the first record, always
     cmdFirst.Value = True
  Else
     cmdRefresh.Value = True
     MsgBox "'" & txtFilter.Text & "' not found " & _
            "in this table.", _
            vbExclamation, "No Result"
  End If
  
  Exit Sub
Message:
  'MsgBox Err.Number & " - " & Err.Description
  MsgBox "'" & txtFilter.Text & "' not found " & Chr(13) & _
         "in this log DB.", _
         vbExclamation, "No Result"
End Sub

Private Sub cmdSort_Click()
ProgramActivation
Dim TipeSort As String
   If optSort(0).Value = True Then
      TipeSort = "ASC"
   Else
      TipeSort = "DESC"
   End If
   Set adoSort = New ADODB.Recordset
   adoSort.Open "SHAPE " & _
     "{SELECT * FROM T_Log " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ParentCMD APPEND " & _
     "({SELECT * FROM T_Log " & _
     "ORDER BY " & cboField.Text & " " & TipeSort & "} " & _
     "AS ChildCMD RELATE User_ID TO User_ID) " & _
     "AS ChildCMD", _
     db, adOpenStatic, adLockOptimistic
     On Error Resume Next
     If adoSort.RecordCount > 0 Then
        Set grdDataGrid.DataSource = adoSort.DataSource
        Set adoPrimaryRS = adoSort.DataSource
        Dim oTextData As TextBox
        For Each oTextData In Me.txtFields
            Set oTextData.DataSource = adoSort.DataSource
        Next
        Set adoPrimaryRS = adoSort
        cmdFirst.Value = True
     End If
   
   Exit Sub
Message:
     MsgBox Err.Number & " - " & _
            Err.Description, _
            vbExclamation, "No Result"

End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
ProgramActivation
'Validate every character that user typed in txtUserID
Dim strValid As String
'This is the valid string user can type to this textbox
'It's up to you, if you want to add another character...
strValid = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz0123456789"
  If KeyAscii = 27 Then 'If user hit Esc button in keyboard
     cmdClose_Click    'Exit from login
  ElseIf KeyAscii = 13 Then 'If user hit Enter
     cmdGo.SetFocus   'move to next field (Password)
     SendKeys "{Home}+{End}" 'Highlight Password
  End If
  If InStr(strValid, Chr(KeyAscii)) = 0 And _
     KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete And _
     KeyAscii <> vbKeySpace Then
     KeyAscii = 0  '
  End If
End Sub

Private Sub txtFilter_Change()
  ProgramActivation
  If Len(Trim(txtFilter.Text)) = 0 Then
     cmdFilter.Enabled = False
  Else
     cmdFilter.Enabled = True
  End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
  ProgramActivation
  cmdFilter.Default = True
End Sub

Private Sub txtFind_Change()
  ProgramActivation
  If Len(Trim(txtFind.Text)) = 0 Then
     cmdFindFirst.Enabled = False
     cmdFindNext.Enabled = False
  Else
     cmdFindFirst.Enabled = True
     cmdFindFirst.Default = True
  End If
End Sub
