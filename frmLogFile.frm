VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Activity Log File"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmLogFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   5040
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Bro&wse..."
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fraFind 
      Caption         =   "Find:"
      Height          =   1095
      Left            =   6480
      TabIndex        =   18
      Top             =   240
      Width           =   3975
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtFind 
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdFindFirst 
         Caption         =   "&Find First"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.ProgressBar prgBar1 
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4320
      Picture         =   "frmLogFile.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Last record"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3960
      Picture         =   "frmLogFile.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Next record"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1425
      Picture         =   "frmLogFile.frx":098E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Previous record"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      Picture         =   "frmLogFile.frx":0CD0
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "First record"
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "frmLogFile.frx":1012
      Left            =   240
      List            =   "frmLogFile.frx":1014
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblFileName 
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   11295
   End
   Begin VB.Label lblPersen 
      Height          =   255
      Left            =   10560
      TabIndex        =   17
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User_ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Activity:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmLogFile.frm
'Description : You can see user activity in this
'              program from log file (.txt), find,
'              and delete its record.
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Sunday, June 15, 2003
'Location    : Jakarta, INDONESIA
'-----------------------------------------------------

Option Explicit

Dim CurrentRec As Integer
Dim FoundPos As Integer
Dim strKriteria As String
Dim bCancel As Boolean

Private Sub cmdBrowse_Click()
ProgramActivation
Dim Jawab As Integer
On Error GoTo Batal
   With CD
     .DialogTitle = "Open Log Text File..."
     .Filter = "*.txt|*.txt"
     .ShowOpen
     lblFileName.Caption = "File Name: " & CD.FileName
     Dim Contrl As Control
     For Each Contrl In Me.Controls
         If (TypeOf Contrl Is TextBox) Then Contrl.Text = ""
     Next Contrl
     lblStatus.Caption = ""
     If CD.FileName <> "" Then
        GetLogFromTextFile
     Else
        Exit Sub
     End If
     If txtUserID.Text <> "" Then
        IsValidLogFile True
     Else
        IsValidLogFile False
        lblFileName.Caption = ""
        MsgBox "File you opened is not log file for this program!", _
               vbExclamation, "Not Log File"
     End If
   End With
   Exit Sub
Batal:
   Exit Sub
End Sub

Private Sub GetLogFromTextFile()
    Dim sFile As String
    Dim NextLine As String
    Dim fNum1 As Integer
    On Error GoTo Pesan
    List1.Clear
    sFile = CD.FileName
    ' the FreeFile function assign unique number to the Filenum variable,
    ' to avoid collision with other opened file. But remember,
    ' you have to use a different variable for assign this
    ' FreeFile, even they were used in different procedure!
    fNum1 = FreeFile
    Open sFile For Input As #fNum1
    ' do until the file reach to its end
    Do Until EOF(fNum1)
    ' read one line from the file to the NextLine String
        Line Input #fNum1, NextLine
    ' add the line to the List Box
        List1.AddItem NextLine
    Loop
    ' Close the file
    Close #fNum1
    If List1.ListCount > 0 Then _
       List1.Selected(0) = True
    Exit Sub
Pesan:
    MsgBox Err.Number & " - " & Err.Description

End Sub

Private Sub cmdCancel_Click()
  ProgramActivation
  bCancel = True
  cmdCancel.Enabled = False
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDelete_Click()
ProgramActivation
Dim i, j As Integer
Dim sFileName As String
Dim fileNo As Integer
On Error Resume Next
  j = 0
  If List1.ListCount > 0 Then
    For i = 0 To List1.ListCount - 1
      If List1.Selected(i) = True Then
         j = j + 1
      End If
    Next i
    If j = 0 Then
       MsgBox "Please select record that you want to delete!", _
              vbExclamation, "Select Record"
       Exit Sub
    End If
    If MsgBox("Are you sure you want to delete the selected record?", _
            vbQuestion + vbYesNo, _
            "Delete Record") = vbYes Then
       For i = 0 To List1.ListCount - 1
         If List1.Selected(i) = True Then
            List1.RemoveItem (i)
         End If
       Next i
     
       SaveChanges
       If List1.ListCount = 0 Then
         txtUserID.Text = ""
         txtDescription.Text = ""
         txtDate.Text = ""
         txtTime.Text = ""
       Else
         'List1.Selected(0) = True
         'List1.SetFocus
       End If
        
       If List1.ListCount > 0 Then
         'List1.Selected(0) = True
       End If
      
       Exit Sub
    End If
  
  Else
    
    MsgBox "Data is empty!", vbCritical, "Empty"
  End If
  Exit Sub
Pesan:
    MsgBox "Please select record that you want to delete!", _
           vbExclamation, "Select Record"
End Sub

Private Sub SaveChanges()
'WRITE A LINE TO THE FILE
   Dim fileNo As Integer
   Dim sFileName As String
   Dim sPassword As String
   Dim sDesc As String
   Dim sLocation As String
   Dim sExpiry As String
   Dim i As Integer
   Dim panjang As Integer
   
   'On Error GoTo Pesan

   panjang = Len(Trim(txtUserID.Text))
   
  'retrieve the typed-in values
   sPassword = txtUserID.Text
   sDesc = txtDescription.Text
   sLocation = txtDate.Text
   sExpiry = txtTime.Text

   'this is the file to save to
   sFileName = App.Path & "\LogLogin3.txt"

  'get the next free file handle from Windows
   fileNo = FreeFile

  'save to disk from List1
   prgBar1.Visible = True
   prgBar1.Min = 0
   prgBar1.Max = List1.ListCount - 1
   Open sFileName For Output As #fileNo
     For i = 0 To List1.ListCount - 1
       DoEvents
       prgBar1.Value = i
       DoEvents
       lblPersen.Caption = _
          Format(CInt((((List1.ListCount - 1) - i) _
                / (List1.ListCount - 1)) * 100), "###") & "%"
       DoEvents
       Print #fileNo, List1.List(i)
     Next i
  Close #fileNo
  prgBar1.Value = 0
  prgBar1.Visible = False
  lblPersen.Caption = ""
  List1.Enabled = True
  Exit Sub
Pesan:
   MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then
     cmdDelete_Click
  End If
End Sub

Private Sub cmdFindFirst_Click()
ProgramActivation
Dim i As Integer
Dim ada As Integer
  bCancel = False
  strKriteria = txtFind.Text
  If strKriteria = "" Then
     cmdCancel.Enabled = False
     MsgBox "Type your word to find in this log file!", _
            vbExclamation, "Find"
     txtFind.SetFocus
     Exit Sub
  End If
  cmdCancel.Enabled = True
  If List1.ListCount < 1 Then Exit Sub
  List1.Selected(0) = True
  For i = 0 To List1.ListCount - 1
     DoEvents
     If bCancel = True Then Exit Sub
     List1.Selected(i) = True
     DoEvents
     ada = InStr(1, UCase(List1.Text), UCase(strKriteria))
     DoEvents
     If ada > 0 Then
        FoundPos = i
        cmdCancel.Enabled = False
        cmdFindNext.Enabled = True
        Exit Sub
     End If
     DoEvents
  Next i
  FoundPos = 0
  MsgBox "'" & strKriteria & "' not found!", _
         vbExclamation, "Finished searching"
End Sub

Private Sub cmdFindNext_Click()
ProgramActivation
Dim i As Integer
Dim ada As Integer
  If List1.ListCount < 1 Then Exit Sub
  bCancel = False
  cmdCancel.Enabled = True
  'This is necessary because if user change criteria
  'whenever still click find next button.
  strKriteria = txtFind.Text
  If FoundPos = -1 Or strKriteria = "" Then
     cmdFindFirst_Click
     Exit Sub
  End If
  For i = FoundPos + 1 To List1.ListCount - 1
     DoEvents
     If bCancel = True Then Exit Sub
     DoEvents
     List1.Selected(i) = True
     DoEvents
     ada = InStr(1, UCase(List1.Text), UCase(strKriteria))
     DoEvents
     FoundPos = i
     If ada > 0 Then
        cmdCancel.Enabled = False
        Exit Sub
     End If
     DoEvents
  Next i
  DoEvents
  cmdFindFirst.Enabled = True
  cmdFindNext.Enabled = False
  cmdCancel.Enabled = False
  MsgBox "'" & strKriteria & "' not found!", _
         vbExclamation, "Finished searching"
End Sub

Private Sub cmdFirst_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub cmdLast_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub cmdNext_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub cmdPrevious_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub Form_Load()
    Dim sNamaFile As String
    Dim NextLine As String
    Dim FileNum As Integer
    Dim i As Integer
    On Error GoTo Pesan
    List1.Clear
    FinishWaiting
    Exit Sub
Pesan:
   MsgBox Err.Number & "->" & Err.Description, _
          vbCritical, "Error, euy..."
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Set frmLogFile = Nothing
End Sub

Private Sub Form_Resize()
   ProgramActivation
   List1.Move 0, 2100, Me.ScaleWidth, Me.ScaleHeight - 2100
End Sub

Private Sub List1_Click()
   ProgramActivation
   Dim fileNo As Integer
   Dim sFileName As String
   Dim tmp As String
   Dim pos As Integer
   Dim i As Integer
   Dim sPassword As String
   Dim sDesc As String
   Dim sLocation As String
   Dim sExpiry As String
   
   On Error GoTo Pesan

     'If tmp = "" Then Exit Sub
     tmp = List1.Text
     'find the first comma
     pos = InStr(tmp, ",")
     'extract the string up to the comma
     sPassword = Left$(tmp, pos - 1)
     'shorten the string by removing the item
    'ready to find the next comma
     tmp = Mid$(tmp, pos + 1, Len(tmp))
     'do it again
     pos = InStr(tmp, ",")
     sDesc = Left$(tmp, pos - 1)
     tmp = Mid$(tmp, pos + 1, Len(tmp))
     'do it again
     pos = InStr(tmp, ",")
     sLocation = Left$(tmp, pos - 1)
     tmp = Mid$(tmp, pos + 1, Len(tmp))
     'the remainder is the expiry
     sExpiry = tmp
     'display the retrieved values
     txtUserID.Text = sPassword
     txtDescription.Text = sDesc
     txtDate.Text = sLocation
     txtTime.Text = sExpiry
     CurrentRec = List1.ListIndex
     lblStatus.Caption = "Record # " & CurrentRec + 1 & " of " & List1.ListCount & "."
     IsValidLogFile True
     Exit Sub
Pesan:
     Select Case Err.Number
            Case 5
                 MsgBox "File you opened is not log file for this program!", _
                        vbExclamation, "Not Log File"
                 List1.Clear
                 IsValidLogFile False
                 lblStatus.Caption = ""
                 lblFileName.Caption = ""
            Case Else
                 MsgBox Err.Number & " - " & _
                        Err.Description, vbCritical, "Error"
                 List1.Clear
                 lblStatus.Caption = ""
     End Select
End Sub

Private Sub IsValidLogFile(bVal As Boolean)
  txtUserID.Enabled = bVal
  txtDescription.Enabled = bVal
  txtDate.Enabled = bVal
  txtTime.Enabled = bVal
  cmdDelete.Enabled = bVal
  cmdClose.Enabled = bVal
  cmdFindFirst.Enabled = bVal
  cmdFindNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdNext.Enabled = bVal
  cmdPrevious.Enabled = bVal
  cmdLast.Enabled = bVal
  txtFind.Enabled = bVal
End Sub

Private Sub cmdFirst_Click()
On Error Resume Next
  ProgramActivation
  lblStatus.Caption = "First record"
  List1.Selected(0) = True
End Sub

Private Sub cmdLast_Click()
On Error Resume Next
  ProgramActivation
  lblStatus.Caption = "Last record"
  List1.Selected(List1.ListCount - 1) = True
End Sub

Private Sub cmdNext_Click()
On Error Resume Next
  ProgramActivation
  If List1.Selected(List1.ListCount - 1) = True Then
     MsgBox "This is the last record!", _
            vbInformation, "Last Record"
  End If
  If CurrentRec >= 0 And _
     CurrentRec < List1.ListCount Then
     List1.Selected(CurrentRec + 1) = True
  End If
  CurrentRec = List1.ListIndex
  lblStatus.Caption = "Record number " & CurrentRec + 1 & " of " & List1.ListCount & "."
End Sub

Private Sub cmdPrevious_Click()
  ProgramActivation
  If List1.Selected(0) = True Then
     MsgBox "This is the first record!", _
            vbInformation, "First Record"
  End If
  If CurrentRec > 0 And _
     CurrentRec < List1.ListCount Then
     List1.Selected(CurrentRec - 1) = True
  End If
  CurrentRec = List1.ListIndex
  lblStatus.Caption = "Record number " & CurrentRec + 1 & " of " & List1.ListCount & "."
End Sub
    
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub txtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
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

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub
    
Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
  ProgramActivation
  If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub
