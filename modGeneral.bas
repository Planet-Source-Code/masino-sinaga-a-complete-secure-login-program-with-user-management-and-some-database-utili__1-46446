Attribute VB_Name = "modGeneral"
'File Name   : Module1.bas
'Description : Declare global variable, procedure, and
'              function
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Sunday, June 15, 2003
'Location    : Jakarta, INDONESIA
'-----------------------------------------------------

Public Type arrSetting
  LindungLayar As Integer
  PassLindungLayar As String
  MenitDelay As String
  IntervalMenit As String
End Type
Public gloSet As arrSetting

Public db As ADODB.Connection
Public m_UserID As String
Public m_Level As String
Public m_blnLogin As Boolean
Public m_blnCancel As Boolean
Public strPassword As String, Temp As String
Public rsUser As ADODB.Recordset
Public StaKonek As Boolean
Public LindungLayar As Byte
Public Awal As Date
Public Gerak As Boolean
Public Aksi As Boolean

Public Type tUser
  UserID As String
  password As String
End Type
Public tabUser() As tUser  'Array dari record tipe tUser

Public Type cUser
  User As String
End Type
Public cekUser() As cUser  'Array dari record tipe cUser

'This will open connection to database. We use database
'MSAccess97 password protected. The password is 'masino2002'
Public Sub OpenConnection()
On Error GoTo ErrMess
  Screen.MousePointer = vbHourglass
  Set db = New Connection
  'Always use cursorlocation in client side
  'because we use client database, not server database
  db.CursorLocation = adUseClient
  'First to remember is: we use PROVIDER MSDataShape
  'because we want to make a recordset with "master-detail"
  'style in order that we want to show all data in sub-
  'recordset (child) in DataGrid control below and display
  'the selected record (parent) in Textbox control above
  'the form. But this is not truly master-detail recordset,
  'we just use this for display a complete data in a form!
  'In this example, I use a Access database protected
  'password: masino2002
  db.Open "PROVIDER=MSDataShape;Data PROVIDER=" & _
          "Microsoft.Jet.OLEDB.4.0;Data Source=" _
          & App.Path & "\Data.mdb;Jet OLEDB:" & _
          "Database Password=masino2002;"
  StaKonek = True
  Screen.MousePointer = vbDefault
  Exit Sub
ErrMess:
  StaKonek = False
  Screen.MousePointer = vbDefault
End Sub

'This will open User table to get the user information
Public Sub OpenTableUser()
On Error GoTo MessErr
  If db Is Nothing Then OpenConnection
  Set rsUser = New ADODB.Recordset
  DoEvents
  'See below.. there are two recordset in a recordset.
  'The first sub recordset we called: "ParentCMD", and
  'the second sub recordset we called: "ChildCMD".
  'If we use PROVIDER MSDataShape, we also can use
  'a simple SELECT statement (not using SHAPE).
  rsUser.Open _
      "SHAPE {SELECT * FROM T_User " & _
      "Order by User_ID} AS ParentCMD APPEND " & _
      "({SELECT * FROM T_User " & vbCrLf & _
      "Order by User_ID } AS ChildCMD RELATE User_ID TO " & _
      "User_ID) AS ChildCMD", _
      db, adOpenStatic, adLockOptimistic
  DoEvents
  Exit Sub
MessErr:
  Select Case Err.Number
         Case 3709
              MsgBox "Failed to connect to database!", _
                     vbExclamation, "Failed"
              If FormLoadedByName("frmLogin") = True Then
                 Unload frmLogin
              End If
         Case Else
              MsgBox Err.Number & " - " & _
                     Err.Description, _
                     vbCritical, "Error"
  End Select
End Sub

'This will get number of record in table T_User
Public Function NumOfUser() As Integer
   NumOfUser = rsUser.RecordCount
End Function

'This will encrypt/decrypt password field in database
Public Sub EncryptDecrypt()
Dim i As Integer
Dim intLocation As Integer
Dim Code As String
Code = "1234567890" 'This is key for encrypting/decrypting
  Temp$ = ""
  For i% = 1 To Len(strPassword)
      intLocation% = (i% Mod Len(Code)) + 1
      'Use XOR logic combination for encrypting/decrypting
      Temp$ = Temp$ + Chr$(Asc(Mid$(strPassword, i%, 1)) Xor _
              Asc(Mid$(Code, intLocation%, 1)))
  Next i%
End Sub

'This will check whether a form was loaded in memory or not.
Public Function FormLoadedByName(FormName As String) As Boolean
Dim i As Integer, fnamelc As String
fnamelc = LCase$(FormName)
FormLoadedByName = False
For i = 0 To Forms.Count - 1
If LCase$(Forms(i).Name) = fnamelc Then
  FormLoadedByName = True
  Exit Function
End If
Next
End Function

'This will save what user do in this program to log file
Public Sub SaveActivityToLogFile(Aktivitas As String)
Dim sFName As String
Dim nFile As Integer
Dim strTime As String
Dim strDate As String
  If m_UserID = "" Then m_UserID = "Unknown user"
  'In this example, I use Indonesia date setting
  strDate = Format(Date, "dd/mm/yyyy")
  With frmMain
    'And so is the time..   .
    strTime = Format(Time, "hh:mm:ss")
  End With
  'Log filename is in the same directory with app is
  sFName = App.Path & "\LogLogin3.txt"
  'We use FreeFile from OS Windows in order that to prevent
  'clash with another text file (if there is)
   nFile = FreeFile
  'Save to log file with Append mode (add new record to the
  'last line in file). We use comma separator to separate
  'one item (field) to another.
   Open sFName For Append As #nFile
     'Save it to file, now.....!
     Print #nFile, m_UserID & "," & Aktivitas & "," & strDate & "," & strTime
   Close #nFile  'Don't forget to close the logfile
End Sub

'This will save what user do in this program to log
'table in database...
Public Sub SaveActivityToLogDB(Aktivitas As String)
Dim sFName As String
Dim nFile As Integer
Dim strTime As String
Dim strDate As String
  If m_UserID = "" Then m_UserID = "Unknown user"
  'In this example, I use Indonesia date setting
  strDate = Format(Date, "dd/mm/yyyy")
  With frmMain
    'And so is the time..   .
    strTime = Format(Time, "hh:mm:ss")
  End With
  If db Is Nothing Then OpenConnection
  db.Execute "INSERT INTO T_Log " & _
       "VALUES('" & Replace(m_UserID, "'", "''") & "'," & _
       "'" & Replace(Aktivitas, "'", "''") & "'," & _
       "'" & strDate & "'," & _
       "'" & strTime & "')"
End Sub

'Close all forms in this project
Public Sub CloseAllForms()
Dim Form As Form
   For Each Form In Forms
       Unload Form
       Set Form = Nothing
   Next Form
   End
End Sub

'Show a message in statusbar on frmMain
Public Sub Message(strMessage As String)
  frmMain.StatusBar1.Panels(1).Text = strMessage
End Sub

'Begin to wait a process...
Public Sub StartWaiting(strMess As String)
  Screen.MousePointer = vbHourglass
  DoEvents
  With frmWait
    DoEvents
    .lblProses.Move 120, .prgBar1.Top
    .lblProses = strMess
    DoEvents
    .prgBar1.Visible = False
    .lblAngka.Visible = False
    .Show , frmMain
  End With
End Sub

'Mark end to a process, unload the form...
Public Sub FinishWaiting()
  Unload frmWait
  Set frmWait = Nothing
  Screen.MousePointer = vbDefault
End Sub

'This will close all forms except the form that you
'mention in parameter named "FormToStay" below
Public Sub UnloadAllExceptOne(FormToStay As String)
Dim oFrm As Form
For Each oFrm In Forms
    If oFrm.Name <> FormToStay And Not _
       (TypeOf oFrm Is MDIForm) Then
       Unload oFrm
       Set oFrm = Nothing
    End If
Next
End Sub


'This procedure will adjust DataGrids column width
'based on longest field in underlying source
Public Sub AdjustDataGridColumns _
           (DG As DataGrid, _
           adoData As ADODB.Recordset, _
           intRecord As Integer, _
           intField As Integer, _
           Optional AccForHeaders As Boolean)

'DG = DataGrid
'adoData = Adodc control
'intRecord = Number of record
'intField = Number of field
'AccForHeaders = True or False
    Screen.MousePointer = vbHourglass
    Dim row As Long, col As Long
    Dim width As Single, maxWidth As Single
    Dim saveFont As StdFont, saveScaleMode As Integer
    Dim cellText As String, i As Integer
    
    'If number of records = 0 then exit from the sub
    If intRecord = 0 Then Exit Sub
    'Save the form's font for DataGrid's font
    'We need this for form's TextWidth method
    Set saveFont = DG.Parent.Font
    Set DG.Parent.Font = DG.Font
    'Adjust ScaleMode to vbTwips for the form (parent).
    saveScaleMode = DG.Parent.ScaleMode
    DG.Parent.ScaleMode = vbTwips
    'Always from first record...
    adoData.MoveFirst
    maxWidth = 0
    frmWait.Show , frmMain
    maks = intField * intRecord
    With frmWait
      .prgBar1.Visible = True
      .cmdCancel.Visible = True
      .prgBar1.Max = maks
      .lblProses.Caption = _
        "Wait, adjusting datagrid's columns width..."
    End With
    'We begin from the first column until the last column
    For col = 0 To intField - 1
        If m_blnCancel = True Then Exit For
        DoEvents
        frmWait.lblField.Caption = _
           "Column: " & DG.Columns(col).DataField
        DoEvents
        adoData.MoveFirst
        'Optional param, if true, set maxWidth to
        'width of DG.Parent
        If AccForHeaders Then
            maxWidth = DG.Parent.TextWidth(DG.Columns(col).Text) + 200
        End If
        'Repeat from first record again after we have
        'finished process the last record in
        'former column...
        adoData.MoveFirst
        For row = 0 To intRecord - 1
            If m_blnCancel = True Then Exit For
            DoEvents
            'Get the text from the DataGrid's cell
            If intField = 1 Then
            Else  'If number of field more than one
                cellText = DG.Columns(col).Text
            End If
            'Fix the border...
            'Not for "multiple-line text"...
            width = DG.Parent.TextWidth(cellText) + 200
            'Update the maximum width if we found
            'the wider string...
            If width > maxWidth Then
               maxWidth = width
               DG.Columns(col).width = maxWidth
            End If
            'Process next record...
            adoData.MoveNext
            i = i + 1
            DoEvents
            frmWait.lblAngka.Caption = _
              "Finished " & Format((i / maks) * 100, "0") & "%"
            DoEvents
            frmWait.prgBar1.Value = i
            DoEvents
            
        Next row
        'Change the column width...
        DG.Columns(col).width = maxWidth 'kolom terakhir!
    Next col
    'Change the DataGrid's parent property
    Set DG.Parent.Font = saveFont
    DG.Parent.ScaleMode = saveScaleMode
    If m_blnCancel = True Then
       Screen.MousePointer = vbDefault
       Unload frmWait
       Set frmWait = Nothing
       Exit Sub
    End If
    'If finished, then move pointer to first record again
    adoData.MoveFirst
    Unload frmWait
    Set frmWait = Nothing
    Screen.MousePointer = vbDefault
End Sub  'End of AdjustDataGridColumns

'Untuk mengupdate status pergerakan mouse/keypress di program
Public Sub ProgramActivation()
   Awal = Time 'Jika ada pergerakan, set waktu saat itu
   Aksi = True 'Update status...
End Sub

'Disable Ctrl-Alt-Del
Public Sub DisableCtrlAltDelete(bDisabled As Boolean)
Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

'Hide program from tasklist
Public Sub HideApp(Hide As Boolean)
    Dim ProcessID As Long
    ProcessID = GetCurrentProcessId()
    If Hide Then
        retval = RegisterServiceProcess(ProcessID, RSP_SIMPLE_SERVICE)
    Else
        retval = RegisterServiceProcess(ProcessID, RSP_UNREGISTER_SERVICE)
    End If
End Sub

'Untuk mengambil nilai setting dari frmSetting
Public Sub AmbilSetting()
  With frmSetting
    gloSet.LindungLayar = .ScreenSaver.Value
    gloSet.PassLindungLayar = .ScreenSaverPassword.Value
    gloSet.IntervalMenit = .MinuteStandBy.Text
  End With
End Sub


