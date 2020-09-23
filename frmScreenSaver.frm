VERSION 5.00
Begin VB.Form frmScreenSaver 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6870
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmScreenSaver.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox fraAnimasi 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   2655
      TabIndex        =   1
      Top             =   480
      Width           =   2655
      Begin VB.Timer tmrQuartz 
         Interval        =   500
         Left            =   1080
         Top             =   600
      End
      Begin VB.Label lblDigital 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   15
         Top             =   300
         Width           =   1695
      End
      Begin VB.Line LineHour 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   1250
         X2              =   1700
         Y1              =   1250
         Y2              =   1020
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   220
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   250
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   60
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1570
         Width           =   375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   550
         TabIndex        =   10
         Top             =   1930
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   2050
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1580
         TabIndex        =   8
         Top             =   1930
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1580
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2130
         TabIndex        =   6
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2000
         TabIndex        =   5
         Top             =   580
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1590
         TabIndex        =   4
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1100
         TabIndex        =   3
         Top             =   80
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2400
         Left            =   50
         Shape           =   3  'Circle
         Top             =   60
         Width           =   2445
      End
      Begin VB.Line LineMinute 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         X1              =   1250
         X2              =   375
         Y1              =   1250
         Y2              =   780
      End
      Begin VB.Line LineSecond 
         BorderColor     =   &H000000FF&
         X1              =   1250
         X2              =   705
         Y1              =   1250
         Y2              =   2035
      End
      Begin VB.Label lblTanggal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   100
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   3240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   675
      Left            =   2970
      TabIndex        =   19
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   675
      Left            =   2970
      TabIndex        =   18
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN Program Ver 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Index           =   0
      Left            =   2980
      TabIndex        =   16
      Top             =   570
      Width           =   8175
   End
   Begin VB.Label lblTekan 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Press any key or click your mouse on this screen to exit..."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4035
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN Program Ver 1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1095
      Index           =   1
      Left            =   3015
      TabIndex        =   20
      Top             =   600
      Width           =   8160
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Your monitor screen is protected by:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3015
      TabIndex        =   17
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI = 3.14159
Dim sHari As String
Dim aHari

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  frmMain.Timer1.Enabled = True
  ProgramActivation
  Set frmScreenSaver = Nothing
End Sub

Private Sub Timer1_Timer()
   lblTekan = ""
End Sub

Private Sub Form_Load()
Dim res
   'I made this array to translate name of day to Indonesia
   'language. You can change this array elemen to your
   'country language. Start from Sunday...
   aHari = Array("Sunday", "Monday", "Tuesday", "Wednesday", _
                "Thursday", "Friday", "Saturday")
   sHari = aHari(Abs(Weekday(Date) - 1))
   frmMain.Caption = "LOGIN PROGRAM Versi 1.0"
   frmMain.Timer1.Enabled = False
   Timer1.Enabled = True
   DoEvents
   Label1.Caption = "" & sHari & ", " _
                   & Format(Date, "mmmm dd yyyy")
   Label2.Caption = Format(Time, "hh:mm:ss")
   DoEvents
   App.HelpFile = ""
   res = SetWindowPos(frmScreenSaver.hWnd, _
                      HWND_TOPMOST, 0, 0, 0, 0, _
                      flags)
End Sub

Private Sub Form_Click()
   If frmSetting.ScreenSaverPassword.Value = 1 Then
      frmPassword.Show 1
   Else
      Unload Me
      Set frmScreenSaver = Nothing
   End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If frmSetting.ScreenSaverPassword.Value = 1 Then
      frmPassword.Show 1
   Else
      Unload Me
      Set frmScreenSaver = Nothing
   End If
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblTekan.Caption = "Press any key or click your mouse on this screen to exit..."
End Sub

Private Sub tmrQuartz_Timer()
  Dim Hours As Single, Minutes As Single, Seconds As Single
  Dim TrueHours As Single
  sHari = aHari(Abs(Weekday(Date) - 1))
  Label1.Caption = "" & sHari & ", " _
                   & Format(Date, "mmmm dd yyyy")
  Label2.Caption = Format(Time, "hh:mm:ss")
  Hours = Hour(Time)
  Minutes = Minute(Time)
  Seconds = Second(Time)
  'I got this code (Analog Clock) from
  'M. Thaha Husain.
  'Thanks to Husain!
  TrueHours = Hours + Minutes / 60
  LineHour.X2 = 750 * Cos(PI / 180 * (30 * TrueHours - 90)) + LineHour.X1
  LineHour.Y2 = 750 * Sin(PI / 180 * (30 * TrueHours - 90)) + LineHour.Y1
  LineMinute.X2 = 1050 * Cos(PI / 180 * (6 * Minutes - 90)) + LineHour.X1
  LineMinute.Y2 = 1050 * Sin(PI / 180 * (6 * Minutes - 90)) + LineHour.Y1
  LineSecond.X2 = 1100 * Cos(PI / 180 * (6 * Seconds - 90)) + LineHour.X1
  LineSecond.Y2 = 1100 * Sin(PI / 180 * (6 * Seconds - 90)) + LineHour.Y1
End Sub
