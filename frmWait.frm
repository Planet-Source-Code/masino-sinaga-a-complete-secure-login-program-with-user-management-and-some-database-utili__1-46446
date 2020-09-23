VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1305
   ClientLeft      =   4170
   ClientTop       =   5790
   ClientWidth     =   4155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar prgBar1 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   500
      Visible         =   0   'False
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblField 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblAngka 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblProses 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   3975
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmWait.frm
'Description : Display message based on parameter sent
'              to this form by other form when display
'              this form to wait a process...
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 14, 2003
'Location    : Jakarta, INDONESIA
'--------------------------------------------------------------
Private Sub cmdCancel_Click()
  m_blnCancel = True
End Sub

Private Sub Form_Load()
  m_blnCancel = False
End Sub
