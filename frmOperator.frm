VERSION 5.00
Begin VB.Form frmOperator 
   Caption         =   "Operator"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmOperator.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Back"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"frmOperator.frx":08CA
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   4080
      Width           =   8415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This is Demo Menu for Operator from LOGIN Program"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Width           =   8175
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'File Name   : frmOperator.frm
'Description : Just demo to display Operator menu.
'Author      : Masino Sinaga (masino_sinaga@yahoo.com)
'Web Site    : http://www30.brinkster.com/masinosinaga/
'              http://www.geocities.com/masino_sinaga/
'Date        : Saturday, June 13, 2003
'Location    : Jakarta, INDONESIA
'-----------------------------------------------------------

Private Sub cmdOK_Click()
  ProgramActivation
  Unload Me
End Sub

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
End Sub

Private Sub Form_Load()
  ProgramActivation
  Call SaveActivityToLogDB("Start access Operator menu.")
  Call Message("This menu is for user who has level: 'Operator' and 'Admin' .")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Call SaveActivityToLogDB("Finish access Operator menu.")
End Sub
