VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCloseAbout 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by Shekhar Somani"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4260
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HyperLink Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCloseAbout_Click()
Unload Me
Set frmAbout = Nothing
End Sub

Private Sub Form_Load()
Caption = "About " & App.Title
lblDesc = "by Shekhar Somani" & vbCrLf & "Shekhar_Extreme@yahoo.com -or-" & vbCrLf & "Shekhar_D_S@yahoo.com" & vbCrLf & vbCrLf & "If you find any bug or problem, or have any suggestions, comments, or improvements, then do mail me, I will really appriciate." & vbCrLf & vbCrLf & vbCrLf & "This control is Free for non-commercial use, but any changes in the source or distribution in public requires author's prior permission."
End Sub
