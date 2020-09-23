VERSION 5.00
Object = "{E0EC34B1-F558-11D5-B104-C6A7F59A432C}#2.0#0"; "HYPERLINKCONTROL.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HyperLink Tester"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ControlBox      =   0   'False
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Appearance"
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Tag             =   "Changes foreground color on mouse hover"
      Top             =   600
      Width           =   3255
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Tag             =   "This link uses default colors and hover sttings"
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         Caption         =   "Default Settings"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   16711680
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Tag             =   "Uses different colors"
         Top             =   840
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   450
         Caption         =   "Different Colors"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12640511
         ForeColor       =   16711680
         HoverForeColor  =   255
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Tag             =   "Different colors and Arial Font at 10pt."
         Top             =   1320
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   529
         Caption         =   "Different Colors & Fonts"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632319
         ForeColor       =   12582912
         HoverForeColor  =   255
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Tag             =   "Underline caption on mouse hover"
         Top             =   1800
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   450
         Caption         =   "Underline ON"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   255
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   8
         Tag             =   "Do not underline caption on mouse hover"
         Top             =   1800
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         Caption         =   "Underline OFF"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   255
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   0   'False
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Tag             =   "Changes foreground color on mouse hover"
         Top             =   2160
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   450
         Caption         =   "Hover Colors ON"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   255
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   6
         Left            =   945
         TabIndex        =   10
         Tag             =   "Does not changes foreground color on mouse hover"
         Top             =   2400
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   450
         Caption         =   "Hover Colors OFF"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   255
         UseHoverForeColor=   0   'False
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   11
         Tag             =   "LinkAvailable is set to False, notice the mouse pointer"
         Top             =   2760
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         Caption         =   "Unavailable Link"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483631
         HoverForeColor  =   255
         UseHoverForeColor=   0   'False
         UnderlineOnHover=   0   'False
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Tag             =   "No hover color or underline"
         Top             =   3120
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   450
         Caption         =   "All Special Properties Off"
         Target          =   ""
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   255
         UseHoverForeColor=   0   'False
         UnderlineOnHover=   0   'False
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Target"
      Height          =   3615
      Left            =   3840
      TabIndex        =   14
      Tag             =   "Changes foreground color on mouse hover"
      Top             =   600
      Width           =   3255
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTest.frx":0742
         Left            =   1440
         List            =   "frmTest.frx":0752
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3240
         Width           =   1695
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   15
         Tag             =   "Starts notepad normally"
         Top             =   600
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   450
         Caption         =   "Click Here to start Notepad"
         Target          =   "notepad.exe"
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   16711680
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   16
         Tag             =   "View control readme."
         Top             =   1320
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   450
         Caption         =   "HyperLink Control Readme"
         Target          =   "readme.htm"
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   16711680
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   21
         Tag             =   "Opens this link in your default web browser (Needs IE4 update or Windows 98 or higher)"
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Caption         =   "www.Yahoo.com"
         Target          =   "http://www.yahoo.com"
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   16711680
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin HyperLinkControlProject.HyperLink lnkDemo 
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   22
         Tag             =   "Opens a new message in your default mail client (Needs IE4 update or any other mail client installed)"
         Top             =   2760
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   450
         Caption         =   "Click Here to mail the author now"
         Target          =   "mailto:Shekhar_Extreme@yahoo.com"
         AutoSize        =   -1  'True
         LinkAvailable   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         HoverForeColor  =   16711680
         UseHoverForeColor=   -1  'True
         UnderlineOnHover=   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WindowStyle:"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   3300
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Address :"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web site :"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document :"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application :"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   870
      End
   End
   Begin HyperLinkControlProject.HyperLink lnkAbout 
      Height          =   330
      Left            =   1680
      TabIndex        =   13
      Top             =   120
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   582
      Caption         =   "HyperLink Control Demonstration"
      Target          =   ""
      AutoSize        =   -1  'True
      LinkAvailable   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   16711680
      UseHoverForeColor=   -1  'True
      UnderlineOnHover=   0   'False
   End
   Begin HyperLinkControlProject.HyperLink lnkMinimize 
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   450
      Caption         =   "Minimize"
      Target          =   ""
      AutoSize        =   -1  'True
      LinkAvailable   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   16711680
      UseHoverForeColor=   0   'False
      UnderlineOnHover=   0   'False
   End
   Begin HyperLinkControlProject.HyperLink lnkClose 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   450
      Caption         =   "Close"
      Target          =   ""
      AutoSize        =   -1  'True
      LinkAvailable   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      HoverForeColor  =   16711680
      UseHoverForeColor=   0   'False
      UnderlineOnHover=   0   'False
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WelCome"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   7065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   480
      Y1              =   31
      Y2              =   31
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   8
      X2              =   480
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub

Private Sub lnkAbout_Click()
lnkAbout.About
End Sub

Private Sub lnkAbout_MouseHover()
lblStatus = "See About box of the control"
End Sub

Private Sub lnkAbout_MouseLeave()
lblStatus = ""
End Sub

Private Sub lnkClose_Click()
Unload Me
End
End Sub

Private Sub lnkClose_MouseHover()
lblStatus = "Close Demo"
End Sub

Private Sub lnkClose_MouseLeave()
lblStatus = ""
End Sub

Private Sub lnkDemo_Click(Index As Integer)
Dim ws As APIWindowStyleConstants
If lnkDemo(Index).Target <> "" And lnkDemo(Index).LinkAvailable Then
    Select Case Combo1.ListIndex
    Case 0
        ws = SW_NORMAL
    Case 1
        ws = SW_MAXIMIZE
    Case 2
        ws = SW_MINIMIZE
    Case 3
        ws = SW_HIDE
    Case Else
        ws = SW_NORMAL
    End Select
    lnkDemo(Index).OpenTarget Me.hWnd, ws, App.Path
End If
End Sub

Private Sub lnkDemo_MouseHover(Index As Integer)
lblStatus = lnkDemo(Index).Tag
End Sub

Private Sub lnkDemo_MouseLeave(Index As Integer)
lblStatus = "Mouse leaves " & lnkDemo(Index).Caption
End Sub

Private Sub lnkDemo_RightClick(Index As Integer, Shift As Integer)
MsgBox "Right Click on " & lnkDemo(Index).Caption
End Sub

Private Sub lnkMinimize_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub lnkMinimize_MouseHover()
lblStatus = "Minimize this window"
End Sub

Private Sub lnkMinimize_MouseLeave()
lblStatus = ""
End Sub
