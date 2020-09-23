VERSION 5.00
Begin VB.UserControl HyperLink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   FontTransparent =   0   'False
   MouseIcon       =   "Link.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1140
   ScaleWidth      =   2625
   ToolboxBitmap   =   "Link.ctx":0152
   Begin VB.Image imgNoLink 
      Height          =   480
      Left            =   600
      Picture         =   "Link.ctx":0464
      Top             =   360
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Left            =   0
      Picture         =   "Link.ctx":05B6
      Top             =   360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Height          =   135
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "HyperLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' API Declarations
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

' Defauly property constants
Private Const m_def_AutoSize As Boolean = True

' Enum declaration for windowstyle
Public Enum APIWindowStyleConstants
    SW_HIDE = 0
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
    SW_RESTORE = 9
End Enum

' Property Variables
Private m_Caption               As String       ' Holds the caption to be displayed over it
Private m_Target                As String       ' Holds the target location of the link
Private m_LinkAvailable         As Boolean      ' Link is available or not ??
Private m_AutoSize              As Boolean      ' Whether or not the control will automatically decide its size as per the caption set
Private m_Font                  As StdFont
Private m_BackColor             As OLE_COLOR
Private m_ForeColor             As OLE_COLOR
Private m_HoverForeColor        As OLE_COLOR
Private m_UnderlineOnHover      As Boolean
Private m_UseHoverForeColor     As Boolean

' Other internal private variables
Private AutoChange              As Boolean      ' Flag to prevent unnecessary iterations
Private MouseEntered            As Boolean      ' Internal flag to know that mouse has entered the control area
Private isMouseDown             As Boolean

' Event declarations
Event Click()
Attribute Click.VB_Description = "Event fired when user click over the control with mouse or presses Space or Enter when the control has focus."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event RightClick(Shift As Integer)
Attribute RightClick.VB_Description = "Event fired when user clicks the control with right mouse button."
Event MouseHover()
Attribute MouseHover.VB_Description = "Event fired when mouse pointer enters the control area."
Event MouseLeave()
Attribute MouseLeave.VB_Description = "Event fired when mouse pointer leaves the control area."
Event KeyPress(KeyAscii As Integer)


''''''''''''''''''''''''
'      Properties      '
''''''''''''''''''''''''

' Property set for "Caption"
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Text to be displayed on the link control."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
End Property
' End property set for "Caption"

' Property set for "Target"
Public Property Get Target() As String
Attribute Target.VB_Description = "Specifies the target location of the HyperLink, this property must be initialized before calling OpenTarget method."
Attribute Target.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Target.VB_UserMemId = 0
    Target = m_Target
End Property

Public Property Let Target(New_Target As String)
    m_Target = New_Target
    PropertyChanged "Target"
End Property
' End property set for "Target"

' Property set for "AutoSize"
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Automatically sizes the control as per the caption."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Refresh
End Property
' End property set for "AutoSize"

' Property set for "LinkAvailable"
Public Property Get LinkAvailable() As Boolean
Attribute LinkAvailable.VB_Description = "Returns/Sets the flag whether the link is available or not."
Attribute LinkAvailable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    LinkAvailable = m_LinkAvailable
End Property

Public Property Let LinkAvailable(New_LinkAvailable As Boolean)
    m_LinkAvailable = New_LinkAvailable
    PropertyChanged "LinkAvailable"
    Refresh
End Property
' End property set for "LinkAvailable"

' Property set for "Font"
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/Sets the font to be used to draw the control."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    Set m_Font = New_Font
    PropertyChanged "Font"
    Refresh
End Property
' End property set for "Font"

' Property set for "BackColor"
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Background color of the control."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property
' End property set for "BackColor"

' Property set for "ForeColor"
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Color to draw text of the control."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property
' End property set for "ForeColor"

' Property set for "HoverForeColor"
Public Property Get HoverForeColor() As OLE_COLOR
Attribute HoverForeColor.VB_Description = "Font color to be used when the mouse is on the control."
Attribute HoverForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoverForeColor = m_HoverForeColor
End Property

Public Property Let HoverForeColor(New_HoverForeColor As OLE_COLOR)
    m_HoverForeColor = New_HoverForeColor
    PropertyChanged "HoverForeColor"
End Property
' End property set for "HoverForeColor"

' Property set for "UseHoverForeColor"
Public Property Get UseHoverForeColor() As Boolean
Attribute UseHoverForeColor.VB_Description = "Returns/Sets whether or not the control is redrawn with HoverForeColor when mouse is on the control."
Attribute UseHoverForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UseHoverForeColor = m_UseHoverForeColor
End Property

Public Property Let UseHoverForeColor(New_UseHoverForeColor As Boolean)
    m_UseHoverForeColor = New_UseHoverForeColor
    PropertyChanged "UseHoverForeColor"
End Property
' End property set for "UseHoverForeColor"

' Property set for "UnderlineOnHover"
Public Property Get UnderlineOnHover() As Boolean
Attribute UnderlineOnHover.VB_Description = "Underlines the caption text when mouse is on the control."
Attribute UnderlineOnHover.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UnderlineOnHover = m_UnderlineOnHover
End Property

Public Property Let UnderlineOnHover(New_UnderlineOnHover As Boolean)
    m_UnderlineOnHover = New_UnderlineOnHover
    PropertyChanged "UnderlineOnHover"
End Property
' End property set for "UnderlineOnHover"

''''''''''''''''''''''''''''''''''''''
'      To Read-Write Properties      '
''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'Debug.Print "ReadProperties"
    AutoChange = True

    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_Target = PropBag.ReadProperty("Target", "")
    m_AutoSize = PropBag.ReadProperty("AutoSize", True)
    m_LinkAvailable = PropBag.ReadProperty("LinkAvailable", True)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    m_UseHoverForeColor = PropBag.ReadProperty("UseHoverForeColor", True)
    m_HoverForeColor = PropBag.ReadProperty("HoverForeColor", vbBlue)
    m_UnderlineOnHover = PropBag.ReadProperty("UnderlineOnHover", True)
    
    AutoChange = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'Debug.Print "WriteProperties"
    PropBag.WriteProperty "Caption", m_Caption
    PropBag.WriteProperty "Target", m_Target
    PropBag.WriteProperty "AutoSize", m_AutoSize
    PropBag.WriteProperty "LinkAvailable", m_LinkAvailable
    PropBag.WriteProperty "Font", m_Font
    PropBag.WriteProperty "BackColor", m_BackColor
    PropBag.WriteProperty "ForeColor", m_ForeColor
    PropBag.WriteProperty "HoverForeColor", m_HoverForeColor
    PropBag.WriteProperty "UseHoverForeColor", m_UseHoverForeColor
    PropBag.WriteProperty "UnderlineOnHover", m_UnderlineOnHover
End Sub
' End Read-Write Properties

Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
    m_Target = ""
    m_AutoSize = m_def_AutoSize
    m_LinkAvailable = True
    m_BackColor = Ambient.BackColor
    m_ForeColor = Ambient.ForeColor
    m_HoverForeColor = vbBlue
    m_UseHoverForeColor = True
    m_UnderlineOnHover = True
    m_UseHoverForeColor = True
    Set m_Font = Ambient.Font
End Sub

Private Sub UserControl_Initialize()
AutoChange = True   ' Until the Show event is fired, this flag will prevent unncessary redrawing of the control
End Sub

Private Sub UserControl_Show()
'Debug.Print "Show"
    AutoChange = False
    Refresh
End Sub

Private Sub UserControl_EnterFocus()
    Shape1.Visible = True
End Sub

Private Sub UserControl_ExitFocus()
    Shape1.Visible = False
End Sub

Private Sub UserControl_Resize()
    If AutoChange Then Exit Sub
'Debug.Print "Resize"
    Refresh
End Sub

' Events
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture ' Release the capture on the mouse so that VB can temporarily take control and resolve any possible mouse ownership problems
               ' Courtesy: Stephen Kent the author of GradientButton (www.planetsourcecode.com)
               ' (Without this statement, it is impossible to make a perfect hover event)
If Button = vbRightButton Then
    MouseEntered = False
    RaiseEvent RightClick(Shift)
Else
    isMouseDown = True
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isMouseDown Then
    isMouseDown = False
    MouseEntered = False
    RaiseEvent Click
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not MouseEntered Then
    MouseEntered = True
    SetCapture UserControl.hWnd
    If m_UseHoverForeColor Or m_UnderlineOnHover Then Refresh
    RaiseEvent MouseHover
Else
    If (X < 0) Or (Y < 0) Or (X > ScaleWidth) Or (Y > ScaleHeight) Then
        MouseEntered = False
        ReleaseCapture
        If m_UseHoverForeColor Or m_UnderlineOnHover Then Refresh
        RaiseEvent MouseLeave
    End If
End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
'''''''''''''''''''''''''''''''''''''
'      Other supported methods      '
'''''''''''''''''''''''''''''''''''''
Public Sub Refresh()
Attribute Refresh.VB_Description = "Redraws the control."
Attribute Refresh.VB_UserMemId = -550
'Debug.Print "Refresh"
    
    Set UserControl.Font = m_Font
    UserControl.BackColor = m_BackColor
    If MouseEntered And m_UseHoverForeColor Then
        UserControl.ForeColor = m_HoverForeColor
    Else
        UserControl.ForeColor = m_ForeColor
    End If
    If m_UnderlineOnHover Then
        UserControl.Font.Underline = MouseEntered Or Font.Underline
    End If
    If m_AutoSize Then
        AutoChange = True
        UserControl.Width = UserControl.TextWidth(m_Caption) + 120
        If UserControl.Width < 45 Then UserControl.Width = 45
        UserControl.Height = UserControl.TextHeight(m_Caption) + 60
        If UserControl.Height < 15 Then UserControl.Height = 15
        AutoChange = False
    End If
    
    UserControl.Cls
    UserControl.CurrentX = (UserControl.ScaleWidth - UserControl.TextWidth(m_Caption)) / 2
    UserControl.CurrentY = (UserControl.ScaleHeight - UserControl.TextHeight(m_Caption)) / 2
    UserControl.Print m_Caption
    
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    If m_LinkAvailable = True Then
        Set UserControl.MouseIcon = imgLink.Picture
    Else
        Set UserControl.MouseIcon = imgNoLink.Picture
    End If
End Sub

Public Function OpenTarget(Optional OwnerHWnd As Long = 0, Optional WindowStyle As APIWindowStyleConstants = SW_NORMAL, Optional StartupDirectory As String = "") As Long
Attribute OpenTarget.VB_Description = "Opens target location specified in Target property, will not work if LinkAvailable is set to False."
    If m_LinkAvailable Then
        If m_Target <> "" Then
            OpenTarget = OpenFile(m_Target, OwnerHWnd, StartupDirectory, WindowStyle)
        Else
            Raise_NoTargetError
        End If
    Else
        Raise_NoLinkError
    End If
End Function

Public Sub About()
Attribute About.VB_Description = "Shows the control developer's information and copyright licencing etc."
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

'''''''''''''''''''''''''''''''''''
'      Private subprocedures      '
'''''''''''''''''''''''''''''''''''
Private Sub Raise_NoTargetError()
    Err.Raise vbObjectError + 513, , "Target property must be initialized before calling OpenTarget method"
End Sub

Private Sub Raise_NoLinkError()
    Err.Raise vbObjectError + 514, , "Link not available (LinkAvaible=False)"
End Sub
