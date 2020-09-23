VERSION 5.00
Begin VB.UserControl EyeDropperContainer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "EyeDropperContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
' Module    : EyeDropperContainer
' DateTime  : 08/02/2005
' Author    : Gary Noble   ©2005 Telecom Direct Limited
' Purpose   : Simple Container Control Create Specificly For the Eyedropper Control
' Assumes   :
' Notes     :
' Revision  : 1.0
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  08/02/2005
'//---------------------------------------------------------------------------------------
Option Explicit

Public Event ThemeChanged(ByVal sThemeName As String)

Implements iSubclass

Private m_Subclass As IAPP_Subclass
Private mDC As IAPP_MemDC
Private m_HeaderRect As RECT
Private m_InsideRect As RECT

'-- Draw Colors
Private m_lngColorHeaderColorOne As OLE_COLOR
Private m_lngColorHeaderColorTwo As OLE_COLOR
Private m_lngColorBorder As OLE_COLOR

Private m_lngHeaderHeight As Long

Const m_def_Caption = ""

Dim m_Caption As String
Dim m_HeaderIcon As Picture

Const m_def_ForeColor = vbWindowText

Dim m_ForeColor As OLE_COLOR
'Default Property Values:
Const m_def_UseCustomColors = False
Const m_def_CustomColor = vbButtonFace
'Property Variables:
Dim m_UseCustomColors As Boolean
Dim m_CustomColor As OLE_COLOR




'//---------------------------------------------------------------------------------------
'//--Procedure : pvGetGradientColors
'//--Type      : Sub
'//--DateTime  : 04/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Sets The Theme colours
'//--Returns   :
'//--Notes     : Supports XP And Non XP Themes
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvGetGradientColors()


    m_lngColorHeaderColorOne = 1
    m_lngColorHeaderColorTwo = 1
    
    mGlobal.GetThemeName UserControl.hwnd
    
    If AppThemed And Not UseCustomColors Then
        Select Case m_sCurrentSystemThemename
            Case "HomeStead"
                m_lngColorBorder = RGB(100, 144, 88)
                m_lngColorHeaderColorOne = RGB(228, 235, 200)
                m_lngColorHeaderColorTwo = RGB(175, 194, 142)
            Case "NormalColor"
                m_lngColorBorder = RGB(0, 45, 150)
                m_lngColorHeaderColorOne = RGB(197, 221, 250)
                m_lngColorHeaderColorTwo = RGB(128, 167, 225)
            Case "Metallic"
                m_lngColorBorder = RGB(119, 118, 151)
                m_lngColorHeaderColorOne = RGB(219, 220, 232)
                m_lngColorHeaderColorTwo = RGB(149, 147, 177)
            Case Else
                m_lngColorHeaderColorOne = BlendColor(vbButtonFace, vbWhite, 120)
                m_lngColorHeaderColorTwo = vbButtonFace
                m_lngColorBorder = TranslateColor(vbInactiveTitleBar)
        End Select
    ElseIf Not UseCustomColors Then
        m_lngColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
        m_lngColorHeaderColorOne = BlendColor(vbButtonFace, vbWhite, 120)
        m_lngColorHeaderColorTwo = vbButtonFace
    Else
        m_lngColorBorder = BlendColor(TranslateColor(m_CustomColor), vbBlack, 150)
        m_lngColorHeaderColorOne = BlendColor(TranslateColor(m_CustomColor), vbWhite, 120)
        m_lngColorHeaderColorTwo = BlendColor(TranslateColor(m_CustomColor), vbBlack, 200) 'BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite), 200)
       
    End If



End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    pvDrawControl
End Property

Public Property Get HeaderIcon() As Picture
    Set HeaderIcon = m_HeaderIcon
End Property

Public Property Set HeaderIcon(ByVal New_HeaderIcon As Picture)
    Set m_HeaderIcon = New_HeaderIcon
    PropertyChanged "HeaderIcon"
End Property

Private Sub UserControl_Initialize()
    
    Set mDC = New IAPP_MemDC
    
    Set m_Subclass = New IAPP_Subclass
    
    With m_Subclass
        .UnSubclass
        .Subclass UserControl.hwnd, Me
        .AddMsg WM_DISPLAYCHANGE, MSG_AFTER
        .AddMsg WM_SYSCOLORCHANGE, MSG_AFTER
        .AddMsg WM_THEMECHANGED, MSG_AFTER
    End With
    
    pvGetGradientColors
    pvDrawControl


End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    Set m_HeaderIcon = LoadPicture("")
    Set UserControl.Font = Ambient.Font
    m_ForeColor = m_def_ForeColor
        
    m_UseCustomColors = m_def_UseCustomColors
    m_CustomColor = m_def_CustomColor
    
    pvGetGradientColors
    pvDrawControl
    
    UserControl_Resize

End Sub

Private Sub UserControl_Paint()
    
    UserControl_Resize
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set m_HeaderIcon = PropBag.ReadProperty("HeaderIcon", Nothing)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
        Dim fDC As New IAPP_MemDC
        
        With fDC
            .Init ScaleWidth, ScaleHeight
            Set .Font = Font
            m_lngHeaderHeight = .TextHeight("^,") + 6
        End With
        
    Set fDC = Nothing

    UserControl_Resize
    UserControl.Refresh
    
    m_UseCustomColors = PropBag.ReadProperty("UseCustomColors", m_def_UseCustomColors)
    m_CustomColor = PropBag.ReadProperty("CustomColor", m_def_CustomColor)
    
    pvGetGradientColors
    pvDrawControl
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

    '-- Re Initialise The Memory DC
    mDC.Init ScaleWidth, ScaleHeight, hdc
    mDC.BackStyle = BS_NEWTRANSPARENT
        
    
    Dim CTL As Control
        
        '-- Resize Each Of the controls Within the Usercontrol
        For Each CTL In UserControl.ContainedControls
            CTL.Move 0, (m_lngHeaderHeight * Screen.TwipsPerPixelY) + 1, (ScaleWidth * Screen.TwipsPerPixelX), ScaleHeight * Screen.TwipsPerPixelY - ((m_lngHeaderHeight * Screen.TwipsPerPixelY) + 1)
        Next
        
    
    '-- Do IT
    pvDrawControl
    
On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()
'//-- Clean Up
    On Error Resume Next
    
    m_Subclass.UnSubclass
    
    Set m_Subclass = Nothing

On Error GoTo 0
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("HeaderIcon", m_HeaderIcon, Nothing)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("UseCustomColors", m_UseCustomColors, m_def_UseCustomColors)
    Call PropBag.WriteProperty("CustomColor", m_CustomColor, m_def_CustomColor)
End Sub



Private Sub iSubclass_After(lReturn As Long, _
                            ByVal hwnd As Long, _
                            ByVal uMsg As WinSubHook.eMsg, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long)

        Select Case uMsg
            Case WM_DISPLAYCHANGE
                pvGetGradientColors
                pvDrawControl
            Case WM_SYSCOLORCHANGE
                pvGetGradientColors
                RaiseEvent ThemeChanged(m_sCurrentSystemThemename)
                pvDrawControl
            Case WM_THEMECHANGED
                pvGetGradientColors
                RaiseEvent ThemeChanged(m_sCurrentSystemThemename)
                pvDrawControl
        End Select

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, _
                             lReturn As Long, _
                             hwnd As Long, _
                             uMsg As WinSubHook.eMsg, _
                             wParam As Long, _
                             lParam As Long)

'
End Sub



Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    
    Dim fDC As New IAPP_MemDC
        
        With fDC
            .Init ScaleWidth, ScaleHeight
            Set .Font = New_Font
            m_lngHeaderHeight = .TextHeight("^,") + 6
        End With
        
    Set fDC = Nothing
    
    PropertyChanged "Font"
    
    UserControl_Resize
    
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    pvDrawControl
End Property



'//---------------------------------------------------------------------------------------
' Procedure : pvDrawControl
' Type      : Sub
' DateTime  : 08/02/2005
' Author    : Gary Noble
' Purpose   : Draws The Actual Control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  08/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvDrawControl()
On Error Resume Next

    With mDC
        Set .Font = Font
        .ForeColor = Me.ForeColor
        .BackColor = Me.BackColor
        .Rectangle -2, 0, ScaleWidth + 2, m_lngHeaderHeight, Me.BackColor, , m_lngColorBorder
        .FillGradient -3, 1, ScaleWidth + 3, m_lngHeaderHeight - 1, m_lngColorHeaderColorOne, m_lngColorHeaderColorTwo, True
      
        If Len(RTrim(Me.Caption)) > 0 Then
            .DrawText Me.Caption, 4, (m_lngHeaderHeight \ 2) - ((m_lngHeaderHeight - 6) \ 2), ScaleWidth - 6, m_lngHeaderHeight, DT_LEFT
        End If
        
        .BitBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, , , vbSrcCopy
    
    End With
    
        
On Error GoTo 0
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get UseCustomColors() As Boolean
    UseCustomColors = m_UseCustomColors
End Property

Public Property Let UseCustomColors(ByVal New_UseCustomColors As Boolean)
    m_UseCustomColors = New_UseCustomColors
    PropertyChanged "UseCustomColors"
     
    pvGetGradientColors
    pvDrawControl
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbbuttonface
Public Property Get CustomColor() As OLE_COLOR
    CustomColor = m_CustomColor
End Property

Public Property Let CustomColor(ByVal New_CustomColor As OLE_COLOR)
    m_CustomColor = New_CustomColor
    PropertyChanged "CustomColor"
     
    pvGetGradientColors
    pvDrawControl
    
End Property

