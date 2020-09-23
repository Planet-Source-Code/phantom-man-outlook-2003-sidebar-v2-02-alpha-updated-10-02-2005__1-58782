VERSION 5.00
Begin VB.UserControl EyeDropper 
   Alignable       =   -1  'True
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   Begin VB.Menu mnuMain 
      Caption         =   "mnuMain"
   End
   Begin VB.Menu EDItem 
      Caption         =   ""
      Index           =   0
   End
End
Attribute VB_Name = "EyeDropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
'//--Control   : EyeDropper
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble  ©2004-2005 Telecom Direct Limited
'//--Purpose   : A Tab Control To Simulate The MS Outlook2003 sidebar control
'//--Assumes   : EDItems And EDItem
'//--Notes     : Compelete New Rebuild From Scratch
'//--Revision  : 2.0 Alpha
'//---------------------------------------------------------------------------------------
'
'   This Control Also Uses Code From Other Authors, All The Original Copyrights
'   Credits Can Be Found Where They Put Them.
'
'   Give Credit Where Credit Is Due.
'
'   Special Thanks To: Vlad Vissoultchev  - Memdc Drawing Class
'                      Paul Caton - Subclassing Code
'
'
' --------------------------------------------------------------------------------------
' History:
'           19/04/2004 - Initial Implementation (Gary Noble)
'           21/04/2004 - Added Add Remove item Functionality (Gary Noble)
'           22/04/2004 - Added Visible Property (Gary Noble)
'                      - This Allows You To Hide The Item But Not Delete It
'                      - It Is Then Placed In The Add Remove Menu Items
'           05/05/2004 - Fixed The Virtual Memory Low Error in WinXP (Gary Noble)
'                      - Added Right To Left Support (Gary Noble)
'           06/05/2004 - Added The Ability To Set Mouse Pointer To Hand (Gary Noble)
'                      - Added DrawToolbarItems RightToLeft Just To Make (Gary Noble)
'                        More Like The Original
'                      - Cleaned Up Code and Test Until I Was Blue In The Face!!! (Can't Say My Name Anymore)
'           11/05/2004 - Added Custom Properties Function
'                      - Added Header Colour
'                      - Update The Default Colours To Conincide With The Original Control
'           14/05/2004 - Added Visible Item SaveState when The User First Resize
'                        What This Does Is Return The Number of Visible Items To Its Original
'                        State Before Sizing.
'                      - Updated Caption Drawing - It Now Has It's Own Sub pDrawCaption
'                      - Updated The Paint Order So The Caption Does'nt Paint Under The Splitter
'                        And Over The Items. (Much More Pro!)
'                      - Updated The Drawing Of The Line to use Api Instead Of .line (x,y)-(x1,y1)
'                      - Removed memDC Drawing As It Wasn't Having Any Effect On The control
'
'           01/02/2008 - New Implementation
'                      - Complete New Rebuild (Completely Dependency Free)
' --------------------------------------------------------------------------------------
'
'  *********  If You Use This Control Please Give Credit  *********

'
'//----------------------------------------------------------------------------------------
'//--Copyright (c) 2004-2005 Gary Noble
'//-----------------------------------------------------------------------
'
'//--Redistribution and use in source and binary forms, with or
'//-- without modification, are permitted provided that the following
'//-- conditions are met:
'
'//-- 1. Redistributions of source code must retain the above copyright
'//--    notice, this list of conditions and the following disclaimer.
'
'//-- 2. Redistributions in binary form must reproduce the above copyright
'//--    notice, this list of conditions and the following disclaimer in
'//--    the documentation and/or other materials provided with the distribution.
'
'//-- 3. The end-user documentation included with the redistribution, if any,
'//--    must include the following acknowledgment:
'
'//--  "This product includes software developed by Gary Noble"
'
'//-- Alternately, this acknowledgment may appear in the software itself, if
'//-- and wherever such third-party acknowledgments normally appear.
'
'//-- THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
'//-- INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
'//-- AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
'//-- GARY NOBLE OR ANY CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
'//-- INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
'//-- BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
'//-- USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'//-- THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'//-- (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'//-- THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'//-----------------------------------------------------------------------


Option Explicit


Public Event HoverItem(ByVal oItem As EDItem)
Public Event ItemSelected(ByVal oItem As EDItem)
Public Event MouseRightClick()
Public Event ThemeChanged(ByVal sThemeName As String)

Private m_oTooltip As IAPP_ToolTip

'-- Mouse Tracking

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type


'-- Subclassing
Implements iSubclass
Private m_Subclass As IAPP_Subclass
Private m_SubClassMouse As IAPP_Subclass

Private Const WM_DISPLAYCHANGE As Long = &H7E
Private Const WM_EXITSIZEMOVE As Long = &H232
Private Const WM_ENTERSIZEMOVE As Long = &H231&
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A

Private m_lngActualVisibleItemCount As Long


'-- Mouse Is Hand Cursor
Private m_bHand As Boolean
Private m_bSplitterDown As Boolean


Private bBeginSize As Boolean                       '-- indicates That The Parent is Sizing
Private m_fntMenu As StdFont
Private m_bInMenu As Boolean                        '-- In Menu Button Flag
Private m_bInMenuShown As Boolean                   '-- Is the Menu Shown
Private m_strButtonDownKey As String                '-- Button Down
Private m_lngLastY As Long                          '-- Mouse Last Y CoOrdinate
Private m_RecInsidePanel As RECT
Private m_RecCaption As RECT
Private m_lngCaptionHeight As Long

'-- Visible Item Data
Private m_lngVisItemsMove As Long
Private m_lVisibleItemsMax As Long

'-- Draw Colors
Private m_lngColorOneSelectedNormal As OLE_COLOR
Private m_lngColorTwoSelectedNormal As OLE_COLOR
Private m_lngColorOneNormal As OLE_COLOR
Private m_lngColorTwoNormal As OLE_COLOR
Private m_lngColorOneSelected As OLE_COLOR
Private m_lngColorTwoSelected As OLE_COLOR
Private m_lngColorHeaderColorOne As OLE_COLOR
Private m_lngColorHeaderColorTwo As OLE_COLOR
Private m_lngColorHeaderForeColor As OLE_COLOR
Private m_lngColorHotOne As OLE_COLOR
Private m_lngColorHotTwo As OLE_COLOR
Private m_lngColorBorder As OLE_COLOR

'-- Defaults
Private Const m_const_lngDefToolbarHeight As Integer = 30
Private m_lngDefItemHeight As Long
Private Const m_const_lngDefItemHeightOffSet As Integer = 5
Private mDC As IAPP_MemDC
Private m_strSelectedKey As Variant
Private m_strHoverKey As Variant
Private m_lngFirstItemTop As Long
Private m_lngSplitterTop As Long
Private m_blnLeftButtonDown As Boolean
Private m_recSplitter As RECT
Private m_recMenu As RECT
Private m_EyeDropperItems As EDItems
Private Const m_def_VisibleItems As Integer = 6
Private m_VisibleItems As Long
Private m_HoverFont As Font
Private m_SelectedFont As StdFont
'Default Property Values:
Const m_def_UseCustomColor = False
Const m_def_CustomColor = vbButtonFace
Const m_def_Version = "v1.0"
Const m_def_Redraw = True
Const m_def_RighToLeft = False
Const m_def_DisplayHeader = True
Const m_def_SelectedForeColor = vbBlack
Const m_def_HoverForeColor = vbRed
Const m_def_NormalForeColor = vbBlack
Const m_def_HeaderForeColor = vbWindowText
'Property Variables:
Dim m_UseCustomColor As Boolean
Dim m_CustomColor As OLE_COLOR
Dim m_Version As String
Dim m_Redraw As Boolean
Dim m_RighToLeft As Boolean
Dim m_DisplayHeader As Boolean
Dim m_SelectedForeColor As OLE_COLOR
Dim m_HoverForeColor As OLE_COLOR
Dim m_NormalForeColor As OLE_COLOR
Dim m_HeaderForeColor As OLE_COLOR
Dim m_HeaderFont As Font

'-- Mouse Tracking
Private bTrack As Boolean
Private bTrackUser32 As Boolean
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'Determine if the passed function is supported
'-- Thanks To Paul Caton For This code
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod As Long
    Dim bLibLoaded As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        Call FreeLibrary(hMod)
    End If
End Function

'Track the mouse leaving the indicated window
'-- Thanks To Paul Caton For This code
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT

    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    pvDrawControl

End Property

'//---------------------------------------------------------------------------------------
' Procedure : Clear
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Clears The control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Sub Clear()
    On Error Resume Next

    Set m_EyeDropperItems = Nothing
    Set m_EyeDropperItems = New EDItems
    m_strSelectedKey = ""
    Set m_EyeDropperItems.EyeDropperControl = Me
    m_EyeDropperItems.hwnd = UserControl.hwnd
    VisibleItems = 0

    On Error GoTo 0
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    Dim xItem As EDItem

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    On Error Resume Next
    For Each xItem In m_EyeDropperItems
        xItem.Panel.Enabled = New_Enabled
    Next xItem
    pvDrawControl
    On Error GoTo 0

End Property

Friend Property Get Extender() As Object

    Set Extender = UserControl.Extender

End Property

Public Property Get EyeDropperItems() As EDItems
Attribute EyeDropperItems.VB_Description = "Items Collection"

    Set EyeDropperItems = m_EyeDropperItems

End Property


'//---------------------------------------------------------------------------------------
' Procedure : Font
' Type      : Property
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Default Item Font
' Returns   : Font
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

    Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set UserControl.Font = New_Font
    PropertyChanged "Font"

    '-- Get The Default Item Height

    With mDC
        Set .Font = New_Font
        If .TextHeight("`,Q") > m_lngDefItemHeight Then
            m_lngDefItemHeight = .TextHeight("`,Q") + 5
        End If
    End With    'mDC
    With mDC
        Set .Font = HoverFont
        If .TextHeight("`,Q") > m_lngDefItemHeight Then
            m_lngDefItemHeight = .TextHeight("`,Q") + 5
        End If
    End With    'mDC
    pvDrawControl

End Property

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


    m_lngColorOneSelected = 1
    m_lngColorTwoSelected = 1
    m_lngColorHeaderColorOne = 1
    m_lngColorHeaderColorTwo = 1
    m_lngColorHeaderForeColor = 1
    m_lngColorHotOne = 1
    m_lngColorHotTwo = 1
    mGlobal.GetThemeName UserControl.hwnd
    
    
    If AppThemed And UseCustomColor = False Then
        Select Case m_sCurrentSystemThemename
            Case "HomeStead"
                m_lngColorOneNormal = RGB(228, 235, 200)
                m_lngColorTwoNormal = RGB(175, 194, 142)
                m_lngColorBorder = RGB(100, 144, 88)
                m_lngColorHeaderColorOne = RGB(165, 182, 121)
                m_lngColorHeaderColorTwo = BlendColor(RGB(99, 122, 68), vbBlack, 200)
            Case "NormalColor"
                m_lngColorOneNormal = RGB(197, 221, 250)
                m_lngColorTwoNormal = RGB(128, 167, 225)
                m_lngColorBorder = RGB(0, 45, 150)
                m_lngColorHeaderColorOne = RGB(81, 128, 208)
                m_lngColorHeaderColorTwo = BlendColor(RGB(11, 63, 153), vbBlack, 230)
            Case "Metallic"
                m_lngColorOneNormal = RGB(219, 220, 232)
                m_lngColorTwoNormal = RGB(149, 147, 177)
                m_lngColorBorder = RGB(119, 118, 151)
                m_lngColorHeaderColorOne = RGB(163, 162, 187)
                m_lngColorHeaderColorTwo = BlendColor(RGB(112, 111, 145), vbBlack, 200)
            Case Else
                m_lngColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
                m_lngColorTwoNormal = vbButtonFace
                m_lngColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
                m_lngColorHeaderColorOne = vbButtonFace
                m_lngColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, vbBlack, 200)
                m_lngColorBorder = TranslateColor(vbInactiveTitleBar)
        End Select
        m_lngColorOneSelectedNormal = RGB(248, 216, 126)
        m_lngColorTwoSelectedNormal = RGB(240, 160, 38)
        m_lngColorHotOne = BlendColor(vbWindowBackground, vbButtonFace, 220)
        m_lngColorHotTwo = RGB(248, 216, 126)
        m_lngColorOneSelected = RGB(240, 160, 38)
        m_lngColorTwoSelected = RGB(248, 216, 126)
    ElseIf UseCustomColor = False Then
        m_lngColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
        m_lngColorTwoNormal = vbButtonFace
        m_lngColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
        m_lngColorHeaderColorOne = vbButtonFace
        m_lngColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbBlack, vbButtonFace, 10), 200)
        m_lngColorBorder = TranslateColor(vbInactiveTitleBar)
        m_lngColorHotTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 50), 10)
        m_lngColorHotOne = m_lngColorHotTwo
        m_lngColorOneSelected = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 100)
        m_lngColorTwoSelected = m_lngColorOneSelected
        m_lngColorOneSelectedNormal = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 130)
        m_lngColorTwoSelectedNormal = m_lngColorOneSelectedNormal
    ElseIf UseCustomColor Then
        
        m_lngColorOneNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 150)
        m_lngColorTwoNormal = TranslateColor(m_CustomColor)
        m_lngColorBorder = BlendColor(TranslateColor(m_CustomColor), vbBlack, 150)
        m_lngColorHeaderColorOne = BlendColor(TranslateColor(m_CustomColor), vbWhite, 120)
        m_lngColorHeaderColorTwo = BlendColor(TranslateColor(m_CustomColor), vbBlack, 200) 'BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite), 200)
        m_lngColorHotTwo = BlendColor(TranslateColor(m_CustomColor), vbWhite, 20) ' BlendColor(TranslateColor(m_CustomColor), vbBlack, 230)
        m_lngColorHotOne = BlendColor(TranslateColor(m_CustomColor), vbWhite, 40) ' BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite, 150), 130)
        m_lngColorTwoSelected = BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite, 150), 130)
        m_lngColorOneSelected = BlendColor(TranslateColor(m_CustomColor), vbBlack, 230)
        m_lngColorOneSelectedNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 70)
        m_lngColorTwoSelectedNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 70)
    
    End If

End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : GetNextSelectedItem
'//--Type      : Sub
'//--DateTime  : 04/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Sets The Selected Item If The Selected Item Is Delete Or Made Invisible
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Friend Sub GetNextSelectedItem()

    Dim xItem As EDItem
    Dim CTL As Control

    On Error Resume Next
    For Each CTL In UserControl.ContainedControls
        CTL.Visible = False
    Next CTL
    If Not m_strSelectedKey = Empty Then
        For Each xItem In m_EyeDropperItems
            With xItem
                If .Key = m_strSelectedKey And .Visible Then
                    If .Visible Then
                        If Not .Panel Is Nothing Then
                            .Panel.Visible = True
                        End If
                    Else
                        If Not .Panel Is Nothing Then
                            .Panel.Visible = False
                        End If
                    End If
                ElseIf .Key = m_strSelectedKey And Not .Visible Then
                    If Not .Panel Is Nothing Then
                        .Panel.Visible = False
                    End If
                End If
            End With    'xItem
        Next xItem
        If m_EyeDropperItems(m_strSelectedKey).Visible Then
            If Not m_EyeDropperItems(m_strSelectedKey).Panel Is Nothing Then
                m_EyeDropperItems(m_strSelectedKey).Panel.Visible = True
            End If
        Else
            For Each xItem In m_EyeDropperItems
                If xItem.Key <> m_strSelectedKey And xItem.Visible Then
                    m_strSelectedKey = xItem.Key
                    Exit For
                End If
            Next xItem
        End If
    Else
        For Each CTL In UserControl.ContainedControls
            CTL.Visible = False
        Next CTL
    End If
CleanExit:
    'SetItemRects
    '//--pvDrawControl
    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : HitTest
' Type      : Function
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Indicates If The Mouse Pointer is On An Item/Menu Or The Splitterbar
' Returns   : EDItem
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Function HitTest(Optional IsSplitterBar As Boolean = False) As EDItem
    On Error Resume Next

    Dim HitTestItem As EDItem
    Dim PT As POINTAPI
    Dim RC As RECT

    GetCursorPos PT
    ScreenToClient hwnd, PT
    m_bInMenu = False

    '-- Menu Button
    With RC
        .Top = m_recMenu.Top \ Screen.TwipsPerPixelY
        .Bottom = m_recMenu.Bottom    '//--\ Screen.TwipsPerPixelY
        .Left = m_recMenu.Left
        .Right = m_recMenu.Right
    End With
    If PtInRect(RC, PT.X, PT.Y) Then
        m_bInMenu = True
        Exit Function
    End If

    '-- Splitter
    With RC
        If Me.VisibleItems = 0 Then
            .Top = (m_recSplitter.Top \ Screen.TwipsPerPixelY) - 10
        Else
            .Top = m_recSplitter.Top \ Screen.TwipsPerPixelY
        End If
        .Bottom = m_recSplitter.Bottom \ Screen.TwipsPerPixelY
        .Left = m_recSplitter.Left
        .Right = m_recSplitter.Right
    End With
    If PtInRect(RC, PT.X, PT.Y) Then
        IsSplitterBar = True
        m_strHoverKey = ""
    End If


    '-- Not On The Above - Check To See if We Are On An Item
    For Each HitTestItem In m_EyeDropperItems
        If HitTestItem.Visible Then
            With RC
                .Top = HitTestItem.Top \ Screen.TwipsPerPixelY
                .Bottom = HitTestItem.Bottom \ Screen.TwipsPerPixelY
                .Left = HitTestItem.Left
                .Right = HitTestItem.Right
            End With
            If PtInRect(RC, PT.X, PT.Y) Then
                Set HitTest = HitTestItem
                Exit For
            End If
        End If
    Next HitTestItem

    '-- Reset The Hover Item Key
    If HitTestItem Is Nothing Then
        m_strHoverKey = ""
    End If

    On Error GoTo 0
End Function

Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd

End Property

Private Sub iSubclass_After(lReturn As Long, _
                            ByVal hwnd As Long, _
                            ByVal uMsg As WinSubHook.eMsg, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long)
    Static bInControl As Boolean


    If hwnd = UserControl.hwnd Then
        Select Case uMsg
            Case WM_MOUSELEAVE
            
                bInControl = False
                m_strHoverKey = ""
                
                If Not m_bInMenuShown Then
                    m_bInMenu = False
                End If
                
                If m_bInMenu And m_blnLeftButtonDown Then
                    pvDrawControl
                ElseIf Not m_blnLeftButtonDown Then
                    m_bInMenu = False
                    pvDrawControl
                Else
                    m_bInMenu = False
                    pvDrawControl
                End If
                
                MousePointer = vbDefault

                'Mouse has moved
            Case WM_MOUSEMOVE
                If hwnd = UserControl.hwnd Then
                    If Not bInControl Then
                        bInControl = True
                        Call TrackMouseLeave(UserControl.hwnd)
                        pvDrawControl
                    End If
                Else
                    '
                End If
        End Select

    Else
        Select Case uMsg
            Case WM_ENTERSIZEMOVE
                bBeginSize = True
                m_lngVisItemsMove = VisibleItems
            Case WM_DISPLAYCHANGE
                pvGetGradientColors
                pvDrawControl
            Case WM_EXITSIZEMOVE
            Case WM_SYSCOLORCHANGE
                pvGetGradientColors
                pvDrawControl
            Case WM_THEMECHANGED
                pvGetGradientColors
                RaiseEvent ThemeChanged(m_sCurrentSystemThemename)
                pvDrawControl
        End Select

    End If

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, _
                             lReturn As Long, _
                             hwnd As Long, _
                             uMsg As WinSubHook.eMsg, _
                             wParam As Long, _
                             lParam As Long)

'
End Sub


'//---------------------------------------------------------------------------------------
' Procedure : HoverFont
' Type      : Property
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Hover Font
' Returns   : Font
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Property Get HoverFont() As Font
Attribute HoverFont.VB_Description = "Normal Font"

    Set HoverFont = m_HoverFont

End Property

Public Property Set HoverFont(ByVal New_HoverFont As Font)

    Set m_HoverFont = New_HoverFont
    PropertyChanged "HoverFont"
    With mDC
        Set .Font = New_HoverFont
        If .TextHeight("`,Q") > m_lngDefItemHeight Then
            m_lngDefItemHeight = .TextHeight("`,Q") + 5
        End If
    End With    'mDC
    With mDC
        Set .Font = Font
        If .TextHeight("`,Q") > m_lngDefItemHeight Then
            m_lngDefItemHeight = .TextHeight("`,Q") + 5
        End If
    End With    'mDC

End Property

'//---------------------------------------------------------------------------------------
'//--Procedure : pvDrawControl
'//--Type      : Sub
'//--DateTime  : 04/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Draws The Actual Control
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvDrawControl()

       Dim xItem As EDItem
    Dim RC As RECT

    Dim isToolBarItem As Boolean
    Dim PT As POINTAPI
    Dim I As Integer

    '-- Get The Cursor Pos
    GetCursorPos PT
    ScreenToClient hwnd, PT

    If m_Redraw Then

        '-- Draw The Whole control Border
        With New IAPP_MemDC
            .Init ScaleWidth, ScaleHeight, hdc
            .BackStyle = BS_NEWTRANSPARENT
            .Rectangle 0, 0, ScaleWidth, ScaleHeight, BackColor, , TranslateColor(m_lngColorBorder)
            .BitBlt mDC.hdc, 0, 0, ScaleWidth, ScaleHeight
        End With

        If Me.DisplayHeader Then
            
            '-- Draw the Header
            With New IAPP_MemDC
                .Init ScaleWidth - 1, m_lngCaptionHeight, hdc
                .BackStyle = BS_NEWTRANSPARENT
                .Rectangle 0, 0, ScaleWidth - 1, m_lngCaptionHeight, vbApplicationWorkspace, , TranslateColor(m_lngColorBorder)
                .FillGradient 0, 0, ScaleWidth - 2, m_lngCaptionHeight, m_lngColorHeaderColorOne, m_lngColorHeaderColorTwo, True
                If Not Me.SelectedItem Is Nothing Then
                    Set .Font = Me.HeaderFont
                    If Me.Enabled Then
                        .ForeColor = Me.HeaderForeColor
                    Else
                        .ForeColor = vbGrayText
                    End If
                    If Not Me.SelectedItem.Picture Is Nothing Then
                        If Me.Enabled Then
                            .PaintPicture Me.SelectedItem.Picture, IIf(RighToLeft, ScaleWidth - 23, 3), (m_lngCaptionHeight \ 2) - 8, 16, 16, , , vbSrcCopy, vbBlack
                        Else
                            '==>.PaintDisabledPicture Me.SelectedItem.Picture, IIf(RighToLeft, ScaleWidth - 23, 3), (m_lngCaptionHeight \ 2) - 8, 16, 16, , , vbButtonFace
                                .PaintIconGrayscale .hdc, Me.SelectedItem.Picture, IIf(RighToLeft, ScaleWidth - 23, 3), (m_lngCaptionHeight \ 2) - 8, 16, 16
                        End If
                    End If
                    .DrawText Me.SelectedItem.Caption, IIf(Me.RighToLeft, 0, 24), (m_lngCaptionHeight \ 2) - (.TextHeight("'jQ") \ 2), ScaleWidth - IIf(RightToleft, 24, 28), m_lngCaptionHeight, IIf(RighToLeft, DT_RIGHT, DT_LEFT)
                End If
                .BitBlt mDC.hdc, 1, 1, ScaleWidth - 1, m_lngCaptionHeight, , , vbSrcCopy
            End With

        End If

        '-- Draw The Tool Bar Back Ground
        With New IAPP_MemDC
            .Init ScaleWidth - 1, m_const_lngDefToolbarHeight, hdc
            .BackStyle = BS_NEWTRANSPARENT
            .Rectangle 0, 0, ScaleWidth, m_const_lngDefToolbarHeight, vbApplicationWorkspace, , TranslateColor(m_lngColorBorder)
            .FillGradient 1, 1, ScaleWidth - 2, m_const_lngDefToolbarHeight, m_lngColorOneNormal, m_lngColorTwoNormal, True
            .BitBlt mDC.hdc, 1, (ScaleHeight) - (m_const_lngDefToolbarHeight) - 1, ScaleWidth, m_const_lngDefToolbarHeight, , , vbSrcCopy
        End With

        '-- Draw The Splitter
        With New IAPP_MemDC
            '//-- Splitter
            .Init ScaleWidth - 1, 9, hdc
            .BackStyle = BS_NEWTRANSPARENT
            .Rectangle 0, 0, ScaleWidth, 10, , , TranslateColor(m_lngColorBorder)
            .FillGradient 1, 1, ScaleWidth - 1, 9, m_lngColorHeaderColorOne, BlendColor(m_lngColorHeaderColorTwo, vbBlack), True
            '//-- Splitter Gripper Bar
            For I = 1 To 30 Step 5
                .SetPixel (ScaleWidth \ 2) - (2 + I), 3, vbBlack
                .SetPixel (ScaleWidth \ 2) - (1 + I), 3, vbBlack
                .SetPixel (ScaleWidth \ 2) - (1 + I), 3.5, vbWhite
                .SetPixel (ScaleWidth \ 2) + (2 + I), 3, vbBlack
                .SetPixel (ScaleWidth \ 2) + (1 + I), 3, vbBlack
                .SetPixel (ScaleWidth \ 2) + (2 + I), 3.5, vbWhite
            Next I
            .BitBlt mDC.hdc, 0, (m_lngSplitterTop \ Screen.TwipsPerPixelY) - IIf(m_VisibleItems > 0, 7, 9), ScaleWidth, m_const_lngDefToolbarHeight, , , vbSrcCopy
        End With


        '-- Draw Each Item
        For Each xItem In m_EyeDropperItems

            '-- only Draw What We Can See
            If xItem.Left > ScaleWidth Then GoTo CleanExit

            If xItem.Visible Then
                With RC
                    .Top = xItem.Top \ Screen.TwipsPerPixelY
                    .Bottom = xItem.Bottom \ Screen.TwipsPerPixelY
                    .Left = xItem.Left
                    .Right = xItem.Right
                End With
                With New IAPP_MemDC
                    .Init RC.Right - RC.Left, (RC.Bottom - RC.Top), hdc
                    .BackStyle = BS_NEWTRANSPARENT
                    isToolBarItem = xItem.isToolBarItem
                    If Not m_bSplitterDown And PtInRect(RC, PT.X, PT.Y) Then
                        If xItem.isToolBarItem And RC.Right >= m_recMenu.Left Then
                            GoTo MoveToNextItem
                        End If

                        Set .Font = Me.HoverFont
                        If Me.Enabled Then
                            If m_strHoverKey = xItem.Key Then
                                .ForeColor = Me.SelectedForeColor
                            Else
                                .ForeColor = Me.NormalForeColor
                            End If
                        Else
                            .ForeColor = vbGrayText
                        End If
                        If xItem.isToolBarItem = False Then
                            .Rectangle 0, 0 - IIf(Not isToolBarItem, 1, 0), RC.Right - RC.Left, RC.Bottom - RC.Top, , , TranslateColor(m_lngColorBorder)
                        End If
                        If m_strSelectedKey = xItem.Key Then

                            If m_strSelectedKey = m_strButtonDownKey Then
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneSelected, m_lngColorTwoSelected, True
                            Else
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneSelectedNormal, m_lngColorTwoSelectedNormal, True
                            End If
                        Else
                            If m_strHoverKey = xItem.Key Then .ForeColor = Me.HoverForeColor
                            If m_strButtonDownKey = xItem.Key Then
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneSelected, m_lngColorTwoSelected, True
                            Else
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorHotOne, m_lngColorHotTwo, True
                            End If
                        End If
                        If xItem.isToolBarItem = False Then
                            If Not Enabled Then
                                If Not xItem.Picture Is Nothing Then
                                    '==> .PaintDisabledPicture xItem.Picture, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24
                                    .PaintIconGrayscale .hdc, xItem.Picture.handle, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24
                                       
                                       
                                End If
                            Else
                                If Not xItem.Picture Is Nothing Then
                                    .PaintPicture xItem.Picture, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24, , , vbSrcCopy, vbBlack
                                End If
                            End If
                            
                            .DrawText xItem.Caption, IIf(m_RighToLeft, 0, 35), ((RC.Bottom - RC.Top) / 2) - (.TextHeight("`!,") / 2), (RC.Right - RC.Left) - IIf(m_RighToLeft, 40, 5), RC.Bottom - RC.Top, IIf(m_RighToLeft, DT_RIGHT, DT_LEFT)

                        Else
                            If Not Enabled Then
                                If Not xItem.Picture Is Nothing Then
                                    '==> .PaintDisabledPicture xItem.Picture, ((RC.Right - RC.Left) / 2) - 8, ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16
                                         .PaintIconGrayscale .hdc, xItem.Picture.handle, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16
                                End If
                            Else
                                If Not xItem.Picture Is Nothing Then
                                    .PaintPicture xItem.Picture, ((RC.Right - RC.Left) / 2) - 8, ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16, , , vbSrcCopy, vbBlack
                                End If
                            End If
                        End If
                    Else
                        Set .Font = Me.Font
                        If Me.Enabled Then

                            If m_strHoverKey = xItem.Key Then .ForeColor = Me.HoverForeColor
                            If xItem.Key = m_strSelectedKey And xItem.Key <> m_strHoverKey Then .ForeColor = Me.HoverForeColor
                            If Not xItem.Key = m_strHoverKey And xItem.Key <> m_strSelectedKey Then .ForeColor = Me.NormalForeColor

                        Else
                            .ForeColor = vbGrayText
                        End If
                        If xItem.isToolBarItem And RC.Right >= m_recMenu.Left Then
                            GoTo MoveToNextItem
                        End If
                        If xItem.isToolBarItem = False Then
                            .Rectangle 0, 0 - IIf(Not isToolBarItem, 1, 0), RC.Right - RC.Left, RC.Bottom - RC.Top, , , TranslateColor(m_lngColorBorder)
                        End If
                        If m_strSelectedKey = xItem.Key Then
                            If Me.Enabled Then
                                .ForeColor = Me.SelectedForeColor
                            Else
                                .ForeColor = vbGrayText
                            End If


                            If m_strSelectedKey = m_strButtonDownKey Then
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneSelected, m_lngColorTwoSelected, True
                            Else
                                .FillGradient IIf(Not isToolBarItem, 1, 0), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneSelectedNormal, m_lngColorTwoSelectedNormal, True
                            End If
                        Else
                            .FillGradient IIf(Not isToolBarItem, 1, -5), 0, (RC.Right - RC.Left) - IIf(Not isToolBarItem, 1, 0), (RC.Bottom - RC.Top) - IIf(Not isToolBarItem, 1, 0), m_lngColorOneNormal, m_lngColorTwoNormal, True
                        End If
                        If xItem.isToolBarItem = False Then
                            If Not Enabled Then
                                If Not xItem.Picture Is Nothing Then
                                    '==> .PaintDisabledPicture xItem.Picture, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24
                                         .PaintIconGrayscale .hdc, xItem.Picture.handle, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24

                                End If
                            Else
                                If Not xItem.Picture Is Nothing Then
                                    .PaintPicture xItem.Picture, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 12, 12)), 24, 24, , , vbSrcCopy, vbBlack
                                End If
                            End If
                            .DrawText xItem.Caption, IIf(m_RighToLeft, 0, 35), ((RC.Bottom - RC.Top) / 2) - (.TextHeight("`!,") / 2), (RC.Right - RC.Left) - IIf(m_RighToLeft, 40, 5), RC.Bottom - RC.Top, IIf(m_RighToLeft, DT_RIGHT, DT_LEFT)
                        Else
                            If Not Enabled Then
                                If Not xItem.Picture Is Nothing Then
                                   '==> .PaintDisabledPicture xItem.Picture, ((RC.Right - RC.Left) / 2) - 8, ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16
                                        .PaintIconGrayscale .hdc, xItem.Picture.handle, IIf(m_RighToLeft, ScaleWidth - 35, 3), ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16
                                End If
                            Else
                                If Not xItem.Picture Is Nothing Then
                                    .PaintPicture xItem.Picture, ((RC.Right - RC.Left) / 2) - 8, ((RC.Bottom - RC.Top) / 2) - (IIf(m_blnLeftButtonDown, 8, 8)), 16, 16, , , vbSrcCopy, vbBlack
                                End If
                            End If
                        End If
                    End If
                    .BitBlt mDC.hdc, RC.Left, RC.Top + 1, ScaleWidth, RC.Bottom - RC.Top, , , vbSrcCopy
MoveToNextItem:
                End With
            End If
        Next xItem

CleanExit:


        '//-- Menu Button
        If m_EyeDropperItems.Count > 0 Then

            With New IAPP_MemDC
                .Init 15, m_const_lngDefToolbarHeight, hdc
                .BackStyle = BS_NEWTRANSPARENT
                Set .Font = m_fntMenu
                .ForeColor = IIf(Enabled, vbBlack, vbGrayText)
                .Rectangle 14, 0, 15, m_const_lngDefToolbarHeight, , , TranslateColor(m_lngColorBorder)
                If Not m_bInMenu Then
                    .FillGradient 0, 0, 14, m_const_lngDefToolbarHeight - 1, m_lngColorOneNormal, m_lngColorTwoNormal, True
                Else
                    If m_blnLeftButtonDown Then
                        .FillGradient 0, 0, 14, m_const_lngDefToolbarHeight - 1, m_lngColorOneSelected, m_lngColorTwoSelected, True
                    Else
                        .FillGradient 0, 0, 14, m_const_lngDefToolbarHeight - 1, m_lngColorHotOne, m_lngColorHotTwo, True
                    End If
                End If
                .DrawText "4", -1, 5, 15, m_const_lngDefToolbarHeight, DT_LEFT
                .DrawText "4", 3, 5, 15, m_const_lngDefToolbarHeight, DT_LEFT
                .DrawText "6", 2, 13, 15, m_const_lngDefToolbarHeight, DT_LEFT
                .BitBlt mDC.hdc, ScaleWidth - 15, (ScaleHeight) - (m_const_lngDefToolbarHeight) + 1, ScaleWidth, m_const_lngDefToolbarHeight - 2, , , vbSrcCopy
            End With
        End If
        mDC.BitBlt hdc, 0, 0, ScaleWidth, ScaleHeight, , , vbSrcCopy
        If m_bInMenu Then
            mGlobal.SetScreenCursor True
            m_bHand = True
        End If

    End If


End Sub

'//---------------------------------------------------------------------------------------
' Procedure : RedrawControl
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Global Redraw Call
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Friend Sub RedrawControl()

    VisibleItems = VisibleItems
    SetItemRects
    pvDrawControl

End Sub




'//---------------------------------------------------------------------------------------
' Procedure : SelectedItem
' Type      : Property
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Selected Item
' Returns   : EDItem
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Property Get SelectedItem() As EDItem
    On Error GoTo errSelected


    If m_lngActualVisibleItemCount > 0 Then
        Set SelectedItem = m_EyeDropperItems(m_strSelectedKey)
    End If


CleanExit:
    On Error GoTo 0

    Exit Sub

errSelected:
    Err.Raise vbObjectError + 10001, "EyeDropper.EDItem - Selected Item", Err.Description
    Resume CleanExit

End Property

Public Property Let SelectedItem(ByVal xItem As EDItem)

    On Error GoTo errSelectedItem

    m_strSelectedKey = xItem.Key
    pvDrawControl
    m_strSelectedKey = xItem.Key

CleanExit:
    On Error GoTo 0

    Exit Property
errSelectedItem:
    Err.Raise vbObjectError + 10001, "EyeDropper.SelectedItem", Err.Description
    Resume CleanExit

End Property



'//---------------------------------------------------------------------------------------
' Procedure : SelectedKey
' Type      : Property
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Internal Set Key Call
' Returns   : Variant
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Friend Property Get SelectedKey() As Variant

    SelectedKey = m_strSelectedKey

End Property

Friend Property Let SelectedKey(sKey As Variant)

    m_strSelectedKey = sKey
    pvDrawControl

End Property

'//---------------------------------------------------------------------------------------
' Procedure : SetItemRects
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Sets The Item Rectangle Data For Drawing
' Returns   :
' Notes     : this Only Gets Called If A New Item Is Added / Delete Or The Parent is Sizing
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Friend Sub SetItemRects()

    If Not m_Redraw Then Exit Sub
    
    Dim xItem As EDItem
    Dim lLeft As Long
    Dim lToobarMinusVisibleItemCount As Long
    Dim I As Long
    
        
    With m_recMenu
        .Left = ScaleWidth - 15
        .Top = (ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)
        .Right = ScaleWidth
        .Bottom = ScaleHeight
    End With

    m_lngActualVisibleItemCount = 0


    For Each xItem In m_EyeDropperItems
        If xItem.Visible Then
            lToobarMinusVisibleItemCount = lToobarMinusVisibleItemCount + 1
        End If
    Next

    '-- Set Where The First Toolbar Icon Will Be Drawn
    lToobarMinusVisibleItemCount = lToobarMinusVisibleItemCount - Me.VisibleItems
    If lToobarMinusVisibleItemCount > Round((ScaleWidth - 15) / 23) Then lToobarMinusVisibleItemCount = Round((ScaleWidth - 15) / 23)
    lLeft = (ScaleWidth - 20) - (lToobarMinusVisibleItemCount * 23)

    
    If Me.VisibleItems > 0 Then
        I = 1
        m_lngFirstItemTop = (ScaleHeight * Screen.TwipsPerPixelY - (VisibleItems * (m_lngDefItemHeight * Screen.TwipsPerPixelY))) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)
        m_lngSplitterTop = m_lngFirstItemTop
        m_lngLastY = m_lngFirstItemTop \ Screen.TwipsPerPixelY
        With m_recSplitter
            .Top = m_lngFirstItemTop - 200
            .Bottom = m_lngFirstItemTop
            .Left = 0
            .Right = ScaleWidth
        End With


        For Each xItem In m_EyeDropperItems
            If lLeft > ScaleWidth Then Exit For
            If xItem.Visible Then
                xItem.SetRectItem 0, 0, 0, 0
                If I <= Me.VisibleItems Then
                    m_lngActualVisibleItemCount = m_lngActualVisibleItemCount + 1
                    With xItem
                        .isToolBarItem = False
                        .SetRectItem 0, m_lngFirstItemTop, m_lngFirstItemTop + (m_lngDefItemHeight * Screen.TwipsPerPixelY), ScaleWidth
                        m_lngFirstItemTop = ((m_lngFirstItemTop + .ItemRect.Bottom) + (m_lngDefItemHeight * Screen.TwipsPerPixelY))
                    End With    'xItem
                    ' lLeft = lLeft + (23)
                Else
                    xItem.isToolBarItem = True
                    If lLeft >= ScaleWidth Then Exit For
                    xItem.SetRectItem lLeft, (ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY), (ScaleHeight * Screen.TwipsPerPixelY) - m_const_lngDefToolbarHeight, lLeft + 22
                    lLeft = lLeft + (23)
                End If
                I = I + 1
            End If
        Next xItem
    Else

        '-- Toolbar Items
        m_lngFirstItemTop = (ScaleHeight * Screen.TwipsPerPixelY - (VisibleItems * (m_lngDefItemHeight * Screen.TwipsPerPixelY))) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)
        m_lngSplitterTop = m_lngFirstItemTop
        With m_recSplitter
            m_lngLastY = ((ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)) - 20
            .Top = ((ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)) - 20
            .Bottom = (ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY)
            .Left = 0
            .Right = ScaleWidth
        End With

        For Each xItem In m_EyeDropperItems
            If lLeft > ScaleWidth Then Exit For
            With xItem
                .SetRectItem 0, 0, 0, 0
                If .Visible Then
                    If lLeft >= ScaleWidth Then Exit For
                    m_lngActualVisibleItemCount = m_lngActualVisibleItemCount + 1
                    .isToolBarItem = True
                    .SetRectItem lLeft, (ScaleHeight * Screen.TwipsPerPixelY) - (m_const_lngDefToolbarHeight * Screen.TwipsPerPixelY), (ScaleHeight * Screen.TwipsPerPixelY) - m_const_lngDefToolbarHeight, lLeft + 22
                    lLeft = lLeft + (23)
                    
                End If
            End With    'xItem
        Next xItem
    End If
    With m_RecInsidePanel
        .Top = 10 + IIf(Me.DisplayHeader, m_lngCaptionHeight * Screen.TwipsPerPixelY, 0)
        .Left = 10
        .Right = ScaleWidth - 2
        .Bottom = m_lngSplitterTop - (IIf(VisibleItems = 0, 150, 120) + IIf(Me.DisplayHeader, (m_lngCaptionHeight * Screen.TwipsPerPixelY), 0))
    End With

    '-- Make Sure The Seleted Item Is Visible
    GetNextSelectedItem

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvShowMenu
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Creates A New Menu With The Items Of The Control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvShowMenu()
    On Error Resume Next

    Dim sKey As String
    Dim NewMenu As IAPP_MenuHandler
    Dim xItem As EDItem
    Dim lParentIndex As Long
    Dim lParentIndexADDREM As Long
    Dim rID As Long
    Dim lMaxCount As Long

    Set NewMenu = New IAPP_MenuHandler
    With NewMenu

        .DrawStyle = IIf(mGlobal.AppThemed, mds_XP, mds_3D)
        .CreateFromNothing UserControl.hwnd

        .RightToleft = m_RighToLeft

        lParentIndex = .AddItem(0, Key:="MainMenu")

        '-- Add The Items to The Visible Menu
        For Each xItem In m_EyeDropperItems
            If xItem.Visible Then
                lMaxCount = lMaxCount + 1
                .AddItem lParentIndex, xItem.Caption, , , xItem.Key & "¦", Enabled, IIf(xItem.Key = m_strSelectedKey, True, False), EMCS_Icon, xItem.Picture
            End If
        Next xItem

        '-- Splitter
        If lMaxCount > 0 Then .AddItem lParentIndex, "-", , , "--¬", Enabled

        '--Add Remove Menu
        lParentIndexADDREM = .AddItem(lParentIndex, "Add Or Remove Items ", , , "EDADDREM¬", Enabled)

        '-- Add The Items to The Add Remove Menu
        For Each xItem In m_EyeDropperItems
            .AddItem lParentIndexADDREM, xItem.Caption, , , xItem.Key & "¬¬", Enabled, xItem.Visible, EMCS_Icon, xItem.Picture
        Next xItem

        If lMaxCount > 0 Then .AddItem lParentIndex, "-", , , "-¬", Enabled

        If VisibleItems > 0 Then
            .AddItem lParentIndex, "Less Items", "Ctrl+L", , "EDLess¬", Enabled
        End If

        If lMaxCount > 0 And VisibleItems < lMaxCount Then
            .AddItem lParentIndex, "More Items", "Ctrl+M", , "EDMore¬", Enabled
        End If

        '-- Flag The The Menu Is Shown
        m_bInMenuShown = True

        '-- Show The Menu And Wait For The Return Value
        rID = .PopUpMenu("MainMenu", , , TPM_lEFTALIGN)
        If rID <> 0 Then
            .CurrentMenuIndex = .IndexForID(rID)
            sKey = .ItemKey(.CurrentMenuIndex)
Debug.Print sKey
            If InStr(1, sKey, "¬¬") Then
                m_EyeDropperItems(Replace$(sKey, "¬¬", vbNullString)).Visible = Not m_EyeDropperItems(Replace$(sKey, "¬¬", vbNullString)).Visible
                VisibleItems = VisibleItems
                If VisibleItems = 0 Then
                    If m_EyeDropperItems(Replace$(sKey, "¬¬", "")).Visible Then
                        If LenB(m_strSelectedKey) = 0 Then
                            m_strSelectedKey = Replace$(sKey, "¬¬", vbNullString)
                        End If
                    End If
                End If
            ElseIf InStr(1, sKey, "¦") Then
                m_strSelectedKey = Replace$(sKey, "¦", vbNullString)
                RaiseEvent ItemSelected(Me.EyeDropperItems(m_strSelectedKey))
                pvSizePanel m_EyeDropperItems(m_strSelectedKey)
            ElseIf sKey = "EDMore¬" Then
                VisibleItems = VisibleItems + 1
            ElseIf sKey = "EDLess¬" Then
                VisibleItems = VisibleItems - 1

            End If
            'MsgBox .ItemKey
        End If
    End With

    Set NewMenu = Nothing
    m_bInMenuShown = False
    m_bInMenu = False
    SetItemRects
    GetNextSelectedItem
    pvDrawControl

    On Error GoTo 0
End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvSizePanel
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Sizes The Item Panel To The Inside CoOrdinates Of The control
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvSizePanel(xItem As EDItem)

    Dim vItem As EDItem

    On Error Resume Next
    If Not xItem.Panel Is Nothing Then
        xItem.Panel.ZOrder
        If Not xItem.Key = m_strSelectedKey Then
            With m_RecInsidePanel
                xItem.Panel.Move .Left, .Top, .Right * Screen.TwipsPerPixelX, .Bottom
                If Not xItem.Panel.Visible Then
                    xItem.Panel.Visible = True
                End If
            End With
        Else
            With m_RecInsidePanel
                xItem.Panel.Move .Left, .Top, .Right * Screen.TwipsPerPixelX, .Bottom
                If Not xItem.Panel.Visible Then
                    xItem.Panel.Visible = True
                End If
            End With
        End If
    Else
        For Each vItem In m_EyeDropperItems
            DoEvents
            If Not xItem.Panel = vItem.Panel Then
                If Not vItem.Panel.Visible = False Then
                    vItem.Panel.Visible = False
                End If
            End If
        Next vItem
    End If


    If VisibleItems = 0 Then
        If m_recSplitter.Top < 150 And Me.DisplayHeader = False Then
            xItem.Panel.Visible = False
        ElseIf m_recSplitter.Top < 175 + (m_lngCaptionHeight * Screen.TwipsPerPixelY) And Me.DisplayHeader = True Then
            xItem.Panel.Visible = False
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub UserControl_Click()

    If m_bHand Then
        mGlobal.SetScreenCursor True
    End If

End Sub

Private Sub UserControl_DblClick()

    If m_bHand Then
        mGlobal.SetScreenCursor True
    End If

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next

    Set m_oTooltip = New IAPP_ToolTip

    Set m_fntMenu = New StdFont
    m_fntMenu.Name = "Marlett"
    Set m_Subclass = New IAPP_Subclass
    Set m_SubClassMouse = New IAPP_Subclass


    Set mDC = New IAPP_MemDC
    
    Set m_EyeDropperItems = New EDItems
    Set m_EyeDropperItems.EyeDropperControl = Me
    m_EyeDropperItems.hwnd = UserControl.hwnd


    mGlobal.GetThemeName hwnd
    pvGetGradientColors
    pvDrawControl

    On Error GoTo 0

End Sub
Public Sub Initialise()
    On Error Resume Next

    With m_Subclass
        .UnSubclass
        .Subclass Parent.hwnd, Me
        .AddMsg WM_DISPLAYCHANGE, MSG_AFTER
        .AddMsg WM_SYSCOLORCHANGE, MSG_AFTER
        .AddMsg WM_THEMECHANGED, MSG_AFTER
        .AddMsg WM_ENTERSIZEMOVE, MSG_AFTER
    End With


    '-- Determine what level of window/mouse tracking support is available

    bTrack = True

    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "user32")

    If Not bTrackUser32 Then
        If Not IsFunctionExported("_TrackMouseEvent", "comctl32") Then
            bTrack = False
        End If
    End If

    If bTrack Then

        m_SubClassMouse.UnSubclass

        '-- OS supports mouse leave, so subclass for it
        With UserControl
            'Start subclassing the PictureBox
            Call m_SubClassMouse.Subclass(.hwnd, Me)
            Call m_SubClassMouse.AddMsg(WM_MOUSELEAVE, MSG_AFTER)
            Call m_SubClassMouse.AddMsg(WM_MOUSEMOVE, MSG_AFTER)
        End With
    End If

    On Error GoTo 0

End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    Set m_HoverFont = Ambient.Font
    Set m_SelectedFont = Ambient.Font
    m_lngDefItemHeight = 40
    m_VisibleItems = m_def_VisibleItems
    Set UserControl.Font = Ambient.Font

    m_SelectedForeColor = m_def_SelectedForeColor
    m_HoverForeColor = m_def_HoverForeColor
    m_NormalForeColor = m_def_NormalForeColor
    m_HeaderForeColor = m_def_HeaderForeColor
    Set m_HeaderFont = Ambient.Font

    m_DisplayHeader = m_def_DisplayHeader
    m_RighToLeft = m_def_RighToLeft
    m_Redraw = m_def_Redraw
    m_Version = "v" & App.Major & "." & App.Minor & App.Revision
    m_CustomColor = m_def_CustomColor
    m_UseCustomColor = m_def_UseCustomColor
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 12 Then VisibleItems = VisibleItems - 1
    If KeyAscii = 13 Then VisibleItems = VisibleItems + 1

    On Error GoTo 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    If Button = vbRightButton Then
        '--Raise The Right Mouse Click Event
        RaiseEvent MouseRightClick
        If m_bHand Then
            If m_EyeDropperItems.Count > 0 Then mGlobal.SetScreenCursor True
        End If
    Else

    Dim xItem As EDItem
    Dim bSplitter As Boolean

    m_blnLeftButtonDown = True
    Set xItem = HitTest(bSplitter)

    '-- Normal Mouse Down
    If Not bSplitter And Not m_bInMenu Then
        If Not xItem Is Nothing Then
            m_strHoverKey = xItem.Key
            m_strButtonDownKey = xItem.Key
            m_blnLeftButtonDown = True
            If m_bHand Then
                If m_EyeDropperItems.Count > 0 Then mGlobal.SetScreenCursor True
            End If
            pvDrawControl
            If m_bHand Then
                If m_EyeDropperItems.Count > 0 Then mGlobal.SetScreenCursor True
            End If
        End If
        '-- Menu Mouse Down
    ElseIf m_bInMenu Then
        If m_EyeDropperItems.Count > 0 Then pvShowMenu

        '-- Other
    Else
        m_bSplitterDown = True
        m_strHoverKey = ""
    End If
    
    End If
    
    pvDrawControl
    Set xItem = Nothing
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    Dim PT As POINTAPI
    Dim xItem As EDItem
    Dim bSplitter As Boolean

    On Error Resume Next

    Dim oTooltip As IAPP_ToolTip
    Set oTooltip = New IAPP_ToolTip

    '-- Bail
    If m_EyeDropperItems.Count <= 0 Then
        Exit Sub
    End If

    GetCursorPos PT
    ScreenToClient hwnd, PT


    '-- Check That The Splitter Is not Down
    If Not m_bSplitterDown Then
        Set xItem = HitTest(bSplitter)
    End If

    '-- Wrer In The Menu Button Params So Just Draw And Bail
    If m_bInMenu Then
        pvDrawControl
        Exit Sub
    End If

    '-- Get The Item The Mouse is Over
    If Not bSplitter And Not m_bSplitterDown Then
        UserControl.MousePointer = vbDefault
        If Not xItem Is Nothing Then
            If m_strHoverKey <> xItem.Key Then

                '-- Set The Hover Key
                m_strHoverKey = xItem.Key

                RaiseEvent HoverItem(xItem)

                pvDrawControl

                m_oTooltip.Destroy


                '-- Show Tooltip If The item Is In The Toolbar
                If xItem.isToolBarItem Then

                    With m_oTooltip
                        .Destroy
                        .Style = TTStandard
                        .TipText = xItem.Caption
                        .VisibleTime = 2000
                        .DelayTime = 300
                        .Create UserControl.hwnd
                    End With

                End If
            End If

            '-- Check The Hand Cursor
            If m_bHand Then
                mGlobal.SetScreenCursor True
            End If
            m_bHand = True
        Else


            pvDrawControl

            '-- Kill The Tooltip
            m_oTooltip.Destroy

            '-- Check The Hand Cursor
            If m_bHand Then
                mGlobal.SetScreenCursor False
            End If
            m_bHand = False
        End If
    ElseIf m_bSplitterDown Then

        '-- Show More items
        If PT.Y < m_lngLastY - (m_lngDefItemHeight - 5) Then
            If VisibleItems <= m_EyeDropperItems.Count Then
                If VisibleItems <= m_EyeDropperItems.Count Then
                    If Y > (ScaleHeight) - (m_const_lngDefToolbarHeight + IIf(VisibleItems = 0, m_const_lngDefToolbarHeight, 0)) Then
                        Exit Sub
                    End If
                    VisibleItems = VisibleItems + 1
                End If
            End If
        Else
            '-- Show Less items
            If PT.Y >= (m_lngLastY) + (m_lngDefItemHeight - 5) Then
                If Y > (ScaleHeight) - (m_const_lngDefToolbarHeight) Then
                    VisibleItems = 0
                    Exit Sub
                End If
                If VisibleItems > 0 Then
                    VisibleItems = VisibleItems - 1
                End If
            End If
        End If
    Else

        '-- Reset The Cursor
        If UserControl.MousePointer <> 7 Then
            UserControl.MousePointer = 7
        End If
        If Button <> vbLeftButton Then
            pvDrawControl
        End If

    End If
    On Error GoTo 0

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    Dim xItem As EDItem

    m_blnLeftButtonDown = False
    m_bInMenu = False
    
    If Not m_bSplitterDown Then
        m_bSplitterDown = False
        Set xItem = HitTest()
        pvDrawControl
        If m_bHand Then
            mGlobal.SetScreenCursor True
        End If

        '-- Set The Selected Item
        If Not xItem Is Nothing Then
            If m_strButtonDownKey = m_strHoverKey Then
                m_strSelectedKey = xItem.Key
                RaiseEvent ItemSelected(xItem)
                '-- Size The Selected Panel
                mGlobal.SetScreenCursor True
                pvDrawControl
                pvSizePanel xItem
            End If


        End If
    End If

    '-- Reset The Values
    m_bSplitterDown = False
    m_strButtonDownKey = ""

End Sub

Private Sub UserControl_Paint()

    pvDrawControl

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_lngDefItemHeight = 36
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set m_HoverFont = PropBag.ReadProperty("HoverFont", Ambient.Font)
    Set m_SelectedFont = PropBag.ReadProperty("SelectedFont", Ambient.Font)
    m_VisibleItems = PropBag.ReadProperty("VisibleItems", m_def_VisibleItems)
    Set mDC.Font = m_HoverFont
    If mDC.TextHeight("`,Q") > m_lngDefItemHeight Then
        m_lngDefItemHeight = mDC.TextHeight("`,Q") + m_const_lngDefItemHeightOffSet
    End If

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)

    m_SelectedForeColor = PropBag.ReadProperty("SelectedForeColor", m_def_SelectedForeColor)
    m_HoverForeColor = PropBag.ReadProperty("HoverForeColor", m_def_HoverForeColor)
    m_NormalForeColor = PropBag.ReadProperty("NormalForeColor", m_def_NormalForeColor)
    m_HeaderForeColor = PropBag.ReadProperty("HeaderForeColor", m_def_HeaderForeColor)
    Set m_HeaderFont = PropBag.ReadProperty("HeaderFont", Ambient.Font)
    m_DisplayHeader = PropBag.ReadProperty("DisplayHeader", m_def_DisplayHeader)

    With mDC
        Set .Font = m_HeaderFont
        m_lngCaptionHeight = .TextHeight("`,Q") + m_const_lngDefItemHeightOffSet
        If m_lngCaptionHeight < 24 Then m_lngCaptionHeight = 24
    End With


    m_RighToLeft = PropBag.ReadProperty("RighToLeft", m_def_RighToLeft)
    m_Redraw = PropBag.ReadProperty("Redraw", m_def_Redraw)
    m_Version = "v" & App.Major & "." & App.Minor & App.Revision
    
    m_CustomColor = PropBag.ReadProperty("CustomColor", m_def_CustomColor)
    m_UseCustomColor = PropBag.ReadProperty("UseCustomColor", m_def_UseCustomColor)
    
    If m_UseCustomColor Then
        pvGetGradientColors
        pvDrawControl
    End If
    
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    '//-- Sets The Max Visible Items

    m_VisibleItems = m_lngVisItemsMove
    m_lVisibleItemsMax = Round(((ScaleHeight - (40 + IIf(Me.DisplayHeader, m_lngCaptionHeight, 0))) \ m_lngDefItemHeight))

    If m_VisibleItems >= m_lVisibleItemsMax Then
        VisibleItems = m_lVisibleItemsMax
    End If

    '-- Re Initialise The Memory DC
    mDC.Init ScaleWidth, ScaleHeight, hdc
    mDC.BackStyle = BS_NEWTRANSPARENT

    '-- Do IT
    SetItemRects
    pvDrawControl

    If m_strSelectedKey > "" Then
        pvSizePanel m_EyeDropperItems(m_strSelectedKey)
    End If
    On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()

'//-- Clean Up
    On Error Resume Next
    m_Subclass.UnSubclass
    
    Set m_Subclass = Nothing
    
    Set m_fntMenu = Nothing
    
    
    Set m_EyeDropperItems = Nothing
    Set mDC = Nothing

    m_oTooltip.Destroy

    Set m_oTooltip = Nothing

    m_SubClassMouse.UnSubclass


    Set m_SubClassMouse = Nothing

    '
    On Error GoTo 0

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)


    With PropBag
        .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "HoverFont", m_HoverFont, Ambient.Font
        .WriteProperty "SelectedFont", m_SelectedFont, m_HoverFont
        .WriteProperty "VisibleItems", m_VisibleItems, m_def_VisibleItems
        .WriteProperty "Font", UserControl.Font, Ambient.Font
    End With    'PropBag

    Call PropBag.WriteProperty("SelectedForeColor", m_SelectedForeColor, m_def_SelectedForeColor)
    Call PropBag.WriteProperty("HoverForeColor", m_HoverForeColor, m_def_HoverForeColor)
    Call PropBag.WriteProperty("NormalForeColor", m_NormalForeColor, m_def_NormalForeColor)
    Call PropBag.WriteProperty("HeaderForeColor", m_HeaderForeColor, m_def_HeaderForeColor)
    Call PropBag.WriteProperty("HeaderFont", m_HeaderFont, Ambient.Font)
    Call PropBag.WriteProperty("DisplayHeader", m_DisplayHeader, m_def_DisplayHeader)
    Call PropBag.WriteProperty("RighToLeft", m_RighToLeft, m_def_RighToLeft)
    Call PropBag.WriteProperty("Redraw", m_Redraw, m_def_Redraw)
    Call PropBag.WriteProperty("Version", m_Version = "v" & App.Major & "." & App.Minor & App.Revision, m_Version = "v" & App.Major & "." & App.Minor & App.Revision)
    
    Call PropBag.WriteProperty("CustomColor", m_CustomColor, m_def_CustomColor)
    Call PropBag.WriteProperty("UseCustomColor", m_UseCustomColor, m_def_UseCustomColor)
End Sub

Public Property Get VisibleItems() As Long
Attribute VisibleItems.VB_Description = "Visible Tab Item Count"
    On Error Resume Next

    If m_EyeDropperItems.Count > 0 Then
        VisibleItems = m_VisibleItems
    Else
        VisibleItems = 0
    End If

    On Error GoTo 0
End Property

'//---------------------------------------------------------------------------------------
' Procedure : VisibleItems
' Type      : Property
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Sets The Max/Min Visible Items
' Returns   : Long
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Property Let VisibleItems(ByVal New_VisibleItems As Long)

    Dim xItem As EDItem
    Dim lMax As Long

    '-- Chack To See If The Parent is Sizing
    If Not bBeginSize Then
        m_lngVisItemsMove = New_VisibleItems
    End If

    m_VisibleItems = New_VisibleItems
    If m_VisibleItems < 0 Then
        m_VisibleItems = 0
    End If
    If m_VisibleItems > 1 Then
        m_lngLastY = m_recSplitter.Top
    Else
        m_lngLastY = (ScaleHeight) - (m_const_lngDefToolbarHeight)
    End If

    Dim ll As Long
    
    '-- Get The Max Of The Visible items
    For Each xItem In m_EyeDropperItems
        If xItem.Visible Then
               lMax = lMax + 1
        End If
    Next xItem

    If lMax > Round(((ScaleHeight - (40 + IIf(Me.DisplayHeader, m_lngCaptionHeight, 0))) \ m_lngDefItemHeight)) Then
        lMax = Round(((ScaleHeight - (40 + IIf(Me.DisplayHeader, m_lngCaptionHeight, 0))) \ m_lngDefItemHeight))
    End If
    
    '-- Set The Max
    If m_VisibleItems >= lMax Then
        m_VisibleItems = lMax
    End If


    PropertyChanged "VisibleItems"

    '-- Do it
    SetItemRects
    
    
    If m_strSelectedKey > "" Then
        pvSizePanel m_EyeDropperItems(m_strSelectedKey)
    End If

    pvDrawControl

End Property


Public Property Get SelectedForeColor() As OLE_COLOR
    SelectedForeColor = m_SelectedForeColor
End Property

Public Property Let SelectedForeColor(ByVal New_SelectedForeColor As OLE_COLOR)
    m_SelectedForeColor = New_SelectedForeColor
    PropertyChanged "SelectedForeColor"
End Property

Public Property Get HoverForeColor() As OLE_COLOR
    HoverForeColor = m_HoverForeColor
End Property

Public Property Let HoverForeColor(ByVal New_HoverForeColor As OLE_COLOR)
    m_HoverForeColor = New_HoverForeColor
    PropertyChanged "HoverForeColor"
End Property

Public Property Get NormalForeColor() As OLE_COLOR
    NormalForeColor = m_NormalForeColor
End Property

Public Property Let NormalForeColor(ByVal New_NormalForeColor As OLE_COLOR)
    m_NormalForeColor = New_NormalForeColor
    PropertyChanged "NormalForeColor"
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_HeaderForeColor
End Property

Public Property Let HeaderForeColor(ByVal New_HeaderForeColor As OLE_COLOR)
    m_HeaderForeColor = New_HeaderForeColor
    PropertyChanged "HeaderForeColor"
End Property

Public Property Get HeaderFont() As Font
    Set HeaderFont = m_HeaderFont
End Property

Public Property Set HeaderFont(ByVal New_HeaderFont As Font)
    Set m_HeaderFont = New_HeaderFont
    PropertyChanged "HeaderFont"

    With mDC
        Set .Font = m_HeaderFont
        m_lngCaptionHeight = .TextHeight("`,Q") + m_const_lngDefItemHeightOffSet
        If m_lngCaptionHeight < 24 Then m_lngCaptionHeight = 24
    End With

    SetItemRects
    pvDrawControl


End Property

Public Property Get DisplayHeader() As Boolean
    DisplayHeader = m_DisplayHeader
End Property

Public Property Let DisplayHeader(ByVal New_DisplayHeader As Boolean)

    m_DisplayHeader = New_DisplayHeader
    PropertyChanged "DisplayHeader"
    SetItemRects
    pvDrawControl
    pvSizePanel Me.SelectedItem

End Property

Public Property Get RighToLeft() As Boolean
    RighToLeft = m_RighToLeft
End Property

Public Property Let RighToLeft(ByVal New_RighToLeft As Boolean)
    m_RighToLeft = New_RighToLeft
    PropertyChanged "RighToLeft"
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

'//---------------------------------------------------------------------------------------
' Procedure : Redraw
' Type      : Property
' DateTime  : 07/02/2005
' Author    : Gary Noble
' Purpose   : Quicker Drawing Of Control
' Returns   : Boolean
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  07/02/2005
'//---------------------------------------------------------------------------------------
Public Property Let Redraw(ByVal New_Redraw As Boolean)
    m_Redraw = New_Redraw
    PropertyChanged "Redraw"

    If m_Redraw Then

        With mDC
            Set .Font = UserControl.Font
            If .TextHeight("`,Q") > m_lngDefItemHeight Then
                m_lngDefItemHeight = .TextHeight("`,Q") + 5
            End If
        End With

        With mDC
            Set .Font = HoverFont
            If .TextHeight("`,Q") > m_lngDefItemHeight Then
                m_lngDefItemHeight = .TextHeight("`,Q") + 5
            End If
        End With

        With mDC
            Set .Font = m_HeaderFont
            m_lngCaptionHeight = .TextHeight("`,Q") + m_const_lngDefItemHeightOffSet
            If m_lngCaptionHeight < 24 Then m_lngCaptionHeight = 24
        End With

        UserControl_Resize
        SetItemRects
        GetNextSelectedItem
        Me.VisibleItems = Me.VisibleItems
        pvDrawControl
    End If

End Property

Public Property Get Version() As String
    Version = m_Version
End Property
Public Property Let Version(ByVal sVersion As String)
    m_Version = "v" & App.Major & "." & App.Minor & App.Revision
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get UseCustomColor() As Boolean
    UseCustomColor = m_UseCustomColor

End Property

Public Property Let UseCustomColor(ByVal New_UseCustomColor As Boolean)
    m_UseCustomColor = New_UseCustomColor
    PropertyChanged "UseCustomColor"
    pvGetGradientColors
    pvDrawControl
End Property

