VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_MenuHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
'//--Module    : IAPP_MenuHandler
'//--DateTime  : 04/08/2004
'//--Author    : Gary Noble   ©2005 Telecom Direct Limited
'//--Purpose   : Custom Drawn Menu Handler
'//--Assumes   : IAPP_StoreMenu
'//--Notes     :
'//--Revision  : 1.2
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Option Explicit
Private Const BF_BOTTOM As Long = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_RIGHT As Long = &H4
Private Const BF_TOP As Long = &H2
Private Const BF_RECT As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2
Private Const DT_NOCLIP As Long = &H100
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const ODT_MENU As Integer = 1
Private Const ODS_SELECTED As Long = &H1
Private Const ODS_DISABLED As Long = &H4
Private Const ODS_CHECKED As Long = &H8
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_TYPE As Long = &H10
Private Const MF_ItemOwnerDraw As Long = &H100
Private Const MF_SEPARATOR As Long = &H800
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_ERASEBKGND As Long = &H14
Private Const WM_INITMENUPOPUP As Long = &H117
Private Const WM_COMMAND As Long = &H111
Private Const WM_MENUCHAR As Long = &H120
Private Const MF_MENUBREAK As Long = &H40
Private Const MF_CHECKED As Long = &H8
Private Const MF_DEFAULT As Long = &H1000
Private Const MFS_CHECKED As Long = MF_CHECKED
Private Const MFS_GRAYED As Long = &H3
Private Const MFS_DISABLED As Long = MFS_GRAYED
Private Const MF_BYCOMMAND As Long = &H0
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MF_SYSMENU As Long = &H2000
Private Const MIIM_DATA As Long = &H20
Private Const m_RightIndent As Long = 9
Private Const m_RightIndentXP As Long = 20
Private Const m_IconSize As Long = 21
Private Const m_IconSizeXP As Long = 24
Private Const m_CaptionIndent As Long = 2
Private Const m_CaptionIndentXP As Long = 8
Private Const m_SubTriSize As Long = 15
'//------ Enums
Public Enum EMenuCheckedStyle
    EMCS_Normal = 0
    EMCS_Radio = 1
    EMCS_Icon = 2
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
Private EMCS_Normal, EMCS_Radio, EMCS_Icon
#End If
Public Enum MenuDrawStyle
    mds_3D = 0
    mds_XP = 1
    'EMCS_Icon = 2
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
Private mds_3D, mds_XP
#End If
Public Enum SpecialColors
    SC_XPBack = 16251903
    SC_XPBorder = 7021576
    SC_XPHighlight = 14073525
    SC_XPFill = 14604246
    SC_XPShadow = 10260108
    SC_XPMargin = 14080990
    SC_XPDarkFill = 11899524
    SC_XPDisabled = 10856101
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
Private SC_XPBack, SC_XPBorder, SC_XPHighlight, SC_XPFill, SC_XPShadow, SC_XPMargin, SC_XPDarkFill, SC_XPDisabled
#End If
Public Enum MenuOwnerDrawFlags
    modf_Highlighted = &H0
    modf_Checked = &H1
    modf_Enabled = &H2
    modf_TopMenu = &H3
    modf_Clicked = &H4
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
Private modf_Highlighted, modf_Checked, modf_Enabled, modf_TopMenu, modf_Clicked
#End If
Public Enum ShowPopupMenuConstants
    '//--Track popup menu constants:
    TPM_CENTERALIGN = &H4&
    TPM_lEFTALIGN = &H0&
    TPM_LEFTBUTTON = &H0&
    TPM_RIGHTALIGN = &H8&
    TPM_RIGHTBUTTON = &H2&
    TPM_HORIZONTAL = &H0                                   '/* Horz alignment matters more */
    TPM_VERTICAL = &H40                                    '/* Vert alignment matters more */
    '//--Win98/2000 menu animation and menu within menu options:
    TPM_RECURSE = &H1&
    TPM_HORPOSANIMATION = &H400&
    TPM_HORNEGANIMATION = &H800&
    TPM_VERPOSANIMATION = &H1000&
    TPM_VERNEGANIMATION = &H2000&
    '//--Win2000 only:
    TPM_NOANIMATION = &H4000&
End Enum
#If False Then    'Trick preserves Case of Enums when typing in IDE
Private TPM_CENTERALIGN, TPM_lEFTALIGN, TPM_LEFTBUTTON, TPM_RIGHTALIGN, TPM_RIGHTBUTTON, TPM_HORIZONTAL, TPM_VERTICAL, TPM_RECURSE
Private TPM_HORPOSANIMATION, TPM_HORNEGANIMATION, TPM_VERPOSANIMATION, TPM_VERNEGANIMATION, TPM_NOANIMATION
#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    Itemid As Long
    itemAction As Long
    itemState As Long
    hWndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    Itemid As Long
    ItemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private m_SubClass As IAPP_Subclass
Implements iSubclass
Private m_VBMenus As Collection
Private m_CollMenus As Collection
Private m_MemDC As IAPP_MemDC
Private m_TmpMemDC As IAPP_MemDC
Private m_hMenu As Long
Private m_hWnd As Long
Private m_MenusCount As Long
Private m_ItemHeight As Long
Private m_ImageList As Object
Private m_CurrentMenuIndex As Long
Private m_hSystemMenu As Long
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_HighlightColor As OLE_COLOR
Private m_HighlightForeColor As OLE_COLOR
Private m_Font As StdFont
Private m_ItemOwnerDrawAll As Boolean
Private m_BackGroundPicture As StdPicture
Private m_BackGroundWidth As Long
Private m_BackGroundHeight As Long
Private m_DrawStyle As MenuDrawStyle
Private m_TopMenus As Boolean
Private m_DrawBackGround As Boolean
Private m_RightToLeft As Boolean
Private m_LastID As Long
Private m_IsCreated As Boolean
Private m_FromForm As Boolean
Private m_AllSupportedKeys As String
Private m_hWndKeyOwner As Long
Private m_ConsumeKeys As Boolean

'//--  Events

Public Event MenuHighlight(ByVal Index As Long)        ', ByVal IsSeparator As Boolean)
Public Event MenuExit()
Public Event DrawItem(ByRef Cancel As Boolean, ByVal Index As Long, ByVal hdc As Long, ByVal bSelected As Boolean, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Public Event MeasureItem(ByRef Cancel As Boolean, ByVal Index As Long, ByRef rWidth As Long, ByRef rHeight As Long)
Public Event Click(ByVal Index As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal uItem As Long, _
                                                                                ByVal byPosition As Long, _
                                                                                lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
                                                                                ByVal un As Long, _
                                                                                ByVal bool As Boolean, _
                                                                                lpcMenuItemInfo As Any) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, _
                                                  ByVal nPos As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                                                     ByVal bRevert As Long) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, _
                                                                              ByVal un As Long, _
                                                                              ByVal bool As Boolean, _
                                                                              lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, _
                                                      ByVal wFlags As Long, _
                                                      ByVal X As Long, _
                                                      ByVal Y As Long, _
                                                      ByVal nReserved As Long, _
                                                      ByVal hwnd As Long, _
                                                      lprc As RECT) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
                                                  ByVal nPosition As Long, _
                                                  ByVal wFlags As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hMenu As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

'//---------------------------------------------------------------------------------------
'//--Procedure : AddItem
'//--Type      : Function
'//--DateTime  : 04/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Adds A Dynamic Menu Item
'//--Returns   : Long
'//--Notes     : The Menu Item Relates To The Parent Inde
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Function AddItem(ByVal lngParent As Long, _
                        Optional ByVal strCaption As String, _
                        Optional KeyAccel, _
                        Optional Before, _
                        Optional Key, _
                        Optional ByVal blnEnabled As Boolean = True, _
                        Optional ByVal blnChecked As Boolean, _
                        Optional CheckedStyle As EMenuCheckedStyle, _
                        Optional stdImage As StdPicture, _
                        Optional ByVal HelpText, _
                        Optional ByVal OwnerDraw As Boolean, _
                        Optional ByVal OwnerDrawID, _
                        Optional ByVal OwnerDrawData, _
                        Optional ByVal Break As Boolean, _
                        Optional ByVal Description) As Long

'add

    Dim tMParent As MENUITEMINFO
    Dim tMNew As MENUITEMINFO

    'Dim cIndex As Long
    Dim currMenu As IAPP_StoreMenu
    Dim NewMenu As IAPP_StoreMenu
    Dim parentMenu As IAPP_StoreMenu
    'Dim lR As Long
    Dim hMenu As Long
    Dim bTopMenu As Boolean
    Dim NewIndex As Long
    Dim IsSep As Boolean
    Dim tstIndex As Long
    Dim tstMenu As IAPP_StoreMenu
    Dim iCount As Long
    On Error Resume Next
    IsSep = (Trim$(strCaption) = "-")
    If lngParent = 0 Then
        'top menu
        hMenu = m_hMenu
        If m_FromForm Then
            bTopMenu = True
        End If
    Else
        Set currMenu = m_CollMenus.Item(lngParent)
        If (currMenu Is Nothing) Then
            Exit Function
        End If
        tMParent.fMask = MIIM_SUBMENU
        tMParent.cbSize = LenB(tMParent)
        GetMenuItemInfo currMenu.hMenu, currMenu.Index, True, tMParent
        If tMParent.hSubMenu = 0 Then
            hMenu = CreatePopupMenu()
            tMParent.hSubMenu = hMenu
            Set parentMenu = m_CollMenus.Item(lngParent)
            parentMenu.hSubMenu = hMenu
            Set parentMenu = Nothing
        Else
            hMenu = tMParent.hSubMenu
        End If
        '//--Now set the parent item so it has a popup menu:
        tMParent.fMask = MIIM_SUBMENU
        tMParent.cbSize = LenB(tMParent)
        'tMParent.hSubMenu = hMenu
        SetMenuItemInfo currMenu.hMenu, currMenu.Index, True, tMParent
    End If
    If hMenu <> 0 Then
        With tMNew
            .cbSize = Len(tMNew)
            .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_DATA Or MIIM_STATE
            .wID = pvFreeID
        End With
        tMNew.fType = tMNew.fType Or MF_ItemOwnerDraw
        If IsSep Then
            tMNew.fType = tMNew.fType Or MF_SEPARATOR
        End If
        If blnChecked Then
            tMNew.fState = tMNew.fState Or MFS_CHECKED
        End If
        If Not blnEnabled Then
            tMNew.fState = tMNew.fState Or MFS_DISABLED
        End If
        If Break Then
            tMNew.fType = tMNew.fType Or MF_MENUBREAK
        End If
        Set NewMenu = New IAPP_StoreMenu
        With NewMenu
            .ID = tMNew.wID
            If Not bTopMenu Then
                .ParentID = currMenu.ID
            End If
            .hMenu = hMenu
            .TopMenu = bTopMenu
            '.Caption = Caption
            '//--           .Description = Description
            .Separator = IsSep
            If Not (IsMissing(Key)) Then
                .Name = CStr(Key)
                .vKey = Key
            End If
            .CheckedStyle = CheckedStyle
            NewIndex = -1
            If Not (IsMissing(Before)) Then
                tstIndex = IndexForKey(Before)
                If tstIndex > 0 Then
                    Set tstMenu = m_CollMenus.Item(tstIndex)
                    If (Not (tstMenu Is Nothing)) Then
                        If tstMenu.hMenu = hMenu Then
                            pvMoveIndex tstMenu.hMenu, .Index
                            NewIndex = .Index
                        End If
                    End If
                End If
            End If
            'Debug.Print Caption
            If NewIndex = -1 Then
                iCount = GetMenuItemCount(hMenu)
                .Index = iCount    '+ 1
            Else
                .Index = NewIndex
            End If
            .OwnerDraw = OwnerDraw
            If Not (IsMissing(HelpText)) Then
                .HelpText = CStr(HelpText)
            End If
            If IsMissing(Description) Then
                .Description = .HelpText
            Else
                .Description = CStr(Description)
            End If
            If Not (IsMissing(stdImage)) Then
                Set .Image = stdImage
            End If
            If Not (IsMissing(OwnerDrawID)) Then
                .OwnerDrawID = CLng(OwnerDrawID)
            End If
            If Not (IsMissing(OwnerDrawData)) Then
                .OwnerDrawData = CLng(OwnerDrawData)
            End If
            m_CollMenus.Add NewMenu, "#" & .ID
            pvSetCaption m_CollMenus.Count, strCaption
            If Not (IsMissing(KeyAccel)) Then
                pvSetKeyAccel m_CollMenus.Count, CStr(KeyAccel)
            End If
            .vKey = Key
        End With
        Set NewMenu = Nothing
        If InsertMenuItem(hMenu, NewIndex, True, tMNew) > 0 Then
            AddItem = m_CollMenus.Count
        End If
    End If
    Set currMenu = Nothing
    If bTopMenu Then
        DrawMenuBar m_hWnd
    End If
    On Error GoTo 0

End Function

'//-- Get\Set the BackColor of menus
Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)

    m_BackColor = vNewValue

End Property

'//-- Get the backgorund of menus
Public Property Get BackGroundPicture() As StdPicture

    Set BackGroundPicture = m_BackGroundPicture

End Property

Public Property Let BackGroundPicture(ByVal vNewValue As StdPicture)

    Set m_BackGroundPicture = vNewValue
    If (m_BackGroundPicture Is Nothing) Then
        m_DrawBackGround = False
    Else
        With m_TmpMemDC
            m_BackGroundWidth = .HM2Pix(m_BackGroundPicture.Width)
            m_BackGroundHeight = .HM2Pix(m_BackGroundPicture.Height)
        End With
        m_DrawBackGround = True
    End If

End Property



'//--  Class Events

Private Sub Class_Initialize()

    Set m_SubClass = New IAPP_Subclass
    Set m_VBMenus = New Collection
    Set m_CollMenus = New Collection
    Set m_MemDC = New IAPP_MemDC
    Set m_TmpMemDC = New IAPP_MemDC
    '//----
    m_BackColor = vbMenuBar
    m_ForeColor = vbMenuText
    m_HighlightColor = vbHighlight
    m_HighlightForeColor = vbWhite
    With m_TmpMemDC
        .Init 100, 100
        Set m_Font = .SystemMenuFont
        Set .Font = m_Font
        pvCalculateItemHeight
    End With
    'hook keyboard
    HookKeyboard Me
    m_ConsumeKeys = True

End Sub

Private Sub Class_Terminate()

    Dim I As Long

    For I = 1 To m_CollMenus.Count
        DestroyMenu m_CollMenus.Item(I).hMenu
    Next I
    m_SubClass.UnSubclass
    Set m_SubClass = Nothing
    Set m_VBMenus = Nothing
    Set m_CollMenus = Nothing
    Set m_MemDC = Nothing
    Set m_TmpMemDC = Nothing
    Set m_ImageList = Nothing
    Set m_Font = Nothing
    Set m_BackGroundPicture = Nothing
    m_AllSupportedKeys = vbNullString
    RemoveHookKeyboard Me

End Sub

'//-- Get\Set
Public Property Get ConsumeKeys() As Boolean

    ConsumeKeys = m_ConsumeKeys

End Property

Public Property Let ConsumeKeys(ByVal vNewValue As Boolean)

    m_ConsumeKeys = vNewValue

End Property

'//---------------------------------------------------------------------------------------
'//--Procedure : CreateFromNothing
'//--Type      : Function
'//--DateTime  : 04/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Creates A New Popup Menu
'//--Returns   : Boolean
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Public Function CreateFromNothing(Optional ByVal hwnd) As Boolean


    If m_IsCreated Then
        Exit Function
    End If
    m_hMenu = CreatePopupMenu()
    If Not (IsMissing(hwnd)) Then
        m_hWnd = CLng(hwnd)
        pvSubClass
        m_IsCreated = True
    End If

End Function


'//--  Properties

'//-- Move pointer to a menu so that we can
'call properties without telling the index
Public Property Get CurrentMenuIndex() As Long

    CurrentMenuIndex = m_CurrentMenuIndex

End Property

Public Property Let CurrentMenuIndex(ByVal vNewValue As Long)

    m_CurrentMenuIndex = vNewValue

End Property

'//-- Get\Set the Style of menus
Public Property Get DrawStyle() As MenuDrawStyle

    DrawStyle = m_DrawStyle

End Property

Public Property Let DrawStyle(ByVal vNewValue As MenuDrawStyle)

    If m_DrawStyle <> vNewValue Then
        m_DrawStyle = vNewValue
        pvCalculateItemHeight
        pvReplaceAllItems
    End If

End Property

'//-- Get\Set the Font of menus
Public Property Get Font() As StdFont

    Set Font = m_Font

End Property

Public Property Set Font(ByVal vNewValue As StdFont)

    With m_TmpMemDC
        If (vNewValue Is Nothing) Then
            Set m_Font = .SystemMenuFont
        ElseIf (vNewValue = m_Font) Then
            vNewValue.Bold = False
            Exit Property
        Else
            vNewValue.Bold = False
            Set m_Font = vNewValue
        End If
        Set .Font = m_Font
        pvCalculateItemHeight
        pvReplaceAllItems
        If m_FromForm Then
            DrawMenuBar m_hWnd
        End If
    End With

End Property

'//-- Get\Set the ForeColor of menus
Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)

    m_ForeColor = vNewValue

End Property

Public Property Get handle() As Long

    handle = m_hMenu

End Property

'//-- Just to store the HELP-TEXT for an item
Public Property Get HelpText(ByVal Index) As String


    Dim ItemClass As IAPP_StoreMenu

    On Error Resume Next
    Set ItemClass = m_CollMenus.Item(Index)
    If (ItemClass Is Nothing) Then
        Exit Property
    End If
    HelpText = ItemClass.HelpText
    On Error GoTo 0

End Property

Public Property Let HelpText(ByVal Index, _
                             ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu

    On Error Resume Next
    Set ItemClass = m_CollMenus.Item(Index)
    If (ItemClass Is Nothing) Then
        Exit Property
    End If
    ItemClass.HelpText = vNewValue
    On Error GoTo 0

End Property

'//-- Get\Set the HighlightColor of menus
Public Property Get HighlightColor() As OLE_COLOR

    HighlightColor = m_HighlightColor

End Property

Public Property Let HighlightColor(ByVal vNewValue As OLE_COLOR)

    m_HighlightColor = vNewValue

End Property

'//-- Get\Set the HighlightForeColor of menus
Public Property Get HighlightForeColor() As OLE_COLOR

    HighlightColor = m_HighlightForeColor

End Property

Public Property Let HighlightForeColor(ByVal vNewValue As OLE_COLOR)

    m_HighlightForeColor = vNewValue

End Property

Private Sub HLSToRGB(ByVal h As Single, _
                     ByVal S As Single, _
                     ByVal l As Single, _
                     R As Long, _
                     G As Long, _
                     B As Long)

    Dim rR As Single
    Dim rG As Single
    Dim rB As Single

    Dim Min As Single
    Dim Max As Single
    If S = 0 Then
        '//--Achromatic case:
        rR = l
        rG = l
        rB = l
    Else
        '//--Chromatic case:
        '//--delta = Max-Min
        If l <= 0.5 Then
            's = (Max - Min) / (Max + Min)
            '//--Get Min value:
            Min = l * (1 - S)
        Else
            's = (Max - Min) / (2 - Max - Min)
            '//--Get Min value:
            Min = l - S * (1 - l)
        End If
        '//--Get the Max value:
        Max = 2 * l - Min
        '//--Now depending on sector we can evaluate the h,l,s:
        If h < 1 Then
            rR = Max
            If h < 0 Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If h < 2 Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If h < 4 Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255
    G = rG * 255
    B = rB * 255

End Sub

Public Property Get hWndKeyOwner() As Long

    hWndKeyOwner = m_hWndKeyOwner

End Property

Public Property Let hWndKeyOwner(ByVal vNewValue As Long)

    m_hWndKeyOwner = vNewValue

End Property

'//-- Source of imaged
Public Property Get ImageList() As Object

    Set ImageList = m_ImageList

End Property

Public Property Set ImageList(ByVal vNewValue As Object)

    Set m_ImageList = vNewValue

End Property

Public Function IndexForID(ByVal ID As Long) As Long

    Dim I As Long

    For I = 1 To m_CollMenus.Count
        If (m_CollMenus.Item(I).ID) = ID Then
            IndexForID = I
            Exit For
        End If
    Next I

End Function

'//-- Get the menu index of a menu name
Public Function IndexForKey(Key) As Long


    Dim I As Long

    If IsMissing(Key) Then
        IndexForKey = m_CurrentMenuIndex
    Else
        If IsNumeric(Key) Then
            IndexForKey = CLng(Key)
            If (IndexForKey < 1 Or IndexForKey > m_CollMenus.Count) Then
                IndexForKey = 0
            End If
        Else
            For I = 1 To m_CollMenus.Count
                If StrComp(Key, m_CollMenus.Item(I).Name, vbTextCompare) = 0 Then
                    IndexForKey = I
                End If
            Next I
        End If
    End If

End Function

Public Function IndexForKeyAccel(ByVal KeyAccel As String) As Long

    Dim I As Long

    If Not (pvKeyAccelExists(KeyAccel)) Then
        Exit Function
    End If
    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If (InStr(1, "|" & .KeyAccel & "|", "|" & KeyAccel & "|")) Then
                IndexForKeyAccel = I
                Exit For
            End If
        End With
    Next I

End Function

Private Sub iSubclass_After(lReturn As Long, _
                            ByVal hwnd As Long, _
                            ByVal uMsg As WinSubHook.eMsg, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long)

'


End Sub

Private Sub iSubclass_Before(bHandled As Boolean, _
                             lReturn As Long, _
                             hwnd As Long, _
                             uMsg As WinSubHook.eMsg, _
                             wParam As Long, _
                             lParam As Long)

    On Error Resume Next
    Select Case uMsg
        Case WM_MEASUREITEM
            If (pvMeasureItem(wParam, lParam)) Then
                bHandled = True
            End If
        Case WM_DRAWITEM

            'Debug.Print "DrawItem"
            If (pvDrawItem(wParam, lParam)) Then
                bHandled = True
            End If
        Case WM_ERASEBKGND

            'Debug.Print "TILE"
            bHandled = True
        Case WM_INITMENUPOPUP

            pvInitMenuPopUp wParam
        Case WM_MENUSELECT

            pvMenuSelect wParam, lParam
        Case WM_COMMAND

            If (pvCommand(wParam, lParam)) Then
                bHandled = True
            End If
        Case WM_MENUCHAR

            If (((wParam \ &H10000) And MF_SYSMENU) <> MF_SYSMENU) Then
                lReturn = pvMenuChar(lParam, wParam And &HFFFF&)
                bHandled = True
            End If
    End Select
    On Error GoTo 0

End Sub


'//--  SubClassing

Private Sub ISuperClass_After(ByVal lReturn As Long, _
                              ByVal lngHwnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long)

'hold place


End Sub

Private Sub ISuperClass_Before(ByVal lHandled As Long, _
                               ByVal lReturn As Long, _
                               ByVal lngHwnd As Long, _
                               ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long)



End Sub

'//-- Get\Set if the item is a menu-break
Public Property Get ItemBreak(Optional Index) As Boolean


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        ItemBreak = ((MII.fType And MF_MENUBREAK) = MF_MENUBREAK)
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemBreak(Optional Index, _
                              ByVal vNewValue As Boolean)


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Break = vNewValue
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        If vNewValue Then
            MII.fType = MII.fType Or MF_MENUBREAK
        Else
            MII.fType = MII.fType And Not MF_MENUBREAK
        End If
        SetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        Set ItemClass = Nothing
    End If

End Property

'//-- Get\Set caption of an item
Public Property Get ItemCaption(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemCaption = ItemClass.Caption
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemCaption(Optional Index, _
                                ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.Caption = vNewValue
        pvSetCaption cIndex, vNewValue
        'remove item - insert again
        ReplaceItem cIndex
        Set ItemClass = Nothing
    End If

End Property

'//-- Return If Item Is Checked
Public Property Get ItemChecked(Optional Index) As Boolean


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        ItemChecked = ((MII.fState And MFS_CHECKED) = MFS_CHECKED)
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemChecked(Optional Index, _
                                ByVal vNewValue As Boolean)


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.X = vNewValue
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        If vNewValue Then
            MII.fState = MII.fState Or MFS_CHECKED
        Else
            MII.fState = MII.fState And Not MFS_CHECKED
        End If
        SetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        Set ItemClass = Nothing
    End If

End Property

'//-- Get\Set the checked style of an item
Public Property Get ItemCheckedStyle(Optional Index) As EMenuCheckedStyle


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    On Error Resume Next
    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemCheckedStyle = ItemClass.CheckedStyle
        Set ItemClass = Nothing
    End If
    On Error GoTo 0

End Property

Public Property Let ItemCheckedStyle(Optional Index, _
                                     ByVal vNewValue As EMenuCheckedStyle)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    On Error Resume Next
    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.CheckedStyle = vNewValue
        Set ItemClass = Nothing
    End If
    On Error GoTo 0

End Property

Public Property Get ItemCount() As Long

    ItemCount = m_CollMenus.Count

End Property

Public Property Get ItemDescription(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemDescription = ItemClass.Description
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemDescription(Optional Index, _
                                    ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Description = vNewValue
        Set ItemClass = Nothing
    End If

End Property

Public Property Get ItemEnabled(Optional Index) As Boolean


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        ItemEnabled = Not ((MII.fState And MFS_DISABLED) = MFS_DISABLED)
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemEnabled(Optional Index, _
                                ByVal vNewValue As Boolean)


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.X = vNewValue
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        If vNewValue Then
            MII.fState = MII.fState And Not MFS_DISABLED
        Else
            MII.fState = MII.fState Or MFS_DISABLED
        End If
        SetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        If ItemClass.TopMenu Then
            DrawMenuBar m_hWnd
        End If
        Set ItemClass = Nothing
    End If

End Property

'//-- Set\Get the help text for a menu
Public Property Get ItemHelpText(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemHelpText = ItemClass.HelpText
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemHelpText(Optional Index, _
                                 ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.HelpText = vNewValue
        Set ItemClass = Nothing
    End If

End Property

Public Property Get Itemid(Optional Index) As Long


    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Itemid = m_CollMenus.Item(cIndex).ID
    End If

End Property

'//-- Set\Clear an image for a menu
Public Property Get ItemImage(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemImage = ItemClass.Image
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemImage(Optional Index, _
                              ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Image = vNewValue
        Set ItemClass = Nothing
    End If

End Property

Public Property Get ItemKey(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemKey = ItemClass.vKey
        Set ItemClass = Nothing
    End If

End Property

'//-- Get\Set the shortcut of a menu
'this is a great property not included in Others classes
Public Property Get ItemKeyAccel(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemKeyAccel = ItemClass.KeyAccel
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemKeyAccel(Optional Index, _
                                 ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'ItemClass.KeyAccel = vNewValue
        If Not (ItemClass.TopMenu) Then
            If (StrComp(ItemClass.KeyAccel, vNewValue) <> 0) Then
                If pvSetKeyAccel(cIndex, vNewValue) Then
                    'remove item - insert again
                    ReplaceItem cIndex
                End If
            End If
        End If
        Set ItemClass = Nothing
    End If

End Property

'//-- Get\Set if the item is ItemOwnerDraw by you
'if true the event MeasureItem & DrawItem
'will be called before processing
Public Property Get ItemOwnerDraw(Optional Index) As Boolean


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDraw = ItemClass.OwnerDraw
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemOwnerDraw(Optional Index, _
                                  ByVal vNewValue As Boolean)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        If ItemClass.OwnerDraw <> vNewValue Then
            ItemClass.OwnerDraw = vNewValue
            'remove item - insert again
            ReplaceItem cIndex
        End If
        Set ItemClass = Nothing
    End If

End Property

'//-- Get the menu index of a menu name
Public Property Get ItemOwnerDrawAll() As Boolean

    ItemOwnerDrawAll = m_ItemOwnerDrawAll

End Property

Public Property Let ItemOwnerDrawAll(ByVal vNewValue As Boolean)

    m_ItemOwnerDrawAll = vNewValue

End Property

'//-- Just to store the number value which will
'help you to draw your item
Public Property Get ItemOwnerDrawData(Optional Index) As Long


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDrawData = ItemClass.OwnerDrawData
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemOwnerDrawData(Optional Index, _
                                      ByVal vNewValue As Long)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.OwnerDrawData = vNewValue
        Set ItemClass = Nothing
    End If

End Property

'//-- Just to store a number which descripe the
'way you want to draw your item
Public Property Get ItemOwnerDrawID(Optional Index) As Long


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemOwnerDrawID = ItemClass.OwnerDrawID
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemOwnerDrawID(Optional Index, _
                                    ByVal vNewValue As Long)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.OwnerDrawID = vNewValue
        Set ItemClass = Nothing
    End If

End Property

Public Property Get ItemParentID(Optional Index) As Long


    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        ItemParentID = m_CollMenus.Item(cIndex).ParentID
    End If

End Property

Public Property Get ItemParentIndex(Optional Index) As Boolean


    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        ItemParentIndex = IndexForID(m_CollMenus.Item(cIndex).ParentID)
    End If

End Property

'//-- Get\Set if the item is a menu-break
Public Property Get ItemSeparator(Optional Index) As Boolean


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        'apply
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
        ItemSeparator = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemSeparator(Optional Index, _
                                  ByVal vNewValue As Boolean)


    Dim MII As MENUITEMINFO
    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long
    Dim sBuffer As String * 80

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        If ItemClass.Separator <> vNewValue Then
            ItemClass.Separator = vNewValue
            'apply
            With MII
                .cbSize = Len(MII)
                .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
                .fType = 0
                .dwTypeData = sBuffer
                .cch = 80
            End With
            GetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
            If vNewValue Then
                MII.fType = MII.fType Or MF_SEPARATOR
            Else
                MII.fType = MII.fType And Not MF_SEPARATOR
            End If
            SetMenuItemInfo ItemClass.hMenu, ItemClass.Index, True, MII
            ReplaceItem cIndex
        End If
        Set ItemClass = Nothing
    End If

End Property

Public Property Get ItemTag(Optional Index) As String


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemTag = ItemClass.Tag
        Set ItemClass = Nothing
    End If

End Property

Public Property Let ItemTag(Optional Index, _
                            ByVal vNewValue As String)


    Dim ItemClass As IAPP_StoreMenu
    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set ItemClass = m_CollMenus.Item(cIndex)
        If (ItemClass Is Nothing) Then
            Exit Property
        End If
        ItemClass.Tag = vNewValue
        Set ItemClass = Nothing
    End If

End Property

Public Property Get ItemTopMenu(Optional Index) As Boolean


    Dim cIndex As Long

    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        ItemTopMenu = m_CollMenus.Item(cIndex).TopMenu
    End If

End Property

Public Function KeyAccelPressed(ByVal lStr As String) As Boolean

    Dim cIndex As Long

    If GetActiveWindow() = hWndKeyOwner Then
        cIndex = IndexForKeyAccel(lStr)
        If cIndex <> 0 Then
            If ItemEnabled(cIndex) Then
                If Not (ItemSeparator(cIndex)) Then
                    RaiseEvent Click(cIndex)
                    KeyAccelPressed = True
                End If
            End If
        End If
    End If

End Function

Private Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long

    Dim lC As Long
    Dim h As Single
    Dim S As Single
    Dim l As Single

    Dim lR As Long
    Dim lG As Long
    Dim lB As Long
    Static s_lColLast As Long
    Static s_lLightColLast As Long
    lC = m_TmpMemDC.TranslateColor(oColor)
    If lC <> s_lColLast Then
        s_lColLast = lC
        RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, S, l
        If l > 0.99 Then
            l = l * 0.8
        Else
            l = l * 1.2
            If l > 1 Then
                l = 1
            End If
        End If
        HLSToRGB h, S, l, lR, lG, lB
        s_lLightColLast = RGB(lR, lG, lB)
    End If
    LighterColour = s_lLightColLast

End Property

Private Function Maximum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single

    If rR > rG Then
        If rR > rB Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If rB > rG Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If

End Function

Friend Sub MenuClicked(ByVal MenuName As String)

    RaiseEvent Click(MenuName)

End Sub

Private Function Minimum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single

    If rR < rG Then
        If rR < rB Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If rB < rG Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If

End Function

Public Function PopUpMenu(Optional Index, _
                          Optional X, _
                          Optional Y, _
                          Optional ByVal Options As ShowPopupMenuConstants = TPM_lEFTALIGN Or TPM_HORIZONTAL) As Long


    Dim PO As POINTAPI
    Dim CPO As POINTAPI
    Dim CPB As Boolean
    Dim hMenu As Long
    Dim currMenu As IAPP_StoreMenu
    Dim cIndex As Long
    Dim rID As Long
    Dim tRect As RECT
    Const TPM_RETURNCMD As Long = &H100

    If IsMissing(Index) Then
        hMenu = m_hMenu
    Else
        cIndex = IndexForKey(Index)
        If cIndex < 1 Then
            Exit Function
        End If
    End If
    If IsMissing(X) Then
        GetCursorPos CPO
        PO.X = CPO.X
        CPB = True
    Else
        PO.X = CLng(X)
    End If
    If IsMissing(Y) Then
        If Not (CPB) Then
            GetCursorPos CPO
        End If
        PO.Y = CLng(CPO.Y)
    Else
        PO.Y = CLng(Y)
    End If
    If hMenu = 0 Then
        Set currMenu = m_CollMenus.Item(cIndex)
        hMenu = currMenu.hSubMenu
    End If
    If hMenu <> 0 Then
        Options = Options Or TPM_RETURNCMD
        rID = TrackPopupMenu(hMenu, Options, PO.X, PO.Y, 0, m_hWnd, tRect)
    End If
    'MsgBox rID
    pvRaiseClick , rID
    Set currMenu = Nothing
    PopUpMenu = rID

End Function


Private Sub pvAddSupportedKeys(ByVal lKeys As String)


    If m_AllSupportedKeys = vbNullString Then
        m_AllSupportedKeys = "|" & lKeys & "|"
    Else
        m_AllSupportedKeys = Replace$(m_AllSupportedKeys, "|", "|" & lKeys & "|", , 1)
    End If

End Sub

Private Sub pvCalculateItemHeight()

    If m_DrawStyle = mds_3D Then
        m_ItemHeight = m_TmpMemDC.FontHeight + 6
    ElseIf (m_DrawStyle = mds_XP) Then
        m_ItemHeight = m_TmpMemDC.FontHeight + 9
    End If

End Sub

Private Function pvCheckKey(ByRef lStr As String) As Boolean

    pvCheckKey = True
    'Dim keyArr() As String
    'Dim I As Long
    'Dim repStr As String
    'keyArr = Split(lStr, "+")
    'For I = 0 To UBound(keyArr)
    '
    'Next
    'Erase keyArr

End Function

Private Function pvCheckKeys(ByRef lStr As String) As Boolean

    Dim keyArr() As String
    Dim I As Long
    Dim repStr As String

    keyArr = Split(lStr, "|")
    pvCheckKeys = True
    For I = 0 To UBound(keyArr)
        If pvCheckKey(keyArr(I)) Then
            If I = 0 Then
                repStr = keyArr(I)
            Else
                repStr = repStr & "|" & keyArr(I)
            End If
        Else
            pvCheckKeys = False
        End If
    Next I
    Erase keyArr

End Function

Private Function pvCommand(ByVal wParam As Long, _
                           ByVal lParam As Long) As Boolean

'Dim lHiWord    As Long

    Dim lMenuId As Long

    Dim currMenu As IAPP_StoreMenu
    On Error Resume Next
    If lParam = 0 Then
        '//--Low order word of the wParam item is the menu item id:
        lMenuId = (wParam And &HFFFF&)
        Set currMenu = m_CollMenus.Item("#" & lMenuId)
        RaiseEvent Click(IndexForKey(currMenu.Name))
        Set currMenu = Nothing
    End If
    On Error GoTo 0

End Function

Private Function pvCreateShadowPicture(ByVal picThis As StdPicture) As StdPicture


    Dim sX As Long
    Dim sY As Long
    Dim xMemDC As IAPP_MemDC

    Set xMemDC = New IAPP_MemDC
    With xMemDC
        .Init .HM2Pix(picThis.Width), .HM2Pix(picThis.Height)
        .PaintPicture picThis
        For sX = 0 To .Width
            For sY = 0 To .Height
                If (.GetPixel(sX, sY)) Then
                    .SetPixel sX, sY, vbButtonShadow
                End If
            Next sY
        Next sX
        Set pvCreateShadowPicture = .Image
    End With
    Set xMemDC = Nothing

End Function

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawItem
' Type      : Function
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Draw Item  Calling Procedure
' Returns   : Boolean
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Function pvDrawItem(ByVal wParam As Long, _
                            ByVal lParam As Long) As Boolean

    Dim DrawInfo As DRAWITEMSTRUCT

    CopyMemory DrawInfo, ByVal lParam, LenB(DrawInfo)    'get the UDT for the menuitem's appearance
    If DrawInfo.CtlType <> ODT_MENU Then
        Exit Function
    End If
    pvDrawItem = True
    If m_DrawStyle = mds_3D Then
        pvDrawItem3D DrawInfo
    ElseIf (m_DrawStyle = mds_XP) Then
        pvDrawItemXP DrawInfo
    End If

End Function

Private Sub pvDrawItem3D(DrawInfo As DRAWITEMSTRUCT)  'ByVal wParam As Long, ByVal lParam As Long)

'Dim DrawInfo   As DRAWITEMSTRUCT

    Dim HasIcon As Boolean

    ''Dim HasSub     As Boolean
    Dim currMenu As IAPP_StoreMenu
    Dim bCancel As Boolean
    Dim xMemDC As IAPP_MemDC
    Dim cTop As Long
    Dim cLeft As Long
    Dim cText As String
    Dim cAccel As String
    Dim sX As Long
    Dim sY As Long
    Dim ti As Long
    Dim tII As Long
    Dim fLeft As Long
    Dim fRight As Long
    Dim pDisableDrawIcon As Boolean
    Dim cS As EMenuCheckedStyle
    Dim bChar As String
    Dim New_Font As StdFont
    Dim oPic As StdPicture
    Dim IsSelected As Boolean
    Dim IsDisabled As Boolean
    Dim IsChecked As Boolean

    'bail - not 'ItemOwnerDrawn'
    If DrawInfo.CtlType <> ODT_MENU Then
        GoTo Ignore
    End If

    IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED)    'selected ?
    IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED)    'disabled ?
    IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED)    'checked ?
    Set currMenu = m_CollMenus.Item("#" & DrawInfo.Itemid)  'get menu details from the class
    With currMenu
        HasIcon = Not .Image Is Nothing

        If (.OwnerDraw Or m_ItemOwnerDrawAll) Then
            RaiseEvent DrawItem(bCancel, IndexForID(.ID), DrawInfo.hdc, IsSelected, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Right, DrawInfo.rcItem.Bottom)
            If bCancel Then
                GoTo Ignore
            End If
        End If
    End With    'currMenu

    Set xMemDC = New IAPP_MemDC
    cText = currMenu.Caption
    cAccel = pvExtractAccel(currMenu.KeyAccel)
    With xMemDC
        .Init DrawInfo.rcItem.Right - DrawInfo.rcItem.Left, DrawInfo.rcItem.Bottom - DrawInfo.rcItem.Top
        Set .Font = m_Font    '//--.SystemMenuFont
        .BackStyle = BS_TRANSPARENT
        .FillRect 0, 0, .Width, .Height, m_BackColor

        '//-- if top menu
        If currMenu.TopMenu Then
            .FillRect 0, 0, .Width, .Height, vbButtonFace
            If IsDisabled Then
                .ForeColor = vbWhite
                .DrawText currMenu.Caption, 1, 1, .Width + 1, .Height + 1, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
                .ForeColor = vbButtonShadow
                .DrawText currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
            Else
                If ((DrawInfo.itemState And &H1) = &H1) Then
                    'clicked
                    .DrawEdge 0, 0, .Width, .Height, BDR_SUNKENOUTER
                ElseIf ((DrawInfo.itemState And &H40) = &H40) Then
                    'hilghited
                    .DrawEdge 0, 0, .Width, .Height, BDR_RAISEDINNER
                Else
                    'normal
                End If
                .ForeColor = vbBlack
                .DrawText currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
            End If
            GoTo Display
        End If
        '//--------------
        If m_DrawBackGround Then
            sX = -(DrawInfo.rcItem.Left Mod m_BackGroundWidth)
            sY = -(DrawInfo.rcItem.Top Mod m_BackGroundHeight)

            For ti = sX To .Width Step m_BackGroundWidth
                For tII = sY To .Height Step m_BackGroundHeight
                    .PaintPicture m_BackGroundPicture, ti, tII
                Next tII
            Next ti
        End If
        If currMenu.Separator Then
            .DrawLine 0, 4, .Width, 4, vbButtonShadow
            .DrawLine 0, 5, .Width, 5, vbWhite
            GoTo Display
        End If
        '//-- ======= TEXT ==============
        'cLeft = m_IconSize + m_CaptionIndent
        If m_RightToLeft Then
            cLeft = .Width - (m_IconSize + m_CaptionIndent)
        Else
            cLeft = m_IconSize + m_CaptionIndent
        End If
        If IsDisabled Then
            'draw menu disabled
            .BackColor = m_BackColor    '//--vbButtonFace
            .ForeColor = vbWhite
            'caption
            '//---
            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                .DrawText cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
            '//---
            .ForeColor = vbButtonShadow
            If m_RightToLeft Then
                .DrawText cText, -1, cTop - 1, .Width - m_IconSize - m_CaptionIndent - 1, .Height - 1, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft - 1, cTop - 1, .Width = 1, .Height - 1, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    .DrawText cAccel, 8, -1, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 10, -1, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        ElseIf (IsSelected) Then
            'draw menu selected
            .BackColor = m_HighlightColor    '//--vbHighlight
            .ForeColor = m_HighlightForeColor    '//--vbWhite
            If HasIcon Then
                If m_RightToLeft Then
                    fLeft = 0
                    fRight = .Width - m_IconSize    '//--- m_SubTriSize - 1
                Else
                    fLeft = m_IconSize
                    fRight = .Width    '//--- m_SubTriSize - 1
                End If
            Else
                fLeft = 0
                fRight = .Width    '//--- m_SubTriSize - 1
            End If
            .FillRect fLeft, 0, fRight, .Height, m_HighlightColor    '//--vbHighlight)

            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                .DrawText cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        Else
            'draw menu normally
            .BackColor = 0    '//--m_BackColor '//--vbButtonFace
            .ForeColor = m_ForeColor    '//--vbBlack
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                'Call .DrawText(cText, .Width - cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP)
                .DrawText cText, 0, cTop, .Width - m_IconSize - m_CaptionIndent, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    'Call .DrawText(cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER)
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        End If
        '//-- ======= ICON\CHECKED ============
        If IsChecked Then
            If IsSelected Then
                If m_RightToLeft Then
                    .FillRect .Width - (m_IconSize - 1), 0, .Width - 1, .Height, LighterColour(m_BackColor)
                Else
                    .FillRect 1, 0, m_IconSize - 1, .Height, LighterColour(m_BackColor)
                End If
            Else
                If m_RightToLeft Then
                    .FillRect .Width - (m_IconSize - 1), 0, .Width - 1, .Height, vbButtonFace
                Else
                    .FillRect 1, 0, m_IconSize - 1, .Height, vbButtonFace
                End If
            End If
            If m_RightToLeft Then
                .DrawEdge .Width - (m_IconSize - 1), 0, .Width - 1, .Height, BDR_SUNKENOUTER, BF_RECT
            Else
                .DrawEdge 1, 0, m_IconSize - 1, .Height, BDR_SUNKENOUTER, BF_RECT
            End If
        End If
        If IsChecked Then
            cS = currMenu.CheckedStyle
            If cS = EMCS_Normal Then
                bChar = "b"
            ElseIf (cS = EMCS_Radio) Then
                bChar = "h"
            End If
            If bChar <> vbNullString Then
                pDisableDrawIcon = True
                Set New_Font = New StdFont
                With New_Font
                    .Name = "Marlett"
                    .Size = 12
                End With
                Set .Font = New_Font
                .ForeColor = vbBlack
                If m_RightToLeft Then
                    .DrawText bChar, .Width - (m_IconSize - 1), 0, .Width - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
                Else
                    .DrawText bChar, 1, 0, m_IconSize - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
                End If
                Set New_Font = Nothing
            End If
        End If
        If (HasIcon And (Not (pDisableDrawIcon))) Then
            Set oPic = currMenu.Image
            If Not (oPic Is Nothing) Then
                If Not IsDisabled Then
                    If IsSelected Then
                        If Not IsChecked Then
                            If m_RightToLeft Then
                                .DrawEdge .Width - (m_IconSize - 1), 0, .Width - 1, .Height, BDR_RAISEDINNER
                            Else
                                .DrawEdge 1, 0, m_IconSize - 1, .Height, BDR_RAISEDINNER
                            End If
                            'Call .DrawEdge(1, 0, m_IconSize - 1, .Height, BDR_RAISEDINNER)
                        End If
                    End If
                End If
                If IsDisabled Then
                    If m_RightToLeft Then
                        .PaintDisabledPicture oPic, .Width - (16 + 3), (.Height - 16) / 2, 16, 16    '//--2)
                    Else
                        .PaintDisabledPicture oPic, 3, (.Height - 16) / 2, 16, 16  '//--2)
                    End If
                Else
                    If m_RightToLeft Then
                        .PaintPicture oPic, .Width - (16 + 3), (.Height - 16) / 2, 16, 16  '//--2)
                    Else
                        .PaintPicture oPic, 3, (.Height - 16) / 2, 16, 16    '//--2)
                    End If
                End If
                Set oPic = Nothing
            End If
        End If

Display:
        .BitBlt DrawInfo.hdc, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, .Width, .Height, 0, 0
        'Debug.Print "Owner Draw"
    End With
Ignore:
    Set xMemDC = Nothing
    Set currMenu = Nothing

End Sub

'//---------------------------------------------------------------------------------------
' Procedure : pvDrawItemXP
' Type      : Sub
' DateTime  : 04/02/2005
' Author    : Gary Noble
' Purpose   : Draws An XP Style Menu
' Returns   :
' Notes     :
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------
Private Sub pvDrawItemXP(DrawInfo As DRAWITEMSTRUCT)  'ByVal wParam As Long, ByVal lParam As Long)


    Dim pDisableDrawIcon As Boolean
    Dim cS As EMenuCheckedStyle
    Dim bChar As String
    Dim New_Font As StdFont
    Dim oPic As StdPicture
    Dim IsSelected As Boolean
    Dim IsDisabled As Boolean
    Dim IsChecked As Boolean
    Dim HasIcon As Boolean

    ''Dim HasSub     As Boolean
    Dim currMenu As IAPP_StoreMenu
    Dim bCancel As Boolean
    Dim xMemDC As IAPP_MemDC
    Dim cTop As Long
    Dim cLeft As Long
    Dim cText As String
    Dim cAccel As String

    'bail - not 'ItemOwnerDrawn'
    If DrawInfo.CtlType <> ODT_MENU Then
        GoTo Ignore
    End If

    IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED)    'selected ?
    IsDisabled = ((DrawInfo.itemState And ODS_DISABLED) = ODS_DISABLED)    'disabled ?
    IsChecked = ((DrawInfo.itemState And ODS_CHECKED) = ODS_CHECKED)    'checked ?
    Set currMenu = m_CollMenus.Item("#" & DrawInfo.Itemid)  'get menu details from the class

    With currMenu
        HasIcon = Not .Image Is Nothing
        'HasSub = .HasSub

        If (.OwnerDraw Or m_ItemOwnerDrawAll) Then
            RaiseEvent DrawItem(bCancel, IndexForID(.ID), DrawInfo.hdc, IsSelected, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, DrawInfo.rcItem.Right, DrawInfo.rcItem.Bottom)
            If bCancel Then
                GoTo Ignore
            End If
        End If
    End With    'currMenu

    Set xMemDC = New IAPP_MemDC
    cText = currMenu.Caption    '.vbMenuID.Caption
    cAccel = pvExtractAccel(currMenu.KeyAccel)
    With xMemDC
        '//-- Init
        .Init DrawInfo.rcItem.Right - DrawInfo.rcItem.Left, DrawInfo.rcItem.Bottom - DrawInfo.rcItem.Top
        .BackStyle = BS_TRANSPARENT
        Set .Font = m_Font
        '//-- if top menu draw and good bye
        If currMenu.TopMenu Then
            .ForeColor = vbBlack
            If ((DrawInfo.itemState And &H1) = &H1) Then
                'Clicked
                .FillRect 0, 0, .Width, .Height, SC_XPMargin
                .DrawLine 0, 0, .Width, 0, vbButtonShadow
                .DrawLine 0, 0, 0, .Height, vbButtonShadow
                .DrawLine .Width - 1, 0, .Width - 1, .Height, vbButtonShadow
            ElseIf ((DrawInfo.itemState And &H40) = &H40) Then
                'Highlighted
                .Rectangle 0, 0, .Width, .Height, SC_XPHighlight, , SC_XPBorder
            Else
                'Normal
                .FillRect 0, 0, .Width, .Height, vbButtonFace
            End If
            'draw the caption
            .DrawText currMenu.Caption, 0, 0, .Width, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
            'display
            GoTo Display
        End If
        '//--------------
        'draw the background
        .FillRect 0, 0, .Width, .Height, SC_XPBack
        'margin
        If m_RightToLeft Then
            .FillGradient .Width - m_IconSizeXP, 0, .Width, .Height, BlendColor(vbInactiveTitleBar, vbWhite), BlendColor(vbActiveTitleBar, vbInactiveTitleBar, 30), False
        Else
            .FillGradient 0, 0, m_IconSizeXP, .Height, BlendColor(vbInactiveTitleBar, vbWhite), BlendColor(vbActiveTitleBar, vbInactiveTitleBar, 30), False
        End If
        '//-- if is separator draw and good bye
        If currMenu.Separator Then
            If m_RightToLeft Then
                .DrawLine .Width - (m_IconSizeXP + 5), 4, 0, 4, vbButtonShadow
            Else
                .DrawLine m_IconSizeXP + 5, 4, .Width, 4, vbButtonShadow
            End If
            'Call .DrawLine(0, 5, .Width, 5, vbWhite)
            GoTo Display
        End If
        '//-- ======= TEXT ==============
        If m_RightToLeft Then
            cLeft = .Width - (m_IconSizeXP + m_CaptionIndentXP)
        Else
            cLeft = m_IconSizeXP + m_CaptionIndentXP
        End If
        If IsDisabled Then
            'draw menu disabled
            .ForeColor = vbButtonShadow
            'caption
            '//---
            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                .DrawText cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        ElseIf (IsSelected) Then
            'draw menu selected
            .ForeColor = vbBlack    '//--vbWhite

            .Rectangle 1, 0, .Width - 1, .Height, BlendColor(RGB(248, 216, 126), vbWhite), , vbActiveTitleBar
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                .DrawText cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        Else
            'draw menu normally
            .ForeColor = vbBlack    '//--vbBlack
            'caption
            cTop = (.Height - .TextHeight(cText)) / 2
            If m_RightToLeft Then
                .DrawText cText, 0, cTop, .Width - m_IconSizeXP - m_CaptionIndentXP, .Height, DT_RIGHT Or DT_NOCLIP
            Else
                .DrawText cText, cLeft, cTop, .Width, .Height, DT_LEFT Or DT_NOCLIP
            End If
            'accel
            If LenB(cAccel) Then
                If m_RightToLeft Then
                    .DrawText cAccel, 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                Else
                    .DrawText cAccel, .Width - .TextWidth(cAccel) - 9, 0, .Width, .Height, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
                End If
            End If
        End If
        '//-- ======= ICON\CHECKED ============
        If IsChecked Then
            If Not IsDisabled Then
                If IsSelected Then
                    If m_RightToLeft Then
                        .Rectangle .Width - (m_IconSizeXP - 1), 1, .Width - 2, .Height - 1, RGB(240, 160, 38), , BlendColor(vbActiveTitleBar, vbBlack)
                    Else
                        .Rectangle 2, 1, m_IconSizeXP - 1, .Height - 1, RGB(240, 160, 38), , BlendColor(vbActiveTitleBar, vbBlack)
                    End If
                Else
                    If m_RightToLeft Then
                        .Rectangle .Width - (m_IconSizeXP - 1), 1, .Width - 2, .Height - 1, RGB(248, 216, 126), , BlendColor(vbActiveTitleBar, vbBlack)
                    Else
                        .Rectangle 2, 1, m_IconSizeXP - 1, .Height - 1, RGB(248, 216, 126), , BlendColor(vbActiveTitleBar, vbBlack)
                    End If
                End If
            End If
        End If
        If IsChecked Then
            cS = currMenu.CheckedStyle
            If cS = EMCS_Normal Then
                bChar = "b"
            ElseIf (cS = EMCS_Radio) Then
                bChar = "h"
            End If
            If bChar <> vbNullString Then
                pDisableDrawIcon = True
                Set New_Font = New StdFont
                With New_Font
                    .Name = "Marlett"
                    .Size = 12
                End With
                Set .Font = New_Font
                If IsDisabled Then
                    .ForeColor = vbButtonShadow
                Else
                    .ForeColor = vbBlack
                End If
                If m_RightToLeft Then
                    .DrawText bChar, .Width - m_IconSize - 1 - 2, 0, .Width - 2, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
                Else
                    .DrawText bChar, 2, 0, m_IconSize - 1, .Height, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP
                End If
                Set New_Font = Nothing
            End If
        End If
        If (HasIcon And (Not (pDisableDrawIcon))) Then
            Set oPic = currMenu.Image
            If Not (oPic Is Nothing) Then
                If IsDisabled Then
                    If m_RightToLeft Then
                        .PaintDisabledPicture oPic, .Width - 20, (.Height - 16) / 2, 16, 16    '//--2)
                    Else
                        .PaintDisabledPicture oPic, 4, (.Height - 16) / 2, 16, 16    '//--2)
                    End If
                Else
                    If m_RightToLeft Then
                        .PaintPicture oPic, .Width - 20, (.Height - 16) / 2, 16, 16    '//--2)
                    Else
                        If IsSelected Then
                            .PaintPicture oPic, 4, (.Height - 16) / 2, 16, 16   '//--2)
                        Else
                            .PaintPicture oPic, 4, (.Height - 16) / 2, 16, 16    '//--2)
                        End If
                        'End If
                    End If
                    '//--Call .PaintPicture(oPic, 4, (.Height - 16) / 2, 16, 16) '//--2)
                End If
                Set oPic = Nothing
            End If
        End If
Display:
        .BitBlt DrawInfo.hdc, DrawInfo.rcItem.Left, DrawInfo.rcItem.Top, .Width, .Height, 0, 0
    End With
Ignore:
    Set xMemDC = Nothing
    Set currMenu = Nothing

End Sub

'//------------------------------
'//--  PRIVATE
'//------------------------------
Private Sub pvEnumMenus(ByVal hMenu As Long, _
                        ByVal ParentID As Long, _
                        Optional ItemOwnerDraw As Boolean = True)

    Dim I As Long
    Dim hSubMenu As Long
    Dim MenuCount As Long
    Dim Store_Menu As IAPP_StoreMenu
    Dim Chk As Long

    Dim cKey As String
    Dim VBMenu As Menu
    On Error Resume Next
    MenuCount = GetMenuItemCount(hMenu)
    For I = 0 To MenuCount - 1
        Set Store_Menu = pvGetMenu(hMenu, I, ParentID, ItemOwnerDraw)   '//--New IAPP_StoreMenu
        With Store_Menu
            'Set .Parent = Me
            m_MenusCount = m_MenusCount + 1
            If m_FromForm Then
                Set VBMenu = m_VBMenus(m_MenusCount)
                .Tag = VBMenu.Tag
                cKey = VBMenu.Name
                Chk = VBMenu.Index
            End If
            If Err.Number = 0 Then
                cKey = cKey & "(" & Chk & ")"
            End If
            .Name = cKey
            m_CollMenus.Add Store_Menu, "#" & .ID
            hSubMenu = GetSubMenu(hMenu, I)
            If hSubMenu > 0 Then
                .HasSub = True
                pvEnumMenus hSubMenu, .ID, True
            End If
        End With
    Next I
    Set VBMenu = Nothing
    Set Store_Menu = Nothing
    On Error GoTo 0

End Sub

Private Function pvExtractAccel(ByVal lStr As String) As String

    If (InStr(1, lStr, "|")) Then
        pvExtractAccel = Split(lStr, "|")(0)
    Else
        pvExtractAccel = lStr
    End If

End Function

Private Function pvFreeID() As Long

    If m_LastID = 0 Then
        m_LastID = 99
    End If
    m_LastID = m_LastID + 1
    pvFreeID = m_LastID

End Function

Private Function pvGetMenu(ByVal hMenu As Long, _
                           ByVal lngIndex As Long, _
                           ByVal ParentID As Long, _
                           Optional ByVal ItemOwnerDraw As Boolean = True) As IAPP_StoreMenu


    Dim New_Class As IAPP_StoreMenu
    Dim MII As MENUITEMINFO

    'Dim MI(0 To 1023) As Byte
    Dim sBuffer As String * 80
    Dim Temp As String
    Dim IsSep As Boolean
    Set New_Class = New IAPP_StoreMenu
    With New_Class
        '********************
        With MII
            .cbSize = Len(MII)
            .fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_ID Or MIIM_SUBMENU
            .fType = 0
            .dwTypeData = sBuffer
            .cch = 80
        End With
        GetMenuItemInfo hMenu, lngIndex, True, MII
        'mnuID = MII.wID
        .ID = MII.wID
        .ParentID = ParentID
        IsSep = ((MII.fType And MF_SEPARATOR) = MF_SEPARATOR)
        .Separator = IsSep
        Temp = Left$(Replace$(MII.dwTypeData, Chr$(0), vbNullString), MII.cch)
        If InStr(1, Temp, vbTab) Then
            .Caption = Split(Temp, vbTab)(0)
            .KeyAccel = Split(Temp, vbTab)(1)
        Else
            .Caption = Temp
            .KeyAccel = ""
        End If
        .hSubMenu = MII.hSubMenu
        .hMenu = hMenu
        .Index = lngIndex
        .TopMenu = (hMenu = m_hMenu)
        .CheckedStyle = EMCS_Normal
        'If it's not a main menu or a separator then convert to 'ItemOwnerDrawn'
        'If ItemOwnerDraw And IsSep = False Then
        If ItemOwnerDraw Then
            MII.fType = MII.fType Or MF_ItemOwnerDraw
            SetMenuItemInfo hMenu, lngIndex, True, MII
        End If
        '********************
    End With
    Set pvGetMenu = New_Class
    Set New_Class = Nothing

End Function

Private Sub pvInitMenuPopUp(ByVal hMenu As Long)

'Call pvEnumMenus(hMenu)


End Sub

Private Function pvIsFileExists(ByVal FileName As String) As Boolean

    On Local Error GoTo errHandle
    FileLen FileName
    pvIsFileExists = True
errHandle:

End Function

Private Function pvKeyAccelExists(ByVal KeyAccel As String) As Boolean

    Dim Fnd As Long

    Fnd = InStr(1, m_AllSupportedKeys, "|" & KeyAccel & "|")
    pvKeyAccelExists = (Fnd <> 0)

End Function

Private Function pvMeasureItem(ByVal wParam As Long, _
                               ByVal lParam As Long) As Boolean


'Set menu size

    Dim rWidth As Long
    Dim rHeight As Long
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim currMenu As IAPP_StoreMenu
    Dim bCancel As Boolean

    'Dim bSeparator As Boolean
    CopyMemory MeasureInfo, ByVal lParam, Len(MeasureInfo)
    'get the UDT for the menuitem's dimensions
    'bail - not 'ItemOwnerDrawn'
    If MeasureInfo.CtlType <> ODT_MENU Then
        GoTo Ignore
    End If
    'If (MeasureInfo.ItemID = 100) Then
    '//--   MsgBox "Hi"
    'End If
    pvMeasureItem = True
    Set currMenu = m_CollMenus.Item("#" & MeasureInfo.Itemid)  'get menu details from the class
    'If (MeasureInfo.ItemID = 100) Then
    '//--   Debug.Print "hey"
    'End If
    'Adjust dimensions as neccessary
    If (currMenu.OwnerDraw Or m_ItemOwnerDrawAll) Then
        rHeight = m_ItemHeight
        RaiseEvent MeasureItem(bCancel, IndexForKey(currMenu.Name), rWidth, rHeight)
        If bCancel Then
            MeasureInfo.ItemWidth = rWidth
            MeasureInfo.ItemHeight = rHeight
        Else
            GoTo AutoMesaure
        End If
    Else
AutoMesaure:
        If m_DrawStyle = mds_3D Then
            pvMeasureItem3D currMenu, MeasureInfo.ItemWidth, MeasureInfo.ItemHeight
        ElseIf (m_DrawStyle = mds_XP) Then
            pvMeasureItemXP currMenu, MeasureInfo.ItemWidth, MeasureInfo.ItemHeight
        End If
    End If
    'Return the UDT with the new values
    CopyMemory ByVal lParam, MeasureInfo, Len(MeasureInfo)
Ignore:
    Set currMenu = Nothing

End Function

Private Sub pvMeasureItem3D(ByVal currMenu As IAPP_StoreMenu, _
                            ByRef rWidth As Long, _
                            ByRef rHeight As Long)

    If currMenu.TopMenu Then
        rHeight = m_ItemHeight    '//--ParForm.TextHeight(BB.Caption) + 6
        rWidth = m_TmpMemDC.TextWidth(currMenu.Caption)   '+ 2
    ElseIf (currMenu.Separator) Then
        rWidth = 0
        '//--m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        rHeight = 9
    Else
        rHeight = m_ItemHeight    '//--ParForm.TextHeight(BB.Caption) + 6
        If LenB(currMenu.KeyAccel) = 0 Then
            rWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent
            '+ m_IconSize 'm_ExtraSize
        Else
            rWidth = m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent + m_TmpMemDC.TextWidth(pvExtractAccel(currMenu.KeyAccel)) + m_RightIndent
        End If
    End If

End Sub

Private Sub pvMeasureItemXP(ByVal currMenu As IAPP_StoreMenu, _
                            ByRef rWidth As Long, _
                            ByRef rHeight As Long)

    If currMenu.TopMenu Then
        rHeight = m_ItemHeight    '//--ParForm.TextHeight(BB.Caption) + 6
        'Debug.Print m_TmpMemDC.Font.Name
        rWidth = m_TmpMemDC.TextWidth(currMenu.Caption)   '+ 2
    ElseIf (currMenu.Separator) Then
        rWidth = 0
        '//--m_IconSize + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndent '+ m_IconSize 'm_ExtraSize
        rHeight = 9
    Else
        rHeight = m_ItemHeight    '//--ParForm.TextHeight(BB.Caption) + 6
        If LenB(currMenu.KeyAccel) = 0 Then
            rWidth = m_IconSizeXP + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndentXP
            '+ m_IconSize 'm_ExtraSize
        Else
            rWidth = m_IconSizeXP + m_TmpMemDC.TextWidth(currMenu.Caption) + m_CaptionIndentXP + m_TmpMemDC.TextWidth(pvExtractAccel(currMenu.KeyAccel)) + m_RightIndentXP
        End If
    End If

End Sub

Private Function pvMenuChar(ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    Dim I As Long
    Dim cChar As String

    cChar = UCase$(Chr$(lParam))
    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If .hMenu = wParam Then
                If .AccelChar = cChar Then
                    pvMenuChar = &H20000 Or .Index
                    Exit For
                End If
            End If
        End With
    Next I

End Function

Private Function pvMenuSelect(ByVal wParam As Long, _
                              ByVal lParam As Long) As Boolean

    Dim lHiWord As Long
    Dim lMenuId As Long

    Dim currMenu As IAPP_StoreMenu
    On Error Resume Next
    lHiWord = wParam \ &H10000
    lMenuId = wParam And &HFFFF&

    '//--Now check if the message is a menu item higlight,
    '//--or whether it is indicating exit from the menu:
    Set currMenu = m_CollMenus.Item("#" & lMenuId)
    If (currMenu Is Nothing) Then
        RaiseEvent MenuExit
    Else
        'Debug.Print currMenu.Name
        RaiseEvent MenuHighlight(IndexForKey(currMenu.Name))
        Set currMenu = Nothing
    End If
    On Error GoTo 0

End Function

Private Sub pvMoveIndex(ByVal hMenu As Long, _
                        ByVal lStart As Long, _
                        Optional ByVal lValue As Integer = 1)


    Dim I As Long
    Dim currMenu As IAPP_StoreMenu

    For I = 1 To m_CollMenus.Count
        Set currMenu = m_CollMenus.Item(I)
        With currMenu
            If hMenu = .hMenu Then
                If .Index >= lStart Then
                    .Index = .Index + lValue
                End If
            End If
        End With
    Next I
    Set currMenu = Nothing

End Sub

Private Sub pvRaiseClick(Optional poiIndex As Long, _
                         Optional ID As Long)


    Dim cID As Long
    Dim I As Long

    If Not (IsMissing(poiIndex)) Then
    ElseIf (Not (IsMissing(ID))) Then
        cID = CLng(ID)
        For I = 1 To m_CollMenus.Count
            With m_CollMenus.Item(I)
                If .ID = cID Then
                    'MsgBox .Name
                    RaiseEvent Click(I)
                End If
            End With
        Next I
    End If

End Sub

Private Function pvRemoveSubMenus(ByVal ID As Long)

'm_MenusToDestroyCnt

    Dim I As Long

    For I = 1 To m_CollMenus.Count
        With m_CollMenus.Item(I)
            If .ParentID = ID Then
                .DestroyMe = True
                pvRemoveSubMenus .ID
            End If
        End With
    Next I

End Function

Private Sub pvRemoveSupportedKeys(ByVal lKeys As String)


    m_AllSupportedKeys = Replace$(m_AllSupportedKeys, "|" & lKeys & "|", "|")
    If Len(m_AllSupportedKeys) = 1 Then
        m_AllSupportedKeys = vbNullString
    End If

End Sub

Private Sub pvReplaceAllItems()

    Dim I As Long

    For I = 1 To m_CollMenus.Count
        ReplaceItem I
    Next I

End Sub

Private Sub pvSetCaption(ByVal lngIndex As Long, _
                         ByVal vNewValue As String)


    Dim Fnd As Long
    Dim lLen As Long
    Dim Chk As String

    With m_CollMenus.Item(lngIndex)
        .Caption = vNewValue
        Fnd = InStr(1, vNewValue, "&")
        If Fnd Then
            lLen = Len(vNewValue)
            If Fnd <> lLen Then
                Chk = Mid$(vNewValue, Fnd + 1, 1)
                If Chk <> "&" Then
                    .AccelChar = UCase$(Chk)
                End If
            End If
        End If
    End With

End Sub

Private Function pvSetKeyAccel(ByVal lngIndex As Long, _
                               ByVal vNewValue As String) As Boolean


    If Not (pvCheckKeys(vNewValue)) Then
        Exit Function
    End If
    With m_CollMenus.Item(lngIndex)
        pvRemoveSupportedKeys .KeyAccel
        If vNewValue <> vbNullString Then
            pvAddSupportedKeys vNewValue
        End If
        .KeyAccel = vNewValue
    End With
    pvSetKeyAccel = True

End Function

Private Sub pvSubClass()

    With m_SubClass
        .AddMsg WM_MEASUREITEM, MSG_BEFORE
        .AddMsg WM_DRAWITEM, MSG_BEFORE
        .AddMsg WM_ERASEBKGND, MSG_BEFORE
        .AddMsg WM_MENUSELECT, MSG_BEFORE
        .AddMsg WM_COMMAND, MSG_BEFORE
        .AddMsg WM_MENUCHAR, MSG_BEFORE
        'Call .AddMsg(WM_INITMENUPOPUP, True)
        .Subclass m_hWnd, Me
    End With

End Sub

Public Function ReplaceItem(Index) As Boolean


    Dim cIndex As Long
    Dim currMenu As IAPP_StoreMenu

    'Dim I As Long
    Dim tMI As MENUITEMINFO
    'Dim hMenu As Long
    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set currMenu = m_CollMenus.Item(cIndex)
    Else
        Exit Function
    End If
    tMI.cbSize = Len(tMI)
    tMI.fMask = MIIM_SUBMENU Or MIIM_TYPE Or MIIM_DATA Or MIIM_ID Or MIIM_CHECKMARKS Or MIIM_STATE    '//--MIIM_SUBMENU
    GetMenuItemInfo currMenu.hMenu, currMenu.Index, True, tMI
    RemoveMenu currMenu.hMenu, currMenu.ID, MF_BYCOMMAND
    InsertMenuItem currMenu.hMenu, currMenu.Index, True, tMI
    Set currMenu = Nothing

End Function

Private Sub RGBToHLS(ByVal R As Long, _
                     ByVal G As Long, _
                     ByVal B As Long, _
                     h As Single, _
                     S As Single, _
                     l As Single)

    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single
    Dim rG As Single
    Dim rB As Single

    rR = R / 255
    rG = G / 255
    rB = B / 255
    '{Given: rgb each in [0,1].
    '//--Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        S = 0
        h = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
        '{First calculate the saturation.}
        If l <= 0.5 Then
            S = (Max - Min) / (Max + Min)
        Else
            S = (Max - Min) / (2 - Max - Min)
        End If
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta    '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta    '{Resulting color is between magenta and cyan}
        End If
        'end {Chromatic Case}
    End If

End Sub

'//-- Get\Set the Right to Left
Public Property Get RightToleft() As OLE_COLOR

    RightToleft = m_RightToLeft

End Property

Public Property Let RightToleft(ByVal vNewValue As OLE_COLOR)

    m_RightToLeft = vNewValue

End Property

Public Function SaveShortcuts(ByVal FileName As String) As Boolean


    Dim I As Long
    Dim hFile As Long
    Dim currMenu As IAPP_StoreMenu

    If pvIsFileExists(FileName) Then
        Kill FileName
    End If
    hFile = FreeFile
    Open FileName For Binary Lock Write As hFile
    Put #hFile, , vbNullChar
    Put #hFile, , "[Boody-Menus-Shortcuts]"
    Put #hFile, , vbNullChar
    Put #hFile, , "[V-1-0-0]"
    Put #hFile, , vbNullChar
    For I = 1 To m_CollMenus.Count
        Set currMenu = m_CollMenus.Item(I)
        With currMenu
            If .Name = vbNullString Then
                Put #hFile, , "I"
                Put #hFile, , vbNullChar
                Put #hFile, , CStr(I)
                Put #hFile, , vbNullChar
                Put #hFile, , .KeyAccel
                Put #hFile, , vbNullChar
            Else
                Put #hFile, , "K"
                Put #hFile, , vbNullChar
                Put #hFile, , .Name
                Put #hFile, , vbNullChar
                Put #hFile, , .KeyAccel
                Put #hFile, , vbNullChar
            End If
        End With
    Next I
    Close hFile
    Set currMenu = Nothing
    SaveShortcuts = True

End Function

Public Function AccelPressed(ByVal lLong As Long) As Boolean



End Function

Public Sub BuildOn(ByVal lngHwnd As Long)


    If m_IsCreated Then
        Exit Sub
    End If
    m_hWnd = lngHwnd
    pvSubClass
    SetMenu m_hWnd, m_hMenu
    DrawMenuBar m_hWnd
    m_IsCreated = True

End Sub


'//--  Functions

Public Sub CreateFromForm(ByVal oForm As Object, _
                          Optional ByVal TopMenus As Boolean = True)


    Dim CTL As Control

    'Dim cIndex As Long
    'Dim hMenu As Long
    If m_IsCreated Then
        Exit Sub
    End If
    m_TopMenus = TopMenus
    If Not (TypeOf oForm Is Form) Then
        Exit Sub
    End If
    m_hWnd = oForm.hwnd
    hWndKeyOwner = m_hWnd
    m_FromForm = True
    For Each CTL In oForm.Controls
        If TypeOf CTL Is Menu Then
            m_VBMenus.Add CTL
        End If
    Next CTL
    m_hMenu = GetMenu(m_hWnd)
    m_hSystemMenu = GetSystemMenu(m_hWnd, 0)
    m_RightToLeft = oForm.RightToleft
    pvEnumMenus m_hMenu, 0
    pvSubClass
    m_IsCreated = True

End Sub

Public Function LoadShortcuts(ByVal FileName As String) As Boolean


    Dim hFile As Long
    Dim lStr As String
    Dim lStrArr() As String
    Dim I As Long
    Dim Cnt As Long
    Dim oldIndex As Long
    Dim currStep As Byte
    Dim useKey As Boolean

    On Error GoTo errHandle
    If Not (pvIsFileExists(FileName)) Then
        Exit Function
    End If
    hFile = FreeFile
    Open FileName For Binary As hFile
    lStr = String$(LOF(hFile), 0)
    Get #hFile, , lStr
    Close hFile
    lStrArr = Split(lStr, vbNullChar)
    Cnt = UBound(lStrArr)
    If lStrArr(0) <> vbNullString Then
        GoTo errHandle
    End If
    If lStrArr(1) <> "[Boody-Menus-Shortcuts]" Then
        GoTo errHandle
    End If
    If lStrArr(2) <> "[V-1-0-0]" Then
        GoTo errHandle
    End If
    On Error GoTo IgnoreError
    For I = 3 To Cnt
        If currStep = 0 Then
            'type
            If lStrArr(I) = "I" Then
                useKey = False
            ElseIf (lStrArr(I) = "K") Then
                useKey = True
            Else
                GoTo IgnoreError
            End If
            currStep = currStep + 1
        ElseIf (currStep = 1) Then
            'index
            If useKey Then
                oldIndex = IndexForKey(lStrArr(I))
            Else
                oldIndex = CLng(lStrArr(I))
            End If
            currStep = currStep + 1
        ElseIf (currStep = 2) Then
            'accel
            If oldIndex <> 0 Then
                ItemKeyAccel(oldIndex) = CStr(lStrArr(I))
            End If
            currStep = 0
        End If
    Next I
IgnoreError:
    LoadShortcuts = True
errHandle:
    lStr = vbNullString

End Function

Public Function RemoveItem(Index) As Boolean


    Dim cIndex As Long
    Dim currMenu As IAPP_StoreMenu
    Dim I As Long

    'Dim B As Long
    'Dim hMenu As Long
    Dim bTopMenu As Boolean
    cIndex = IndexForKey(Index)
    If cIndex > 0 Then
        Set currMenu = m_CollMenus.Item(cIndex)
    Else
        Exit Function
    End If
    With currMenu
        RemoveMenu .hMenu, .ID, MF_BYCOMMAND
        .DestroyMe = True
        bTopMenu = .TopMenu
        pvRemoveSubMenus .ID
        pvMoveIndex .hMenu, .Index, -1
    End With
    Set currMenu = Nothing
    Do
        I = I + 1
        If I > m_CollMenus.Count Then
            Exit Do
        End If
        With m_CollMenus.Item(I)
            If .DestroyMe Then
                DestroyMenu .hMenu
                'Call DestroyMenu(.hSubMenu)
                'MsgBox .Name
                m_CollMenus.Remove I
                I = I - 1
            End If
            'B = B + 1
        End With
    Loop
    If bTopMenu Then
        DrawMenuBar m_hWnd
    End If

End Function


