VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
'//--Module    : EDItem
'//--DateTime  : 03/02/2005
'//--Author    : Gary Noble   ©2005 Telecom Direct Limited
'//--Purpose   : Data Holder For The EyeDropper Control
'//--Assumes   :
'//--Notes     :
'//--Revision  : 1.0
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  03/02/2005
'//---------------------------------------------------------------------------------------

Option Explicit

Private mvarCaption As String
Private mvarKey As Variant
Private mvarPicture As Picture
Private mvarEnabled As Boolean
Private mvarItemRect As RECT
Private mvarIsToolbarItem As Boolean
Private mvarToolBarPic As StdPicture
Private mvarPanel As Object
Private m_lngLeft As Long
Private m_lngRight As Long
Private m_lngTop As Long
Private m_lngBottom As Long
Private mvarVisible As Boolean
Private mvarEyeDropperControl As EyeDropper

Friend Property Set EyeDropperControl(ByVal vData As EyeDropper)
    Set mvarEyeDropperControl = vData
End Property

Friend Property Get Bottom() As Long

    Bottom = m_lngBottom

End Property

Public Property Get Caption() As String

    Caption = mvarCaption

End Property

Public Property Let Caption(ByVal vData As String)

    mvarCaption = vData
    If Not mvarEyeDropperControl Is Nothing Then mvarEyeDropperControl.RedrawControl

End Property

Private Sub Class_Terminate()

    On Error Resume Next
    If Not Panel Is Nothing Then
        Panel.Visible = False
    End If


    On Error GoTo 0

End Sub

Public Property Get Enabled() As Boolean

    Enabled = mvarEnabled

End Property

Public Property Let Enabled(ByVal vData As Boolean)

    mvarEnabled = vData
    If Not mvarEyeDropperControl Is Nothing Then mvarEyeDropperControl.RedrawControl

End Property

Friend Property Get isToolBarItem() As Boolean

    isToolBarItem = mvarIsToolbarItem

End Property

Friend Property Let isToolBarItem(ByVal vData As Boolean)

    mvarIsToolbarItem = vData

End Property

Friend Property Get ItemRect() As RECT

    LSet ItemRect = mvarItemRect

End Property

Friend Property Let ItemRect(vData As RECT)

    LSet mvarItemRect = vData

End Property

Public Property Get Key() As Variant

    If IsObject(mvarKey) Then
        Set Key = mvarKey
    Else
        Key = mvarKey
    End If

End Property

Friend Property Let Key(ByVal vData As Variant)

    mvarKey = vData

End Property

Friend Property Set Key(ByVal vData As Variant)

    Set mvarKey = vData
    If Not mvarEyeDropperControl Is Nothing Then mvarEyeDropperControl.RedrawControl

End Property

Friend Property Get Left() As Long

    Left = m_lngLeft

End Property

Public Property Get Panel() As Object

    Set Panel = mvarPanel

End Property

Public Property Set Panel(ByVal vData As Object)


    If Not Me.Panel Is Nothing Then
        Me.Panel.Visible = False
    End If

    Set mvarPanel = vData

    Me.Panel.Visible = Me.Visible


End Property

Public Property Get Picture() As Picture

    Set Picture = mvarPicture

End Property



Public Property Set Picture(ByVal vData As StdPicture)

On Error Resume Next

    If vData.Height \ Screen.TwipsPerPixelY <> 56 Or vData Is Nothing Or vData.Type <> 3 Then
        Err.Raise vbObjectError + 12011, "Add (EDItem)", "Image Not Correct Size Or Is Not An Icon"
        Exit Property
    End If


    Set mvarPicture = vData

    If Not mvarEyeDropperControl Is Nothing Then mvarEyeDropperControl.RedrawControl

On Error GoTo 0


End Property

Friend Property Get Right() As Long

    Right = m_lngRight

End Property

Friend Sub SetRectItem(ByVal lngLeft As Long, _
                       ByVal lngTop As Long, _
                       ByVal Bottom As Long, _
                       ByVal Right As Long)


    m_lngLeft = lngLeft
    m_lngRight = Right
    m_lngBottom = Bottom
    m_lngTop = lngTop

End Sub

Public Property Get ToolBarPic() As StdPicture

    Set ToolBarPic = mvarToolBarPic

End Property


Friend Property Get Top() As Long

    Top = m_lngTop

End Property

Public Property Get Visible() As Boolean

    Visible = mvarVisible

End Property

Public Property Let Visible(ByVal vData As Boolean)

    mvarVisible = vData
    If Not mvarEyeDropperControl Is Nothing Then
        mvarEyeDropperControl.SetItemRects

        mvarEyeDropperControl.RedrawControl
        mvarEyeDropperControl.GetNextSelectedItem

    End If


End Property
