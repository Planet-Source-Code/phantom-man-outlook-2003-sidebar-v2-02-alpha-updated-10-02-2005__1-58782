VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"EDItem"
Attribute VB_Ext_KEY = "Member0" ,"EDItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
'//--Module    : EDItems
'//--DateTime  : 03/02/2005
'//--Author    : Gary Noble   ©2005 Telecom Direct Limited
'//--Purpose   : Interface For Adding/Editing Items From The EyeDropper Control
'//--Assumes   : EDItem
'//--Notes     :
'//--Revision  : 1.0
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  03/02/2005
'//---------------------------------------------------------------------------------------
Option Explicit
Private mCol As Collection
Private mvarEyeDropperControl As EyeDropper
Private mHWND As Long

Public Function Add(ByVal strCaption As String, _
                    Key As Variant, _
                    Large32x32Picture As StdPicture, _
                    Optional ByVal PanelObject As Object, _
                    Optional ByVal bVisible As Boolean = True) As EDItem

    On Error GoTo errAdd

    Dim objNewMember As EDItem

    Set objNewMember = New EDItem

    If Not KeyExists(Key) Then

        objNewMember.Caption = strCaption

        Set objNewMember.EyeDropperControl = mvarEyeDropperControl

        If Not PanelObject Is Nothing Then
            If Not mvarEyeDropperControl Is Nothing Then
                If LenB(mvarEyeDropperControl.SelectedKey) = 0 Then
                    mvarEyeDropperControl.SelectedKey = Key
                End If
                Set PanelObject.Container = mvarEyeDropperControl.Extender
                Set objNewMember.Panel = PanelObject
            End If
        End If

        If IsObject(Key) Then
            Set objNewMember.Key = Key
        Else
            objNewMember.Key = Key
        End If

        With objNewMember

            Set .Picture = Large32x32Picture
            'Set .ToolBarPic = Small16x16ToolBarPicture
            .Visible = bVisible

        End With

        If Len(Key) = 0 Then
            mCol.Add objNewMember, Key
        Else
            mCol.Add objNewMember, Key
        End If

    Else
        Err.Raise vbObjectError + 1001, "Eye Dropper Item", "Key: " & Key & " Already Exists"
    End If

    'return the object created
    Set Add = objNewMember

CleanExit:

    Set objNewMember = Nothing
    Exit Function
errAdd:
    Err.Raise vbObjectError + 10001, "EyeDropper.EDItem - Add", Err.Description
    Resume CleanExit


End Function

Private Sub Class_Initialize()

'creates the collection when this class is created

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()

'destroys collection when this class is terminated

    Set mCol = Nothing

End Sub

Public Property Get Count() As Long
    On Error Resume Next

    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count

    Count = mCol.Count

    On Error GoTo 0
End Property

Friend Property Get EyeDropperControl() As EyeDropper

    Set EyeDropperControl = mvarEyeDropperControl

End Property

Friend Property Set EyeDropperControl(ByVal vData As EyeDropper)

    Set mvarEyeDropperControl = vData

End Property

Friend Property Let hwnd(ByVal lHwnd As Long)


    mHWND = lHwnd

End Property

Public Property Get Item(vntIndexKey As Variant) As EDItem
Attribute Item.VB_UserMemId = 0

    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    On Error GoTo 0

End Property

Private Function KeyExists(Key As Variant) As Boolean

    On Error Resume Next
    KeyExists = Not mCol(Key) Is Nothing
    On Error GoTo 0

End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mCol.[_NewEnum]

End Property

Public Sub Remove(vntIndexKey As Variant)
    On Error GoTo errRemove

    'used when removing an element from the collection

    With Me.EyeDropperControl
        If mCol.Count = 1 Then
            mCol.Remove vntIndexKey
            .SelectedKey = ""
        Else
            mCol.Remove vntIndexKey
            If vntIndexKey = .SelectedKey Then
                .SelectedKey = IIf(mCol.Count > 0, mCol(1).Key, vbNullString)
            End If
        End If
        .SetItemRects
        .VisibleItems = .VisibleItems
        .RedrawControl
        .GetNextSelectedItem
    End With

CleanExit:
    On Error GoTo 0
    Exit Sub
errRemove:
    Err.Raise vbObjectError + 10001, "EyeDropper.EDItem - Remove", Err.Description
    Resume CleanExit
End Sub


