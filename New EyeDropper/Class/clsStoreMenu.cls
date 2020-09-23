VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAPP_StoreMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'//---------------------------------------------------------------------------------------
'EyeDropperTab
'//---------------------------------------------------------------------------------------
' Module    : IAPP_StoreMenu
' DateTime  : 04/02/2005
' Author    : Gary Noble   ©2005 Telecom Direct Limited
' Purpose   : Menu Item Holder
' Assumes   : IAPP_MenuHandler
' Notes     :
' Revision  : 1.3
'//---------------------------------------------------------------------------------------
' History   : Initial Implementation    Gary Noble  04/02/2005
'//---------------------------------------------------------------------------------------

Option Explicit

Public ID As Long
Public OwnerDrawID As Long
Public Caption As String
Public Name As String
Public HelpText As String
Public Image As StdPicture
Public ParentHandle As Long
Public KeyAccel As String
Public OwnerDraw As Boolean
Public CheckedStyle As EMenuCheckedStyle
Public HasSub As Boolean
Public AtRunTime As Boolean
Public OwnerDrawData As Long
Public Index As Long
Public Break As Long
Public hMenu As Long
Public Separator As Boolean
Public TopMenu As Boolean
Public hSubMenu As Long
Public ParentID As Long
Public DestroyMe As Boolean
Public AccelChar As String
Public Description As String
Public Tag As String
Public InVisible As Boolean
Public vKey As String

Private Sub Class_Terminate()

    KeyAccel = vbNullString
    Caption = vbNullString
    Name = vbNullString
    Set Image = Nothing
    HelpText = vbNullString
    AccelChar = vbNullString
    Description = vbNullString
    Tag = vbNullString
    vKey = vbNullString

End Sub

