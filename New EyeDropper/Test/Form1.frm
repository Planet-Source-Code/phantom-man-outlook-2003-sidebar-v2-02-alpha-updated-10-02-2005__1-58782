VERSION 5.00
Object = "*\A..\EyeDropper.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "EyeDropper Example"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   8640
      ScaleHeight     =   4215
      ScaleWidth      =   3015
      TabIndex        =   10
      Top             =   960
      Width           =   3015
      Begin EyeDropperTab.EyeDropperContainer EyeDropperContainer3 
         CausesValidation=   0   'False
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3836
         Caption         =   "Date Select"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         CustomColor     =   33023
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1889
            Index           =   2
            Left            =   0
            Picture         =   "Form1.frx":0000
            ScaleHeight     =   1890
            ScaleWidth      =   3015
            TabIndex        =   14
            Top             =   286
            Width           =   3015
         End
      End
      Begin EyeDropperTab.EyeDropperContainer EyeDropperContainer2 
         CausesValidation=   0   'False
         Height          =   2175
         Left            =   0
         TabIndex        =   11
         Top             =   2520
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3836
         BackColor       =   16777215
         Caption         =   "My Calenders"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         CustomColor     =   33023
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1889
            Left            =   0
            Picture         =   "Form1.frx":E80E
            ScaleHeight     =   1890
            ScaleWidth      =   4095
            TabIndex        =   12
            Top             =   286
            Width           =   4095
         End
      End
   End
   Begin EyeDropperTab.EyeDropperContainer EyeDropperContainer1 
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Caption         =   "Inbox Items"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      CustomColor     =   33023
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   404
         Index           =   7
         Left            =   0
         Picture         =   "Form1.frx":1D1E0
         ScaleHeight     =   405
         ScaleWidth      =   2655
         TabIndex        =   9
         Top             =   331
         Width           =   2655
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4440
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":320AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32988
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33262
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33B3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34416
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":355CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":35EA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3677E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":37058
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   6
      Left            =   10080
      Picture         =   "Form1.frx":37932
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   5
      Left            =   10080
      Picture         =   "Form1.frx":38B10
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   10080
      Picture         =   "Form1.frx":39EBA
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   10080
      Picture         =   "Form1.frx":3B00C
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   10080
      Picture         =   "Form1.frx":3CF92
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin EyeDropperTab.EyeDropper EyeDropper1 
      Align           =   3  'Align Left
      Height          =   6270
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11060
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisibleItems    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderForeColor =   16777215
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomColor     =   33023
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   3735
      Left            =   3000
      TabIndex        =   16
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton Command18 
         Caption         =   "Use Custom Color"
         Height          =   555
         Left            =   1440
         TabIndex        =   34
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Custom Colour"
         Height          =   555
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Selected Forecolor"
         Height          =   495
         Left            =   4320
         TabIndex        =   32
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Fore Color"
         Height          =   555
         Left            =   1440
         TabIndex        =   31
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Hover Forecolor"
         Height          =   555
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Header Forecolor"
         Height          =   555
         Left            =   2880
         TabIndex        =   29
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Hover Item Font"
         Height          =   555
         Left            =   2880
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Item Font"
         Height          =   555
         Left            =   4320
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Less items"
         Height          =   555
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disable/Enable"
         Height          =   555
         Left            =   4320
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Hide Selected Item"
         Height          =   555
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove Selected Item"
         Height          =   555
         Left            =   1440
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reload"
         Height          =   555
         Left            =   2880
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Header"
         Height          =   555
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Righ To Left Support"
         Height          =   555
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Header Font"
         Height          =   555
         Left            =   4320
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "More Items"
         Height          =   555
         Left            =   1440
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Clear"
         Height          =   555
         Left            =   2880
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Events "
      Height          =   1455
      Left            =   3000
      TabIndex        =   6
      Top             =   4680
      Width           =   5655
      Begin MSComctlLib.ListView ListView1 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1931
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Event"
            Object.Width           =   8467
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Me.EyeDropper1.VisibleItems = Me.EyeDropper1.VisibleItems - 1

End Sub

Private Sub Command10_Click()
Me.EyeDropper1.VisibleItems = Me.EyeDropper1.VisibleItems + 1
End Sub

Private Sub Command11_Click()
 Me.EyeDropper1.Clear
End Sub

Private Sub Command12_Click()
 With New ICDLG_FontDialogHandler
    .Init Me.EyeDropper1.HoverFont, "Choose Font", Me.hwnd
        .Show
        Set Me.EyeDropper1.HoverFont = .Font
        Me.EyeDropper1.Redraw = True
End With

End Sub

Private Sub Command13_Click()
    With New ICDLG_ColorDialogHandler
        .Init Me.EyeDropper1.HeaderForeColor, "Choose", Me.hwnd
        .Show
        
        
        Me.EyeDropper1.HeaderForeColor = .SelectedColour
        
        Me.EyeDropper1.Redraw = True
    End With

End Sub

Private Sub Command14_Click()

    With New ICDLG_ColorDialogHandler
        .Init Me.EyeDropper1.Hoverforecolor, "Choose", Me.hwnd
        .Show
        Me.EyeDropper1.Hoverforecolor = .SelectedColour
        Me.EyeDropper1.Redraw = True
    End With

End Sub

Private Sub Command15_Click()
    With New ICDLG_ColorDialogHandler
        .Init Me.EyeDropper1.NormalForeColor, "Choose", Me.hwnd
        .Show
        
        
        Me.EyeDropper1.NormalForeColor = .SelectedColour
        
        Me.EyeDropper1.Redraw = True
    End With

End Sub

Private Sub Command16_Click()
    With New ICDLG_ColorDialogHandler
        .Init Me.EyeDropper1.selectedforecolor, "Choose", Me.hwnd
        .Show
        Me.EyeDropper1.selectedforecolor = .SelectedColour
        Me.EyeDropper1.Redraw = True
    End With

End Sub

Private Sub Command17_Click()
    With New ICDLG_ColorDialogHandler
        .Init Me.EyeDropper1.customcolor, "Choose", Me.hwnd
        .Show
        Me.EyeDropper1.customcolor = .SelectedColour
        
        Me.EyeDropperContainer1.customcolor = .SelectedColour
        Me.EyeDropperContainer2.customcolor = .SelectedColour
        Me.EyeDropperContainer3.customcolor = .SelectedColour
        
        Me.EyeDropper1.Redraw = True
    End With
End Sub

Private Sub Command18_Click()
    Me.EyeDropper1.UseCustomColor = Not Me.EyeDropper1.UseCustomColor
    Me.EyeDropperContainer1.UseCustomColors = Not Me.EyeDropperContainer1.UseCustomColors
    Me.EyeDropperContainer2.UseCustomColors = Not Me.EyeDropperContainer2.UseCustomColors
    Me.EyeDropperContainer3.UseCustomColors = Not Me.EyeDropperContainer3.UseCustomColors
    
End Sub

Private Sub Command2_Click()

    Me.EyeDropper1.Enabled = Not Me.EyeDropper1.Enabled

End Sub

Private Sub Command3_Click()
On Error Resume Next

    Me.EyeDropper1.SelectedItem.Visible = Not Me.EyeDropper1.SelectedItem.Visible
    

On Error GoTo 0
End Sub

Private Sub Command4_Click()
On Error Resume Next

    Me.EyeDropper1.EyeDropperItems.Remove (Me.EyeDropper1.SelectedItem.Key)

On Error GoTo 0
End Sub

Private Sub Command5_Click()

    Form_Load

End Sub

Private Sub Command6_Click()
    Me.EyeDropper1.DisplayHeader = Not Me.EyeDropper1.DisplayHeader
    Me.EyeDropper1.Redraw = True
    
    
    
End Sub

Private Sub Command7_Click()
    Me.EyeDropper1.RighToLeft = Not Me.EyeDropper1.RighToLeft
    Me.EyeDropper1.Redraw = True
End Sub

Private Sub Command8_Click()

 
 With New ICDLG_FontDialogHandler
    .Init Me.EyeDropper1.HeaderFont, "Choose Font", Me.hwnd
    
        .Show
        
        Set Me.EyeDropper1.HeaderFont = .Font
        
        Me.EyeDropper1.Redraw = True
    
End With

    
End Sub

Private Sub Command9_Click()

 With New ICDLG_FontDialogHandler
    .Init Me.EyeDropper1.HeaderFont, "Choose Font", Me.hwnd
        .Show
        Set Me.EyeDropper1.Font = .Font
        Me.EyeDropper1.Redraw = True
End With

End Sub

Private Sub EyeDropper1_HoverItem(ByVal oItem As EyeDropperTab.EDItem)
    
    
    Me.ListView1.ListItems.Add , , "Hovering:   " & oItem.Caption
    Me.ListView1.ListItems(Me.ListView1.ListItems.Count).EnsureVisible
    Me.ListView1.SelectedItem = Me.ListView1.ListItems(Me.ListView1.ListItems.Count)
    
End Sub

Private Sub EyeDropper1_ItemSelected(ByVal oItem As EyeDropperTab.EDItem)
    
    Me.ListView1.ListItems.Add , , "ItemSelected: " & oItem.Caption
    Me.ListView1.ListItems(Me.ListView1.ListItems.Count).EnsureVisible
    Me.ListView1.SelectedItem = Me.ListView1.ListItems(Me.ListView1.ListItems.Count)
    
End Sub

Private Sub EyeDropper1_MouseRightClick()
    
    Me.ListView1.ListItems.Add , , "Right Mouse Down"
    Me.ListView1.ListItems(Me.ListView1.ListItems.Count).EnsureVisible
    Me.ListView1.SelectedItem = Me.ListView1.ListItems(Me.ListView1.ListItems.Count)
    
End Sub

Private Sub EyeDropper1_ThemeChanged(ByVal sThemeName As String)
    
    Me.ListView1.ListItems.Add , , "Windows Theme Changed: " & sThemeName
    Me.ListView1.ListItems(Me.ListView1.ListItems.Count).EnsureVisible
    Me.ListView1.SelectedItem = Me.ListView1.ListItems(Me.ListView1.ListItems.Count)
    
End Sub

Private Sub Form_DblClick()

    Set Me.EyeDropper1.EyeDropperItems(3).Picture = Me.EyeDropper1.EyeDropperItems(1).ToolBarPic

End Sub

Private Sub Form_Load()

Dim i     As Long
Dim xItem As EDItem


    Me.Label1.Caption = "EyeDropper " & Me.EyeDropper1.version & " Alpha Example"
    
    '-- This MUST Be Called First
    Me.EyeDropper1.Initialise
    
    '-- Set The Redraw To False For Faster Drawing
    Me.EyeDropper1.Redraw = False
    
    '-- Clear The control
    Me.EyeDropper1.Clear
    
    
    '-- Add 7 Items - Its That Simple
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Mail", "Mail", Me.ImageList2.ListImages(1).ExtractIcon, Me.Picture1(1))
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Calender", "Calender", Me.ImageList2.ListImages(2).ExtractIcon, Me.Picture2)
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Contacts", "Contacts", Me.ImageList2.ListImages(3).ExtractIcon, Me.Picture1(3))
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Journal", "Journal", Me.ImageList2.ListImages(4).ExtractIcon, Me.Picture1(4))
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Notes", "Notes", Me.ImageList2.ListImages(5).ExtractIcon, Me.Picture1(5))
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Tasks", "Tasks", Me.ImageList2.ListImages(6).ExtractIcon, Me.Picture1(6))
    Set xItem = Me.EyeDropper1.EyeDropperItems.Add("Folder List", "Folder List", Me.ImageList2.ListImages(7).ExtractIcon, EyeDropperContainer1)
    
    '-- Set Hom Many Visible Items You Want To see
    Me.EyeDropper1.VisibleItems = 4
    
    '-- Redraw The Control
    Me.EyeDropper1.Redraw = True
    
End Sub

Private Sub Picture2_Resize()
    On Error Resume Next
        
        DoEvents
        
        EyeDropperContainer3.Move 0, 0, Picture2.ScaleWidth, Picture2.Height \ 2
        
        EyeDropperContainer2.Move 0, Picture2.Height \ 2, Picture2.ScaleWidth, Picture2.Height \ 2
        
End Sub
