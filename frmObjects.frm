VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmObjects 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4110
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5821
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "SCRL"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   2
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5895
      TabIndex        =   4
      Top             =   600
      Width           =   5895
      Begin ComctlLib.Toolbar tbrToolbar 
         Height          =   600
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         ButtonWidth     =   609
         ButtonHeight    =   953
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   " "
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar tbrToolbar 
         Height          =   600
         Index           =   0
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         ButtonWidth     =   609
         ButtonHeight    =   953
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   " "
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.Toolbar tbrToolbar 
         Height          =   600
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         ButtonWidth     =   609
         ButtonHeight    =   953
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   " "
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblToolbar 
         Caption         =   "Label1"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   3
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   600
      Width           =   5895
      Begin VB.ListBox lstMode 
         Height          =   450
         Left            =   2040
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblStatusBar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   1
      Left            =   240
      ScaleHeight     =   3255
      ScaleWidth      =   5895
      TabIndex        =   2
      Top             =   600
      Width           =   5895
      Begin ComctlLib.ListView lstListView 
         Height          =   3255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tbsTabStrip 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Dim i As Integer

clsTrans.SetTranslation Me

For i = 0 To 20
    lstListView.ListItems.Add , , clsTrans.GetString(100) & " " & (i + 1)
    lstListView.ListItems(i + 1).SubItems(1) = clsTrans.GetString(101) & " " & (i + 1)
    lstListView.ListItems(i + 1).SubItems(2) = Date
Next
lstMode.AddItem clsTrans.GetString(102)
lstMode.AddItem clsTrans.GetString(103)
lstMode.ListIndex = 0
picContainer(1).ZOrder
End Sub

Private Sub lstMode_Click()
If lstMode.ListIndex = 0 Then
    sbStatusBar.Style = sbrNormal
Else
    sbStatusBar.Style = sbrSimple
End If
End Sub

Private Sub tbsTabStrip_Click()
picContainer(tbsTabStrip.SelectedItem.Index).ZOrder
End Sub
