VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdObjects 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cmbCombo 
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cmbCombo 
      Height          =   315
      Index           =   3
      Left            =   3480
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cmbCombo 
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox cmbCombo 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox picBox 
      Height          =   135
      Left            =   120
      ScaleHeight     =   75
      ScaleWidth      =   6195
      TabIndex        =   9
      Top             =   360
      Width           =   6255
   End
   Begin VB.TextBox txtTextBox 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtTextBox 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Option1"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.CheckBox chkCheckBox 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   3735
   End
   Begin VB.Frame fraIndex 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdIndexed 
         Caption         =   "Command1"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdIndexed 
         Caption         =   "Command1"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton cmdIndexed 
         Caption         =   "Command1"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdIndexed 
         Caption         =   "Command1"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label lblReminder 
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Menu"
      Begin VB.Menu mnuSupported 
         Caption         =   "Menu1"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMic 
         Caption         =   "Menu2"
      End
      Begin VB.Menu mnuTesting 
         Caption         =   "Menu3"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndexed 
         Caption         =   "Menu4"
         Begin VB.Menu mnuItem 
            Caption         =   "Menu5"
            Index           =   0
         End
         Begin VB.Menu mnuItem 
            Caption         =   "Menu6"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
clsTrans.Translation = App.Path & "\Translations\English.lng"
clsTrans.LoadStrings
clsTrans.SetTranslation Me
End Sub

Private Sub chkCheckBox_Click()
If chkCheckBox.Value = 1 Then
    MsgBox clsTrans.GetString(1)
Else
    MsgBox clsTrans.GetString(2)
End If
End Sub

Private Sub optOption_Click()
MsgBox clsTrans.GetString(3), vbExclamation
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdObjects_Click()
frmObjects.Show
End Sub
