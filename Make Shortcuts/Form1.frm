VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Making Shortcuts Without VB runtimes"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox FolderNameTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "My App"
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "StartMenu/Programs"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "StartMenu"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desktop"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox UninLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "C:\Program Files\MyApp\uninstall.exe"
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox HelpLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "C:\Program Files\MyApp\MyApp.hlp"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox AppLocTxt 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "C:\Program Files\MyApp\MyApp.exe"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Go 
      Caption         =   "Go"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "StartMenu Folder Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Create Shortcuts on.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Help File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Uninstall File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Go_Click()
If Check1.Value = 1 Then
MakeDesktopShortcut AppLocTxt
MakeDesktopShortcut UninLocTxt
MakeDesktopShortcut HelpLocTxt
End If

If Check2.Value = 1 Then
MakeStartMenuShortcut AppLocTxt
MakeStartMenuShortcut UninLocTxt
MakeStartMenuShortcut HelpLocTxt
End If

If Check3.Value = 1 Then
MakeStartMenuFolderShortcut AppLocTxt, FolderNameTxt
MakeStartMenuFolderShortcut UninLocTxt, FolderNameTxt
MakeStartMenuFolderShortcut HelpLocTxt, FolderNameTxt
End If

End Sub

