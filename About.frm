VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About PasswordFinder"
   ClientHeight    =   2955
   ClientLeft      =   3975
   ClientTop       =   3630
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1290
      TabIndex        =   5
      Top             =   2445
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2175
      TabIndex        =   4
      Top             =   735
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Simple program to catch passwords from password fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   165
      TabIndex        =   3
      Top             =   120
      Width           =   4230
   End
   Begin VB.Label Label5 
      Caption         =   "www.raygon.8m.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   810
      MouseIcon       =   "About.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1830
      Width           =   2955
   End
   Begin VB.Label Label4 
      Caption         =   "raygon@epatra.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1238
      TabIndex        =   1
      Top             =   1410
      Width           =   2115
   End
   Begin VB.Label Label3 
      Caption         =   "RAYGON B.THOMAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   990
      TabIndex        =   0
      Top             =   1095
      Width           =   2610
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlue
End Sub


Private Sub Label5_Click()
ShellExecute 0, "open", "www.raygon.8m.com", "", vbNull, SW_SHOWNORMAL
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbRed
End Sub
