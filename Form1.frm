VERSION 5.00
Begin VB.Form PWD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PasswordFinder BY raygon@epatra.com"
   ClientHeight    =   390
   ClientLeft      =   4185
   ClientTop       =   5385
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   -30
      Top             =   645
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   4695
   End
End
Attribute VB_Name = "PWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************** PasswordFinder *****************
'Thank you for triying this
' This software can extract the password in password fields which are probably like this "**************"
' run this code and move the mouse pointer over the password field , you can see the password in PasswordFinder in plain text
'Please mail me your feedback and doubts
'I am ready to help you

'***************** DECLARATIONS *****************

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim pt As POINTAPI
Dim wnd As Long
Dim buff As String
Dim num As Long

Private Sub Form_Unload(Cancel As Integer)
Load About
End Sub

Private Sub Timer1_Timer()

GetCursorPos pt      'get the current position of mouse pointer
wnd = WindowFromPoint(pt.X, pt.Y) ' Get the window under mouse pointer
buff = Space$(50)
SendMessage wnd, &HD, 40, ByVal buff 'sending WM_GETTEXT message to that window
Label1.Caption = buff 'Now buff contains the text on the window , simply display it on a label
End Sub
