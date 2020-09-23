VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.swift.dk3.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Intro 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2000 Swift Publications Limited. All rights reserved."
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SJOTS Perfect 5.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmSplash.frx":0292
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SJOTS Perfect 5.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4320
      Picture         =   "frmSplash.frx":06D4
      Top             =   2880
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image X_Down 
      Height          =   255
      Left            =   4320
      Picture         =   "frmSplash.frx":0C6A
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image X_Up 
      Height          =   255
      Left            =   4320
      Picture         =   "frmSplash.frx":1200
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
SP5.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X_Up.Picture = Image1.Picture
End Sub

Private Sub Intro_Click()
Unload Me
SP5.Show
End Sub

Private Sub Label2_Click()
Unload Me
SP5.Show
End Sub

Private Sub Label3_Click()
Unload Me
SP5.Show
End Sub

Private Sub Label4_Click()
Unload Me
SP5.Show
End Sub

Private Sub X_Up_Click()
X_Up.Picture = X_Down.Picture
Unload Me
End Sub

Private Sub X_Up_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X_Up.Picture = X_Down.Picture
End Sub
