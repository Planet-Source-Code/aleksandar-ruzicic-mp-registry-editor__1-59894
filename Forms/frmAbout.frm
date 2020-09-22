VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About MP Registry Editor"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   880
      ScaleHeight     =   1185
      ScaleWidth      =   4305
      TabIndex        =   2
      Top             =   720
      Width           =   4335
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   50
         ScaleHeight     =   5775
         ScaleWidth      =   4335
         TabIndex        =   4
         Top             =   0
         Width           =   4335
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":0E42
            Height          =   5775
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   4335
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":168A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MAD Pinguins Registry Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Picture3.Top = Picture3.Top - 5
    If -Picture3.Top >= 5200 Then Picture3.Top = 0
End Sub
