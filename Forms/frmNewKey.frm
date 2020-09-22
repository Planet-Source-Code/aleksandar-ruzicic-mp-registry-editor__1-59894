VERSION 5.00
Begin VB.Form frmNewKey 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Key"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmNewKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   68
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "New Key #n"
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   280
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmNewKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sKey As String
Dim isOK As Boolean
Dim mu As Boolean
Public Property Get KeyName() As String
    KeyName = sKey
End Property
Public Property Let KeyName(ByVal sName As String)
    sKey = sName
    Text1.Text = sName
End Property
Public Property Get Canceled() As Boolean
    Canceled = Not isOK
End Property
Private Sub cmdCancel_Click()
    isOK = False
    mu = True
    Unload Me
End Sub
Private Sub cmdSave_Click()
    isOK = True
    sKey = Text1.Text
    mu = True
    Unload Me
End Sub
Private Sub Form_Activate()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    isOK = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mu Then isOK = False
End Sub
