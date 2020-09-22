VERSION 5.00
Begin VB.Form frmString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New String"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmString.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   83
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "New String Value #n"
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Value Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Value Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sName As String
Dim sValue As String
Dim isOK As Boolean
Dim mu As Boolean
Public Property Get StringName() As String
    StringName = sName
End Property
Public Property Let StringName(ByVal sn As String)
    sName = sn
    Text1.Text = sn
End Property
Public Property Get StringValue() As String
    StringValue = sValue
End Property
Public Property Let StringValue(ByVal sv As String)
    sValue = sv
    Text2.Text = sv
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
    sName = Text1.Text
    sValue = Text2.Text
    mu = True
    Unload Me
End Sub
Private Sub Form_Activate()
    If Not Text1.Locked Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Text1.SetFocus
        Text1.BackColor = Text2.BackColor
    Else
        Text2.SelStart = 0
        Text2.SelLength = Len(Text2.Text)
        Text2.SetFocus
        Text1.BackColor = Me.BackColor
    End If
    If LCase$(Text1.Text) = "(default)" And _
       LCase$(Text2.Text) = "value not set" Then
       Text2.Text = ""
       sValue = ""
    End If
End Sub
Private Sub Form_Load()
    isOK = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mu Then isOK = False
End Sub

