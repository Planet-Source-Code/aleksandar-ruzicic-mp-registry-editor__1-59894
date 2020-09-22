VERSION 5.00
Begin VB.Form frmDWORD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmDWORD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   83
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame2 
         Caption         =   "Base:"
         Height          =   975
         Left            =   2760
         TabIndex        =   7
         Top             =   840
         Width           =   1815
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            ScaleHeight     =   615
            ScaleWidth      =   1455
            TabIndex        =   8
            Top             =   280
            Width           =   1455
            Begin VB.OptionButton Option2 
               Caption         =   "Decimal"
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   280
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Hexadecimal"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   0
               Width           =   1335
            End
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "New DWORD Value #n"
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
Attribute VB_Name = "frmDWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sName As String
Dim sValue As String
Dim isOK As Boolean
Dim mu As Boolean
Public Property Get DWORDName() As String
    DWORDName = sName
End Property
Public Property Let DWORDName(ByVal sn As String)
    sName = sn
    Text1.Text = sn
End Property
Public Property Get DWORDValue() As String
    If Len(Trim$(sValue)) = 0 Then sValue = "0"
    DWORDValue = sValue
End Property
Public Property Let DWORDValue(ByVal sv As String)
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
    
    If Option1.Value = True Then
        sValue = CStr(Val("&H" & Text2.Text)) 'convert to decimal
    Else
        sValue = Text2.Text
    End If
    
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

Private Sub Option1_Click()
    If Option1.Value And Text2.Tag <> "HEX" And Len(Trim$(Text2.Text)) > 0 Then
        Text2.Text = Hex$(CLng(Text2.Text))
        Text2.Tag = "HEX"
        Text2.SelStart = Len(Text2.Text)
        Text2.SetFocus
    End If
End Sub
Private Sub Option2_Click()
On Error Resume Next
    If Option2.Value And Text2.Tag <> "DEC" And Len(Trim$(Text2.Text)) > 0 Then
        Text2.Text = CStr(Val("&H" & Text2.Text))
        Text2.Tag = "DEC"
        Text2.SelStart = Len(Text2.Text)
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    Dim tmp As String
    
    If Option1.Value = True Then
        tmp = "0123456789ABCDEF "
    Else
        tmp = "0123456789"
    End If
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii))) 'make it uppercase
    If InStr(1, tmp, Chr$(KeyAscii)) < 1 Then
        If Not KeyAscii = vbKeyBack Then KeyAscii = 0
    End If
    
End Sub

