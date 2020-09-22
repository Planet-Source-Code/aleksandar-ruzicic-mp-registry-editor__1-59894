VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to Registry File"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdExport 
      Left            =   3480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "reg"
      DialogTitle     =   "&Export to file..."
      Filter          =   "Registration Files (*.reg)|*.reg|All Files (*.*)|*.*"
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Include subkeys"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   5175
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   3960
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Export from key:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private eKey As String
Private eFile As String
Private eInclude As Boolean
Private isOK As Boolean
Private mu As Boolean
Public Property Let StartKey(sKey As String)
    eKey = sKey
    Text2.Text = eKey
End Property
Public Property Get StartKey() As String
    StartKey = eKey
End Property
Public Property Get RegFile() As String
    RegFile = eFile
End Property
Public Property Get Include() As Boolean
    Include = eInclude
End Property
Public Property Get Canceled() As Boolean
    Canceled = Not isOK
End Property

Private Sub cmdBrowse_Click()
    On Error Resume Next
    
    cdExport.ShowSave
    
    If Err.Number = cdlCancel Then Exit Sub
    
    Err.Clear
    
    Text1.Text = cdExport.FileName
    
End Sub

Private Sub cmdCancel_Click()
    isOK = False
    mu = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim reg As New clsRegistryAccess
    
    If reg.KeyExists(Text2.Text) = False Then
        MsgBox "Key don't exists!", vbCritical, "Error"
        Exit Sub
    Else
        If Dir(Text1.Text) <> "" Then
            If (MsgBox("File allready exists. Do you want to overwrite it?", _
                    vbExclamation + vbYesNo, "Confirm overwrite") = vbYes) Then
                On Error Resume Next
                Kill Text1.Text
                
                If Err.Number <> 0 Then
                    MsgBox "Error overwriting file. Disk is maybe write-protected or file is in use.", _
                            vbCritical, "Error"
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
    End If

    isOK = True
    eKey = Text2.Text
    eFile = Text1.Text
    eInclude = Check1.Value = vbChecked
    mu = True
    Unload Me
End Sub

Private Sub Form_Activate()
    Text2.Text = eKey
    Text1.Text = ""
End Sub

Private Sub Form_Load()
    isOK = True
    mu = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mu Then isOK = False
End Sub
