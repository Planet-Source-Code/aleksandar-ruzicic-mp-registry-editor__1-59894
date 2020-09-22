VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Registry Editor"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9465
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   631
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdImport 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "reg"
      DialogTitle     =   "Import from file..."
      Filter          =   "Registration Files (*.reg)|*.reg|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txtExport 
      Height          =   525
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOutput 
      Height          =   315
      Left            =   3900
      TabIndex        =   4
      Top             =   900
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   5445
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   "My Computer"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      Height          =   1065
      Left            =   3600
      ScaleHeight     =   1005
      ScaleWidth      =   15
      TabIndex        =   2
      Top             =   150
      Width           =   75
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3900
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":197A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwData 
      Height          =   1065
      Left            =   1950
      TabIndex        =   1
      Top             =   150
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   6615
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwKeys 
      Height          =   1065
      Left            =   0
      TabIndex        =   0
      Top             =   150
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1879
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Registry"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Registry File..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Registry File..."
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuModify 
         Caption         =   "Modify"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuNewKey 
            Caption         =   "&Key"
            Enabled         =   0   'False
         End
         Begin VB.Menu sep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNewString 
            Caption         =   "&String Value"
         End
         Begin VB.Menu mnuNewBinary 
            Caption         =   "&Binary value"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuNewDWORD 
            Caption         =   "&DWORD"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKillKey 
         Caption         =   "&Delete Key"
         Enabled         =   0   'False
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuKillValue 
         Caption         =   "Delete &Value"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyName 
         Caption         =   "&Copy Key Name"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStatus 
         Caption         =   "&Status bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu popKey 
      Caption         =   "PopUpKey"
      Visible         =   0   'False
      Begin VB.Menu pop1New 
         Caption         =   "New"
         Begin VB.Menu popNewKey 
            Caption         =   "&Key"
         End
         Begin VB.Menu popSep1 
            Caption         =   "-"
         End
         Begin VB.Menu popNewString 
            Caption         =   "&String Value"
         End
         Begin VB.Menu popNewBinary 
            Caption         =   "&Binary Value"
         End
         Begin VB.Menu popNewDWORD 
            Caption         =   "&DWORD Value"
         End
      End
      Begin VB.Menu popSep2 
         Caption         =   "-"
      End
      Begin VB.Menu popCopyPath 
         Caption         =   "&Copy Key Name"
      End
      Begin VB.Menu popKillKey 
         Caption         =   "&Delete Key"
      End
      Begin VB.Menu popExport 
         Caption         =   "&Export from this key"
      End
   End
   Begin VB.Menu popData 
      Caption         =   "PopUpData"
      Visible         =   0   'False
      Begin VB.Menu popModify 
         Caption         =   "&Modify"
      End
      Begin VB.Menu poopSep3 
         Caption         =   "-"
      End
      Begin VB.Menu pop2New 
         Caption         =   "New"
         Begin VB.Menu pop2NewString 
            Caption         =   "&String Value"
         End
         Begin VB.Menu pop2NewBinary 
            Caption         =   "&Binary Value"
         End
         Begin VB.Menu pop2NewDWORD 
            Caption         =   "&DWORD Value"
         End
      End
      Begin VB.Menu popSep4 
         Caption         =   "-"
      End
      Begin VB.Menu popKillValue 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents cSplit   As clsSplitter 'splitter
Attribute cSplit.VB_VarHelpID = -1
Private Regy                As clsRegistryAccess 'registry thing

'global node object
Private nodX                As Node '(for adding nodes/keys)
Private lastNode            As Node 'saves last selected node
'global listitem object
Private lstX                As ListItem '(for adding values)
Private lastListItem        As ListItem
'in this string will be saved names of all loaded keys (speeds up code very much)
Private opened              As String
Private curSubkeys          As String
Private curValues           As String
Private nCount              As Long
Private sCount              As Long
Private bCount              As Long
Private dCount              As Long
Private Sub Form_Load()
    On Error GoTo errTrap ' just in case ;)
    
    Set Regy = New clsRegistryAccess 'init class module
    
    'first to see if we got any command line arguments
    If Len(Command) > 0 Then
        Dim cmd As String, ret As Long, par() As String
        cmd = Trim$(Command)
        
        If Left$(LCase$(cmd), 2) = "/s" Then 'silent import mode
            cmd = Mid$(cmd, 3, Len(cmd) - 2)
            cmd = Replace(cmd, Chr$(34), vbNullString) 'strip quotes
            ret = Regy.ImportFromReg(cmd)
        ElseIf Left$(LCase$(cmd), 2) = "/e" Then 'export mode
            cmd = Mid$(cmd, 3, Len(cmd) - 2)
            If InStr(1, cmd, Chr$(34)) > 0 Then
                par = Split(cmd, Chr$(34) & " " & Chr$(34))
                ret = Regy.ExportToReg(Replace(par(0), Chr$(34), ""), _
                                       Replace(par(1), Chr$(34), ""))
            Else
                
            End If
        Else 'normal import mode
            If (MsgBox("Are you sure you want to add" & vbNewLine & cmd & _
                vbNewLine & "registry information in registry database?", _
                                                vbYesNo, "Confirm") = vbYes) Then
                ret = Regy.ImportFromReg(cmd)
                If ret = 0 Then
                    MsgBox "Error importing file to registry.", vbExclamation, "Error"
                Else
                    MsgBox "Successufuly entered data into registry.", vbInformation, "Registry Editor"
                End If
            End If
        End If
        
        On Error Resume Next
        End 'close
    End If
    
    Set cSplit = New clsSplitter
    
    If InitXPStyles = False Then
        MsgBox "Error instalizing XP visual styles.", vbInformation, "Error"
    End If
    
    cSplit.Initialise picSplitter, Me  'init splitter
    
    Set nodX = tvwKeys.Nodes.Add(, , "COMP", "My Computer", 5) 'add 'My Computer' node
    nodX.Expanded = True 'and set it expanded
    
    Set lastNode = nodX
    
    'add main hkeys
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT", 1)
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CURRENT_USER", "HKEY_CURRENT_USER", 1)
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE", 1)
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_USERS", "HKEY_USERS", 1)
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG", 1)
    Set nodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_DYN_DATA", "HKEY_DYN_DATA", 1)
    'add empty nodes, so hkeys may be expandeable
    Set nodX = tvwKeys.Nodes.Add("HKEY_CLASSES_ROOT", tvwChild)
    Set nodX = tvwKeys.Nodes.Add("HKEY_CURRENT_USER", tvwChild)
    Set nodX = tvwKeys.Nodes.Add("HKEY_LOCAL_MACHINE", tvwChild)
    Set nodX = tvwKeys.Nodes.Add("HKEY_USERS", tvwChild)
    Set nodX = tvwKeys.Nodes.Add("HKEY_CURRENT_CONFIG", tvwChild)
    Set nodX = tvwKeys.Nodes.Add("HKEY_DYN_DATA", tvwChild)
    
    opened = ":"
    
    Exit Sub
errTrap:
    Dim msg As String
    msg = "An error ocured!" & vbCrLf
    msg = msg & "Description: " & Err.Description & String(2, vbCrLf)
    MsgBox msg, vbExclamation, "Error: " & Err.Number
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

'---------[ Resizing & Splitter thing... ]
Private Sub Form_Resize()
    On Error Resume Next
    Dim s As Single
    
    tvwKeys.Left = 3: tvwKeys.Top = 0
    
    s = Me.ScaleHeight + stbStatus.Height * mnuStatus.Checked
    'dunno why vb raises error here if s < 0???
    If s > 0 Then tvwKeys.Height = s
    
    tvwKeys.Width = picSplitter.Left
    picSplitter.Top = 0
    
    s = Me.ScaleHeight + stbStatus.Height * mnuStatus.Checked
    If s > 0 Then picSplitter.Height = s
    
    lvwData.Top = -1
    lvwData.Left = picSplitter.Left + 6
    
    s = Me.ScaleHeight + 2 + stbStatus.Height * mnuStatus.Checked
    If s > 0 Then lvwData.Height = s
    
    s = Me.ScaleWidth - picSplitter.Left - 7
    If s > 0 Then lvwData.Width = s
    
    Call txtOutput_Change
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cSplit.MouseMove X
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cSplit.MouseUp X
End Sub
Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cSplit.MouseDown X
End Sub
Private Sub cSplit_SplitComplete()
    Call Form_Resize
End Sub
'---------------------------------------

'---------[ Nodes... ]
Private Sub tvwKeys_Collapse(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    If Node.Key <> "COMP" Then Node.Image = 1 'set "close folder" image
    txtOutput.Text = Node.Key
    Call tvwKeys_NodeClick(Node)
End Sub
Private Sub tvwKeys_Expand(ByVal Node As MSComctlLib.Node)
    On Error GoTo errTrap
    Dim rKey       As String
    Dim keyCnt      As Long
    Dim Keys()      As String
    Dim k           As Long
    
    Screen.MousePointer = vbHourglass 'set pointer
    
    rKey = Node.Key
    
    'if this key wasn't loaded yet then load it!
    If InStr(1, opened, ":" & rKey & ":") < 1 Then

        keyCnt = Regy.EnumKeys(rKey, Keys()) 'get subkeys
        
        Node.Expanded = False
        DoEvents
        
        tvwKeys.Nodes.Remove Node.Child.Index 'removes empty node
        
        curSubkeys = ":"
        For k = 0 To keyCnt - 1 'for each subkey in subkeys
            'add new node
            Set nodX = tvwKeys.Nodes.Add(rKey, tvwChild, rKey & "\" & Keys(k), Keys(k), 1)
            curSubkeys = curSubkeys & Keys(k) & ":"
            'add empty node if this key has atleast 1 subkey
            If Regy.HaveSubkey(rKey & "\" & Keys(k)) Then
                Set nodX = tvwKeys.Nodes.Add(rKey & "\" & Keys(k), tvwChild)
            End If
        Next

        'remember this key (so we won't try to load it twice!)
        opened = opened & rKey & ":"
        
        Node.Sorted = True 'sort added keys ;)
        
        Node.Expanded = True 'on end, expande it!
        
        txtOutput.Text = rKey
    End If
    
    If Node.Key <> "COMP" Then Node.Image = 2 'set "open folder" image
    Screen.MousePointer = vbNormal
    
    Set lastListItem = Nothing
    mnuModify.Enabled = False
    mnuKillValue.Enabled = False
    
    Exit Sub
errTrap:
    Dim msg As String
    msg = "An error ocured!" & vbCrLf
    msg = msg & "Description: " & Err.Description & String(2, vbCrLf)
    MsgBox msg, vbExclamation, "Error: " & Err.Number
End Sub
Private Sub tvwKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu popKey
    End If
End Sub

Private Sub tvwKeys_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo errTrap
    Dim valCnt      As Long
    Dim valName()   As String
    Dim valData()   As Variant
    Dim k           As Long
    Dim tmp         As String
    
    Set lastNode = Node 'remmeber this node, so we can refresh it
    
    txtOutput.Text = Node.Key
    
    'enable menus
    mnuCopyName.Enabled = Node.Key <> "COMP"
    mnuNewKey.Enabled = Node.Key <> "COMP"
    mnuNewString.Enabled = Node.Key <> "COMP"
    mnuNewBinary.Enabled = Node.Key <> "COMP"
    mnuNewDWORD.Enabled = Node.Key <> "COMP"
    mnuKillKey.Enabled = Node.Key <> "COMP"
    
    lvwData.ListItems.Clear 'first clear all values from listview
    
    If Node.Key = "COMP" Then Exit Sub 'exit if is clicked on 'My Computer'

    'then add all string values from this key to listview
    valCnt = Regy.EnumValues(Node.Key, valName(), valData(), REG_SZ)
    
    curValues = ":(Default):"
    
    'add all data of string type
    For k = 0 To valCnt - 1
        If Len(valName(k)) > 0 Then
            Set lstX = lvwData.ListItems.Add(, , valName(k), , 3)
            lstX.Tag = REG_SZ
            lstX.ListSubItems.Add , , Chr$(34) & valData(k) & Chr$(34)
            curValues = curValues & valName(k) & ":"
        End If
    Next
    'now add '(Default)' value
    Set lstX = lvwData.ListItems.Add(, , "(Default)", , 3)
    lstX.Tag = REG_SZ
    'and set its data:
    tmp = Regy.ReadString(Node.Key, "")
    If Len(tmp) = 0 Then tmp = "(value not set)" Else tmp = Chr$(34) & tmp & Chr$(34)
    lstX.ListSubItems.Add , , tmp
    
    'reset arrays
    Erase valName
    Erase valData
    
    'now lets get all DWORD's
    valCnt = Regy.EnumValues(Node.Key, valName(), valData(), REG_DWORD)
    For k = 0 To valCnt - 1
        If Len(valName(k)) > 0 Then
            Set lstX = lvwData.ListItems.Add(, , valName(k), , 4)
            lstX.Tag = REG_DWORD
            lstX.ListSubItems.Add , , formatDWORD(valData(k))
            curValues = curValues & valName(k) & ":"
        End If
    Next
    'reset arrays
    Erase valName
    Erase valData

    'after that, we'll read all Binary data from this key
    valCnt = Regy.EnumValues(Node.Key, valName(), valData(), REG_BINARY)
    For k = 0 To valCnt - 1
        If Len(valName(k)) > 0 Then
            Set lstX = lvwData.ListItems.Add(, , valName(k), , 4)
            lstX.Tag = REG_BINARY
            tmp = valData(k)
            If Len(tmp) = 0 Then tmp = "(zero-length binary value)"
            lstX.ListSubItems.Add , , tmp
            curValues = curValues & valName(k) & ":"
        End If
    Next
    
    Set lastListItem = Nothing
    mnuModify.Enabled = False
    mnuKillValue.Enabled = False
    
    Exit Sub 'exit before error handler
errTrap:
    Dim msg As String
    msg = "An error ocured!" & vbCrLf
    msg = msg & "Description: " & Err.Description & String(2, vbCrLf)
    MsgBox msg, vbExclamation, "Error: " & Err.Number
End Sub

Private Sub txtExport_Change()
    Me.Caption = "Exporting - " & txtExport.Text
End Sub
Private Sub txtOutput_Change()
    On Error Resume Next
    Dim txt As String
    
    txt = lastNode.FullPath
    
    'i know, this not work....
    If Me.TextWidth(txt) > stbStatus.Width Then
        Do While Not Me.TextWidth(txt) = stbStatus.Width ' - Me.TextWidth("...")
            If Len(txt) > 1 Then
                txt = Left$(txt, Len(txt) - 1)
            Else
                Exit Do
            End If
        Loop
        txt = txt & "..."
    End If
    
    stbStatus.SimpleText = txt '& " [" & lastNode.Index & "]"
End Sub

'this will format DWORD value in same way as regedit does, e.g. '0x000055be (21950)'
Function formatDWORD(dValue As Variant) As String
    On Error GoTo errTrap
    
    formatDWORD = "0x" & Right$("00000000" & Hex$(dValue), 8) & " (" & dValue & ")"
    
errTrap:
End Function
'-------------------------------------------------

'----------[ Data View.... ]
Private Sub lvwData_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set lastListItem = Item
    mnuModify.Enabled = True
    mnuKillValue.Enabled = True
End Sub
Private Sub lvwData_DblClick()
    If IsObject(lastListItem) Then
        Call mnuModify_Click
    End If
End Sub
Private Sub lvwData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then 'right click
        Set lastListItem = lvwData.SelectedItem
        mnuModify.Enabled = True
        mnuKillValue.Enabled = True
        
        popModify.Enabled = IsObject(lastListItem)
        popKillValue.Enabled = IsObject(lastListItem)
        
        Me.PopupMenu popData, , , , popModify
    End If
End Sub
'-------------------------------------------------

'----------[ Menus.... ]
Private Sub mnuStatus_Click()
'shows/hides status bar
    mnuStatus.Checked = Not mnuStatus.Checked
    stbStatus.Visible = mnuStatus.Checked
    Call Form_Resize 'resize controls
End Sub
Private Sub mnuExit_Click()
    Unload Me 'Quit!
End Sub
Private Sub mnuRefresh_Click()
    'reload contents of selected key
    Call tvwKeys_NodeClick(lastNode)
End Sub
Private Sub mnuCopyName_Click() 'copy key path to clipboard
    On Error Resume Next
    Clipboard.Clear 'clear clipboard
    Clipboard.SetText lastNode.Key 'and set text
End Sub
Private Sub mnuKillKey_Click() 'delete key
    On Error Resume Next
    Dim Index As Integer, sRem As String
    
    If (InStr(1, lastNode.Key, "\") < 1) Then
        MsgBox "Cannot delete root keys!", vbExclamation, "Error"
    Else
        If (MsgBox("Are you sure you want to delete current key?", _
                            vbYesNo + vbQuestion, "Confirm") = vbYes) Then
            Regy.KillKey (lastNode.Key)
            
            sRem = lastNode.Key
            Index = lastNode.Index
            Call tvwKeys_NodeClick(lastNode.Parent) 'set focus to its parent
            tvwKeys.Nodes.Remove Index 'remove node
            opened = Replace(opened, sRem, "")
        End If
    End If
End Sub
Private Sub mnuKillValue_Click()
    On Error Resume Next
    
    If Not IsObject(lastListItem) Then Exit Sub
    
    If (MsgBox("Are you sure you want to delete this value?", _
                                vbYesNo + vbQuestion, "Confirm") = vbYes) Then
            Regy.KillValue lastNode.Key, lastListItem.Text
            curValues = Replace(curValues, ":" & lastListItem.Text & ":", ":")
            lvwData.ListItems.Remove lastListItem.Index
            Set lastListItem = Nothing
            mnuModify.Enabled = False
            mnuKillValue.Enabled = False
    End If
End Sub
Private Sub mnuNewKey_Click() 'create key
    On Error Resume Next

    lastNode.Expanded = True
    DoEvents
    
    nCount = nCount + 1
    Load frmNewKey
    frmNewKey.KeyName = "New Key #" & CStr(nCount)
    frmNewKey.Caption = "New Key"
    frmNewKey.Show vbModal
    Unload frmNewKey
    
check:
    If frmNewKey.Canceled = False Then
        If InStr(1, opened, ":" & lastNode.Key & "\" & frmNewKey.KeyName & _
            ":", 1) > 1 Or InStr(1, curSubkeys, frmNewKey.KeyName, 1) Then
            MsgBox "Key with that name allready exists. Try something else.", vbExclamation, "New Key"
            Load frmNewKey
            frmNewKey.KeyName = frmNewKey.KeyName
            frmNewKey.Show vbModal
            Unload frmNewKey
            GoTo check
        Else
            Set nodX = tvwKeys.Nodes.Add(lastNode.Key, tvwChild, lastNode.Key & "\" & frmNewKey.KeyName, frmNewKey.KeyName, 1)
            curSubkeys = curSubkeys & frmNewKey.KeyName & ":"
            Regy.CreateKey nodX.Key
            nodX.Selected = True
        End If
    End If
    
End Sub
Private Sub mnuNewString_Click() 'write string
    
    sCount = sCount + 1
    
    Load frmString
    frmString.StringName = "New String Value #" & CStr(sCount)
    frmString.Caption = "New String"
    frmString.Show vbModal
    Unload frmString
    
check:
    If frmString.Canceled = False Then
        If InStr(1, curValues, frmString.StringName, 1) Then
            MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "New String"
            Load frmString
            frmString.StringName = frmString.StringName
            frmString.Show vbModal
            Unload frmString
            GoTo check
        Else
            Set lstX = lvwData.ListItems.Add(, , frmString.StringName, , 3)
            lstX.Tag = REG_SZ
            lstX.ListSubItems.Add , , Chr$(34) & frmString.StringValue & Chr$(34)
            curValues = curValues & frmString.StringName & ":"
            Regy.WriteString lastNode.Key, lstX.Text, frmString.StringValue
        End If
    End If
    
End Sub
Private Sub mnuNewBinary_Click() 'write binary
    
    Dim tmp As String
    
    bCount = bCount + 1
    
    Load frmBinary
    frmBinary.BinaryName = "New Binary Value #" & CStr(bCount)
    frmBinary.Caption = "New Binary"
    frmBinary.Show vbModal
    Unload frmBinary
    
check:
    If frmBinary.Canceled = False Then
        If InStr(1, curValues, frmBinary.BinaryName, 1) Then
            MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "New Binary"
            Load frmBinary
            frmBinary.BinaryName = frmBinary.BinaryName
            frmBinary.Show vbModal
            Unload frmBinary
            GoTo check
        Else
            Set lstX = lvwData.ListItems.Add(, , frmBinary.BinaryName, , 4)
            lstX.Tag = REG_BINARY
            tmp = frmBinary.BinaryValue
            If Len(tmp) = 0 Then
                tmp = "(zero-length binary value)"
            Else
                tmp = setSpaces(tmp)
            End If
            lstX.ListSubItems.Add , , tmp
            curValues = curValues & frmBinary.BinaryName & ":"
            Regy.WriteBinary lastNode.Key, lstX.Text, frmBinary.BinaryValue
        End If
    End If
    
End Sub
Private Sub mnuNewDWORD_Click() 'write dword
    
    dCount = dCount + 1
    
    Load frmDWORD
    frmDWORD.DWORDName = "New DWORD Value #" & CStr(dCount)
    frmDWORD.Caption = "New DWORD"
    frmDWORD.Show vbModal
    Unload frmDWORD
    
check:
    If frmDWORD.Canceled = False Then
        If InStr(1, curValues, frmDWORD.DWORDName, 1) Then
            MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "New DWORD"
            Load frmDWORD
            frmDWORD.DWORDName = frmDWORD.DWORDName
            frmDWORD.Show vbModal
            Unload frmDWORD
            GoTo check
        Else
            Set lstX = lvwData.ListItems.Add(, , frmDWORD.DWORDName, , 4)
            lstX.Tag = REG_DWORD
            lstX.ListSubItems.Add , , formatDWORD(frmDWORD.DWORDValue)
            curValues = curValues & frmDWORD.DWORDName & ":"
            Regy.WriteDWORD lastNode.Key, lstX.Text, CLng(frmDWORD.DWORDValue)
        End If
    End If
    
End Sub
Private Sub mnuModify_Click() 'modify value
    Dim tmp As String, def As Boolean, tmp2 As String
    
    On Error GoTo errh 'will not raise error when compiled
    
    If IsObject(lastListItem) Then
        curValues = Replace(curValues, ":" & lastListItem.Text & ":", ":")
        Select Case lastListItem.Tag
            Case REG_SZ
                def = lastListItem.Text = "(Default)"
                tmp = lastListItem.ListSubItems(1).Text
                Load frmString
                frmString.StringName = lastListItem.Text
                frmString.StringValue = Mid$(tmp, 2, Len(tmp) - 2)
                frmString.Caption = "Modify Value"
                frmString.Text1.Locked = def
                frmString.Show vbModal
                Unload frmString
checkSZ:
                If frmString.Canceled = False Then
                    If InStr(1, curValues, frmString.StringName, 1) Then
                        MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "Modify String"
                        Load frmString
                        frmString.StringName = frmString.StringName
                        frmString.StringValue = frmString.StringValue
                        frmString.Text1.Locked = def
                        frmString.Show vbModal
                        Unload frmString
                        GoTo checkSZ
                    Else
                        tmp = frmString.StringName
                        tmp2 = frmString.StringValue
                        If def Then
                            tmp = ""
                            If Len(Trim$(tmp2)) = 0 Then
                                tmp2 = ""
                                lastListItem.ListSubItems(1).Text = "(value not set)"
                                Regy.KillValue lastNode.Key, ""
                            Else
                                lastListItem.ListSubItems(1).Text = Chr$(34) & tmp2 & Chr$(34)
                                Regy.WriteString lastNode.Key, tmp, tmp2
                            End If
                        Else
                            Regy.WriteString lastNode.Key, tmp, tmp2
                            lastListItem.ListSubItems(1).Text = Chr$(34) & tmp2 & Chr$(34)
                        End If
                        lastListItem.Text = frmString.StringName
                        
                    End If
                End If
            Case REG_BINARY
                tmp = lastListItem.ListSubItems(1).Text
                Load frmBinary
                frmBinary.BinaryName = lastListItem.Text
                frmBinary.BinaryValue = tmp
                frmBinary.Caption = "Modify Value"
                frmBinary.Show vbModal
                Unload frmBinary
checkBIN:
                If frmBinary.Canceled = False Then
                    If InStr(1, curValues, frmBinary.BinaryName, 1) Then
                        MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "Modify Binary"
                        Load frmBinary
                        frmBinary.BinaryName = frmBinary.BinaryName
                        frmBinary.BinaryValue = frmBinary.BinaryValue
                        frmBinary.Show vbModal
                        Unload frmBinary
                        GoTo checkBIN
                    Else
                        tmp = frmBinary.BinaryValue
                        If Len(Trim$(tmp)) = 0 Then
                            tmp = ""
                            lastListItem.ListSubItems(1).Text = "(zero-length binary value)"
                        Else
                            lastListItem.ListSubItems(1).Text = setSpaces(tmp)
                        End If
                        
                        lastListItem.Text = frmBinary.BinaryName
                        Regy.WriteBinary lastNode.Key, frmBinary.BinaryName, tmp
                    End If
                End If
            Case REG_DWORD
                tmp = lastListItem.ListSubItems(1).Text
                Load frmDWORD
                frmDWORD.DWORDName = lastListItem.Text
                frmDWORD.DWORDValue = Mid$(tmp, 13, Len(tmp) - 13) 'only dec
                frmDWORD.Option1.Value = False
                frmDWORD.Option2.Value = True 'force decimal display
                frmDWORD.Caption = "Modify Value"
                frmDWORD.Show vbModal
                Unload frmDWORD
checkDWORD:
                If frmDWORD.Canceled = False Then
                    If InStr(1, curValues, frmDWORD.DWORDName, 1) Then
                        MsgBox "Value with that name allready exists. Try something else.", vbExclamation, "Modify DWORD"
                        Load frmDWORD
                        frmDWORD.DWORDName = frmDWORD.DWORDName
                        frmDWORD.DWORDValue = frmDWORD.DWORDValue
                        frmDWORD.Show vbModal
                        Unload frmDWORD
                        GoTo checkDWORD
                    Else
                        tmp = frmDWORD.DWORDValue
                        If Len(Trim$(tmp)) = 0 Then tmp = "0"
                        lastListItem.Text = frmDWORD.DWORDName
                        lastListItem.ListSubItems(1).Text = formatDWORD(CLng(tmp))
                        Regy.WriteDWORD lastNode.Key, frmDWORD.DWORDName, CLng(tmp)
                    End If
                End If
        End Select
        
        curValues = curValues & lastListItem.Text & ":"
        
    End If
errh:
End Sub
Private Sub mnuExport_Click()
    On Error Resume Next
    
    If lastNode.Key = "COMP" Then Exit Sub
    
    Dim ret As Long
    
    Load frmExport
    frmExport.StartKey = lastNode.Key
    frmExport.Show vbModal
    Unload frmExport
    
    If frmExport.Canceled = False Then
        Screen.MousePointer = vbHourglass
        ret = Regy.ExportToReg(frmExport.StartKey, frmExport.RegFile, frmExport.Include, txtExport)
        Screen.MousePointer = vbNormal
        Me.Caption = "Registry Editor"
        If ret = 0 Then 'error
            MsgBox "Error exporting to file.", vbCritical, "Error"
        End If
    End If
    
End Sub
Private Sub mnuImport_Click()
    On Error Resume Next
    
    Dim ret As Long
    
    cdImport.ShowOpen
    
    If Dir(cdImport.FileName) = "" Then
        MsgBox "File not exists!", vbExclamation, "Error"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    ret = Regy.ImportFromReg(cdImport.FileName)
    Screen.MousePointer = vbNormal
    
    If ret = 0 Then 'error
        MsgBox "Error importing file.", vbCritical, "Error"
    End If
End Sub
Private Sub popModify_Click()
    Call mnuModify_Click
End Sub
Private Sub popKillValue_Click()
    Call mnuKillValue_Click
End Sub
Private Sub pop2NewString_Click()
    Call mnuNewString_Click
End Sub
Private Sub pop2NewBinary_Click()
    Call mnuNewBinary_Click
End Sub
Private Sub pop2NewDWORD_Click()
    Call mnuNewDWORD_Click
End Sub
Private Sub popNewString_Click()
    Call mnuNewString_Click
End Sub
Private Sub popNewBinary_Click()
    Call mnuNewBinary_Click
End Sub
Private Sub popNewDWORD_Click()
    Call mnuNewDWORD_Click
End Sub
Private Sub popNewKey_Click()
    Call mnuNewKey_Click
End Sub
Private Sub popCopyPath_Click()
    Call mnuCopyName_Click
End Sub
Private Sub popKillKey_Click()
    Call mnuKillKey_Click
End Sub
Private Sub popExport_Click()
    Call mnuExport_Click
End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub
Private Function setSpaces(sIn As String) As String
    Dim k As Long
    
    sIn = Replace(sIn, " ", "")
    If Len(sIn) Mod 2 <> 0 Then sIn = Left$(sIn, Len(sIn) - 1)
    
    For k = 1 To Len(sIn) Step 2
        setSpaces = setSpaces & Mid$(sIn, k, 2) & " "
    Next
    
    setSpaces = Left$(setSpaces, Len(setSpaces) - 1)
End Function

