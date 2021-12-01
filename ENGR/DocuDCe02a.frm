VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form DocuDCe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document List"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvList 
      Height          =   3615
      Left            =   5280
      TabIndex        =   29
      Top             =   3840
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   6376
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvDoc 
      Height          =   3615
      Left            =   180
      TabIndex        =   28
      Top             =   3840
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   6376
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDoc 
      Height          =   315
      Left            =   8100
      Picture         =   "DocuDCe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "New Part Numbers"
      Top             =   3120
      Width           =   350
   End
   Begin VB.TextBox txtCmt 
      Height          =   1185
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Tag             =   "9"
      Text            =   "DocuDCe02a.frx":049B
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   2100
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CheckBox chkUpChild 
      Height          =   255
      Left            =   7860
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtApp 
      Height          =   285
      Left            =   2100
      TabIndex        =   22
      Tag             =   "2"
      ToolTipText     =   "Approval Name"
      Top             =   1680
      Width           =   2085
   End
   Begin VB.ComboBox txtAte 
      Height          =   315
      Left            =   6300
      TabIndex        =   21
      Tag             =   "4"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Available Documents for"
      Height          =   1095
      Left            =   2100
      TabIndex        =   15
      Top             =   2280
      Width           =   5535
      Begin VB.TextBox txtSearchDoc 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         ToolTipText     =   "Type in a partial document number to narrow down your search"
         Top             =   600
         Width           =   4455
      End
      Begin VB.ComboBox cmbCls 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "9"
         ToolTipText     =   "Select Class From List"
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Number"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblCls 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Top             =   240
         Width           =   2640
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "DocuDCe02a.frx":04A2
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6300
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Text            =   "cmbRev"
      ToolTipText     =   "Document Revision-Select From List"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   420
      Left            =   7740
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Update The Current Selections And, Optionally, The Part Number"
      Top             =   600
      Width           =   870
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<=>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4620
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Highlite Selection And Press To Move"
      Top             =   5340
      Width           =   495
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "9"
      ToolTipText     =   "Select Part Number For List (Type 7 For Service)"
      Top             =   960
      Width           =   3345
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "9"
      Text            =   "cmbTyp"
      ToolTipText     =   "Select Type From List"
      Top             =   600
      Width           =   2000
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   420
      Left            =   7740
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5820
      Top             =   240
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   7710
      FormDesignWidth =   9675
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Approved By"
      Height          =   285
      Index           =   1
      Left            =   900
      TabIndex        =   24
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Date"
      Height          =   285
      Index           =   5
      Left            =   5460
      TabIndex        =   23
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Assigned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   9
      Left            =   5580
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   4
      Left            =   5580
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6300
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   780
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2100
      TabIndex        =   7
      Top             =   1320
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "List Type"
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "DocuDCe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'9/1/04 omit tools
'10/3/06 Corrected DOC vs BOM Revision
Option Explicit
Dim bGoodPart As Byte
Dim bOnLoad As Byte
Dim bListChg As Byte
Dim bFormActivated As Byte
Dim bDocSec As Byte
Dim bCrtNewRev As Byte



Dim sClass As String
Dim sDocument As String
Dim sOldPart As String
Dim sRevision As String
Dim sSheet As String
Private doc As ClassDoc
Private priorDocListSql As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private Const CB_FINDSTRING = &H14C
Private Const CB_SHOWDROPDOWN = &H14F
Private Const LB_FINDSTRING = &H18F
Private Const CB_ERR = (-1)

Private Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) _
    As Long

Private Sub cmbCls_Click()
    If bFormActivated Then
        GetSomeClass
        FillDocuments
   End If
End Sub


Private Sub cmbCls_GotFocus()
   sOldPart = cmbPrt
   
End Sub


Private Sub cmbCls_LostFocus()
   'On Error Resume Next
   'If cmbCls = "" Then cmbCls = cmbCls.List(0)
   'GetSomeClass
   'FillDocuments
   
End Sub
Private Sub cmbPrt_GotFocus()
    SendMessage cmbPrt.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End Sub

Private Sub cmbPrt_KeyPress(KeyAscii As Integer)

'    Dim CB As Long
'    Dim FindString As String
'
'    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
'
'    If cmbPrt.SelLength = 0 Then
'        FindString = cmbPrt.Text & Chr$(KeyAscii)
'    Else
'        FindString = Left$(cmbPrt.Text, cmbPrt.SelStart) & Chr$(KeyAscii)
'    End If
'
'    SendMessage cmbPrt.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
'
'    CB = SendMessage(cmbPrt.hwnd, CB_FINDSTRING, -1, ByVal FindString)
'
'    If CB <> CB_ERR Then
'        cmbPrt.ListIndex = CB
'        cmbPrt.SelStart = Len(FindString)
'        cmbPrt.SelLength = Len(cmbPrt.Text) - cmbPrt.SelStart
'    End If
'
'    KeyAscii = 0
    
End Sub


Private Sub cmbPrt_Click()
   Dim bResponse As Byte
   If Not bOnLoad Then
      If bListChg = 1 Then
         bResponse = MsgBox("You haven't updated the part list for part " & sOldPart & "." & vbCrLf _
                     & "Update the list now?", ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            UpdateThisList sOldPart
         Else
            cmdUpd.Enabled = False
         End If
      End If
   End If
   FillDocumentRevisions Me
   cmdUpd.Enabled = False
   bGoodPart = GetPart()
   GetDocumentList
End Sub


Private Sub cmbPrt_LostFocus()
   sOldPart = cmbPrt
   bGoodPart = GetPart()
   FillDocumentRevisions Me
   If bGoodPart Then GetDocumentList      'this causes it
End Sub

Private Sub cmbRev_Click()
   GetDocumentList
   SetApproval
End Sub

Private Sub cmbRev_LostFocus()
   On Error GoTo DiaErr1
   
   Dim b As Byte
   Dim iList As Integer
   
   'cmbRev = CheckLen(cmbRev, 6)
   For iList = 0 To (cmbRev.ListCount - 1)
      If cmbRev = cmbRev.list(iList) Then b = True
   Next
   If Not b Then
      
      If (bDocSec = True) Then
         Dim bret As Boolean
         bret = CheckForDocLstSec(cmbRev.Text, False)
         If (bret = False) Then
            Exit Sub
         End If
      End If
      
      If MsgBox("Create a new revision " & cmbRev & " of the document list" & vbCrLf _
         & "for " & cmbPrt & "?", vbQuestion + vbYesNo) <> vbYes Then
         Exit Sub
      End If
      
      If Me.lvList.ListItems.count > 0 Then
         If MsgBox("Copy documents from previous list?", vbQuestion + vbYesNo) <> vbYes Then
            lvList.ListItems.Clear
         End If
      End If
      
      If lvList.ListItems.count > 0 Then
         UpdateThisList
      End If
      
      CreateDocLstHeader
      ResetApproval
      'Beep
      'If cmbRev.ListCount > 0 Then cmbRev = cmbRev.List(0)
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "cmbRev_LostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub cmbTyp_Click()
   Dim bResponse As Byte
   If Val(Left(cmbTyp, 1)) = 1 Then
      cmbRev.Visible = True
      z1(9).Visible = True
   Else
      cmbRev.Visible = False
      z1(9).Visible = False
   End If
   FillParts
   If Not bOnLoad Then
      If cmdUpd.Enabled Then
         bResponse = MsgBox("You Haven't Updated The Current List." & vbCrLf _
                     & "Update The List Now?", ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            cmdUpd_Click
         Else
            cmdUpd.Enabled = False
         End If
      End If
   End If
   
End Sub

Private Sub cmbTyp_LostFocus()
   If Val(Left(cmbTyp, 1)) = 0 Then cmbTyp = cmbTyp.list(0)
   
End Sub


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If Not bOnLoad Then
      If bListChg = 1 Then
         bResponse = MsgBox("You Haven't Updated The Current List." & vbCrLf _
                     & "Update The List Now?", ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            cmdUpd_Click
            Exit Sub
         Else
            bListChg = 0
         End If
      End If
   End If
   MouseCursor 13
   Unload Me
   
End Sub

Private Sub cmdDoc_Click()
   
   DocuDCe01a.chkFromDocLst = vbChecked
   DocuDCe01a.Show
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3302
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub cmdLst_Click()
   Dim b As Byte
   Dim iList As Integer
   Dim lRow As Long
'   If lvdoc.listitems.count > 0 Then
'      If Left(lstDoc.List(0), 5) = "*** N" Then
'         cmdLst.Enabled = False
'         Exit Sub
'      End If
'   End If
'   On Error Resume Next
   For lRow = lvDoc.ListItems.count To 1 Step -1
      If lvDoc.ListItems(lRow).Selected Then
         For iList = 1 To lvList.ListItems.count
            'If lvDoc.ListItems(lRow) = lvList.ListItems(iList) Then b = True
            If AreListItemsEqual(lvDoc.ListItems(lRow), lvList.ListItems(iList)) Then
               b = True
            End If
         Next
         If Not b Then
            cmdUpd.Enabled = True
            CopyListItem lvDoc.ListItems(lRow), lvList
            lvDoc.ListItems.Remove lRow
         End If
         b = False
      End If
   Next
   For lRow = lvList.ListItems.count To 1 Step -1
      If lvList.ListItems(lRow).Selected Then
         cmdUpd.Enabled = True
         If lvList.ListItems(lRow).SubItems(3) = cmbCls Then
            'CopyListItem lvList.ListItems(lRow), lvDoc
         End If
         lvList.ListItems.Remove lRow
      End If
   Next
   bListChg = 1
   'cmdLst.Enabled = False

End Sub

Private Sub cmdLst_GotFocus()
   cmdUpd.Enabled = True
   
End Sub




Private Sub cmdUpd_Click()
   
   If ((bDocSec = True) And (Trim(txtApp) <> "")) Then
            
      Dim bret As Boolean
      Dim b As Byte
      Dim iList As Integer
      Dim strLastRev As String
      
      strLastRev = cmbRev.list(cmbRev.ListCount - 1)
      bret = CheckForDocLstSec(strLastRev, True)
      If (bret = True) Then
         
         For iList = 1 To cmbRev.ListCount
            If cmbRev = cmbRev.list(iList) Then b = True
         Next
         
         If Not b Then
            If MsgBox("Create a new revision " & cmbRev & " of the document list" & vbCrLf _
               & "for " & cmbPrt & "?", vbQuestion + vbYesNo) <> vbYes Then
               Exit Sub
            End If
            
            If Me.lvList.ListItems.count > 0 Then
               If MsgBox("Copy documents from previous list?", vbQuestion + vbYesNo) <> vbYes Then
                  lvList.ListItems.Clear
               End If
            End If
            
            If lvList.ListItems.count > 0 Then
               UpdateThisList
            End If
            
            ResetApproval
         End If
      End If
   Else
      ' If
      UpdateThisList
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Set doc = New ClassDoc
      
      cmbTyp.AddItem "1 - Parts                   " 'Type 1 Bom Revision
      cmbTyp.AddItem "2 - Service Parts           " 'Type 2"
      cmbTyp = cmbTyp.list(0)
      
'      lvDoc.ColumnHeaders.Clear
'      lvDoc.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'      lvDoc.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'      lvDoc.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'      lvDoc.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width
'
'      lvList.ColumnHeaders.Clear
'      lvList.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'      lvList.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'      lvList.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'      lvList.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width

      doc.FillClasses Me.cmbCls, True
      cmbCls.AddItem ("**ALL**")
'      cmbCls_Click
      
      'FillClasses
      FillParts
      bDocSec = GetDocLstSecurity
      bOnLoad = 0
      bListChg = 0
      
   End If
   MouseCursor 0
    bFormActivated = 1
    cmbCls = cmbCls.list(0) 'Force to reload

End Sub

Private Sub Form_Load()
    bFormActivated = 0
'   lvDoc.ColumnHeaders.Clear
'   lvDoc.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width
   
   lvDoc.View = lvwReport '@@@ new
   lvDoc.ColumnHeaders.Clear
   lvDoc.ColumnHeaders.Add , , "Document", 0.45 * lvDoc.Width
   lvDoc.ColumnHeaders.Add , , "Rev", 0.15 * lvDoc.Width
   lvDoc.ColumnHeaders.Add , , "Sht", 0.15 * lvDoc.Width
   lvDoc.ColumnHeaders.Add , , "Class", 0.15 * lvDoc.Width
   
'   lvList.ColumnHeaders.Clear
'   lvList.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width

   lvList.View = lvwReport
   lvList.ColumnHeaders.Clear
   lvList.ColumnHeaders.Add , , "Document", 0.45 * lvDoc.Width
   lvList.ColumnHeaders.Add , , "Rev", 0.15 * lvDoc.Width
   lvList.ColumnHeaders.Add , , "Sht", 0.15 * lvDoc.Width
   lvList.ColumnHeaders.Add , , "Class", 0.15 * lvDoc.Width
   
   FormLoad Me
   FormatControls
'do in form activate
'   cmbTyp.AddItem "1 - Parts                   " 'Type 1 Bom Revision
'   cmbTyp.AddItem "2 - Service Parts           " 'Type 2"
'   cmbTyp = cmbTyp.List(0)
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   'Clean house
   sSql = "DELETE FROM DlstTable WHERE DLSREF=''"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   MouseCursor 0
   Set DocuDCe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

''Private Sub FillClasses()
''   Dim RdoCmb As ADODB.Recordset
''   Dim bByte As Byte
''
''   On Error GoTo DiaErr1
''   sSql = "Qry_FillDocumentClasses"
''   bSqlRows = clsADOCon.GetDataSet(sSql,RdoCmb, ES_FORWARD)
''   If bSqlRows Then
''      With RdoCmb
''         bByte = True
''         'If Trim(sLastDocClass) = "" Then
''         '   cmbCls = "" & Trim(!DCLNAME)
''         'Else
''         '   cmbCls = sLastDocClass
''         'End If
''         Do Until .EOF
''            AddComboStr cmbCls.hWnd, "" & Trim(!DCLNAME)
''            .MoveNext
''         Loop
''         ClearResultSet RdoCmb
''      End With
''   End If
''   If Not bByte Then
''      On Error Resume Next
''      MouseCursor 0
''      MsgBox "Please Install At Least One Document Class.", vbExclamation, Caption
''      Unload Me
''   Else
''      On Error GoTo 0
''      GetSomeClass
''   End If
''   Set RdoCmb = Nothing
''   Exit Sub
''
''DiaErr1:
''   sProcName = "fillclass"
''   CurrError.Number = Err.Number
''   CurrError.Description = Err.Description
''   DoModuleErrors Me
''
''End Sub
''
Private Sub GetSomeClass()
   Dim RdoCls As ADODB.Recordset
   Dim sClass As String
   
   sClass = Compress(cmbCls)
   On Error GoTo DiaErr1
   sSql = "Qry_GetDocumentClass '" & sClass & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCls, ES_KEYSET)
   If bSqlRows Then
      With RdoCls
         'cmbCls = "" & Trim(!DCLNAME)
         lblCls = "" & Trim(!DCLDESC)
         ClearResultSet RdoCls
      End With
   Else
      lblCls = ""
   End If
   Set RdoCls = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsomecl"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillParts()
   Dim iType As Integer
   cmbPrt.Clear
   iType = Val(Left(cmbTyp, 1))
   If iType < 1 Then iType = 1
   On Error GoTo DiaErr1
   MouseCursor 13
   If iType = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
             & "WHERE (PATOOL=0 AND PAPRODCODE<>'BID') ORDER BY PARTREF"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable WHERE " _
             & "(PALEVEL=7 AND PAPRODCODE<>'BID') ORDER BY PARTREF"
   End If
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      'MM cmbPrt = cmbPrt.List(0)
      cmbCls_Click
      cmbPrt_Click
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillDocuments()
   Dim RdoCmb As ADODB.Recordset
   'Dim sClass As String
   
   sClass = Compress(cmbCls)
   'lstDoc.Clear
   lvDoc.ListItems.Clear
    lvDoc.View = lvwReport '@@@ new
'   lvDoc.ColumnHeaders.Clear
'   lvDoc.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'   lvDoc.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width
   
   On Error GoTo DiaErr1
   'Retrieve only valid documents
   sSql = "SELECT DOREF,DONUM,DOREV,DOSHEET,DOCLASS FROM DdocTable " _
          & "WHERE (DOOBSOLETE IS NULL OR " _
          & "DOOBSOLETE>'" & Format(ES_SYSDATE, "mm/dd/yyyy") & "') AND " _
          & "(DOEFFECTIVE IS NULL OR DOEFFECTIVE <='" & Format(ES_SYSDATE, "mm/dd/yyyy") & "')"
   If (sClass <> "**ALL**") Then sSql = sSql & " AND DOCLASS='" & sClass & "'"
   If (Compress(txtSearchDoc) <> "") Then sSql = sSql & " AND DONUM LIKE '" & txtSearchDoc.Text & "%'"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            
'            Dim li as mscomctllib.listitem
'            Set li = lvDoc.ListItems.Add
'            li.Text = Trim(!DONUM)
'            li.SubItems(1) = Trim(!DOREV)
'            li.SubItems(2) = Trim(!DOSHEET)
'            li.SubItems(3) = Trim(!DOCLASS)
'
            lvDoc.ListItems.Add , , Trim(!DONUM)
            lvDoc.ListItems.item(lvDoc.ListItems.count).SubItems(1) = Trim(!DOREV)
            lvDoc.ListItems.item(lvDoc.ListItems.count).SubItems(2) = Trim(!DOSHEET)
            lvDoc.ListItems.item(lvDoc.ListItems.count).SubItems(3) = Trim(!DOCLASS)
            
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      'cmdLst.Enabled = False
      'lstDoc.AddItem "*** No Matching Documents Found ***"
   End If
   Set RdoCmb = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "filldocum"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   MouseCursor 0
End Sub

'Private Sub lvList_Click()
'   On Error Resume Next
'   If lstCur.Selected(lstCur.ListIndex) Then
'      cmdLst.Enabled = True
'   Else
'      cmdLst.Enabled = False
'   End If
'
'End Sub
'
Private Sub lvList_DblClick()
   cmdLst_Click
   
End Sub


Private Sub lvList_GotFocus()
   Dim lList As Long
   'If lvList.ListItems.Count > 0 Then cmdUpd.Enabled = True
   'cmdLst.Enabled = False
   sOldPart = cmbPrt

   For lList = 1 To lvDoc.ListItems.count
      lvDoc.ListItems(lList).Selected = False
   Next
End Sub


Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then cmdLst_Click
   
End Sub

Private Sub lvList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then GetThisDocument 1, lvList.HitTest(X, Y)
   
End Sub

Private Sub lvDoc_Click()
'   On Error Resume Next
'   If lvDoc.ListItems.Count > 0 Then
'      cmdLst.Enabled = True
'   Else
'      cmdLst.Enabled = False
'      Exit Sub
'   End If
   
'   If lstDoc.Selected(lstDoc.ListIndex) Then
'      cmdLst.Enabled = True
'   Else
'      cmdLst.Enabled = False
'   End If
   
End Sub

Private Sub lvDoc_DblClick()
   cmdLst_Click
   
End Sub


Private Sub lvDoc_GotFocus()
   Dim lList As Long
   'cmdLst.Enabled = False
   sOldPart = cmbPrt
   For lList = 1 To lvList.ListItems.count
      lvList.ListItems(lList).Selected = False
   Next
 
End Sub


Private Sub lvDoc_KeyDown(KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyReturn Then cmdLst_Click
   
End Sub

Private Sub lvDoc_LostFocus()
   'cmdLst.Enabled = False
   
End Sub


Private Sub lvDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      GetThisDocument 0, lvDoc.HitTest(X, Y)
   End If
End Sub



Private Sub GetThisDocument(bList As Byte, li As mscomctllib.ListItem)
   Dim RdoCmb As ADODB.Recordset
'   If bList Then
'      ParseListItem lvList.ListItems(lstCur.ListIndex)
'   Else
'      ParseListItem lstDoc.List(lstDoc.ListIndex)
'   End If
   ParseListItem li
   On Error GoTo DiaErr1
   sSql = "SELECT DODESCR,DOCUST,DOCLASS,DOOBSOLETE FROM DdocTable WHERE " _
          & "DOCLASS='" & sClass & "' AND DOREF='" & sDocument & "' AND DOREV='" _
          & sRevision & "' AND DOSHEET='" & sSheet & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         On Error Resume Next
         InfoShow.lblDsc = "" & Trim(!DODESCR) & ", Cust: " & Trim(!DOCUST) _
                           & vbCrLf & "Class: " & Trim(!DOCLASS) _
                           & ", Obsolete: " & Format(!DOOBSOLETE, "mm/dd/yy")
         ClearResultSet RdoCmb
         InfoShow.Show
         If bList Then InfoShow.Left = InfoShow.Left + 3000
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisdoc"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ParseListItem(li As mscomctllib.ListItem)
   'Dim sString As String
   On Error GoTo DlistPl1
   
   'sString = sLisTItem
'   sDocument = Left$(sString, 22)
'   sRevision = Mid$(sString, 24, 6)
'   sSheet = Right$(sString, 6)
'   sDocument = Compress(sDocument)
'   sRevision = Compress(sRevision)
'   sSheet = Compress(sSheet)
   
   sDocument = li.Text
   sRevision = li.SubItems(1)
   sSheet = li.SubItems(2)
   sClass = li.SubItems(3)
   Exit Sub
   
DlistPl1:
   Resume DlistPl2
DlistPl2:
   sDocument = ""
   sRevision = ""
   sSheet = ""
   On Error GoTo 0
   
End Sub

Private Sub UpdateThisList(Optional PartNum As String)
   Dim bResponse As Byte
   Dim iRow As Integer
   Dim iType As Integer
   Dim sPartNum As String
   Dim sPartRev As String
   
   iType = Val(Left(cmbTyp, 1))
   If iType < 1 Then iType = 1
   If PartNum = "" Then
      sPartNum = Compress(cmbPrt)
   Else
      sPartNum = Compress(PartNum)
   End If
   
   sPartRev = Compress(cmbRev)
   sOldPart = ""
   cmdUpd.Enabled = False
   
   'if different revision, ask whether to change
   sSql = "select PADOCLISTREF from PartTable where PARTREF = '" & "'"
   Dim rdo As ADODB.Recordset
   bResponse = vbNo
   If clsADOCon.GetDataSet(sSql, rdo) Then
      If rdo!PADOCLISTREF <> sPartRev Then
         bResponse = MsgBox("Update The Part Number To This Revision As Well?", _
            ES_YESQUESTION, Caption)
      End If
   End If
   
   MouseCursor 13
   
   'Remove current listing and rebuild
   clsADOCon.BeginTrans
   sSql = "DELETE FROM DlstTable WHERE DLSREF='" & sPartNum & "' " _
          & "AND DLSREV='" & sPartRev & "' AND DLSTYPE=" & iType & " "
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   On Error GoTo DlistUl1
   For iRow = 1 To lvList.ListItems.count
      ParseListItem lvList.ListItems(iRow)
      'sClass = lstCls.List(iRow)
      'sClass = Compress(sClass)
      sSql = "INSERT INTO DlstTable (DLSREF,DLSREV,DLSTYPE," _
             & "DLSDOCREF,DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS) VALUES('" _
             & sPartNum & "','" & sPartRev & "'," & iType & ",'" _
             & Compress(sDocument) & "','" & sRevision & "','" & sSheet & "','" _
             & sClass & "')"
'Debug.Print sSql
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   Next
   MouseCursor 0
   If bResponse = vbYes Then
      sSql = "UPDATE PartTable SET PADOCLISTREF='" & sPartNum & "'," _
             & "PADOCLISTREV='" & sPartRev & "' WHERE " _
             & "PARTREF='" & sPartNum & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
   clsADOCon.CommitTrans
   
   'if new revision, add it to the combobox
   Dim i As Integer
   For i = 0 To cmbRev.ListCount - 1
      If cmbRev.list(i) = sPartRev Then
         Exit For
      End If
   Next
   If i = cmbRev.ListCount Then
      cmbRev.AddItem sPartRev
   End If
   
   If bResponse = vbYes Then
      MsgBox "Drawing List And Part Number Were Successfully Updated.", vbInformation, Caption
   Else
      MsgBox "Drawing List Successfully Was Updated.", vbInformation, Caption
   End If
   bListChg = 0
   Exit Sub
   
DlistUl1:
   clsADOCon.RollbackTrans
   bListChg = 0
   CurrError.Description = Err.Description
   Resume DlistUl2
DlistUl2:
   On Error GoTo 0
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf _
      & "Couldn't Update List.", vbExclamation, Caption
   
End Sub

Private Sub GetDocumentList()
   Dim RdoLst As ADODB.Recordset
   Dim iType As Integer
   Dim sPartNum As String
   Dim sRevision As String
   
   cmdUpd.Enabled = False
   iType = Left(cmbTyp, 1)
   If iType < 1 Then iType = 1
   sPartNum = Compress(cmbPrt)
   If iType = 1 Then sRevision = Compress(cmbRev)
   
   On Error GoTo DiaErr1
   
   'if no change, just leave everything alone
   sSql = "SELECT DISTINCT DLSTYPE,DLSDOCREF,DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS,DOREF,DONUM" & vbCrLf _
      & "FROM DlstTable" & vbCrLf _
      & "JOIN DdocTable ON DLSDOCREF=DOREF" & vbCrLf _
      & "WHERE DLSREF='" & sPartNum & "'" & vbCrLf _
      & "AND DLSTYPE=" & iType & " AND DLSREV='" & sRevision & "'"
   If sSql = priorDocListSql Then
      Exit Sub
   Else
      priorDocListSql = sSql
   End If
   
   lvList.ListItems.Clear
'   lvList.ColumnHeaders.Clear
'   lvList.ColumnHeaders.Add , , "Document", 0.3 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Rev", 0.08 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Sht", 0.08 * lvDoc.Width
'   lvList.ColumnHeaders.Add , , "Class", 0.1 * lvDoc.Width
   
'   cmdUpd.Enabled = False
'   iType = Left(cmbTyp, 1)
'   If iType < 1 Then iType = 1
'   sPartNum = Compress(cmbPrt)
'   If iType = 1 Then sRevision = Compress(cmbRev)
'
'   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT DLSTYPE,DLSDOCREF,DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS," _
'          & "DOREF,DONUM FROM DlstTable,DdocTable WHERE DLSREF='" & sPartNum & "' " _
'          & "AND DLSTYPE=" & iType & " AND DLSDOCREF=DOREF AND DLSREV='" & sRevision & "' " & vbCrLf _

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst)
   If bSqlRows Then
      With RdoLst
         Do Until .EOF
            'lvlist.ListItems.Add "" & Left(!DONUM, 22) & vbTab & !DLSDOCREV & vbTab & !DLSDOCSHEET
            'lstCls.AddItem "" & Trim(!DLSDOCCLASS)
            
            Dim li As mscomctllib.ListItem
            Set li = lvList.ListItems.Add
            li.Text = Trim(!DONUM)
            li.SubItems(1) = Trim(!DLSDOCREV)
            li.SubItems(2) = Trim(!DLSDOCSHEET)
            li.SubItems(3) = Trim(!DLSDOCCLASS)
            
            .MoveNext
         Loop
         ClearResultSet RdoLst
      End With
   End If
   bListChg = 0
   'If lvList.ListItems.Count > 0 Then cmdUpd.Enabled = True Else _
                                                 cmdUpd.Enabled = False
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getdocumlst"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim iType As Integer
   Dim sGetPart As String
   Dim sRevision As String
   sGetPart = Compress(cmbPrt)
   
   On Error GoTo DiaErr1
   iType = Val(Left(cmbTyp, 1))
   If iType < 1 Then iType = 1
   If iType = 1 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM PartTable WHERE " _
             & "PARTREF='" & sGetPart & "'"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL FROM PartTable WHERE " _
             & "PARTREF='" & sGetPart & "' AND PALEVEL=7"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         'cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         If iType = 1 Then sRevision = "" & Trim(!PABOMREV)
         GetPart = True
      End With
   Else
      MsgBox "That Part Number Wasn't Found. " & vbCrLf _
         & "Part Type 7 For Service Items.", vbExclamation, Caption
      cmbRev = ""
      lblDsc = ""
      lblTyp = ""
      GetPart = False
      On Error Resume Next
      cmbPrt.SetFocus
      cmbPrt = cmbPrt.list(0)
   End If
   If GetPart Then
      If iType = 1 Then
         'cmbRev.Clear
         'FillBomhRev cmbPrt
         FillDocumentRevisions Me
         
      End If
   End If
   On Error Resume Next
   RdoPrt.Close
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CopyListItem(li As mscomctllib.ListItem, ByRef lvTo As mscomctllib.ListView)
   ' copy a listview item to another listviews
   
   Dim liNew As mscomctllib.ListItem
'   Set liNew = lvTo.ListItems.Add
'   liNew.Text = li.Text
'   Dim i As Integer
'   For i = 1 To lvTo.ColumnHeaders.count - 1
'      liNew.SubItems(i) = li.SubItems(i)
'   Next
   Set liNew = lvTo.ListItems.Add(, , li.Text)
   'liNew.Text = li.Text
   Dim i As Integer
   For i = 1 To lvTo.ColumnHeaders.count - 1
      liNew.SubItems(i) = li.SubItems(i)
   Next
   liNew.Selected = False
End Sub

Private Sub CopyListItem2(li As mscomctllib.ListItem, lvTo As ListView)
   ' copy a listview item to another listviews
   
   Dim liNew As mscomctllib.ListItem
   Set liNew = lvTo.ListItems.Add
   liNew.Text = li.Text
   Dim i As Integer
   For i = 1 To lvTo.ColumnHeaders.count - 1
      liNew.SubItems(i) = li.SubItems(i)
   Next
End Sub

Private Function AreListItemsEqual(li1 As mscomctllib.ListItem, li2 As mscomctllib.ListItem) As Boolean
   If li1.Text <> li2.Text Then
      Exit Function
   End If
   
   Dim i As Integer
   For i = 1 To lvDoc.ColumnHeaders.count - 1
      If li1.SubItems(i) <> li2.SubItems(i) Then
         Exit Function
      End If
   Next
   AreListItemsEqual = True
End Function


Private Sub txtSearchDoc_Change()
    If bFormActivated Then
        GetSomeClass
        FillDocuments
    End If
End Sub
Private Function GetDocLstSecurity()
   On Error GoTo DiaErr1
   Dim RdoDoc As ADODB.Recordset
   Err = 0
   sSql = "SELECT ISNULL(CODOCLSTSEC, 0) CODOCLSTSEC FROM ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         GetDocLstSecurity = IIf((!CODOCLSTSEC = 0), False, True)
         ClearResultSet RdoDoc
      End With
   Else
      GetDocLstSecurity = False
   End If
   Set RdoDoc = Nothing
   Exit Function
   
   
DiaErr1:
   sProcName = "GetDocLstSecurity"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function CheckForDocLstSec(strRev As String, bIncRev As Boolean)
   
   If (bDocSec = True) Then
      
      Dim strNewRev As String
      strNewRev = strRev
      If (Trim(strNewRev) = "") Then strNewRev = "0"

      If (bIncRev = True) Then
         If (IsNumeric(strNewRev)) Then
            strNewRev = CStr(CDbl(strNewRev) + 1)
         Else
            strNewRev = Chr$(Asc(strNewRev) + 1)
         End If
      End If
         
      DocuDCe02b.txtRev = strNewRev
      chkUpChild = 0
      DocuDCe02b.Show vbModal
      
      If (chkUpChild = 0) Then
         cmbRev = strRev
         MsgBox "Please change the revision number to modify the Operations.", vbCritical
         CheckForDocLstSec = False
      Else
         ' update the comments
         Dim strCmt As String
         strCmt = txtCmt.Text
         UpdateComments strCmt
         ' The option is not enabled
         CheckForDocLstSec = True
      End If
   Else
      ' The option is not enabled
      CheckForDocLstSec = True
   End If
End Function



Private Function SetApproval()
   If bGoodPart Then
      On Error Resume Next
       
      Dim RdoRel As ADODB.Recordset
      sSql = "SELECT DLSTAPPDATE,DLSTAPPBY FROM DlsthdTable WHERE " _
              & "DLSTREF ='" & Compress(cmbPrt) & "' " _
              & " AND DLSTREV='" & Trim(cmbRev) & "'" _
              & " AND DLSTYPE ='" & CStr(Left(cmbTyp, 1)) & "'"
              
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoRel, ES_FORWARD)
      If bSqlRows Then
         txtApp = RdoRel!DLSTAPPBY
         txtAte = RdoRel!DLSTAPPDATE
      Else
         txtApp = ""
         txtAte = ""
      End If
      Set RdoRel = Nothing
      
    End If
End Function

Private Function ResetApproval()
    If bGoodPart Then
      txtApp = ""
      txtAte = ""
      On Error Resume Next
       
       sSql = "UPDATE DlsthdTable SET DLSTAPPDATE = NULL, DLSTAPPBY = NULL WHERE " _
              & "DLSTREF ='" & Compress(cmbPrt) & "' " _
              & " AND DLSTREV='" & Trim(cmbRev) & "'" _
              & " AND DLSTYPE ='" & CStr(Left(cmbTyp, 1)) & "'"
              
       clsADOCon.ExecuteSql sSql 'rdExecDirect
    End If
End Function



Private Sub txtApp_LostFocus()
   
   On Error GoTo DiaErr1
   Dim strApp As String
   strApp = txtApp.Text
   
   If (strApp = "") Then
      MsgBox "Please entry Approval Name.", vbInformation
   Else
      ' Check and if needed create new record
      CreateDocLstHeader
      
      sSql = "UPDATE DlsthdTable SET DLSTAPPBY ='" _
             & strApp & "' WHERE " _
             & "DLSTREF='" & Compress(cmbPrt) & "' " _
             & " AND DLSTREV='" & Trim(cmbRev) & "'" _
             & " AND DLSTYPE = '" & CStr(Left(cmbTyp, 1)) & "'"
             
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If

   Exit Sub
DiaErr1:
   sProcName = "txtApp_LostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub txtAte_DropDown()
   ShowCalendar Me
End Sub


Private Sub txtAte_LostFocus()
   
   On Error GoTo DiaErr1
   
   Dim strDate As String
   If Trim(txtAte) = "" Then txtAte = CheckDate(txtAte)
   
   strDate = ""
   If Len(txtAte) > 0 Then
      strDate = Format(txtAte, "mm/dd/yy")
   End If
   
   If (strDate = "") Then
      MsgBox "Please entry Approval date.", vbInformation
   Else
   
      ' Check and if needed create new record
      CreateDocLstHeader
      
      sSql = "UPDATE DlsthdTable SET DLSTAPPDATE ='" _
             & strDate & "' WHERE " _
             & " DLSTREF='" & Compress(cmbPrt) & "' " _
             & " AND DLSTREV='" & Trim(cmbRev) & "'" _
             & " AND DLSTYPE = '" & CStr(Left(cmbTyp, 1)) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
   End If
   
   Exit Sub
DiaErr1:
   sProcName = "txtAte_LostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Function UpdateComments(strCmt As String)
   
   On Error GoTo DiaErr1
   
   ' Check and if needed create new record
   CreateDocLstHeader
   
   sSql = "UPDATE DlsthdTable SET DLSTREVNOTES ='" _
          & strCmt & "' WHERE " _
          & " DLSTREF='" & Compress(cmbPrt) & "' " _
          & " AND DLSTREV='" & Trim(cmbRev) & "'" _
          & " AND DLSTYPE = '" & CStr(Left(cmbTyp, 1)) & "'"
   
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   
   Exit Function
DiaErr1:
   sProcName = "txtAte_LostFocus"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Function CreateDocLstHeader()

   On Error GoTo DiaErr1
   
   Dim RdoRel As ADODB.Recordset
   sSql = "SELECT DLSTAPPDATE,DLSTAPPBY FROM DlsthdTable WHERE " _
           & "DLSTREF ='" & Compress(cmbPrt) & "' " _
           & " AND DLSTREV='" & Trim(cmbRev) & "'" _
           & " AND DLSTYPE ='" & CStr(Left(cmbTyp, 1)) & "'"
           
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRel, ES_FORWARD)
   If Not bSqlRows Then
      sSql = "INSERT INTO DlsthdTable(DLSTREF, DLSTREV, DLSTYPE) " _
             & " VALUES('" & Compress(cmbPrt) & "','" _
             & Trim(cmbRev) & "','" & CStr(Left(cmbTyp, 1)) & "')"
      
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
   Set RdoRel = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CreateDocLstHeader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function



