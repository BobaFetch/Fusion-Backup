VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form InvcINe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Aliases"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   5103
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINe03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optEcom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.ComboBox cmbCls 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "8"
      ToolTipText     =   "Select Product Class From List"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtAdsc 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Tag             =   "2"
      ToolTipText     =   "Alias Part Description  (30 Max)"
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Width           =   5775
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Tag             =   "3"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtMfg 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Tag             =   "2"
      ToolTipText     =   "Alias Part Manufacturer  (30 Max)"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      ToolTipText     =   "Remove The Current Alias"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Add A New Alias"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txtAls 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Alias Part Number (30 Max)"
      Top             =   3840
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select From List Or Enter Part Number"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5775
      FormDesignWidth =   5985
   End
   Begin MSFlexGridLib.MSFlexGrid lstAls 
      Height          =   1935
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Double Click To Select Entry"
      Top             =   1680
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   8404992
      FocusRect       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Commerce"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   285
      Index           =   15
      Left            =   2640
      TabIndex        =   23
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   1440
      Y1              =   1665
      Y2              =   1665
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      Height          =   285
      Index           =   8
      Left            =   3225
      TabIndex        =   22
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alias Part Numbers"
      Height          =   285
      Index           =   7
      Left            =   1455
      TabIndex        =   21
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Alias Number"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliases"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "InvcINe03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
'*********************************************************************************
' InvcINe03a - Assign aliases to part numbers.
'
' Created: 12/26/01 (nth)
' Revisions: 1/8/02 cjs, 2/1/20
' 12/29/04 Fixed FillBoxes
'
'*********************************************************************************
Dim RdoAls As ADODB.Recordset
Dim bOnLoad As Byte
Dim bAddNew As Byte
Dim bGoodPart As Byte
Dim iIndex As Integer

Dim sPart As String
Dim sAlias As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCls_Validate(Cancel As Boolean)
   Dim b As Byte
   Dim C As Byte
   Dim sClass As String
   
   On Error Resume Next
   For b = 0 To cmbCls.ListCount - 1
      If cmbCls = cmbCls.List(b) Then C = 1
   Next
   If C = 0 Then
      Beep
      cmbCls = cmbCls.List(0)
   End If
   
   If cmbCls <> "NONE" Then sClass = Compress(cmbCls)
   If bGoodPart Then
      sSql = "UPDATE PartTable SET PACLASS='" _
             & sClass & "' WHERE PARTREF='" _
             & Compress(cmbPrt) & "' "
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub


Private Sub cmbPrt_Click()
   cmbPrt = CheckLen(cmbPrt, 30)
   lblDsc.ForeColor = Me.ForeColor
   bGoodPart = GetAliasedPart()
   GetAlias
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   lblDsc.ForeColor = Me.ForeColor
   bGoodPart = GetAliasedPart()
   iIndex = 0
   GetAlias True
   
End Sub

Private Sub cmbVnd_Change()
   If Trim(cmbVnd) = "" Then
      txtNme.ForeColor = ES_RED
      txtNme = "*** No Valid Vendor Selected ***"
   Else
      txtNme.ForeColor = Me.ForeColor
   End If
End Sub

Private Sub cmbVnd_Click()
   Dim b As Byte
   txtNme.ForeColor = ForeColor
   b = FindVendor()
   
End Sub

Private Sub cmbVnd_LostFocus()
   Dim b As Byte
   b = FindVendor()
   txtNme.ForeColor = Me.ForeColor
   On Error Resume Next
   If Len(Trim(txtAls)) > 0 Then
      With RdoAls
         !ALVENDOR = Compress(cmbVnd)
         .Update
      End With
      If Err <> 0 Then
         If Err = 40026 Then
            sAlias = Compress(txtAls)
            GetThisAlias
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub

Private Sub cmdAdd_Click()
   'AddAlias
   bAddNew = 1
   txtAls = " "
   txtMfg = " "
   txtAdsc = " "
   cmbVnd = " "
   
   txtAls.BackColor = Es_TextBackColor
   txtMfg.BackColor = Es_TextBackColor
   cmbVnd.BackColor = Es_TextBackColor
   txtAdsc.BackColor = Es_TextBackColor
   
   txtAls.Enabled = True
   txtMfg.Enabled = True
   txtAdsc.Enabled = True
   cmbVnd.Enabled = True
   txtAls.SetFocus
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Len(Trim(txtAls)) > 0 Then
      sMsg = "Are You Sure That You Want To Delete" & vbCr _
             & lstAls.Text & " From The List?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         RemoveAlias
      Else
         CancelTrans
      End If
   Else
      MsgBox "An Alias Must Be Selected.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5103"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillProductClasses
      MouseCursor 13
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   With lstAls
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(1) = flexAlignLeftCenter
      .Row = 0
      '.Col = 0
      '.Text = "Alias Part Number"
      '.Col = 1
      '.Text = "Manufacturer"
      .ColWidth(0) = 1650
      .ColWidth(1) = 1650
      .Rows = 1
   End With
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set RdoAls = Nothing
   Set InvcINe03a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   txtAls.BackColor = Es_FormBackColor
   txtMfg.BackColor = Es_FormBackColor
   cmbVnd.BackColor = Es_FormBackColor
   txtAdsc.BackColor = Es_FormBackColor
   txtAls.Enabled = False
   txtMfg.Enabled = False
   txtAdsc.Enabled = False
   cmbVnd.Enabled = False
   
End Sub

Private Sub GetAlias(Optional bFirstRow As Boolean)
   On Error Resume Next
   'RdoAls.Close
   Set RdoAls = Nothing
   On Error GoTo DiaErr1
   lstAls.Clear
   lstAls.Rows = 1
   cmdDel.Enabled = False
   
   sPart = Compress(cmbPrt)
   sSql = "SELECT * FROM PaalTable " _
          & "WHERE ALPARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAls, ES_KEYSET)
   If bSqlRows Then
      With RdoAls
         Do Until .EOF
            lstAls.Row = lstAls.Rows - 1
            iIndex = lstAls.Row
            lstAls.Col = 0
            lstAls.Text = "" & Trim(!ALALIASNUM)
            lstAls.Col = 1
            lstAls.Text = "" & Trim(!ALMFG)
            If Not .EOF Then lstAls.Rows = lstAls.Rows + 1
            .MoveNext
         Loop
         lstAls.Rows = lstAls.Rows - 1
      End With
      If bFirstRow Then
         lstAls.Row = 0
         lstAls.Col = 0
         If Len(Trim(lstAls.Text)) > 0 Then
            sAlias = Trim(lstAls.Text)
            GetThisAlias
         End If
      End If
      txtAls.BackColor = Es_TextBackColor
      txtMfg.BackColor = Es_TextBackColor
      cmbVnd.BackColor = Es_TextBackColor
      txtAdsc.BackColor = Es_TextBackColor
      
      txtAdsc.Enabled = True
      txtAls.Enabled = True
      txtMfg.Enabled = True
      cmbVnd.Enabled = True
   Else
      txtAls.BackColor = Es_FormBackColor
      txtMfg.BackColor = Es_FormBackColor
      cmbVnd.BackColor = Es_FormBackColor
      txtAdsc.BackColor = Es_FormBackColor
      txtAls = " "
      txtAdsc = " "
      txtMfg = " "
      cmbVnd = " "
      txtNme = ""
      txtAls.Enabled = False
      txtMfg.Enabled = False
      txtAdsc.Enabled = False
      cmbVnd.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getAlias"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillCombo()
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   FillVendors
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      'bGoodPart = GetAliasedPart()
      GetAlias True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "FillCombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub RemoveAlias()
   Dim sRemove As String
   Dim bUsed As Byte
   
   On Error GoTo DiaErr1
   If lstAls.Rows > 1 And lstAls = "" Then
      MsgBox "No Alias Selected.", _
         vbInformation, Caption
   Else
      sRemove = Compress(lstAls)
      sSql = "DELETE FROM PaalTable WHERE " _
             & "ALALIASREF='" & sRemove & "' "
      clsADOCon.ExecuteSQL sSql
      bUsed = clsADOCon.RowsAffected
      If bUsed Then
         SysMsg "Alias Was Deleted.", True
         GetAlias
         txtAls = " "
         txtAdsc = " "
         txtMfg = " "
         cmbVnd = " "
      Else
         MsgBox "Alias In Use. Couldn't Delete.", _
            vbInformation, Caption
      End If
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "removealias"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillBoxes()
   Dim b As Byte
   On Error Resume Next
   If lstAls.Rows > 0 Then
      With RdoAls
         .MoveFirst
         .Move iIndex
         txtAls = "" & Trim(!ALALIASNUM)
         txtAdsc = "" & Trim(!ALALIASDESC)
         txtMfg = "" & Trim(!ALMFG)
         cmbVnd = "" & Trim(!ALVENDOR)
         b = FindVendor()
      End With
   End If
   
End Sub

Private Sub lblDsc_Change()
   If Trim(lblDsc) = "*** Not A Valid Part Number" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = Es_TextForeColor
   End If
   
   
End Sub

Private Sub lstAls_Click()
   iIndex = lstAls.Row
   
End Sub

Private Sub lstAls_DblClick()
   On Error Resume Next
   lstAls.Col = 0
   iIndex = lstAls.Row
   'FillBoxes
   sAlias = Compress(lstAls.Text)
   GetThisAlias
   txtAls.SetFocus
   
End Sub


Private Sub lstAls_GotFocus()
   lstAls.Col = 0
   iIndex = lstAls.Row
   If Len(Trim(lstAls.Text)) > 0 Then cmdDel.Enabled = True
   
End Sub

Private Sub optEcom_Validate(Cancel As Boolean)
   If bGoodPart Then
      On Error Resume Next
      sSql = "UPDATE PartTable SET PAECOMMERCE=" _
             & optEcom.Value & " WHERE PARTREF='" _
             & Compress(cmbPrt) & "' "
      clsADOCon.ExecuteSQL sSql
   End If
   
End Sub


Private Sub txtAdsc_LostFocus()
   txtAdsc = CheckLen(txtAdsc, 30)
   txtAdsc = StrCase(txtAdsc)
   
   On Error Resume Next
   If Len(Trim(txtAls)) > 0 Then
      With RdoAls
         !ALALIASDESC = "" & Trim(txtAdsc)
         .Update
      End With
      If Err <> 0 Then
         If Err = 40026 Then
            sAlias = Compress(txtAls)
            GetThisAlias
         Else
            ValidateEdit
         End If
      End If
   End If
   
End Sub


Private Sub txtAls_LostFocus()
   txtAls = CheckLen(txtAls, 30)
   On Error Resume Next
   If bAddNew Then
      If Len(Trim(txtAls)) > 0 Then
         If Len(Trim(txtAdsc)) = 0 Then txtAdsc = lblDsc
         On Error GoTo DiaErr1
         With RdoAls
            .AddNew
            !ALPARTREF = sPart
            !ALALIASREF = Compress(txtAls)
            !ALALIASNUM = "" & Trim(txtAls)
            !ALALIASDESC = "" & Trim(txtAdsc)
            .Update
         End With
         bAddNew = 0
         ' Refresh list box
         GetAlias
         ' Move cursor back to new alias
         'FillBoxes
         sAlias = Compress(txtAls)
         GetThisAlias
      End If
   Else
      If Len(Trim(txtAls)) > 0 Then
         With RdoAls
            !ALALIASREF = Compress(txtAls)
            !ALALIASNUM = txtAls
            .Update
         End With
      End If
      If Err <> 0 Then
         If Err = 40026 Then
            sAlias = Compress(txtAls)
            GetThisAlias
         Else
            ValidateEdit
         End If
      End If
   End If
   
   Exit Sub
   
DiaErr1:
   On Error Resume Next
   MsgBox "Error. Possibly A Duplicate Alias.", _
      vbExclamation, Caption
   FillBoxes
   
End Sub

Private Sub txtMfg_LostFocus()
   Dim b As Byte
   Dim A As Integer
   
   txtMfg = CheckLen(txtMfg, 30)
   txtMfg = StrCase(txtMfg)
   On Error Resume Next
   If Len(Trim(txtAls)) > 0 Then
      With RdoAls
         !ALMFG = "" & Trim(txtMfg)
         .Update
         lstAls.Col = 0
         For A = 0 To lstAls.Rows - 1
            lstAls.Row = A
            If Trim(lstAls.Text) = Trim(txtAls) Then b = A
         Next
         iIndex = b
         lstAls.Row = iIndex
         lstAls.Col = 1
         lstAls.Text = txtMfg
      End With
   End If
   
End Sub

Private Sub GetThisAlias()
   On Error Resume Next
   RdoAls.Close
   
   On Error GoTo DiaErr1
   sPart = Compress(cmbPrt)
   sSql = "SELECT * FROM PaalTable " _
          & "WHERE ALPARTREF = '" & sPart & "' AND ALALIASREF='" _
          & sAlias & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAls, ES_KEYSET)
   If bSqlRows Then
      With RdoAls
         txtAls = "" & Trim(!ALALIASNUM)
         txtAdsc = "" & Trim(!ALALIASDESC)
         txtMfg = "" & Trim(!ALMFG)
         cmbVnd = "" & Trim(!ALVENDOR)
      End With
      FindVendor
      txtAls.BackColor = Es_TextBackColor
      txtMfg.BackColor = Es_TextBackColor
      cmbVnd.BackColor = Es_TextBackColor
      txtAdsc.BackColor = Es_TextBackColor
      
      txtAdsc.Enabled = True
      txtAls.Enabled = True
      txtMfg.Enabled = True
      cmbVnd.Enabled = True
   Else
      txtAls.BackColor = Es_FormBackColor
      txtMfg.BackColor = Es_FormBackColor
      cmbVnd.BackColor = Es_FormBackColor
      txtAdsc.BackColor = Es_FormBackColor
      txtAls = " "
      txtAdsc = " "
      txtMfg = " "
      cmbVnd = " "
      txtAls.Enabled = False
      txtMfg.Enabled = False
      txtAdsc.Enabled = False
      cmbVnd.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getAlias"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetAliasedPart() As Byte
   Dim RdoApt As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PACLASS," _
          & "PAECOMMERCE FROM PartTable WHERE PARTREF='" _
          & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoApt, ES_FORWARD)
   If bSqlRows Then
      With RdoApt
         lblDsc = "" & Trim(!PADESC)
         If Trim(!PACLASS) = "" Then
            cmbCls = "NONE"
         Else
            cmbCls = "" & Trim(!PACLASS)
         End If
         optEcom.Value = !PAECOMMERCE
         GetAliasedPart = 1
         ClearResultSet RdoApt
      End With
   Else
      lblDsc = "*** Not A Valid Part Number"
      optEcom.Value = vbUnchecked
      GetAliasedPart = 0
   End If
   Set RdoApt = Nothing
End Function
