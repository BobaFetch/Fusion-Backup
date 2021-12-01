VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form ToolTLe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   3401
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "ToolTLe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   2100
      TabIndex        =   14
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   3885
      Width           =   1555
   End
   Begin VB.TextBox txtAvl 
      Height          =   285
      Left            =   4020
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "In Useful Condition Or Not In Use"
      Top             =   2400
      Width           =   615
   End
   Begin VB.CheckBox optDef 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   6060
      TabIndex        =   40
      ToolTipText     =   "Out Of Service"
      Top             =   5655
      Width           =   720
   End
   Begin VB.CheckBox optObs 
      Caption         =   "____"
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4140
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "This Tool May Not Be Used"
      Top             =   5655
      Width           =   715
   End
   Begin VB.ComboBox txtObs 
      Height          =   315
      Left            =   2100
      TabIndex        =   16
      Tag             =   "4"
      ToolTipText     =   "Don't Use After"
      Top             =   5655
      Width           =   1095
   End
   Begin VB.TextBox txtCmt 
      Height          =   1215
      Left            =   2100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      ToolTipText     =   "1000 Chars Max"
      Top             =   4320
      Width           =   4695
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   2100
      TabIndex        =   12
      Tag             =   "4"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4740
      TabIndex        =   13
      Tag             =   "4"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtCyc 
      Height          =   285
      Left            =   2100
      TabIndex        =   11
      Tag             =   "1"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CheckBox optExp 
      Alignment       =   1  'Right Justify
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5700
      TabIndex        =   3
      Top             =   1320
      Width           =   715
   End
   Begin VB.TextBox txtCap 
      Height          =   285
      Left            =   5820
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Units To Build (One Cycle)"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   2100
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Standard Or Actual Cost"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   5820
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Location"
      Top             =   2400
      Width           =   615
   End
   Begin VB.ComboBox cmbCls 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "12 Char Class - Retrieved From Previous Entries"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   2100
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Stock"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtEco 
      Height          =   285
      Left            =   2100
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "20 Char Max"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   2100
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Tool Revision"
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox cmbTol 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter a New Tool or Select From List (30 chars)"
      Top             =   540
      Width           =   3255
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Tag             =   "2"
      ToolTipText     =   "Tool Description (40 chars)"
      Top             =   930
      Width           =   3225
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5940
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7380
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   6180
      FormDesignWidth =   7050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer/Tool Owner"
      Height          =   255
      Index           =   22
      Left            =   240
      TabIndex        =   42
      Top             =   3885
      Width           =   1755
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3780
      TabIndex        =   41
      Top             =   3885
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Defective"
      Height          =   255
      Index           =   21
      Left            =   4980
      TabIndex        =   18
      Top             =   5655
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete "
      Height          =   255
      Index           =   20
      Left            =   3300
      TabIndex        =   39
      Top             =   5655
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Obsolete Date"
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   38
      Top             =   5655
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   37
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expendable"
      Height          =   255
      Index           =   16
      Left            =   4740
      TabIndex        =   36
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   15
      Left            =   4740
      TabIndex        =   35
      Top             =   1560
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   14
      Left            =   4860
      TabIndex        =   34
      Top             =   1560
      Width           =   15
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Next Inspection"
      Height          =   255
      Index           =   13
      Left            =   3420
      TabIndex        =   33
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Inspection"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Units)"
      Height          =   255
      Index           =   11
      Left            =   3060
      TabIndex        =   31
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Insp Cycle"
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   30
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity"
      Height          =   255
      Index           =   9
      Left            =   4740
      TabIndex        =   29
      ToolTipText     =   "Number Of Units Mounted"
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Cost"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   28
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   7
      Left            =   4740
      TabIndex        =   27
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Avail"
      Height          =   255
      Index           =   5
      Left            =   2820
      TabIndex        =   25
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Order/Serial #"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   23
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tool Number"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   930
      Width           =   1155
   End
End
Attribute VB_Name = "ToolTLe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/19/04 New
Option Explicit
Dim RdoTool As ADODB.Recordset

Dim bCancel As Byte
Dim bGoodTool As Byte
Dim bOnLoad As Byte
Dim bPartExists As Byte

Dim sOldClass As String
Dim sOldLast As Variant
Dim sOldNext As Variant

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbCls_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbCls = CheckLen(cmbCls, 12)
   cmbCls = StrCase(cmbCls)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CLASS = Trim(cmbCls)
         .Update
      End With
   End If
   If cmbCls.ListCount > 0 Then
      For iList = 0 To cmbCls.ListCount - 1
         If UCase$(cmbCls) = UCase$(cmbCls.list(iList)) Then bByte = 1
      Next
   End If
   If bByte = 0 Then cmbCls.AddItem cmbCls
   If sOldClass <> Trim(cmbCls) Then
      sSql = "UPDATE TlitTable SET TOOLLISTIT_CLASS='" & Trim(cmbCls) _
             & "' WHERE TOOLLISTIT_TOOLREF='" & Compress(cmbTol) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   sOldClass = Trim(cmbCls)
   
End Sub


Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   
End Sub


Private Sub cmbCst_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   cmbCst = CheckLen(cmbCst, 10)
   For iList = 0 To cmbCst.ListCount - 1
      If Trim(cmbCst) = Trim(cmbCst.list(iList)) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbCst = cmbCst.list(0)
   End If
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CUST = Compress(cmbCst)
         .Update
      End With
   End If
   FindCustomer Me, cmbCst
   
End Sub


Private Sub cmbTol_Click()
   bGoodTool = GetThisTool()
   
End Sub

Private Sub cmbTol_LostFocus()
   cmbTol = CheckLen(cmbTol, 30)
   If bCancel = 1 Then Exit Sub
   bGoodTool = GetThisTool()
   If bGoodTool = 0 Then AddNewTool
   
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3401
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set RdoTool = Nothing
   Set ToolTLe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillToolCombo "
   LoadComboBox cmbTol, -1
   If cmbTol.ListCount > 0 Then
      cmbTol = cmbTol.list(0)
      sSql = "Qry_FillToolClasses "
      LoadComboBox cmbCls, -1
      bGoodTool = GetThisTool()
   End If
   AddComboStr cmbCst.hwnd, "         "
   sSql = "Qry_FillSortedCustomers"
   LoadComboBox cmbCst
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub




Private Sub optDef_Click()
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_DEFECTIVE = optDef.Value
         .Update
      End With
   End If
   
End Sub

Private Sub optExp_Click()
   Dim bLevel As Byte
   If optExp.Value = vbChecked Then
      z1(22).Enabled = False
      cmbCst.Enabled = False
      cmbCst = ""
      lblNme = ""
      lblNme.Enabled = False
      txtBeg.Enabled = False
      txtEnd.Enabled = False
      txtBeg = ""
      txtEnd = ""
      bLevel = 5
   Else
      z1(22).Enabled = True
      cmbCst.Enabled = True
      lblNme.Enabled = True
      txtBeg = sOldLast
      txtEnd = sOldNext
      txtBeg.Enabled = True
      txtEnd.Enabled = True
      bLevel = 8
   End If
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_EXPENDABLE = optExp.Value
         !TOOL_CUST = Compress(cmbCst)
         If optExp.Value = vbChecked Then
            !TOOL_NEXTINSP = Null
            !TOOL_LASTINSP = Null
         Else
            If sOldNext <> "" Then !TOOL_NEXTINSP = sOldNext
            If sOldLast <> "" Then !TOOL_LASTINSP = sOldLast
         End If
         .Update
      End With
      sSql = "UPDATE PartTable SET PALEVEL=" & bLevel & " " _
             & "WHERE PARTREF='" & Compress(cmbTol) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
End Sub


Private Sub txtAvl_LostFocus()
   txtAvl = CheckLen(txtAvl, 6)
   txtAvl = Format(Abs(Val(txtAvl)), "#####0")
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_QTYAVAIL = Val(txtAvl)
         .Update
      End With
   End If
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) > 0 Then txtBeg = CheckDate(txtBeg)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         If Len(Trim(txtBeg)) Then
            !TOOL_LASTINSP = Format(txtBeg, "mm/dd/yy")
         Else
            !TOOL_LASTINSP = Null
         End If
         .Update
      End With
   End If
   
End Sub


Private Sub txtCap_LostFocus()
   txtCap = CheckLen(txtCap, 10)
   txtCap = Format(Abs(Val(txtCap)), "######0")
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CAPACITY = Val(txtCap)
         .Update
      End With
   End If
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 1020)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   txtCmt = ReplaceString(txtCmt)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_COMMENTS = Trim(txtCmt)
         .Update
      End With
   End If
   
End Sub

Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 10)
   txtCst = Format(Abs(Val(txtCst)), "######0.00")
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_STDCOST = Val(txtCst)
         .Update
      End With
   End If
   
End Sub


Private Sub txtCyc_LostFocus()
   txtCyc = CheckLen(txtCyc, 10)
   txtCyc = Format(Abs(Val(txtCyc)), "######0")
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_INSPCYCLE = Val(txtCyc)
         .Update
      End With
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodTool = 1 Then
      On Error Resume Next
      sSql = "UPDATE PartTable SET PADESC='" & Trim(txtDsc) & "' " _
             & "WHERE PARTREF='" & Compress(cmbTol) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      With RdoTool
         '.Edit
         !TOOL_DESC = Trim(txtDsc)
         .Update
      End With
   End If
End Sub


Private Sub txtEco_LostFocus()
   txtEco = CheckLen(txtEco, 20)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_CHANGEORDER = txtEco
         .Update
      End With
   End If
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub



Private Function GetThisTool() As Byte
   On Error GoTo DiaErr1
'   sSql = "SELECT TOOL_NUM,TOOL_PARTREF,TOOL_DESC,TOOL_REVISION,TOOL_CHANGEORDER," _
'          & "TOOL_CLASS,TOOL_QOH,TOOL_QTYAVAIL,TOOL_EXPENDABLE,TOOL_MAXPARTS," _
'          & "TOOL_CAPACITY,TOOL_AUTOPICK,TOOL_STDCOST,TOOL_LOCATION,TOOL_INSPCYCLE," _
'          & "TOOL_LASTINSP,TOOL_NEXTINSP,TOOL_OBSOLETEDATE,TOOL_OBSOLETE," _
'          & "TOOL_DEFECTIVE,TOOL_CUST,TOOL_COMMENTS,CUREF,CUNICKNAME,CUNAME " _
'          & "FROM TohdTable,CustTable WHERE (TOOL_PARTREF='" _
'          & Compress(cmbTol) & "' AND TOOL_CUST*=CUREF)"
   sSql = "SELECT TOOL_NUM,TOOL_PARTREF,TOOL_DESC,TOOL_REVISION,TOOL_CHANGEORDER," _
          & "TOOL_CLASS,TOOL_QOH,TOOL_QTYAVAIL,TOOL_EXPENDABLE,TOOL_MAXPARTS," _
          & "TOOL_CAPACITY,TOOL_AUTOPICK,TOOL_STDCOST,TOOL_LOCATION,TOOL_INSPCYCLE," _
          & "TOOL_LASTINSP,TOOL_NEXTINSP,TOOL_OBSOLETEDATE,TOOL_OBSOLETE," _
          & "TOOL_DEFECTIVE,TOOL_CUST,TOOL_COMMENTS,CUREF,CUNICKNAME,CUNAME " _
          & "FROM TohdTable" & vbCrLf _
          & "LEFT JOIN CustTable ON TOOL_CUST=CUREF" & vbCrLf _
          & "WHERE TOOL_PARTREF='" & Compress(cmbTol) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTool, ES_KEYSET)
   If bSqlRows Then
      With RdoTool
         cmbTol = "" & Trim(!TOOL_NUM)
         txtDsc = "" & Trim(!TOOL_DESC)
         cmbCls = "" & Trim(!TOOL_CLASS)
         optExp.Value = !TOOL_EXPENDABLE
         txtRev = "" & Trim(!TOOL_REVISION)
         txtEco = "" & Trim(!TOOL_CHANGEORDER)
         txtQty = Format(!TOOL_QOH, "#######0")
         txtAvl = Format(!TOOL_QTYAVAIL, "#######0")
         txtLoc = "" & Trim(!TOOL_LOCATION)
         txtCst = Format(!TOOL_STDCOST, "#######0.00")
         txtCyc = Format(!TOOL_INSPCYCLE, "#######0")
         txtCap = Format(!TOOL_CAPACITY, "#######0")
         cmbCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         If Not IsNull(!TOOL_LASTINSP) Then
            txtBeg = Format(!TOOL_LASTINSP, "mm/dd/yy")
         Else
            txtBeg = ""
         End If
         If Not IsNull(!TOOL_NEXTINSP) Then
            txtEnd = Format(!TOOL_NEXTINSP, "mm/dd/yy")
         Else
            txtEnd = ""
         End If
         sOldLast = txtBeg
         sOldNext = txtEnd
         txtCmt = "" & Trim(!TOOL_COMMENTS)
         If Not IsNull(!TOOL_OBSOLETEDATE) Then
            txtObs = Format(!TOOL_OBSOLETEDATE, "mm/dd/yy")
         Else
            txtObs = ""
         End If
         If optExp.Value = vbChecked Then
            z1(22).Enabled = False
            cmbCst.Enabled = False
            cmbCst = ""
            lblNme = ""
            lblNme.Enabled = False
            txtBeg.Enabled = False
            txtEnd.Enabled = False
         Else
            z1(22).Enabled = True
            cmbCst.Enabled = True
            lblNme.Enabled = True
            txtBeg.Enabled = True
            txtEnd.Enabled = True
         End If
         optObs.Value = !TOOL_OBSOLETE
         optDef.Value = !TOOL_DEFECTIVE
         sOldClass = Trim(cmbCls)
      End With
      GetThisTool = 1
   Else
      GetThisTool = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthisto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddNewTool()
   'revised 6/16/2017 to NOT create a part and to allow tool # = part number
   Dim bResponse As Byte
   Dim sMsg As String
   If Len(cmbTol) < 3 Then Exit Sub
   'On Error Resume Next
   bResponse = MsgBox("Add New Tool " & cmbTol & "?", _
               ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      'bPartExists = GetPartNumber()
      txtDsc = ""
      txtRev = ""
      txtEco = ""
      txtQty = "1"
      txtAvl = "1"
      txtCyc = ""
      txtBeg = ""
      txtEnd = ""
      txtCmt = ""
      txtObs = ""
      optObs = vbUnchecked
      optDef = vbUnchecked
'      If bPartExists = 1 Then
'         MsgBox cmbTol & " Is In Use As A Part Number " & vbCrLf _
'            & "And Cannot Be Duplicated. Pick Another.", _
'            vbInformation, Caption
'         Exit Sub
'      Else
         bResponse = IllegalCharacters(cmbTol)
         If bResponse > 0 Then
            MsgBox "The Part Number Contains An Illegal " & Chr$(bResponse) & ".", _
               vbExclamation, Caption
            Exit Sub
         Else
            'Add it
            'Err = 0
            clsADOCon.BeginTrans
            
            Dim part As New ClassPart
            'If part.CreateNewPart(cmbTol, 8, "Tool", "M") Then
            
'            sSql = "INSERT INTO PartTable (PARTREF,PARTNUM,PALEVEL,PATOOL," _
'                   & "PAQOH) VALUES('" & Compress(cmbTol) & "','" & cmbTol & "',8,1,1)"
'            clsADOCon.ExecuteSQL sSql 'rdExecDirect
            
'               sSql = "UPDATE PartTable" & vbCrLf _
'                  & "SET PATOOL = 1, PAQOH = 1" & vbCrLf _
'                 & "WHERE PARTREF = '" & Compress(cmbTol) & "'"
'               clsADOCon.ExecuteSql sSql 'rdExecDirect
               
               sSql = "INSERT INTO TohdTable (TOOL_NUM,TOOL_PARTREF,TOOL_CLASS) " _
                      & "VALUES('" & Trim(cmbTol) & "','" & Compress(cmbTol) & "','" _
                      & Trim(cmbCls) & "')"
               clsADOCon.ExecuteSql sSql 'rdExecDirect
               
               clsADOCon.CommitTrans
               SysMsg "The Tool Was Created.", True
               cmbTol.AddItem cmbTol
               bGoodTool = GetThisTool()

'            Else
'               clsADOCon.RollbackTrans
'               MsgBox "Could Not Create The Tool.", _
'                  vbExclamation, Caption
'            End If
         End If
      'End If
   Else
      CancelTrans
   End If
   
End Sub

'Private Function GetPartNumber() As Byte
'   Dim RdoPrt As ADODB.Recordset
'   On Error GoTo DiaErr1
'   sSql = "SELECT PARTREF FROM PartTable WHERE PARTREF='" _
'          & Compress(cmbTol) & " '"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
'   If bSqlRows Then GetPartNumber = 1 Else GetPartNumber = 0
'   Set RdoPrt = Nothing
'   Exit Function
'
'DiaErr1:
'   sProcName = "getpartnum"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) > 0 Then txtEnd = CheckDate(txtEnd)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         If Len(Trim(txtEnd)) Then
            !TOOL_NEXTINSP = Format(txtEnd, "mm/dd/yy")
         Else
            !TOOL_NEXTINSP = Null
         End If
         .Update
      End With
   End If
   
End Sub

Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 4)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_LOCATION = txtLoc
         .Update
      End With
   End If
   
End Sub


Private Sub txtObs_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtObs_LostFocus()
   If Len(txtObs) < 0 Then txtObs = CheckDate(txtObs)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         If Len(Trim(txtObs)) Then
            optObs.Value = vbChecked
            !TOOL_OBSOLETEDATE = Format(txtObs, "mm/dd/yy")
            !TOOL_OBSOLETE = 1
         Else
            optObs.Value = vbUnchecked
            !TOOL_OBSOLETEDATE = Null
            !TOOL_OBSOLETE = 0
         End If
         .Update
      End With
   End If
   
End Sub


Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 6)
   txtQty = Format(Abs(Val(txtQty)), "#####0")
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_QOH = Val(txtQty)
         .Update
      End With
      sSql = "UPDATE PartTable SET PAQOH=" & Val(txtQty) & " " _
             & "WHERE PARTREF='" & Compress(cmbTol) & "'"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 6)
   If bGoodTool = 1 Then
      On Error Resume Next
      With RdoTool
         '.Edit
         !TOOL_REVISION = txtRev
         .Update
      End With
   End If
   
End Sub
