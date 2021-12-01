VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MrplMRp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Activity for Part BOM"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Tag             =   "3"
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "MrplMRp03a.frx":0000
      Height          =   315
      Left            =   5760
      Picture         =   "MrplMRp03a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   960
      Width           =   350
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.CheckBox OptCmt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   3120
      Width           =   375
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRp03a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   288
      ItemData        =   "MrplMRp03a.frx":0E32
      Left            =   2160
      List            =   "MrplMRp03a.frx":0E48
      TabIndex        =   3
      Text            =   "7"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Frame z2 
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   2160
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   6
         Top             =   200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   7
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp03a.frx":0E5E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp03a.frx":0FE8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3465
      FormDesignWidth =   7245
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include BOM comments"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Levels Through"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   1152
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   1428
   End
End
Attribute VB_Name = "MrplMRp03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 5/12/04
'2/6/07 Added Ext Desc Group 7.2.5
Option Explicit
Dim bOnLoad As Byte
Dim iOrder As Integer
Dim sIns As String
Dim sBomRev As String

'Private txtKeyPress() As New EsiKeyBd
'Private txtGotFocus() As New EsiKeyBd
'Private txtKeyDown() As New EsiKeyBd

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr03", sOptions)
   If Len(Trim(sOptions)) > 0 Then optDsc.Value = Val(sOptions)
   
End Sub

Private Sub PrintReport()
    Dim bLvl As Byte
    Dim sMbe As String
    Dim sBegDate As String
    Dim sEndDate As String
    
   
    Dim sCustomReport As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
    Dim strIncludes As String
    Dim strDateDev As String
    Dim strMbe As String
   
    MouseCursor 13
    On Error GoTo DiaErr1
    GetMRPCreateDates sBegDate, sEndDate
   
    bLvl = Val(cmbLvl) + 1
   
   
    sCustomReport = GetCustomReport("prdmr03")
    Set cCRViewer = New EsCrystalRptViewer
    cCRViewer.Init
    cCRViewer.SetReportFileName sCustomReport, sReportPath

    cCRViewer.SetReportTitle = sCustomReport
    cCRViewer.ShowGroupTree False
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "PartDescription"
    aFormulaName.Add "RequestBy"
    aFormulaName.Add "DateDeveloped"
    aFormulaName.Add "Mbe"
    aFormulaName.Add "ShowPartDesc"
    aFormulaName.Add "ShowBOMCmt"

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = cmbPart
    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
    aFormulaValue.Add CStr("'" & CStr(lblDsc) & "...'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")
    
    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEndDate & "'"
    aFormulaValue.Add CStr(strDateDev)

   
'    sSql = "{MrplTable.MRP_PARTREF}= '" & Compress(cmbPart) & "' AND "
    sSql = "{MrpbTable.MRPBILL_LEVEL}<" & bLvl & " AND "
    sSql = sSql & "{MrpbTable.MRPBILL_LEVEL} >= 0.00 and not ({PartTable.PALEVEL} in [6, 5])"

    If optMbe(0).Value = True Then
       sMbe = "Make"
       sSql = sSql & "AND {PartTable.PAMAKEBUY} ='M'"
    ElseIf optMbe(1).Value = True Then
       sMbe = "Buy"
       sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
    ElseIf optMbe(2).Value = True Then
       sMbe = "Either"
       sSql = sSql & " AND {PartTable.PAMAKEBUY}='E'"
    Else
       sMbe = "Make, Buy And Either"
    End If
   
    strMbe = "'" & sMbe & " AND BOM Levels " & "Through " & cmbLvl & "'"
    aFormulaValue.Add CStr(strMbe)
   
    ' option for part desc
    aFormulaValue.Add CStr(optDsc)
    aFormulaValue.Add CStr(optCmt)
    ' report parameter
    'cCRViewer.SetReportSection "GroupHeaderSection3", True
    
    ' Set Formula values
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    ' set the report Selection
    cCRViewer.SetReportSelectionFormula (sSql)
    cCRViewer.CRViewerSize Me
    
    ' Set report parameter
    cCRViewer.SetDbTableConnection


    cCRViewer.OpenCrystalReportObject Me, aFormulaName

    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aRptParaType
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue

   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetNextBillLevel4(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel4"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',4)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel5 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel5(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel5"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',5)"
            ' MM 6/19/2010 - Missing record update
            clsADOCon.ExecuteSQL sIns
            
            GetNextBillLevel6 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel6(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel6"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',6)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel7 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub GetNextBillLevel7(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel7"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',7)"
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub GetNextBillLevel3(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel3"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',3)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel4 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub cmbLvl_LostFocus()
   Dim bByte As Byte
   Dim iList As Byte
   cmbLvl = CheckLen(cmbLvl, 1)
   For iList = 0 To cmbLvl.ListCount - 1
      If cmbLvl = cmbLvl.List(iList) Then bByte = 1
   Next
   If bByte = 0 Then
      Beep
      cmbLvl = 7
   End If
   
End Sub


'Private Sub cmbPrt_Click()
''MsgBox cmbPrt
'   'cmbPrt = GetCurrentPart(cmbPrt.Text, lblDsc)   'infinite loop
'   Dim s As String
'   s = GetCurrentPart(cmbPrt.Text, lblDsc)
''MsgBox cmbPrt
'End Sub


'Private Sub cmbPrt_LostFocus()
'   Dim bByte As Byte
'   Dim iList As Integer
'   Dim sDesc As String
'
''   For iList = 0 To cmbPrt.ListCount - 1
''      If iList = 32767 Then
''         Debug.Print "32767"
''      End If
''      If Compress(cmbPrt) = Compress(cmbPrt.List(iList)) Then bByte = 1
''   Next

'   bByte = 1   'just configure as listbox - don't let user type
   
'   If bByte = 0 Then
'      Beep
'      cmbPrt = cmbPrt.List(0)
'   End If
   
'   'ignore error if GetCurrentPart reeturns an invalid
'   On Error Resume Next
'   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
'
'End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



'Private Sub cmdFnd_Click()
'    ViewParts.lblControl = "TXTPRT"
'    ViewParts.txtPrt = txtPrt
'    ViewParts.Show
'End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart.ListIndex = 0
    cmbPart = GetCurrentPart(cmbPart, lblDsc)
'    cmbPart = "ALL"
'    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   
   
   If bOnLoad Then
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombos
   End If
   
   bOnLoad = 0
   MouseCursor 0
'   cmbPart = ""
   cmbPart = GetCurrentPart(cmbPart, lblDsc)
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set MrplMRp03a = Nothing
   
End Sub



Private Sub FormatControls()
   'Dim b As Byte
   'b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   'txtPrt = "ALL"
End Sub

'Private Sub FillCombo()
'   Dim sDesc As String
'   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT MRP_PARTREF,MRP_PARTNUM FROM " _
'          & "MrplTable WHERE MRP_PARTREF <>'' ORDER BY MRP_PARTREF"
'   LoadComboBox cmbPrt
'   If cmbPrt.ListCount <> 0 Then
'      cmbPrt = cmbPrt.List(0)
'      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
'   Else
'      MsgBox "There Is No Current MRP Available.", _
'         vbInformation, Caption
'   End If
'   Exit Sub
'
'DiaErr1:
'   sProcName = "fillcombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub

'Not using recursion here to keep the levels straight and
'make it easy to read

Private Sub GetBill()
   Dim RdoBom As ADODB.Recordset
   MouseCursor 13
   iOrder = 0
   On Error GoTo DiaErr1
   sProcName = "getbill"
   sSql = "truncate table MrpbTable"
   clsADOCon.ExecuteSQL sSql
   sSql = "INSERT INTO MrpbTable (MRPBILL_PARTREF," _
          & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
          & "VALUES('" & Compress(cmbPart) & "','" _
          & "',0)"
   clsADOCon.ExecuteSQL sSql
   
   sBomRev = GetBomRev(Compress(cmbPart))
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM BmplTable " _
          & "WHERE (BMASSYPART='" & Compress(cmbPart) & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            sProcName = "getbomrev"
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',1)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel2 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   PrintReport
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextBillLevel2(sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel2"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',2)"
            clsADOCon.ExecuteSQL sIns
            GetNextBillLevel3 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub optDis_Click()
    If Compress(cmbPart) = "" Then
        MsgBox "You Must Enter a Valid Part Number", vbOKOnly
        Exit Sub
    End If
    
   GetBill
   
End Sub


Private Sub optPrn_Click()
    If Compress(cmbPart) = "" Then
        MsgBox "You Must Enter a Valid Part Number", vbOKOnly
        Exit Sub
    End If
   
   GetBill
   
End Sub



Private Function GetBomRev(sPartNumber) As String
   Dim RdoRev As ADODB.Recordset
   sProcName = "getbomrev"
   sSql = "SELECT PARTREF,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
   If bSqlRows Then
      With RdoRev
         GetBomRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoRev
      End With
   Else
      GetBomRev = ""
   End If
   Set RdoRev = Nothing
End Function

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "Prdmr03", optDsc.Value
   
End Sub

'Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'        ViewParts.lblControl = "TXTPRT"
'        ViewParts.txtPrt = txtPrt
'        ViewParts.Show
'    End If
'End Sub

'Private Sub txtPrt_LostFocus()
'    txtPrt = CheckLen(txtPrt, 30)
'
'   On Error Resume Next
'   txtPrt = GetCurrentPart(txtPrt, lblDsc)
'
''    If Len(txtPrt) = 0 Then txtPrt = "ALL"
'End Sub


Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    On Error Resume Next
    cmbPart = GetCurrentPart(cmbPart, lblDsc)
'      If Trim(cmbPart) = "" Then
'        cmbPart = "ALL"
'        Me.lblDsc = "* ALL PARTS *"
'    End If
    
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPart.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPart.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function



Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPart = txtPrt
   On Error Resume Next
   cmbPart = GetCurrentPart(cmbPart, lblDsc)
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPart = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPart
   ViewParts.Show
End Sub

