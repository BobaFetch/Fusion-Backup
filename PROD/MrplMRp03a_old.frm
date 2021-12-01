VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MrplMRp03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Activity for Part BOM"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRp03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   288
      ItemData        =   "MrplMRp03a.frx":07AE
      Left            =   2160
      List            =   "MrplMRp03a.frx":07C4
      TabIndex        =   1
      Text            =   "7"
      Top             =   1680
      Width           =   615
   End
   Begin VB.Frame z2 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   4
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   5
         Top             =   200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp03a.frx":07DA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp03a.frx":0964
         Style           =   1  'Graphical
         TabIndex        =   10
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Part Number From List (Contains Part Numbers From Last MRP Run"
      Top             =   960
      Width           =   3345
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
      FormDesignHeight=   3180
      FormDesignWidth =   7245
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   288
      Index           =   7
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   2388
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Levels Through"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   2028
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   288
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   1152
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   2160
      TabIndex        =   13
      Top             =   1320
      Width           =   3012
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   12
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
   If Len(Trim(sOptions)) > 0 Then optDsc.value = Val(sOptions)
   
End Sub

Private Sub PrintReport()
    Dim bLvl As Byte
    Dim sMbe As String
    Dim sBegDate As String
    Dim sEnddate As String
   
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
    GetMRPCreateDates sBegDate, sEnddate
    SetMdiReportsize MDISect
   
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

    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

    strIncludes = cmbPrt
    aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
    aFormulaValue.Add CStr("'" & CStr(lblDsc) & "...'")
    aFormulaValue.Add CStr("'" & CStr(sInitials) & "'")
    
    strDateDev = "'MRP Created  " & sBegDate & " For Requirements Through " & sEnddate & "'"
    aFormulaValue.Add CStr(strDateDev)

   
    sSql = "{MrpbTable.MRPBILL_LEVEL}<" & bLvl & " "
    
    If optMbe(0).value = True Then
       sMbe = "Make"
       sSql = sSql & "AND {PartTable.PAMAKEBUY} ='M'"
    ElseIf optMbe(1).value = True Then
       sMbe = "Buy"
       sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
    ElseIf optMbe(2).value = True Then
       sMbe = "Either"
       sSql = sSql & " AND {PartTable.PAMAKEBUY}='E'"
    Else
       sMbe = "Make, Buy And Either"
    End If
   
    strMbe = "'" & sMbe & " AND BOM Levels " & "Through " & cmbLvl & "'"
    aFormulaValue.Add CStr(strMbe)
   
    ' option for part desc
    aFormulaValue.Add CStr(optDsc)
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

Private Sub PrintReport1()
   Dim bLvl As Byte
   Dim sMbe As String
   Dim sBegDate As String
   Dim sEnddate As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   GetMRPCreateDates sBegDate, sEnddate
   SetMdiReportsize MDISect
   bLvl = Val(cmbLvl) + 1
   MDISect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MDISect.Crw.Formulas(1) = "Includes='" & cmbPrt _
                        & "'"
   MDISect.Crw.Formulas(2) = "PartDescription='" & lblDsc & "'"
   MDISect.Crw.Formulas(3) = "RequestBy = 'Requested By: " & sInitials & "'"
   MDISect.Crw.Formulas(4) = "DateDeveloped = 'MRP Created " & sBegDate _
                        & " For Requirements Through " & sEnddate & "'"
   
   sCustomReport = GetCustomReport("prdmr03")
   MDISect.Crw.ReportFileName = sReportPath & sCustomReport
   sSql = "{MrpbTable.MRPBILL_LEVEL}<" & bLvl & " "
   If optMbe(0).value = True Then
      sMbe = "Make"
      sSql = sSql & "AND {PartTable.PAMAKEBUY} ='M'"
   ElseIf optMbe(1).value = True Then
      sMbe = "Buy"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
   ElseIf optMbe(2).value = True Then
      sMbe = "Either"
      sSql = sSql & " AND {PartTable.PAMAKEBUY}='E'"
   Else
      sMbe = "Make, Buy And Either"
   End If
   MDISect.Crw.Formulas(5) = "Mbe='" & sMbe & " AND BOM Levels " _
                        & "Through " & cmbLvl & "'"
   If optDsc.value = vbChecked Then
      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;T;;;"
   Else
      MDISect.Crw.SectionFormat(0) = "GROUPHDR.0.1;F;;;"
   End If
   
   MDISect.Crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextBillLevel4(sPartNumber As String, AssyRef As String)
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel4"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',4)"
            RdoCon.Execute sIns, rdExecDirect
            GetNextBillLevel5 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel5(sPartNumber As String, AssyRef As String)
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel5"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',5)"
            GetNextBillLevel6 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub


Private Sub GetNextBillLevel6(sPartNumber As String, AssyRef As String)
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel6"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',6)"
            RdoCon.Execute sIns, rdExecDirect
            GetNextBillLevel7 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub GetNextBillLevel7(sPartNumber As String, AssyRef As String)
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel7"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
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
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel3"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',3)"
            RdoCon.Execute sIns, rdExecDirect
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


Private Sub cmbPrt_Click()
'MsgBox cmbPrt
   'cmbPrt = GetCurrentPart(cmbPrt.Text, lblDsc)   'infinite loop
   Dim s As String
   s = GetCurrentPart(cmbPrt.Text, lblDsc)
'MsgBox cmbPrt
End Sub


Private Sub cmbPrt_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   Dim sDesc As String
   
'   For iList = 0 To cmbPrt.ListCount - 1
'      If iList = 32767 Then
'         Debug.Print "32767"
'      End If
'      If Compress(cmbPrt) = Compress(cmbPrt.List(iList)) Then bByte = 1
'   Next

   bByte = 1   'just configure as listbox - don't let user type
   
   If bByte = 0 Then
      Beep
      cmbPrt = cmbPrt.List(0)
   End If
   
   'ignore error if GetCurrentPart reeturns an invalid
   On Error Resume Next
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
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
   
End Sub

Private Sub FillCombo()
   Dim sDesc As String
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT MRP_PARTREF,MRP_PARTNUM FROM " _
          & "MrplTable WHERE MRP_PARTREF <>'' ORDER BY MRP_PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount <> 0 Then
      cmbPrt = cmbPrt.List(0)
      cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   Else
      MsgBox "There Is No Current MRP Available.", _
         vbInformation, Caption
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Not using recursion here to keep the levels straight and
'make it easy to read

Private Sub GetBill()
   Dim RdoBom As rdoResultset
   MouseCursor 13
   iOrder = 0
   On Error GoTo DiaErr1
   sProcName = "getbill"
   sSql = "truncate table MrpbTable"
   RdoCon.Execute sSql, rdExecDirect
   sSql = "INSERT INTO MrpbTable (MRPBILL_PARTREF," _
          & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
          & "VALUES('" & Compress(cmbPrt) & "','" _
          & "',0)"
   RdoCon.Execute sSql, rdExecDirect
   
   sBomRev = GetBomRev(Compress(cmbPrt))
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM BmplTable " _
          & "WHERE (BMASSYPART='" & Compress(cmbPrt) & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
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
            RdoCon.Execute sIns, rdExecDirect
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
   Dim RdoBom As rdoResultset
   sBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel2"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & sBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = GetDataSet(RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbTable (MRPBILL_ORDER,MRPBILL_PARTREF," _
                   & "MRPBILL_USEDON,MRPBILL_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "',2)"
            RdoCon.Execute sIns, rdExecDirect
            GetNextBillLevel3 Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
End Sub

Private Sub optDis_Click()
   GetBill
   
End Sub


Private Sub optPrn_Click()
   GetBill
   
End Sub



Private Function GetBomRev(sPartNumber) As String
   Dim RdoRev As rdoResultset
   sProcName = "getbomrev"
   sSql = "SELECT PARTREF,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = GetDataSet(RdoRev, ES_FORWARD)
   If bSqlRows Then
      With RdoRev
         GetBomRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoRev
      End With
   Else
      GetBomRev = ""
   End If
   
End Function

Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiProd", "Prdmr03", optDsc.value
   
End Sub
