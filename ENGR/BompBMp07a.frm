VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BompBMp07a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exploded Used On (Report)"
   ClientHeight    =   1980
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMp07a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6140
      TabIndex        =   7
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "BompBMp07a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "BompBMp07a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6140
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cmbRev 
      ForeColor       =   &H00800000&
      Height          =   288
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   1110
      Width           =   975
   End
   Begin VB.ComboBox cmbPls 
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Only Parts With A Parts List"
      Top             =   1110
      Width           =   3345
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6780
      Top             =   1500
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1980
      FormDesignWidth =   7275
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1920
      TabIndex        =   10
      Top             =   1464
      Width           =   3132
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6240
      TabIndex        =   9
      Top             =   1464
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   288
      Index           =   8
      Left            =   5400
      TabIndex        =   8
      Top             =   1464
      Width           =   852
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   252
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Top             =   1110
      Width           =   972
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   288
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1110
      Width           =   1812
   End
End
Attribute VB_Name = "BompBMp07a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
Option Explicit
'Dim RdoQry As rdoQuery
Dim AdoCmdObj As ADODB.Command
Dim bBol As Boolean

Dim bCheck As Byte
Dim bfirstRun As Byte
Dim bGoodList As Byte
Dim bOnLoad As Byte

Dim iRow As Integer

Dim sPartNum As String
Dim sPartDesc As String
Dim sPartExtDesc As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub
Private Sub cmbPls_Click()
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub

Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   FillBomhRev cmbPls
   GetPartRevision
   
End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 907
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      FillBomhRev cmbPls
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
'   sSql = "SELECT DISTINCT BMASSYPART,PARTREF,PARTNUM FROM " _
'          & "BmplTable,PartTable WHERE BMPARTREF=PARTREF " _
'          & "ORDER BY PARTREF"
'   LoadComboBox cmbPls, 1
   sSql = "SELECT DISTINCT PARTNUM" & vbCrLf _
      & "FROM BmplTable" & vbCrLf _
      & "JOIN PartTable ON BMPARTREF = PARTREF" & vbCrLf _
      & "ORDER BY PARTNUM"
   LoadComboBox cmbPls, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV,BMSEQUENCE," _
          & "BMQTYREQD,BMUNITS,BMCONVERSION,BMPARTREV FROM BmplTable " _
          & "WHERE BMASSYPART= ? AND BMREV= ? ORDER BY BMSEQUENCE,BMPARTREF"
   
   Set AdoCmdObj = New ADODB.Command
   AdoCmdObj.CommandText = sSql
   
   Dim prmPrtRef As ADODB.Parameter
   Set prmPrtRef = New ADODB.Parameter
   prmPrtRef.Type = adChar
   prmPrtRef.Size = 30
   AdoCmdObj.Parameters.Append prmPrtRef
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'GetOptions
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'SaveOptions
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'DeleteDataForThisUser
   Set AdoCmdObj = Nothing
   FormUnload
End Sub

Private Sub DeleteDataForThisUser()
   sSql = "delete from EsReportBomTable where BomUser = '" & sInitials & "'"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
End Sub


Private Sub PrintReport()
   Dim sCustomReport As String
   On Error GoTo whoops
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   
   sCustomReport = GetCustomReport("engbm07")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Exploded Used-On for Part " & Trim(cmbPls) & " Rev " & cmbRev & "'")
   
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   ' Set report parameter
   cCRViewer.SetDbTableConnection True
   ' report parameter
   aRptPara.Add CStr(Compress(cmbPls))
   aRptParaType.Add CStr("String")
   
   If cmbRev = "" Then
      aRptPara.Add CStr(" ")
   Else
      aRptPara.Add CStr(Trim(cmbRev))
   End If
   aRptParaType.Add CStr("String")
   
   ' Set report parameter
   
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.SetReportDBParameters aRptPara, aRptParaType    'must happen AFTER SetDbTableConnection call!
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
   Exit Sub
   
whoops:
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub optDis_Click()
   CheckBom
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub

Private Function GetList() As Byte
   Dim RdoLst As ADODB.Recordset
   Dim sPartNumber As String
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM " _
          & "BmhdTable WHERE BMHREF='" & sPartNumber & "' " _
          & "AND BMHREV='" & Trim(cmbRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst)
   If bSqlRows Then
      With RdoLst
         cmbRev = "" & Trim(!BMHREV)
         GetList = 1
         ClearResultSet RdoLst
      End With
   Else
      GetList = 0
   End If
   Set RdoLst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getlist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function

Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optMat_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   CheckBom
End Sub

Private Sub optRaw_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
End Sub


Private Sub optRef_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Private Sub GetPartRevision()
   Dim RdoRev As ADODB.Recordset
   Dim sPartNumber As String
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PADESC,PALEVEL,PABOMREV FROM " _
          & "PartTable WHERE PARTREF='" & sPartNumber & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev)
   If bSqlRows Then
      With RdoRev
         cmbRev = "" & Trim(!PABOMREV)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = "" & Trim(str(!PALEVEL))
         ClearResultSet RdoRev
      End With
   Else
      cmbRev = ""
   End If
   Set RdoRev = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartrev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetBillOfMaterials() As Boolean
   'return = True if successful
   
   On Error GoTo DiaErr1
   DeleteDataForThisUser
   
   'insert part to be exploded
   Dim assy As String
   assy = Compress(cmbPls)
   sSql = "INSERT INTO EsReportBomTable" & vbCrLf _
      & "(BomUser,BomLevel,BomAssembly,BomPartRef,BomRevision," & vbCrLf _
      & "BomQuantity,BomUnits,BomConversion,BomSequence,BomSortKey," & vbCrLf _
      & "ExplodedQty, MostRecentCost)" & vbCrLf _
      & "SELECT TOP 1 '" & sInitials & "',0,'" & assy & "','" & assy & "','" & Compress(cmbRev) & "'," & vbCrLf _
      & "1,'',0,0,''," & vbCrLf _
      & "1,ISNULL(LOTUNITCOST,0)" & vbCrLf _
      & "FROM PartTable" & vbCrLf _
      & "LEFT JOIN LohdTable on LOTPARTREF = PARTREF" & vbCrLf _
      & "WHERE PARTREF = '" & assy & "'" & vbCrLf _
      & "ORDER BY LOTADATE DESC"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   'keep inserting next level parts until there are no more
   Dim level As Integer
   For level = 0 To 10
      sSql = "INSERT INTO EsReportBomTable" & vbCrLf _
         & "(BomUser,BomLevel,BomAssembly,BomPartRef,BomRevision," & vbCrLf _
         & "BomQuantity,BomUnits,BomConversion,BomSequence,BomSortKey," & vbCrLf _
         & "ExplodedQty, MostRecentCost)" & vbCrLf _
         & "SELECT '" & sInitials & "'," & level + 1 & ",BMASSYPART,BMPARTREF,BMPARTREV," & vbCrLf _
         & "BMQTYREQD,RTRIM(BMUNITS),BMCONVERSION,BMSEQUENCE," & vbCrLf _
         & "BomSortKey + replicate('0',5-len(cast(BMSEQUENCE as varchar(5))))" & vbCrLf _
         & " + cast(BMSEQUENCE as varchar(5))," & vbCrLf _
         & "cast(ExplodedQty * BMQTYREQD as DECIMAL(15,4))," & vbCrLf _
         & "ISNULL((SELECT TOP 1 LOTUNITCOST FROM LohdTable" & vbCrLf _
         & "WHERE LOTPARTREF = BMPARTREF ORDER BY LOTADATE DESC),0)" & vbCrLf _
         & "FROM BmplTable" & vbCrLf _
         & "JOIN EsReportBomTable on BMASSYPART = BomPartRef" & vbCrLf _
         & "AND BMREV = BomRevision AND BomLevel = " & level & vbCrLf _
         & "JOIN PartTable on PARTREF = BMPARTREF" & vbCrLf _
         & "WHERE BomUser = '" & sInitials & "'" & vbCrLf _
         & "ORDER BY BMSEQUENCE,BMPARTREF"
      
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      
      ' TODO: Make sure this works
      If clsADOCon.RowsAffected = 0 Then
         Exit For
      End If
   Next
   
''   'create BomRow = integer sort order to display exploded BOM
''   'DOESN'T WORK WITH SQL2000 - JUST SORT BY BomSortKey
''   sSql = "Update EsReportBomTable" & vbCrLf _
''      & "set BomRow = x.rowno from (select ROW_NUMBER() OVER (ORDER BY BomSortKey) as rowno, BomSortKey as sort" & vbCrLf _
''      & "from EsReportBomTable) as x" & vbCrLf _
''      & "join EsReportBomTable bom2 on x.sort = bom2.BomSortKey" & vbCrLf _
''      & "where BomUser = '" & sInitials & "'"
''   RdoCon.Execute sSql
   bCheck = 1
   GetBillOfMaterials = True
   MouseCursor 0
   Exit Function
   
DiaErr1:
   bCheck = 1
   sProcName = "GetBillOfMaterials"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CheckBom()
   Dim bGoodRout As Byte
   Dim bSProcExist As Boolean
   
   bCheck = 0
   If GetBillOfMaterials Then
      bSProcExist = StoreProcedureExists("UpdateBOMNumbers")
   
      If bSProcExist Then
         sSql = "UpdateBOMNumbers"
         clsADOCon.ExecuteSql sSql ' rdExecDirect
      End If
   
      PrintReport
   End If

'   bGoodList = GetList()
'   If bGoodList = 0 Then
'      MsgBox "Parts List Wasn't Found.", vbExclamation, Caption
'      Exit Sub
'   Else
'      If GetBillOfMaterials Then
'         PrintReport
'      End If
'   End If
   
End Sub
