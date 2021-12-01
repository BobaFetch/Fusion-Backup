VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaSCp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exploded Proposed Standard Cost Analysis (Report)"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      Top             =   600
      Width           =   2775
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   3600
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox cmbPrt1 
      Height          =   285
      Left            =   2040
      TabIndex        =   25
      Tag             =   "3"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdVew 
      Height          =   320
      Left            =   4920
      Picture         =   "diaSCp02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Show BOM Structure"
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   7200
      Top             =   3120
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3585
      FormDesignWidth =   7785
   End
   Begin VB.CheckBox chkB 
      Caption         =   "___"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox chkLab 
      Caption         =   "___"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox chkExp 
      Caption         =   "___"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox chkStd 
      Caption         =   "___"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkSum 
      Caption         =   "___"
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6600
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Save And Exit"
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6600
      TabIndex        =   8
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   11
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCp02a.frx":0342
      PictureDn       =   "diaSCp02a.frx":0488
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Show System Printers"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaSCp02a.frx":05CE
      PictureDn       =   "diaSCp02a.frx":0714
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Treated Like Raw Materials Otherwise)"
      Height          =   285
      Index           =   12
      Left            =   4080
      TabIndex        =   23
      Top             =   2880
      Width           =   3705
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Otherwise, Primary And Secondary Shops Used)"
      Height          =   285
      Index           =   11
      Left            =   4080
      TabIndex        =   22
      Top             =   2520
      Width           =   3585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Otherwise, Proposed Expense Used)"
      Height          =   285
      Index           =   10
      Left            =   4080
      TabIndex        =   21
      Top             =   2160
      Width           =   3345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(At All Levels)"
      Height          =   285
      Index           =   9
      Left            =   4080
      TabIndex        =   20
      Top             =   1800
      Width           =   3345
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Based On BOM For ""B"" Parts?"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   2865
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Labor Cost From Routings?"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   3225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Expense Cost From Routings?"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   3225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update Standard Cost?"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   2625
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(i. e. Bypass Subassembly Analysis)"
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   14
      Top             =   1440
      Width           =   3345
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Summary Information Only? "
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   3225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Explored Proposed Cost Analysis For Part?"
      Height          =   525
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2145
   End
End
Attribute VB_Name = "diaSCp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaSCp02a - Exploded cost analysis for part (Report)
'
' Notes:
'
' Created: 11/30/01 (nth)
' Revisions:
'   06/05/02 (nth) Added CreateJetTable subroutine.
'   06/05/02 (nth) Added standard costing logic.
'   06/05/02 (nth) Updated crystal report.
'   06/05/02 (nth) Added ALL levels costing logic derived from BOM jet report table
'   06/10/02 (nth) Changed report to group by parts rather than level.
'   12/04/02 (nth) Removed part combo for and repalced with part lookup
'   12/04/02 (nth) Added saveoptions and getoptions
'
'*************************************************************************************

Dim bOnLoad As Byte
'Dim dbBOM As Recordset 'jet
Dim RdoPrt As ADODB.Recordset
Dim AdoQry1 As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter

Dim AdoQry2 As ADODB.Command
Dim AdoParameter2 As ADODB.Parameter
Dim AdoParameter3 As ADODB.Parameter

Dim bGoodPart As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmbPrt_Click()
   If Not bCancel Then
      cmbPrt = CheckLen(cmbPrt, 30)
      FindPart Me
   End If
End Sub

Private Sub cmbPrt_Change()
   If Not bCancel Then
      cmbPrt = CheckLen(cmbPrt, 30)
      FindPart Me
   End If
End Sub


Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub optVew_Click()
   If optVew.Value = vbUnchecked Then
      ' Part search is closing refresh form
      'cmbPrt_LostFocus
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "No Subject Help"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      If Len(cUR.CurrentPart) Then
         cmbPrt = cUR.CurrentPart
         FindPart Me
         FillPartCombo cmbPrt
      End If
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   ReopenJet
   CreateQueries
   GetOptions
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If bGoodPart Then
      cUR.CurrentPart = Trim(cmbPrt)
      SaveCurrentSelections
   End If
   Set RdoPrt = Nothing
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoParameter3 = Nothing
   Set AdoQry2 = Nothing
   Set AdoQry1 = Nothing
   'JetDb.Execute "DROP TABLE BomTable"
   FormUnload
   Set diaSCp02a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim bSuccess As Byte
   'Dim dbAllLevels As Recordset 'jet
   Dim RdoAllLevels As ADODB.Recordset
   Dim sWindows As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   GetBillOfMaterial Compress(cmbPrt), ""
   
   If chkStd Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "SELECT DISTINCT BomPartRef,BomLevel FROM EsReportTmpBomTable " _
             & "WHERE BomPartRef IS NOT NULL ORDER BY BomLevel DESC"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoAllLevels, ES_FORWARD)
      If bSqlRows Then
         With RdoAllLevels
            bSuccess = 1
            While Not .EOF And bSuccess = 1
               bSuccess = StdCostPart(Trim(!BomPartRef))
               .MoveNext
            Wend
         End With
      End If
      Set RdoAllLevels = Nothing

      ' Now cost the actual assembly
      If bSuccess Then
         bSuccess = StdCostPart(Compress(cmbPrt))
      End If
      
      ' If everything was successfull then...
      If bSuccess Then
         clsADOCon.CommitTrans
         clsADOCon.ADOErrNum = 0

         SysMsg "Standard Cost Updated At All Levels.", True
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0

         Exit Sub
      End If
   End If
   
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim strSumDetail As String
   
   sCustomReport = GetCustomReport("finpc06")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = "finpc06.rpt"
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Summary/Detail"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")

   If chkSum Then
      strSumDetail = "(Summary)"
   Else
      strSumDetail = "(Detail)"
   End If
   aFormulaValue.Add CStr("'" & strSumDetail & "'")

   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection
   'cCRViewer.SetTableConnection aRptPara
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
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

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

' Fill the Jet table
' Note: Only first level here, see GetLevelNext function
' Source derived from BOM logic

Public Sub GetBillOfMaterial(sPartRef1 As String, sRev1 As String)
   
   On Error GoTo DiaErr1
   Err.Clear
   
   sSql = "DELETE EsReportTmpBomTable WHERE BomParent = '" & sPartRef1 & "'"
   clsADOCon.ExecuteSQL sSql ' rdExecDirect
   
   ' Create the BOM partlist
   sSql = "WITH cte AS " & vbCrLf
   sSql = sSql & "   (select BMASSYPART,BMPARTREF,BMPARTREV, BMQTYREQD , RTrim(BMUNITS) BMUNITS," & vbCrLf
   sSql = sSql & "      BMCONVERSION, BMSEQUENCE, 0 as level," & vbCrLf
   sSql = sSql & "      cast(LTRIM(RTrim(BMASSYPART)) + char(36)+ COALESCE(cast(BMSEQUENCE as varchar(4)), '') + " & vbCrLf
   sSql = sSql & "      LTRIM(RTrim(BMPARTREF)) as varchar(max)) as SortKey" & vbCrLf
   sSql = sSql & "   From BmplTable" & vbCrLf
   sSql = sSql & "   where BMASSYPART = '" & sPartRef1 & "' AND BMPARTREV = '" & sRev1 & "'" & vbCrLf
   sSql = sSql & "   UNION ALL" & vbCrLf
   sSql = sSql & "   SELECT a.BMASSYPART,a.BMPARTREF,a.BMPARTREV, a.BMQTYREQD , RTrim(a.BMUNITS) BMUNITS," & vbCrLf
   sSql = sSql & "      a.BMCONVERSION, a.BMSEQUENCE, level + 1," & vbCrLf
   sSql = sSql & "      cast(COALESCE(SortKey,'') + char(36) + COALESCE(cast(a.BMSEQUENCE as varchar(4)), '') + COALESCE(LTRIM(RTrim(a.BMPARTREF)) ,'') as varchar(max))as SortKey" & vbCrLf
   sSql = sSql & "   FROM cte" & vbCrLf
   sSql = sSql & "      inner join BmplTable a" & vbCrLf
   sSql = sSql & "         on cte.BMPARTREF = a.BMASSYPART" & vbCrLf
   sSql = sSql & ")" & vbCrLf
   sSql = sSql & "   INSERT INTO EsReportTmpBomTable(BomAssembly, BomPartRef, BomLevel,BomQuantity,BomSequence, BomSortKey)" & vbCrLf
   sSql = sSql & "   SELECT BMASSYPART, BMPARTREF, level,BMQTYREQD , BMSEQUENCE,  SortKey" & vbCrLf
   sSql = sSql & "            from cte order by SortKey, BMSEQUENCE" & vbCrLf
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   ' Update the parent parent
   sSql = "UPDATE EsReportTmpBomTable SET BomParent = '" & sPartRef1 & "'"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "GetBillOfMaterial"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

' Standard cost a part at ALL levels
' Assumes a jet db to exists listing all parts in BOM

Public Function StdCostPart(SPartRef As String) As Byte
   Dim ThisPartCost As PartCost
   
   On Error GoTo DiaErr1
   AdoQry1.parameters(0).Value = SPartRef
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry1, ES_KEYSET)
   
   If bSqlRows Then
      With RdoPrt
         ' Call the standard cost class/module
         IniStdCost
         ' Calculate All "PALEV" cost
         ThisPartCost = CostPart(SPartRef)
         ' Calculate all "PABOM" cost load into an array
         CostPartBOM SPartRef, "" & Trim(RdoPrt!PABOMREV), 1
         
         On Error Resume Next
         !PALEVLABOR = ThisPartCost.nLabor
         !PALEVEXP = ThisPartCost.nExpense
         !PALEVMATL = ThisPartCost.nMaterial
         !PALEVOH = ThisPartCost.nOverhead
         !PALEVHRS = ThisPartCost.nHours
         
         !PABOMLABOR = BomCost(0).nLabor
         !PABOMEXP = BomCost(0).nExpense
         !PABOMMATL = BomCost(0).nMaterial
         !PABOMOH = BomCost(0).nOverhead
         !PABOMHRS = BomCost(0).nHours
         
         ' Copy current standard to previous standard
         ' before updating current standard
         !PAPREVLABOR = !PATOTLABOR
         !PAPREVEXP = !PATOTEXP
         !PAPREVMATL = !PATOTMATL
         !PAPREVOH = !PATOTOH
         !PAPREVHRS = !PATOTHRS
         !PAPREVSTDCOST = !PASTDCOST
         
         !PATOTLABOR = ThisPartCost.nLabor + BomCost(0).nLabor
         !PATOTEXP = ThisPartCost.nExpense + BomCost(0).nExpense
         !PATOTMATL = ThisPartCost.nMaterial + BomCost(0).nMaterial
         !PATOTOH = ThisPartCost.nOverhead + BomCost(0).nOverhead
         !PATOTHRS = ThisPartCost.nHours + BomCost(0).nHours
         
         !PASTDCOST = (ThisPartCost.nLabor + BomCost(0).nLabor + _
                      ThisPartCost.nExpense + BomCost(0).nExpense + _
                      ThisPartCost.nMaterial + BomCost(0).nMaterial + _
                      ThisPartCost.nOverhead + BomCost(0).nOverhead + _
                      ThisPartCost.nHours + BomCost(0).nHours)
         .Update
         
         If Err <> 0 Then
            ValidateEdit Me
         Else
            StdCostPart = 1
         End If
      End With
   End If
   Exit Function
   
DiaErr1:
   sProcName = "StdCostPart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub SaveOptions()
   Dim sOptions As String
   'Save by Menu Option
   sOptions = RTrim(chkSum.Value) _
              & RTrim(chkStd.Value) _
              & RTrim(chkExp.Value) _
              & RTrim(chkLab.Value) _
              & RTrim(chkB.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      chkSum.Value = Val(Left(sOptions, 1))
      chkStd.Value = Val(Mid(sOptions, 2, 1))
      chkExp.Value = Val(Mid(sOptions, 3, 1))
      chkLab.Value = Val(Mid(sOptions, 4, 1))
      chkB.Value = Val(Mid(sOptions, 5, 1))
   Else
      chkSum.Value = vbUnchecked
      chkStd.Value = vbUnchecked
      chkExp.Value = vbUnchecked
      chkLab.Value = vbUnchecked
      chkB.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
End Sub

Public Sub CreateQueries()
   ' All cost levels for part
   sSql = "SELECT PARTREF, PARTNUM, PADESC, PALEVEL, PAREVDATE," _
          & "PAEXTDESC, PAMAKEBUY, PALEVLABOR, PALEVEXP, PALEVMATL, PALEVOH," _
          & "PALEVHRS, PASTDCOST, PABOMLABOR, PABOMEXP, PABOMMATL, PABOMOH," _
          & "PABOMHRS, PABOMREV, PAPREVLABOR, PAPREVEXP, PAPREVMATL, PAPREVOH," _
          & "PAPREVHRS,PAPREVSTDCOST, PATOTHRS, PATOTEXP, PATOTLABOR, PATOTMATL," _
          & "PATOTOH,PAROUTING,PARRQ,PAEOQ " _
          & "FROM PartTable WHERE PARTREF = ?"
   Set AdoQry1 = New ADODB.Command
   AdoQry1.CommandText = sSql
   
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adChar
   AdoParameter1.SIZE = 30
   AdoQry1.parameters.Append AdoParameter1
   
   
   ' Operations, workcenters, and shops associated with part
   'sSql = "SELECT OPREF,OPNO,OPSETUP,OPUNIT,WCNSTDRATE,WCNOHFIXED,SHPRATE,SHPOHTOTAL, " _
   '       & "WCNOHPCT,WCNOHFIXED,OPSERVPART,PASTDCOST FROM RtopTable " _
   '       & "INNER JOIN WcntTable ON RtopTable.OPCENTER = WcntTable.WCNREF " _
   '       & "INNER JOIN ShopTable ON RtopTable.OPSHOP = ShopTable.SHPREF " _
   '       & "LEFT JOIN PartTable On RtopTable.OPSERVPART = PartTable.PARTREF " _
   '       & "WHERE (RtopTable.OPREF = ?)"
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   
   ' Recursive BOM for part
   sSql = "SELECT BMASSYPART,BMPARTREF,BMPARTNUM,BMREV,BMSEQUENCE," _
          & "BMQTYREQD,BMUNITS,BMCONVERSION,BMPARTREV FROM BmplTable " _
          & "WHERE BMASSYPART= ? AND BMREV= ? ORDER BY BMSEQUENCE,BMPARTREF"
   Set AdoQry2 = New ADODB.Command
   AdoQry2.CommandText = sSql

   
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adChar
   AdoParameter2.SIZE = 30
   AdoQry2.parameters.Append AdoParameter2
   
   Set AdoParameter3 = New ADODB.Parameter
   AdoParameter3.Type = adChar
   AdoParameter3.SIZE = 2
   AdoQry2.parameters.Append AdoParameter3

End Sub
