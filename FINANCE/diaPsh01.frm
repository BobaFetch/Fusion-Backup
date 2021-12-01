VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPsh01
   BorderStyle = 3 'Fixed Dialog
   Caption = "Manufacturing Orders"
   ClientHeight = 5895
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 7470
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5895
   ScaleWidth = 7470
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton optPrn
      Height = 330
      Left = 6800
      Picture = "diaPsh01.frx":0000
      Style = 1 'Graphical
      TabIndex = 44
      ToolTipText = "Print The Report"
      Top = 600
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.CommandButton optDis
      Height = 330
      Left = 6240
      Picture = "diaPsh01.frx":018A
      Style = 1 'Graphical
      TabIndex = 42
      ToolTipText = "Display The Report"
      Top = 600
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.CheckBox chkCustom
      Height = 240
      Left = 2760
      TabIndex = 41
      Top = 5535
      Width = 255
   End
   Begin VB.CheckBox optLst
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2760
      TabIndex = 7
      ToolTipText = "Pick List For This Part (Printed MO's Only) Status PL"
      Top = 3480
      Width = 726
   End
   Begin VB.CheckBox optDoc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2760
      TabIndex = 6
      ToolTipText = "Document List (Printed MO's Only)"
      Top = 3240
      Width = 735
   End
   Begin VB.CheckBox optBud
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2760
      TabIndex = 5
      Top = 2760
      Width = 735
   End
   Begin VB.CheckBox optFrom
      Height = 255
      Left = 4320
      TabIndex = 36
      Top = 4560
      Visible = 0 'False
      Width = 735
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 6240
      TabIndex = 33
      Top = 480
      Width = 1095
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6240
      TabIndex = 32
      TabStop = 0 'False
      Top = 120
      Width = 1065
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 12
      Left = 2760
      TabIndex = 8
      Top = 3720
      Width = 735
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 11
      Left = 2760
      TabIndex = 14
      Top = 5040
      Width = 735
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 10
      Left = 2760
      TabIndex = 13
      Top = 4800
      Width = 735
   End
   Begin VB.CheckBox optInc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Index = 9
      Left = 3960
      TabIndex = 12
      Top = 5160
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 8
      Left = 2760
      TabIndex = 11
      Top = 4560
      Width = 735
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 7
      Left = 2760
      TabIndex = 10
      Top = 4320
      Width = 735
   End
   Begin VB.CheckBox optInc
      Enabled = 0 'False
      ForeColor = &H00C0C0C0&
      Height = 255
      Index = 5
      Left = 2760
      TabIndex = 9
      Top = 4080
      Width = 735
   End
   Begin VB.CheckBox optInc
      Caption = "____"
      Enabled = 0 'False
      ForeColor = &H8000000F&
      Height = 255
      Index = 4
      Left = 6480
      TabIndex = 15
      Top = 4920
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CheckBox optInc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Index = 3
      Left = 2760
      TabIndex = 4
      Top = 2520
      Width = 735
   End
   Begin VB.CheckBox optInc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Index = 2
      Left = 2760
      TabIndex = 3
      Top = 2280
      Width = 735
   End
   Begin VB.CheckBox optInc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Index = 1
      Left = 2760
      TabIndex = 2
      Top = 2040
      Width = 735
   End
   Begin VB.ComboBox cmbRun
      ForeColor = &H00800000&
      Height = 315
      Left = 6240
      TabIndex = 1
      Tag = "2"
      ToolTipText = "Select Run Number"
      Top = 1080
      Width = 1095
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 1560
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Select Part Number"
      Top = 1080
      Width = 3255
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6720
      Top = 4200
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5895
      FormDesignWidth = 7470
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 43
      ToolTipText = "Subject Help"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 397
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      Autosize = 2
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPsh01.frx":0308
      PictureDn = "diaPsh01.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 45
      ToolTipText = "Show System Printers"
      Top = 0
      Width = 255
      _Version = 65536
      _ExtentX = 450
      _ExtentY = 450
      _StockProps = 65
      BackColor = 12632256
      GroupAllowAllUp = -1 'True
      RoundedCorners = 0 'False
      BevelWidth = 0
      Outline = 0 'False
      PictureUp = "diaPsh01.frx":0594
      PictureDn = "diaPsh01.frx":06DA
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 46
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Use Custom MO Format"
      Enabled = 0 'False
      Height = 255
      Index = 18
      Left = 240
      TabIndex = 40
      Top = 5535
      Width = 2325
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Pick List For This Part"
      Height = 255
      Index = 17
      Left = 480
      TabIndex = 39
      ToolTipText = "Pick List For This Part (Printed MO's Only) Status PL"
      Top = 3480
      Width = 2535
   End
   Begin VB.Label lblSta
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 6840
      TabIndex = 38
      Top = 1440
      Width = 495
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "MO Budgets"
      Height = 255
      Index = 16
      Left = 240
      TabIndex = 37
      Top = 2760
      Width = 2535
   End
   Begin VB.Label lblTyp
      Alignment = 1 'Right Justify
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 6240
      TabIndex = 35
      Top = 1440
      Width = 300
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Type/Status"
      Height = 255
      Index = 15
      Left = 5160
      TabIndex = 34
      Top = 1440
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Tool Information"
      Enabled = 0 'False
      Height = 255
      Index = 14
      Left = 480
      TabIndex = 31
      Top = 3720
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Inspection Document"
      Enabled = 0 'False
      Height = 255
      Index = 13
      Left = 240
      TabIndex = 30
      Top = 5040
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Bar Code"
      Enabled = 0 'False
      Height = 255
      Index = 12
      Left = 240
      TabIndex = 29
      Top = 4800
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Cover Sheet (Printed MO's Only):"
      Height = 255
      Index = 11
      Left = 240
      TabIndex = 28
      Top = 3000
      Width = 2655
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "PO Allocations"
      Enabled = 0 'False
      Height = 255
      Index = 10
      Left = 240
      TabIndex = 27
      Top = 4560
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "MO Comments"
      Enabled = 0 'False
      Height = 255
      Index = 9
      Left = 240
      TabIndex = 26
      Top = 4320
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Document List For This Part"
      Height = 255
      Index = 8
      Left = 480
      TabIndex = 25
      ToolTipText = "Document List (Printed MO's Only)"
      Top = 3240
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "SO Allocations"
      Enabled = 0 'False
      Height = 255
      Index = 7
      Left = 240
      TabIndex = 24
      Top = 4080
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "MO Allocations"
      Enabled = 0 'False
      Height = 255
      Index = 6
      Left = 3960
      TabIndex = 23
      Top = 4920
      Visible = 0 'False
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Outside Service Part Numbers"
      Height = 255
      Index = 5
      Left = 240
      TabIndex = 22
      Top = 2520
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Operation Comments"
      Height = 255
      Index = 4
      Left = 240
      TabIndex = 21
      Top = 2280
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions"
      Height = 255
      Index = 3
      Left = 240
      TabIndex = 20
      Top = 2040
      Width = 2535
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 375
      Index = 2
      Left = 240
      TabIndex = 19
      Top = 1800
      Width = 975
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 255
      Index = 0
      Left = 240
      TabIndex = 18
      Top = 1080
      Width = 1095
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Run"
      Height = 255
      Index = 1
      Left = 5160
      TabIndex = 17
      Top = 1080
      Width = 855
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1560
      TabIndex = 16
      Top = 1440
      Width = 3255
   End
End
Attribute VB_Name = "diaPsh01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RdoQry As rdoQuery
Dim DbDoc As Recordset 'Jet
Dim DbPls As Recordset 'Jet

Dim bGoodDocs As Boolean
Dim bGoodPlst As Boolean

Dim bGoodPart As Byte
Dim bGoodMo As Byte
Dim bOnLoad As Byte

Dim sRunPkstart As String
Dim sPartRef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub GetOptions()
   Dim i As Integer
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh01", sOptions)
   If Len(sOptions) > 0 Then
      For i = 1 To 5
         optInc(i) = Val(Mid(sOptions, i, 1))
      Next
      For i = 7 To 11
         optInc(i) = Val(Mid(sOptions, i, 1))
      Next
      optBud = Val(Mid(sOptions, i, 1))
   End If
   
End Sub


Public Sub SaveOptions()
   Dim i As Integer
   Dim sOptions As String
   'Save by Menu Option
   For i = 1 To 5
      sOptions = sOptions & Trim(Val(optInc(i).Value))
   Next
   For i = 7 To 11
      sOptions = sOptions & Trim(Val(optInc(i).Value))
   Next
   sOptions = sOptions & Trim(Val(optInc(i).Value))
   sOptions = sOptions & Trim(Val(optBud.Value))
   SaveSetting "Esi2000", "EsiProd", "sh01", Trim(sOptions)
   
End Sub




Private Sub cmbPrt_Click()
   bGoodPart = GetRuns()
   
End Sub


Private Sub cmbprt_LostFocus()
   cmbprt = CheckLen(cmbprt, 30)
   bGoodPart = GetRuns()
   
End Sub

Private Sub PrintReport()
   MouseCursor 13
   
   On Error GoTo Psh01
   SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   
   If chkCustom.Value = vbUnchecked Then
      If optInc(1) Then
         MdiSect.crw.ReportFileName = sReportPath & "prdsh01.rpt"
      Else
         MdiSect.crw.ReportFileName = sReportPath & "prdsh01a.rpt"
      End If
      If optInc(2).Value = vbUnchecked Then
         MdiSect.crw.SectionFormat(0) = "DETAIL.0.0;F;;;"
         MdiSect.crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
      Else
         MdiSect.crw.SectionFormat(0) = "DETAIL.0.0;T;;;"
         MdiSect.crw.SectionFormat(1) = "DETAIL.0.1;T;;;"
      End If
      If optInc(3).Value = vbUnchecked Then
         MdiSect.crw.SectionFormat(2) = "GROUPFTR.0.0;F;;;"
         MdiSect.crw.SectionFormat(3) = "GROUPFTR.0.1;F;;;"
      Else
         MdiSect.crw.SectionFormat(2) = "GROUPFTR.0.0;T;;;"
         MdiSect.crw.SectionFormat(3) = "GROUPFTR.0.1;T;;;"
      End If
      If optInc(1).Value = vbChecked Then
         If optInc(2).Value = vbChecked Then
            MdiSect.crw.SectionFormat(1) = "DETAIL.0.1;F;;;"
         End If
      End If
      If optBud.Value = vbChecked Then
         MdiSect.crw.SectionFormat(4) = "REPORTFTR.0.0;T;;;"
      Else
         MdiSect.crw.SectionFormat(4) = "REPORTFTR.0.0;F;;;"
      End If
      
   Else
      MdiSect.crw.ReportFileName = "C:\Custom ES2000 Reports\AUSWAT\mfplan.rpt"
   End If
   
   sSql = "{RunsTable.RUNREF}='" & sPartRef & "' " _
          & "AND {RunsTable.RUNNO}=" & Trim(cmbRun) & " "
   MdiSect.crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   Psh01:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
   Psh02:
   DoModuleErrors Me
   
End Sub

Private Sub cmbRun_Click()
   GetThisRun
   
End Sub


Private Sub cmbRun_LostFocus()
   GetThisRun
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillRuns Me, "<>'CA'"
      If optFrom Then
         cmbprt = diaSrvmo.cmbprt
         cmbRun = diaSrvmo.cmbRun
      Else
         If Len(Trim(Cur.CurrentPart)) Then cmbprt = Cur.CurrentPart
      End If
      bGoodPart = GetRuns()
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PARUN,RUNREF,RUNSTATUS," _
          & "RUNNO FROM PartTable,RunsTable WHERE PARTREF= ? " _
          & "AND PARTREF=RUNREF AND RUNSTATUS<>'CA'"
   Set RdoQry = RdoCon.CreateQuery("", sSql)
   bOnLoad = True
   CreatePlsTable
   CreateCovrTable
   GetOptions
   
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   JetDb.Execute "DROP TABLE CvrTable"
   JetDb.Execute "DROP TABLE PlsTable"
   Set RdoQry = Nothing
   rdoRes.Close
   If optFrom Then diaSrvmo.Show Else FormUnload
   Set diaPsh01 = Nothing
   
End Sub




Public Function GetRuns() As Byte
   Dim RdoRns As rdoResultset
   On Error GoTo DiaErr1
   cmbRun.Clear
   sPartRef = Compress(cmbprt)
   RdoQry(0) = sPartRef
   bSqlRows = GetQuerySet(RdoRns, RdoQry)
   If bSqlRows Then
      With RdoRns
         If optFrom Then
            cmbRun = diaSrvmo.cmbRun
         Else
            cmbRun = Format(!RUNNO, "####0")
         End If
         lblDsc = "" & Trim(!PADESC)
         lblTyp = Format(!PALEVEL, "#")
         Do Until .EOF
            cmbRun.AddItem Format(0 + !RUNNO, "####0")
            .MoveNext
         Loop
         .Cancel
      End With
      GetRuns = True
      GetThisRun
   Else
      sPartRef = ""
      GetRuns = False
   End If
   Set RdoRns = Nothing
   Exit Function
   
   DiaErr1:
   sProcName = "getruns"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optBud_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbprt.SetFocus
      Exit Sub
   Else
      PrintReport
   End If
   
End Sub

Private Sub optDoc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFrom_Click()
   'dummy to check if from Revise mo
   
End Sub



Private Sub optLst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Not bGoodPart Then
      MsgBox "Couldn't Find Part Number, Run.", vbExclamation, Caption
      On Error Resume Next
      cmbprt.SetFocus
      Exit Sub
   Else
      On Error Resume Next
      JetDb.Execute "DELETE * FROM PlsTable"
      JetDb.Execute "DELETE * FROM CvrTable"
      'Doc and Pick List only for printed reports
      If optLst.Value = vbChecked Then
         If lblSta = "SC" Then
            sMsg = "Do You Want To Print The MO Pick " & vbCr _
                   & "List And Move The Run Status To Pl?"
            bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
            If bResponse = vbYes Then
               'Build the pick list and change status
               b = 1
               MouseCursor 13
               BuildPartsList
            Else
               CancelTrans
            End If
         Else
            MouseCursor 13
            b = 1
            BuildPickList
         End If
      End If
      If optDoc = vbChecked Then
         MouseCursor 13
         b = 1
         BuildDocumentList
      End If
      If b = 1 Then PrintCover
      PrintReport
   End If
   
End Sub



Public Sub GetThisRun()
   Dim RdoRun As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNSTATUS,RUNPKSTART FROM RunsTable WHERE " _
          & "RUNREF='" & Compress(cmbprt) & "' AND " _
          & "RUNNO=" & cmbRun & " "
   bSqlRows = GetDataSet(RdoRun, ES_FORWARD)
   If bSqlRows Then
      With RdoRun
         lblSta = "" & Trim(!RUNSTATUS)
         If Not IsNull(!RUNPKSTART) Then
            sRunPkstart = Format(!RUNPKSTART, "mm/dd/yy")
         Else
            sRunPkstart = Format(Now, "mm/dd/yy")
         End If
         .Cancel
      End With
   End If
   Set RdoRun = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "getthisrun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CreateCovrTable()
   'Drop and create a Jet table to
   'run beside SQL Server so Crystal can handle the report
   '(Crystal isn't smart enough to handle the joins)
   Dim NewTb As TableDef
   Dim NewFld As Field
   
   On Error Resume Next
   JetDb.Execute "DROP TABLE CvrTable"
   'Fields. Note that we allow empties
   Set NewTb = JetDb.CreateTableDef("CvrTable")
   With NewTb
      'Documents
      '1
      .Fields.Append .CreateField("DLSPart1", dbText, 30)
      .Fields(0).AllowZeroLength = True
      .Fields.Append .CreateField("DLSRev1", dbText, 6)
      .Fields(1).AllowZeroLength = True
      .Fields.Append .CreateField("DLSType1", dbInteger)
      .Fields.Append .CreateField("DLSDocRef1", dbText, 30)
      .Fields(3).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocRev1", dbText, 6)
      .Fields(4).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocSheet1", dbText, 6)
      .Fields(5).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocClass1", dbText, 16)
      .Fields(6).AllowZeroLength = True
      '2
      .Fields.Append .CreateField("DLSPart2", dbText, 30)
      .Fields(7).AllowZeroLength = True
      .Fields.Append .CreateField("DLSRev2", dbText, 6)
      .Fields(8).AllowZeroLength = True
      .Fields.Append .CreateField("DLSType2", dbInteger)
      .Fields.Append .CreateField("DLSDocRef2", dbText, 30)
      .Fields(10).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocRev2", dbText, 6)
      .Fields(11).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocSheet2", dbText, 6)
      .Fields(12).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocClass2", dbText, 16)
      .Fields(13).AllowZeroLength = True
      '3
      .Fields.Append .CreateField("DLSPart3", dbText, 30)
      .Fields(14).AllowZeroLength = True
      .Fields.Append .CreateField("DLSRev3", dbText, 6)
      .Fields(15).AllowZeroLength = True
      .Fields.Append .CreateField("DLSType3", dbInteger)
      .Fields.Append .CreateField("DLSDocRef3", dbText, 30)
      .Fields(17).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocRev3", dbText, 6)
      .Fields(18).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocSheet3", dbText, 6)
      .Fields(19).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocClass3", dbText, 16)
      .Fields(20).AllowZeroLength = True
      '4
      .Fields.Append .CreateField("DLSPart4", dbText, 30)
      .Fields(21).AllowZeroLength = True
      .Fields.Append .CreateField("DLSRev4", dbText, 6)
      .Fields(22).AllowZeroLength = True
      .Fields.Append .CreateField("DLSType4", dbInteger)
      .Fields.Append .CreateField("DLSDocRef4", dbText, 30)
      .Fields(24).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocRev4", dbText, 6)
      .Fields(25).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocSheet4", dbText, 6)
      .Fields(26).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocClass4", dbText, 16)
      .Fields(27).AllowZeroLength = True
      '5
      .Fields.Append .CreateField("DLSPart5", dbText, 30)
      .Fields(28).AllowZeroLength = True
      .Fields.Append .CreateField("DLSRev5", dbText, 6)
      .Fields(29).AllowZeroLength = True
      .Fields.Append .CreateField("DLSType5", dbInteger)
      .Fields.Append .CreateField("DLSDocRef5", dbText, 30)
      .Fields(31).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocRev5", dbText, 6)
      .Fields(32).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocSheet5", dbText, 6)
      .Fields(33).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocClass5", dbText, 16)
      .Fields(34).AllowZeroLength = True
      
      'added
      .Fields.Append .CreateField("DLSDocDesc1", dbText, 60)
      .Fields(35).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocLoc1", dbText, 4)
      .Fields(36).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocDesc2", dbText, 60)
      .Fields(37).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocLoc2", dbText, 4)
      .Fields(38).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocDesc3", dbText, 60)
      .Fields(39).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocLoc3", dbText, 4)
      .Fields(40).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocDesc4", dbText, 60)
      .Fields(41).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocLoc4", dbText, 4)
      .Fields(42).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocDesc5", dbText, 60)
      .Fields(43).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocLoc5", dbText, 4)
      .Fields(44).AllowZeroLength = True
      
      'More
      .Fields.Append .CreateField("DLSDocEco1", dbText, 2)
      .Fields(45).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocAdcn1", dbText, 20)
      .Fields(46).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocEco2", dbText, 2)
      .Fields(47).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocAdcn2", dbText, 20)
      .Fields(48).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocEco3", dbText, 2)
      .Fields(49).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocAdcn3", dbText, 20)
      .Fields(50).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocEco4", dbText, 2)
      .Fields(51).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocAdcn4", dbText, 20)
      .Fields(52).AllowZeroLength = True
      
      .Fields.Append .CreateField("DLSDocEco5", dbText, 2)
      .Fields(53).AllowZeroLength = True
      .Fields.Append .CreateField("DLSDocAdcn5", dbText, 20)
      .Fields(54).AllowZeroLength = True
      
   End With
   'add the table to Jet. No indexes
   JetDb.TableDefs.Append NewTb
   
End Sub

Public Sub CreatePlsTable()
   'Drop and create a Jet table to
   'run beside SQL Server so Crystal can handle the report
   '(Crystal isn't smart enough to handle the joins)
   Dim NewTb As TableDef
   Dim NewFld As Field
   
   On Error Resume Next
   JetDb.Execute "DROP TABLE PlsTable"
   'Fields. Note that we allow empties
   Set NewTb = JetDb.CreateTableDef("PlsTable")
   With NewTb
      'Documents
      '1
      .Fields.Append .CreateField("PLSPart1", dbText, 30)
      .Fields(0).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc1", dbText, 30)
      .Fields(1).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate1", dbDate)
      .Fields.Append .CreateField("PLSAQty1", dbCurrency)
      .Fields(3).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom1", dbText, 2)
      .Fields(4).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc1", dbText, 4)
      .Fields(5).AllowZeroLength = True
      
      '2
      .Fields.Append .CreateField("PLSPart2", dbText, 30)
      .Fields(6).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc2", dbText, 30)
      .Fields(7).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate2", dbDate)
      .Fields.Append .CreateField("PLSAQty2", dbCurrency)
      .Fields(9).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom2", dbText, 2)
      .Fields(10).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc2", dbText, 4)
      .Fields(11).AllowZeroLength = True
      
      '3
      .Fields.Append .CreateField("PLSPart3", dbText, 30)
      .Fields(12).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc3", dbText, 30)
      .Fields(13).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate3", dbDate)
      .Fields.Append .CreateField("PLSAQty3", dbCurrency)
      .Fields(15).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom3", dbText, 2)
      .Fields(16).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc3", dbText, 4)
      .Fields(17).AllowZeroLength = True
      
      '4
      .Fields.Append .CreateField("PLSPart4", dbText, 30)
      .Fields(18).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc4", dbText, 30)
      .Fields(19).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate4", dbDate)
      .Fields.Append .CreateField("PLSAQty4", dbCurrency)
      .Fields(21).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom4", dbText, 2)
      .Fields(22).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc4", dbText, 4)
      .Fields(23).AllowZeroLength = True
      
      '5
      .Fields.Append .CreateField("PLSPart5", dbText, 30)
      .Fields(24).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc5", dbText, 30)
      .Fields(25).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate5", dbDate)
      .Fields.Append .CreateField("PLSAQty5", dbCurrency)
      .Fields(27).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom5", dbText, 2)
      .Fields(28).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc5", dbText, 4)
      .Fields(29).AllowZeroLength = True
      
      '6
      .Fields.Append .CreateField("PLSPart6", dbText, 30)
      .Fields(30).AllowZeroLength = True
      .Fields.Append .CreateField("PLSDesc6", dbText, 30)
      .Fields(31).AllowZeroLength = True
      .Fields.Append .CreateField("PLSADate6", dbDate)
      .Fields.Append .CreateField("PLSAQty6", dbCurrency)
      .Fields(33).DefaultValue = 0
      .Fields.Append .CreateField("PLSUom6", dbText, 2)
      .Fields(34).AllowZeroLength = True
      .Fields.Append .CreateField("PLSLoc6", dbText, 4)
      .Fields(35).AllowZeroLength = True
      
      'Added
      .Fields.Append .CreateField("PLSPQty1", dbCurrency)
      .Fields(36).DefaultValue = 0
      .Fields.Append .CreateField("PLSPQty2", dbCurrency)
      .Fields(37).DefaultValue = 0
      .Fields.Append .CreateField("PLSPQty3", dbCurrency)
      .Fields(38).DefaultValue = 0
      .Fields.Append .CreateField("PLSPQty4", dbCurrency)
      .Fields(39).DefaultValue = 0
      .Fields.Append .CreateField("PLSPQty5", dbCurrency)
      .Fields(40).DefaultValue = 0
      .Fields.Append .CreateField("PLSPQty6", dbCurrency)
      .Fields(41).DefaultValue = 0
   End With
   'add the table to Jet. No indexes
   JetDb.TableDefs.Append NewTb
   
End Sub

Public Sub BuildDocumentList()
   Dim RdoDoc As rdoResultset
   Dim RdoJet As rdoResultset
   Dim i As Integer
   Dim a As Integer
   
   On Error GoTo DiaErr1
   JetDb.Execute "DELETE * FROM CvrTable"
   Set DbDoc = JetDb.OpenRecordset("CvrTable", dbOpenDynaset)
   '   Well it ain't buying this join. SQL Server will but not the RDO
   '
   '    sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR," _
   '        & "DOECO,DOADCN,DOTYPE,DLSREF,DLSREV,DLSTYPE,DLSDOCREF," _
   '        & "DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS FROM DdocTable,DlstTable" _
   '        & "WHERE (DOREF=DLSDOCREF AND DOSHEET=DLSDOCSHEET AND DOREV=" _
   '        & "DLSREV AND DLSREF='65B845892')"
   '
   '   so
   sSql = "SELECT DLSREF,DLSREV,DLSTYPE,DLSDOCREF," _
          & "DLSDOCREV,DLSDOCSHEET,DLSDOCCLASS FROM DlstTable " _
          & "WHERE DLSREF='" & Compress(cmbprt) & "' ORDER BY DLSDOCREF"
   bSqlRows = GetDataSet(RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         DbDoc.AddNew
         On Error Resume Next
         Do Until .EOF
            i = i + 1
            If i > 5 Then Exit Do
            sSql = "SELECT DOREF,DONUM,DOREV,DOCLASS,DOSHEET,DODESCR," _
                   & "DOECO,DOADCN,DOTYPE,DOLOC FROM DdocTable WHERE " _
                   & "(DOREF='" & Trim(!DLSDOCREF) & "' AND " _
                   & "DOSHEET='" & Trim(!DLSDOCSHEET) & "' AND " _
                   & "DOREV ='" & Trim(!DLSDOCREV) & "')"
            bSqlRows = GetDataSet(RdoJet, ES_FORWARD)
            
            Select Case i
               Case 1
                  DbDoc!DLSDocRef1 = "" & Trim(RdoJet!DONUM)
                  DbDoc!DLSDocRev1 = "" & Trim(!DLSDOCREV)
                  DbDoc!DLSDocSheet1 = "" & Trim(!DLSDOCSHEET)
                  DbDoc!DLSDocClass1 = "" & Trim(!DLSDOCCLASS)
                  DbDoc!DLSDocDesc1 = "" & Trim(RdoJet!DODESCR)
                  DbDoc!DLSDocLoc1 = "" & Trim(RdoJet!DOLOC)
                  DbDoc!DLSDocEco1 = "" & Trim(RdoJet!DOECO)
                  DbDoc!DLSDocAdcn1 = "" & Trim(RdoJet!DOADCN)
               Case 2
                  DbDoc!DLSDocRef2 = "" & Trim(RdoJet!DONUM)
                  DbDoc!DLSDocRev2 = "" & Trim(!DLSDOCREV)
                  DbDoc!DLSDocSheet2 = "" & Trim(!DLSDOCSHEET)
                  DbDoc!DLSDocClass2 = "" & Trim(!DLSDOCCLASS)
                  DbDoc!DLSDocDesc2 = "" & Trim(RdoJet!DODESCR)
                  DbDoc!DLSDocLoc2 = "" & Trim(RdoJet!DOLOC)
                  DbDoc!DLSDocEco2 = "" & Trim(RdoJet!DOECO)
                  DbDoc!DLSDocAdcn2 = "" & Trim(RdoJet!DOADCN)
               Case 3
                  DbDoc!DLSDocRef3 = "" & Trim(RdoJet!DONUM)
                  DbDoc!DLSDocRev3 = "" & Trim(!DLSDOCREV)
                  DbDoc!DLSDocSheet3 = "" & Trim(!DLSDOCSHEET)
                  DbDoc!DLSDocClass3 = "" & Trim(!DLSDOCCLASS)
                  DbDoc!DLSDocDesc3 = "" & Trim(RdoJet!DODESCR)
                  DbDoc!DLSDocLoc3 = "" & Trim(RdoJet!DOLOC)
                  DbDoc!DLSDocEco3 = "" & Trim(RdoJet!DOECO)
                  DbDoc!DLSDocAdcn3 = "" & Trim(RdoJet!DOADCN)
               Case 4
                  DbDoc!DLSDocRef4 = "" & Trim(RdoJet!DONUM)
                  DbDoc!DLSDocRev4 = "" & Trim(!DLSDOCREV)
                  DbDoc!DLSDocSheet4 = "" & Trim(!DLSDOCSHEET)
                  DbDoc!DLSDocClass4 = "" & Trim(!DLSDOCCLASS)
                  DbDoc!DLSDocDesc4 = "" & Trim(RdoJet!DODESCR)
                  DbDoc!DLSDocLoc4 = "" & Trim(RdoJet!DOLOC)
                  DbDoc!DLSDocEco4 = "" & Trim(RdoJet!DOECO)
                  DbDoc!DLSDocAdcn4 = "" & Trim(RdoJet!DOADCN)
               Case 5
                  DbDoc!DLSDocRef5 = "" & Trim(RdoJet!DONUM)
                  DbDoc!DLSDocRev5 = "" & Trim(!DLSDOCREV)
                  DbDoc!DLSDocSheet5 = "" & Trim(!DLSDOCSHEET)
                  DbDoc!DLSDocClass5 = "" & Trim(!DLSDOCCLASS)
                  DbDoc!DLSDocDesc5 = "" & Trim(RdoJet!DODESCR)
                  DbDoc!DLSDocLoc5 = "" & Trim(RdoJet!DOLOC)
                  DbDoc!DLSDocEco5 = "" & Trim(RdoJet!DOECO)
                  DbDoc!DLSDocAdcn5 = "" & Trim(RdoJet!DOADCN)
            End Select
            .MoveNext
         Loop
         DbDoc.Update
         .Cancel
      End With
   Else
      DbDoc.AddNew
      DbDoc!DLSDocRef1 = "*** No Documents Recorded ***"
      DbDoc.Update
   End If
   On Error Resume Next
   DbDoc.Close
   Set RdoDoc = Nothing
   Set RdoJet = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "builddoclist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'SC...no Pick List yet

Public Sub BuildPartsList()
   Dim RdoLst As rdoResultset
   Dim b As Byte
   
   Dim cConversion As Currency
   Dim cQuantity As Currency
   
   sSql = "SELECT DISTINCT BMASSYPART FROM BmplTable " _
          & "WHERE BMASSYPART='" & Compress(cmbprt) & "' "
   bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
   If bSqlRows Then
      b = 1
      RdoLst.Cancel
   Else
      MouseCursor 0
      b = 0
      MsgBox "This Part Does Not Have A Parts List.", vbExclamation, Caption
   End If
   If b = 1 Then
      Set DbPls = JetDb.OpenRecordset("PlsTable", dbOpenDynaset)
      
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
             & "BMASSYPART,BMPARTREF,BMQTYREQD,BMSETUP,BMADDER," _
             & "BMCONVERSION FROM PartTable,BmplTable WHERE " _
             & "PARTREF=BMPARTREF AND BMASSYPART='" & Compress(cmbprt) & "'"
      bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
      If bSqlRows Then
         b = 0
         With RdoLst
            DbPls.AddNew
            On Error Resume Next
            Err = 0
            RdoCon.BeginTrans
            Do Until .EOF
               cQuantity = Format(!BMQTYREQD + !BMADDER + !BMSETUP, "#####0.000")
               cConversion = Format(!BMCONVERSION, "#####0.0000")
               If cConversion = 0 Then cConversion = 1
               cQuantity = cQuantity / cConversion
               b = b + 1
               If b < 7 Then
                  Select Case b
                     Case 1
                        DbPls!PLSPart1 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc1 = "" & Trim(!PADESC)
                        DbPls!PLSPqty1 = cQuantity
                        DbPls!PLSUom1 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc1 = "" & Trim(!PALOCATION)
                     Case 2
                        DbPls!PLSPart2 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc2 = "" & Trim(!PADESC)
                        DbPls!PLSPqty2 = cQuantity
                        DbPls!PLSUom2 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc2 = "" & Trim(!PALOCATION)
                     Case 3
                        DbPls!PLSPart3 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc3 = "" & Trim(!PADESC)
                        DbPls!PLSPqty3 = cQuantity
                        DbPls!PLSUom3 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc3 = "" & Trim(!PALOCATION)
                     Case 4
                        DbPls!PLSPart4 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc4 = "" & Trim(!PADESC)
                        DbPls!PLSPqty4 = cQuantity
                        DbPls!PLSUom4 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc4 = "" & Trim(!PALOCATION)
                     Case 5
                        DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc5 = "" & Trim(!PADESC)
                        DbPls!PLSPqty5 = cQuantity
                        DbPls!PLSUom5 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                     Case 6
                        DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                        DbPls!PLSDesc6 = "" & Trim(!PADESC)
                        DbPls!PLSPqty6 = cQuantity
                        DbPls!PLSUom6 = "" & Trim(!PAUNITS)
                        DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
                  End Select
               End If
               If sRunPkstart = "" Then sRunPkstart = Format(Now, "mm/dd/yy")
               sSql = "INSERT INTO MopkTable (PKPARTREF," _
                      & "PKMOPART,PKMORUN,PKTYPE,PKPDATE," _
                      & "PKPQTY,PKBOMQTY) VALUES('" _
                      & Trim(!PARTREF) & "','" & Compress(cmbprt) & "'," _
                      & cmbRun & ",9,'" & sRunPkstart & "'," & cQuantity _
                      & "," & cQuantity & ") "
               RdoCon.Execute sSql, rdExecDirect
               .MoveNext
            Loop
            sSql = "UPDATE RunsTable SET RUNSTATUS='PL'," _
                   & "RUNPLDATE='" & Format(Now, "mm/dd/yy") & "' " _
                   & "WHERE RUNREF='" & Compress(cmbprt) & "' " _
                   & "AND RUNNO=" & cmbRun & " "
            RdoCon.Execute sSql, rdExecDirect
            If Err = 0 Then
               RdoCon.CommitTrans
               lblSta = "PL"
            Else
               RdoCon.RollbackTrans
            End If
            DbPls.Update
            .Cancel
         End With
      End If
   Else
      DbPls.AddNew
      DbPls!PLSPart1 = "*** No Pick List Recorded ***"
      DbPls.Update
   End If
   On Error Resume Next
   DbPls.Close
   Set RdoLst = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "buildpartslist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub PrintCover()
   Dim sWindows As String
   MouseCursor 13
   
   On Error GoTo DiaErr1
   sWindows = GetWindowsDir()
   SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.ReportFileName = sReportPath & "prdshcvr.rpt"
   MdiSect.crw.DataFiles(0) = sWindows & "\temp\esiprod.mdb"
   MdiSect.crw.Formulas(1) = "Includes='" & cmbprt & " Run " & cmbRun & "'"
   MdiSect.crw.Formulas(2) = "Includes2='" & lblDsc & "'"
   If optDoc.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(0) = "GROUPHDR.0.0;F;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPHDR.0.1;F;;;"
   Else
      MdiSect.crw.SectionFormat(0) = "GROUPHDR.0.0;T;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPHDR.0.T;F;;;"
   End If
   If optLst.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.0.1;F;;;"
      MdiSect.crw.SectionFormat(3) = "GROUPFTR.0.2;F;;;"
   Else
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.0.1;T;;;"
      MdiSect.crw.SectionFormat(3) = "GROUPFTR.0.2;T;;;"
   End If
   SetCrystalAction Me
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02
   Psh02:
   DoModuleErrors Me
   
End Sub

'Pick list is active

Public Sub BuildPickList()
   Dim RdoLst As rdoResultset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
          & "PKPARTREF,PKMOPART,PKPQTY,PKADATE,PKAQTY FROM PartTable," _
          & "MopkTable WHERE (PARTREF=PKPARTREF AND PKMOPART='" _
          & Compress(cmbprt) & "' AND PKMORUN=" & cmbRun & ")"
   bSqlRows = GetDataSet(RdoLst, ES_FORWARD)
   If bSqlRows Then
      JetDb.Execute "DELETE * FROM PlsTable"
      Set DbPls = JetDb.OpenRecordset("PlsTable", dbOpenDynaset)
      With RdoLst
         DbPls.AddNew
         Do Until .EOF
            b = b + 1
            If b < 7 Then
               Select Case b
                  Case 1
                     DbPls!PLSPart1 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc1 = "" & Trim(!PADESC)
                     DbPls!PLSPqty1 = Format(!PKPQTY, "####0.000")
                     DbPls!PLSAqty1 = Format(!PKAQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate1 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSUom1 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc1 = "" & Trim(!PALOCATION)
                  Case 2
                     DbPls!PLSPart2 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc2 = "" & Trim(!PADESC)
                     DbPls!PLSPqty2 = Format(!PKPQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate2 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty2 = Format(!PKAQTY, "####0.000")
                     DbPls!PLSUom2 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc2 = "" & Trim(!PALOCATION)
                  Case 3
                     DbPls!PLSPart3 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc3 = "" & Trim(!PADESC)
                     DbPls!PLSPqty3 = Format(!PKPQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate3 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty3 = Format(!PKAQTY, "####0.000")
                     DbPls!PLSUom3 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc3 = "" & Trim(!PALOCATION)
                  Case 4
                     DbPls!PLSPart4 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc4 = "" & Trim(!PADESC)
                     DbPls!PLSPqty4 = Format(!PKPQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate4 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty4 = Format(!PKAQTY, "####0.000")
                     DbPls!PLSUom4 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc4 = "" & Trim(!PALOCATION)
                  Case 5
                     DbPls!PLSPart5 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc5 = "" & Trim(!PADESC)
                     DbPls!PLSPqty5 = Format(!PKPQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate5 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty5 = Format(!PKAQTY, "####0.000")
                     DbPls!PLSUom5 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc5 = "" & Trim(!PALOCATION)
                  Case 6
                     DbPls!PLSPart6 = "" & Trim(!PARTNUM)
                     DbPls!PLSDesc6 = "" & Trim(!PADESC)
                     DbPls!PLSPqty6 = Format(!PKPQTY, "####0.000")
                     If Not IsNull(!PKADATE) Then
                        DbPls!PLSADate6 = Format(!PKADATE, "mm/dd/yy")
                     End If
                     DbPls!PLSAqty6 = Format(!PKAQTY, "####0.000")
                     DbPls!PLSUom6 = "" & Trim(!PAUNITS)
                     DbPls!PLSLoc6 = "" & Trim(!PALOCATION)
               End Select
            End If
            .MoveNext
         Loop
         DbPls.Update
         .Cancel
      End With
   Else
      DbPls.AddNew
      DbPls!PLSPart1 = "*** No Pick List Recorded ***"
      DbPls.Update
   End If
   On Error Resume Next
   DbPls.Close
   Set RdoLst = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "buildpicklist"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub
