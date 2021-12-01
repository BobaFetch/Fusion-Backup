VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPsh04
   BorderStyle = 3 'Fixed Dialog
   Caption = "Manufacturing Orders By Part"
   ClientHeight = 3555
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 7260
   ControlBox = 0 'False
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3555
   ScaleWidth = 7260
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton optPrn
      Height = 330
      Left = 6680
      Picture = "diaPsh04.frx":0000
      Style = 1 'Graphical
      TabIndex = 26
      ToolTipText = "Print The Report"
      Top = 480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.CommandButton optDis
      Height = 330
      Left = 6120
      Picture = "diaPsh04.frx":018A
      Style = 1 'Graphical
      TabIndex = 25
      ToolTipText = "Display The Report"
      Top = 480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.ComboBox txtEnd
      Height = 315
      Left = 4080
      TabIndex = 9
      Tag = "4"
      Top = 1920
      Width = 1095
   End
   Begin VB.ComboBox txtBeg
      Height = 315
      Left = 2400
      TabIndex = 8
      Tag = "4"
      Top = 1920
      Width = 1095
   End
   Begin VB.ComboBox cmbPrt
      Height = 315
      Left = 2400
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Contains Parts With Runs"
      Top = 960
      Width = 3545
   End
   Begin VB.CheckBox optOps
      Caption = "____"
      Enabled = 0 'False
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 12
      Top = 3000
      Width = 735
   End
   Begin VB.CheckBox optQty
      Caption = "____"
      Enabled = 0 'False
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 11
      Top = 2745
      Width = 735
   End
   Begin VB.CheckBox optCmt
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2400
      TabIndex = 10
      Top = 2535
      Width = 735
   End
   Begin VB.CheckBox optSta
      Caption = "CA"
      Height = 255
      Index = 6
      Left = 6000
      TabIndex = 7
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "CL"
      Height = 255
      Index = 5
      Left = 5400
      TabIndex = 6
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "CO"
      Height = 255
      Index = 4
      Left = 4800
      TabIndex = 5
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PC"
      Height = 255
      Index = 3
      Left = 4200
      TabIndex = 4
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PP"
      Height = 255
      Index = 2
      Left = 3600
      TabIndex = 3
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PL"
      Height = 255
      Index = 1
      Left = 3000
      TabIndex = 2
      Top = 1600
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "SC"
      Height = 255
      Index = 0
      Left = 2400
      TabIndex = 1
      Top = 1600
      Width = 615
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6120
      TabIndex = 14
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 6120
      TabIndex = 13
      Top = 360
      Width = 1095
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 6840
      Top = 3000
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3555
      FormDesignWidth = 7260
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 24
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
      PictureUp = "diaPsh04.frx":0308
      PictureDn = "diaPsh04.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 27
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
      PictureUp = "diaPsh04.frx":0594
      PictureDn = "diaPsh04.frx":06DA
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 28
      Top = 0
      Width = 2760
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 2400
      TabIndex = 23
      Top = 1320
      Width = 3375
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Thru"
      Height = 285
      Index = 5
      Left = 3600
      TabIndex = 22
      Top = 1920
      Width = 585
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 21
      Top = 960
      Width = 2025
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Current Operation Information"
      Height = 285
      Index = 8
      Left = 240
      TabIndex = 20
      Top = 3000
      Width = 2265
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Partial Completion Quantities"
      Height = 285
      Index = 7
      Left = 240
      TabIndex = 19
      Top = 2760
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "MO Comments"
      Height = 285
      Index = 6
      Left = 240
      TabIndex = 18
      Top = 2520
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 15
      Top = 2280
      Width = 705
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Manufacturing Orders From"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 16
      Top = 1920
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Run Status:"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 17
      Top = 1600
      Width = 2025
   End
End
Attribute VB_Name = "diaPsh04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress(3) As New EsiKeyBd
Private txtGotFocus(3) As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd




Private Sub cmbPrt_Click()
   FindPart Me, cmbprt
   
End Sub

Private Sub cmbprt_LostFocus()
   cmbprt = CheckLen(cmbprt, 30)
   FindPart Me, cmbprt
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub FillCombo()
   Dim rdocmb As rdoResultset
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF FROM " _
          & "PartTable,RunsTable WHERE PARTREF=RUNREF"
   bSqlRows = GetDataSet(rdocmb)
   If bSqlRows Then
      With rdocmb
         Do Until .EOF
            cmbprt.AddItem "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If Trim(cmbprt) = "" Then
      If Len(Trim(Cur.CurrentPart)) Then
         cmbprt = Cur.CurrentPart
      Else
         If cmbprt.ListCount > 0 Then
            cmbprt = cmbprt.List(0)
         End If
      End If
   End If
   If Len(Trim(cmbprt)) Then FindPart Me, cmbprt
   Set rdocmb = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   bOnLoad = True
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   If Len(Trim(cmbprt)) Then Cur.CurrentPart = cmbprt
   SaveCurrentSelections
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaPsh03 = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sPart As String
   
   MouseCursor 13
   CheckOptions
   On Error Resume Next
   sBegDate = Format(txtBeg, "yyyy,mm,dd")
   sEndDate = Format(txtEnd, "yyyy,mm,dd")
   On Error GoTo DiaErr1
   
   sPart = Compress(cmbprt)
   
   SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='" & txtBeg & " Through " & txtEnd & "...'"
   MdiSect.crw.ReportFileName = sReportPath & "prdsh04.rpt"
   sSql = "{RunsTable.RUNREF}='" & sPart & "' AND " _
          & "{RunsTable.RUNSCHED} in Date(" & sBegDate & ") to Date(" & sEndDate & ") "
   If optSta(0).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PP' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PC' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CO' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CL' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CA' "
   If optCmt.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
   Else
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
   End If
   MdiSect.crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub














Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
   Set txtGotFocus(0).esCmbGotfocus = txtBeg
   Set txtGotFocus(1).esCmbGotfocus = txtEnd
   Set txtGotFocus(2).esCmbGotfocus = cmbprt
   
   Set txtKeyPress(0).esCmbKeyDate = txtBeg
   Set txtKeyPress(1).esCmbKeyDate = txtEnd
   Set txtKeyPress(2).esCmbKeyCase = cmbprt
   
End Sub

Public Sub SaveOptions()
   Dim i As Integer
   Dim sOptions As String
   Dim sPart As String * 30
   sPart = cmbprt
   'Save by Menu Option
   For i = 0 To 5
      sOptions = sOptions & Trim(Str(optSta(i).Value))
   Next
   sOptions = sOptions & Trim(Str(optSta(i).Value))
   sOptions = sOptions & Trim(Str(optCmt.Value))
   sOptions = sOptions & sPart
   SaveSetting "Esi2000", "EsiProd", "sh04", Trim(sOptions)
   
End Sub

Public Sub GetOptions()
   Dim i As Integer
   Dim sOptions As String
   'Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "sh04", sOptions)
   If Len(sOptions) > 0 Then
      For i = 1 To 6
         optSta(i - 1) = Mid$(sOptions, i, 1)
      Next
      optSta(i - 1) = Mid$(sOptions, i, 1)
      optCmt.Value = Val(Mid(sOptions, i + 1, 1))
   End If
   
End Sub



Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub





Private Sub optOps_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   PrintReport
   
End Sub

Private Sub optQty_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSta_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
   
End Sub

Private Sub txtbeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtEnd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
   
End Sub



Public Sub CheckOptions()
   Dim bByte As Byte
   Dim i As Integer
   
   For i = 0 To 5
      If optSta(i).Value = vbChecked Then
         bByte = True
         Exit For
      End If
   Next
   If optSta(i).Value = vbChecked Then bByte = True
   
   If Not bByte Then
      For i = 0 To 6
         optSta(i).Value = vbChecked
      Next
   End If
   
End Sub
