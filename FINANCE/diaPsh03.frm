VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPsh03
   BorderStyle = 3 'Fixed Dialog
   Caption = "Manufacturing Orders By Date (Report)"
   ClientHeight = 3570
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 7260
   ControlBox = 0 'False
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3570
   ScaleWidth = 7260
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton optPrn
      Height = 330
      Left = 6680
      Picture = "diaPsh03.frx":0000
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
      Picture = "diaPsh03.frx":018A
      Style = 1 'Graphical
      TabIndex = 25
      ToolTipText = "Display The Report"
      Top = 480
      UseMaskColor = -1 'True
      Width = 495
   End
   Begin VB.ComboBox txtBeg
      Height = 315
      Left = 2400
      TabIndex = 7
      Tag = "4"
      Top = 1560
      Width = 1095
   End
   Begin VB.ComboBox txtEnd
      Height = 315
      Left = 4080
      TabIndex = 8
      Tag = "4"
      Top = 1560
      Width = 1095
   End
   Begin VB.CheckBox optOps
      Caption = "____"
      Enabled = 0 'False
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 13
      Top = 3120
      Width = 735
   End
   Begin VB.CheckBox optQty
      Caption = "____"
      Enabled = 0 'False
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 12
      Top = 2860
      Width = 735
   End
   Begin VB.CheckBox optCmt
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2400
      TabIndex = 11
      Top = 2660
      Width = 735
   End
   Begin VB.CheckBox optExt
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 10
      Top = 2400
      Width = 735
   End
   Begin VB.CheckBox optDsc
      Caption = "____"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2400
      TabIndex = 9
      Top = 2160
      Width = 735
   End
   Begin VB.CheckBox optSta
      Caption = "CA"
      Height = 255
      Index = 6
      Left = 6000
      TabIndex = 6
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "CL"
      Height = 255
      Index = 5
      Left = 5400
      TabIndex = 5
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "CO"
      Height = 255
      Index = 4
      Left = 4800
      TabIndex = 4
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PC"
      Height = 255
      Index = 3
      Left = 4200
      TabIndex = 3
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PP"
      Height = 255
      Index = 2
      Left = 3600
      TabIndex = 2
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "PL"
      Height = 255
      Index = 1
      Left = 3000
      TabIndex = 1
      Top = 1080
      Width = 615
   End
   Begin VB.CheckBox optSta
      Caption = "SC"
      Height = 255
      Index = 0
      Left = 2400
      TabIndex = 0
      Top = 1080
      Width = 615
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6120
      TabIndex = 15
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 6120
      TabIndex = 14
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
      FormDesignHeight = 3570
      FormDesignWidth = 7260
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 27
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
      PictureUp = "diaPsh03.frx":0308
      PictureDn = "diaPsh03.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 28
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
      PictureUp = "diaPsh03.frx":0594
      PictureDn = "diaPsh03.frx":06DA
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 29
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Current Operation Information"
      Height = 285
      Index = 8
      Left = 240
      TabIndex = 24
      Top = 3120
      Width = 2265
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Partial Completion Quantities"
      Height = 285
      Index = 7
      Left = 240
      TabIndex = 23
      Top = 2880
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "MO Comments"
      Height = 285
      Index = 6
      Left = 240
      TabIndex = 22
      Top = 2640
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions"
      Height = 285
      Index = 5
      Left = 240
      TabIndex = 21
      Top = 2400
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Descriptions"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 19
      Top = 2160
      Width = 1905
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 16
      Top = 1920
      Width = 705
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Thru"
      Height = 285
      Index = 2
      Left = 3600
      TabIndex = 18
      Top = 1560
      Width = 585
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Manufacturing Orders From"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 17
      Top = 1560
      Width = 2145
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Run Status:"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 20
      Top = 1080
      Width = 2025
   End
End
Attribute VB_Name = "diaPsh03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001, ES/2002) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'************************************************************************************
'
'
'
' Created: 08/22/01 (nth)
' Revions:
'
'
'************************************************************************************
Dim bOnLoad As Byte

Private txtKeyPress(2) As New EsiKeyBd
Private txtGotFocus(2) As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'************************************************************************************

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



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   txtBeg = Format(Now, "mm/01/yy")
   txtEnd = Format(Now, "mm/dd/yy")
   GetOptions
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
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
   
   MouseCursor 13
   CheckOptions
   On Error Resume Next
   sBegDate = Format(txtBeg, "yyyy,mm,dd")
   sEndDate = Format(txtEnd, "yyyy,mm,dd")
   On Error GoTo DiaErr1
   
   SetMdiReportsize MdiSect
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='" & txtBeg & " Through " & txtEnd & "...'"
   MdiSect.crw.ReportFileName = sReportPath & "prdsh03.rpt"
   sSql = ""
   sSql = "{RunsTable.RUNSCHED} in Date(" & sBegDate & ") to Date(" & sEndDate & ") "
   If optSta(0).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'SC' "
   If optSta(1).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PL' "
   If optSta(2).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PP' "
   If optSta(3).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'PC' "
   If optSta(4).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CO' "
   If optSta(5).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CL' "
   If optSta(6).Value = vbUnchecked Then sSql = sSql & "AND {RunsTable.RUNSTATUS}<>'CA' "
   If OptCmt.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPFTR.0.1;F;;;"
   Else
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPFTR.0.1;T;;;"
   End If
   
   If optExt.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.1.0;F;;;"
      MdiSect.crw.SectionFormat(3) = "GROUPFTR.1.1;F;;;"
   Else
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.1.0;T;;;"
      MdiSect.crw.SectionFormat(3) = "GROUPFTR.1.1;T;;;"
   End If
   
   If optDsc.Value = vbUnchecked Then
      MdiSect.crw.SectionFormat(4) = "GROUPFTR.2.0;F;;;"
   Else
      MdiSect.crw.SectionFormat(4) = "GROUPFTR.2.0;T;;;"
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
   
   Set txtKeyPress(0).esCmbKeyDate = txtBeg
   Set txtKeyPress(1).esCmbKeyDate = txtEnd
   
End Sub

Public Sub SaveOptions()
   Dim i As Integer
   Dim sOptions As String
   
   'Save by Menu Option
   For i = 0 To 5
      sOptions = sOptions & Trim(Str(optSta(i).Value))
   Next
   sOptions = sOptions & Trim(Str(optSta(i).Value))
   sOptions = sOptions & Trim(Str(optDsc.Value)) _
              & Trim(Str(optExt.Value)) & Trim(Str(OptCmt.Value))
   SaveSetting "Esi2000", "EsiProd", "sh03", Trim(sOptions)
   
End Sub

Public Sub GetOptions()
   Dim i As Integer
   Dim sOptions As String
   
   On Error Resume Next
   'Get By Menu Option
   sOptions = GetSetting("Esi2000", "EsiProd", "sh03", sOptions)
   If Len(sOptions) > 0 Then
      For i = 1 To 6
         optSta(i - 1) = Mid$(sOptions, i, 1)
      Next
      optSta(i - 1) = Mid$(sOptions, i, 1)
      optDsc.Value = Val(Mid(sOptions, i + 1, 1))
      optExt.Value = Val(Mid(sOptions, i + 2, 1))
      OptCmt.Value = Val(Mid(sOptions, i + 3, 1))
   End If
   
End Sub



Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optExt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
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
