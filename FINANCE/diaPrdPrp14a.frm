VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaPrdPrp14a
   BorderStyle = 1 'Fixed Single
   Caption = "Vendor Performance Analysis"
   ClientHeight = 5535
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6525
   ControlBox = 0 'False
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 5535
   ScaleWidth = 6525
   Begin VB.ComboBox cmbTyp
      Enabled = 0 'False
      Height = 315
      Left = 2040
      TabIndex = 33
      Tag = "4"
      Top = 3120
      Width = 735
   End
   Begin VB.TextBox txtLte
      Height = 315
      Left = 2040
      TabIndex = 3
      Tag = "1"
      Top = 2640
      Width = 735
   End
   Begin VB.CheckBox optSum
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 9
      Top = 5160
      Width = 735
   End
   Begin VB.CheckBox optNYR
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 4
      Top = 3840
      Width = 735
   End
   Begin VB.CheckBox optINV
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 8
      Top = 4800
      Width = 735
   End
   Begin VB.CheckBox optREC
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 7
      Top = 4560
      Width = 735
   End
   Begin VB.ComboBox txteDte
      Height = 315
      Left = 2040
      TabIndex = 2
      Tag = "4"
      Top = 2160
      Width = 1095
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 5400
      TabIndex = 13
      Top = 360
      Width = 1095
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaPrdPrp14a.frx":0000
         Style = 1 'Graphical
         TabIndex = 11
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaPrdPrp14a.frx":018A
         Style = 1 'Graphical
         TabIndex = 10
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 5400
      TabIndex = 12
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.ComboBox txtsDte
      Height = 315
      Left = 2040
      TabIndex = 1
      Tag = "4"
      Top = 1800
      Width = 1095
   End
   Begin VB.CheckBox optODI
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 255
      Left = 2040
      TabIndex = 6
      Top = 4305
      Width = 735
   End
   Begin VB.ComboBox cmbVnd
      Height = 315
      Left = 2040
      TabIndex = 0
      Tag = "3"
      ToolTipText = "Contains Vendors With Invoices"
      Top = 840
      Width = 1555
   End
   Begin VB.CheckBox optODD
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 5
      Top = 4080
      Width = 735
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 14
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
      PictureUp = "diaPrdPrp14a.frx":0308
      PictureDn = "diaPrdPrp14a.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 5760
      Top = 1800
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5535
      FormDesignWidth = 6525
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 15
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
      PictureUp = "diaPrdPrp14a.frx":0594
      PictureDn = "diaPrdPrp14a.frx":06DA
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor Types"
      Height = 285
      Index = 15
      Left = 240
      TabIndex = 34
      Top = 3180
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Allowable Variance)"
      Height = 285
      Index = 14
      Left = 3120
      TabIndex = 32
      Top = 2685
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Late Window"
      Height = 285
      Index = 13
      Left = 240
      TabIndex = 31
      Top = 2700
      Width = 1455
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Summary Only?"
      Height = 285
      Index = 12
      Left = 240
      TabIndex = 30
      Top = 5160
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Not Yet Received"
      Height = 285
      Index = 11
      Left = 360
      TabIndex = 29
      Top = 3840
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Ending"
      Height = 285
      Index = 10
      Left = 600
      TabIndex = 28
      Top = 2160
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Starting"
      Height = 285
      Index = 9
      Left = 600
      TabIndex = 27
      Top = 1860
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Invoiced"
      Height = 285
      Index = 8
      Left = 360
      TabIndex = 26
      Top = 4800
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Received"
      Height = 285
      Index = 7
      Left = 360
      TabIndex = 25
      Top = 4560
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "On Dock (Inspected)"
      Height = 285
      Index = 6
      Left = 360
      TabIndex = 24
      Top = 4320
      Width = 1575
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "PO items Due:"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 23
      Top = 1560
      Width = 1545
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "On Dock (Delivered)"
      Height = 285
      Index = 5
      Left = 360
      TabIndex = 22
      Top = 4080
      Width = 1575
   End
   Begin VB.Label lblNme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 2040
      TabIndex = 21
      Top = 1200
      Width = 3000
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Vendor Name"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 20
      Top = 1200
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Nickname"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 19
      Top = 840
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 0
      Left = 3840
      TabIndex = 18
      Top = 840
      Width = 2025
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 17
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include Items:"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 16
      Top = 3600
      Width = 1455
   End
End
Attribute VB_Name = "diaPrdPrp14a"
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

'*********************************************************************************
' diaPrdpr14a - Vendor Performance analysis
'
' Created: 12/01/03 (JcW!)
'
'
'*********************************************************************************
Dim bOnLoad As Byte
Dim bGoodVendor As Boolean
'*********************************************************************************

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbVnd_Click()
   bGoodVendor = FindVendor(Me)
   cmbTyp = ""
   cmbTyp.Enabled = False
End Sub

Private Sub cmbVnd_LostFocus()
   On Error Resume Next
   
   cmbVnd = CheckLen(cmbVnd, 10)
   bGoodVendor = FindVendor(Me)
   If Not bGoodVendor Or Trim(cmbVnd) = "" Then
      cmbVnd = "ALL"
      lblNme = "***Multiple Vendors Selected***"
      cmbTyp.Enabled = True
   Else
      cmbTyp = ""
      cmbTyp.Enabled = False
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Reports"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      GetOptions
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   bOnLoad = True
   txtsDte = Format(Now, "mm/01/yy")
   txteDte = Format(Now, "mm/dd/yy")
End Sub

Private Sub FillCombo()
   Dim RdoVed As rdoResultset
   Dim RdoTyp As rdoResultset
   
   On Error GoTo DiaErr1
   
   sSql = "Qry_fillvendors"
   bSqlRows = GetDataSet(RdoVed)
   If bSqlRows Then
      With RdoVed
         Do Until .EOF
            AddComboStr cmbVnd.hwnd, "" & Trim(!VENICKNAME)
            .MoveNext
         Loop
      End With
   End If
   
   Set RdoVed = Nothing
   If cmbVnd.ListCount > 0 Then
      cmbVnd.ListIndex = 0
      bGoodVendor = FindVendor(Me)
   End If
   
   sSql = "SELECT DISTINCT VETYPE FROM VndrTable"
   bSqlRows = GetDataSet(RdoTyp)
   If bSqlRows Then
      With RdoTyp
         Do Until .EOF
            If "" & Trim(!VETYPE) <> "" Then
               AddComboStr cmbTyp.hwnd, "" & Trim(!VETYPE)
            End If
            .MoveNext
         Loop
      End With
   End If
   Exit Sub
   DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   SaveOptions
   Set diaPrdPrp14a = Nothing
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sLegend As String
   Dim i As Integer
   MouseCursor 13
   On Error GoTo DiaErr1
   
   i = 1
   sSql = ""
   
   SetMdiReportsize MdiSect
   MdiSect.crw.ReportFileName = sReportPath & "prdpr14.rpt"
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestedBy='Requested By ESI'"
   MdiSect.crw.Formulas(2) = "Title1='Vendor Performance Analysis'"
   MdiSect.crw.Formulas(3) = "Title2='PO items Due From " & Trim(txtsDte) & "  Through " & Trim(txteDte) & "'"
   
   If optSum.Value = vbChecked Then
      MdiSect.crw.Formulas(4) = "Summary='1'"
   Else
      MdiSect.crw.Formulas(4) = "Summary='0'"
   End If
   
   If Trim(txtLte) <> "" Then
      MdiSect.crw.Formulas(5) = "LateWindow=" & txtLte
      MdiSect.crw.Formulas(6) = "Title3='Late Window = " & txtLte & "'"
   Else
      MdiSect.crw.Formulas(5) = "LateWindow=15"
      MdiSect.crw.Formulas(6) = "Title3='Late Window = 15 '"
   End If
   
   If Trim(txtsDte) <> "" Then
      sSql = sSql & "{PoitTable.PIPDATE} >= datetime('" & Trim(txtsDte) & "')  "
   End If
   
   If Trim(txteDte) <> "" Then
      sSql = sSql & " AND {PoitTable.PIPDATE} <= Datetime('" & Trim(txteDte) & "') "
   End If
   
   If Trim(cmbVnd) <> "ALL" Then
      sSql = sSql & " AND {PohdTable.POVENDOR} = '" & Compress(cmbVnd) & "' "
      MdiSect.crw.Formulas(7) = "Vendor ='Vendor = " & cmbVnd & "'"
      MdiSect.crw.Formulas(8) = "MultipleVendors='0'"
   Else
      If Trim(cmbTyp) <> "" Then
         sSql = sSql & " AND {VndrTable.VETYPE} = '" & Trim(cmbTyp) & "' "
         MdiSect.crw.Formulas(7) = "Vendor ='Vendor = ALL  Type = " & cmbTyp & "'"
      Else
         MdiSect.crw.Formulas(7) = "Vendor ='Vendor = ALL  Type = ALL'"
      End If
      MdiSect.crw.Formulas(8) = "MultipleVendors='1'"
   End If
   
   If optNYR.Value = vbUnchecked Then
      sSql = sSql & " AND trim({@status}) <> 'NYR' "
   Else
      MdiSect.crw.Formulas(8 + i) = "Legend" & CStr(i) & " = 'NYR = Not Yet Received'"
      i = i + 1
   End If
   
   If optODD.Value = vbUnchecked Then
      sSql = sSql & " AND trim({@status}) <> 'ODD' "
   Else
      MdiSect.crw.Formulas(8 + i) = "Legend" & CStr(i) & " = 'ODD = On Dock (Delivered)'"
      i = i + 1
   End If
   
   If optODI.Value = vbUnchecked Then
      sSql = sSql & " AND trim({@status}) <> 'ODI' "
   Else
      MdiSect.crw.Formulas(8 + i) = "Legend" & CStr(i) & " = 'ODI = On Dock (Inspected)'"
      i = i + 1
   End If
   
   If optREC.Value = vbUnchecked Then
      sSql = sSql & " AND trim({@status}) <> 'REC' "
   Else
      MdiSect.crw.Formulas(8 + i) = "Legend" & CStr(i) & " = 'REC = Received'"
      i = i + 1
   End If
   
   If optInv.Value = vbUnchecked Then
      sSql = sSql & " AND trim({@status}) <> 'INV' "
   Else
      MdiSect.crw.Formulas(8 + i) = "Legend" & CStr(i) & " = 'INV = Invoiced'"
      i = i + 1
   End If
   
   
   
   If optSum.Value = vbChecked Then
      MdiSect.crw.Formulas(8 + i) = "Summary='1'"
   Else
      MdiSect.crw.Formulas(8 + i) = "Summary='0'"
   End If
   
   MdiSect.crw.SelectionFormula = sSql
   
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "Printreport"
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

Private Sub txtEDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEdte_LostFocus()
   txteDte = CheckDate(txteDte)
End Sub

Private Sub txtSDte_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtSDte_LostFocus()
   txtsDte = CheckDate(txtsDte)
End Sub


Public Sub SaveOptions()
   Dim sOptions As String
   On Error Resume Next
   'Save by Menu Option
   sOptions = RTrim(optNYR.Value) _
              & RTrim(optODD.Value) _
              & RTrim(optODI.Value) _
              & RTrim(optREC.Value) _
              & RTrim(optInv.Value) _
              & Trim(optSum.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, Trim(sOptions)
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter
End Sub

Public Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optNYR.Value = Val(Left(sOptions, 1))
      optODD.Value = Val(Mid(sOptions, 2, 1))
      optODI.Value = Val(Mid(sOptions, 3, 1))
      optREC.Value = Val(Mid(sOptions, 4, 1))
      optInv.Value = Val(Mid(sOptions, 5, 1))
      optSum.Value = Val(Mid(sOptions, 6, 1))
   Else
      optNYR.Value = vbChecked
      optODD.Value = vbChecked
      optODI.Value = vbChecked
      optREC.Value = vbChecked
      optInv.Value = vbChecked
      optSum.Value = vbUnchecked
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & "_Printer", lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = "Default Printer"
   End If
End Sub
