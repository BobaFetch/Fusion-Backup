VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form diaSkillsDemo
   BorderStyle = 3 'Fixed Dialog
   Caption = "Skills Demo"
   ClientHeight = 5985
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 6825
   ClipControls = 0 'False
   ControlBox = 0 'False
   ForeColor = &H80000007&
   LinkTopic = "Form1"
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 5985
   ScaleWidth = 6825
   ShowInTaskbar = 0 'False
   Begin VB.CommandButton cmdSel
      Caption = "Select"
      Height = 315
      Left = 5880
      TabIndex = 4
      ToolTipText = "Refresh Grid"
      Top = 2520
      Width = 875
   End
   Begin MSFlexGridLib.MSFlexGrid grid1
      Height = 2775
      Left = 120
      TabIndex = 5
      Top = 3120
      Width = 6615
      _ExtentX = 11668
      _ExtentY = 4895
      _Version = 393216
      Rows = 1
      Cols = 1
      FixedRows = 0
      FixedCols = 0
      Enabled = 0 'False
   End
   Begin VB.ComboBox cmbSO
      Height = 315
      Left = 1440
      TabIndex = 2
      Top = 1680
      Width = 1575
   End
   Begin VB.ComboBox cmbPO
      Height = 315
      Left = 1440
      TabIndex = 1
      Top = 1200
      Width = 1575
   End
   Begin VB.Frame Frame1
      Height = 30
      Left = 120
      TabIndex = 19
      Top = 3000
      Width = 6615
   End
   Begin VB.TextBox txtPrt
      Height = 285
      Left = 1440
      TabIndex = 3
      Tag = "3"
      Top = 2160
      Width = 2775
   End
   Begin VB.ComboBox cmbCst
      Height = 315
      Left = 1440
      TabIndex = 0
      Tag = "3"
      Top = 360
      Width = 1575
   End
   Begin VB.CommandButton cmdVew
      Height = 320
      Left = 4320
      Picture = "diaSkillsDemo.frx":0000
      Style = 1 'Graphical
      TabIndex = 12
      TabStop = 0 'False
      ToolTipText = "Show BOM Structure"
      Top = 2160
      UseMaskColor = -1 'True
      Width = 350
   End
   Begin VB.CheckBox optVew
      Height = 255
      Left = 3840
      TabIndex = 11
      Top = 0
      Visible = 0 'False
      Width = 735
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 5640
      TabIndex = 9
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 615
      Left = 5640
      TabIndex = 6
      Top = 360
      Width = 1215
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaSkillsDemo.frx":0342
         Style = 1 'Graphical
         TabIndex = 8
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaSkillsDemo.frx":04C0
         Style = 1 'Graphical
         TabIndex = 7
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 13
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
      PictureUp = "diaSkillsDemo.frx":064A
      PictureDn = "diaSkillsDemo.frx":0790
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 4920
      Top = 0
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 5985
      FormDesignWidth = 6825
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 14
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
      PictureUp = "diaSkillsDemo.frx":08D6
      PictureDn = "diaSkillsDemo.frx":0A1C
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 11
      Left = 3120
      TabIndex = 25
      Top = 1680
      Width = 1665
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "(Blank For All)"
      Height = 285
      Index = 9
      Left = 3120
      TabIndex = 24
      Top = 1200
      Width = 1665
   End
   Begin VB.Label lblDsc
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 23
      Top = 2520
      Width = 2775
   End
   Begin VB.Image imgInc
      Height = 180
      Left = 6480
      Picture = "diaSkillsDemo.frx":0B6E
      Top = 2280
      Visible = 0 'False
      Width = 255
   End
   Begin VB.Image imgdInc
      Height = 180
      Left = 6120
      Picture = "diaSkillsDemo.frx":0E20
      Top = 2280
      Visible = 0 'False
      Width = 255
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "or"
      Height = 285
      Index = 5
      Left = 360
      TabIndex = 22
      Top = 1920
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "or"
      Height = 285
      Index = 4
      Left = 360
      TabIndex = 21
      Top = 1440
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Sales Order #"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 20
      Top = 1680
      Width = 1065
   End
   Begin VB.Label lblNme
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 285
      Left = 1440
      TabIndex = 18
      Top = 720
      Width = 2775
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Customer PO #"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 17
      Top = 1200
      Width = 1425
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Customer"
      Height = 285
      Index = 1
      Left = 120
      TabIndex = 16
      Top = 360
      Width = 1665
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Number"
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 15
      Top = 2160
      Width = 1665
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      Caption = "Default Printer"
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 10
      Top = 0
      Width = 2760
   End
End
Attribute VB_Name = "diaSkillsDemo"
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
' diaSkillsDemo - Part Inquiry (lookup)
'
' Notes:
'
' Created: 07/15/03 (nth)
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPart As Byte
Dim bGoodCust As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'Private Sub cmbPO_Click()
'    sRefreshSO
'End Sub

'*********************************************************************************

Private Sub cmbPO_LostFocus()
   If Not bCancel Then
      sRefreshSO
      DoEvents
   End If
End Sub


Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmbCst_Click()
   bGoodCust = FindThisCustomer(Me)
End Sub

Private Sub cmbCst_LostFocus()
   If Not bCancel Then
      bGoodCust = FindThisCustomer(Me)
      If bGoodCust Then
         sFillPO
         sRefreshSO
      End If
   End If
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             X As Single, Y As Single)
   bCancel = True
End Sub

Private Sub cmdSel_Click()
   sRefreshGrid
End Sub

Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   VewParts.Show
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
   With grid1
      .Enabled = False
      .Rows = 0
      .Cols = 0
   End With
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaSkillsDemo = Nothing
End Sub

Private Sub optDis_Click()
   PrintReport
End Sub

Private Sub optPrn_Click()
   PrintReport
End Sub

Private Sub ShowPrinters_Click(Value As Integer)
   SysPrinters.Show
   ShowPrinters.Value = False
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub PrintReport()
   Dim sPart As String
   Dim sPO As String
   Dim sSO As String
   Dim sItem As String
   Dim sRev As String
   Dim sTemp As String
   Dim b As Byte
   Dim I As Integer
   
   MouseCursor 13
   
   On Error GoTo DiaErr1
   
   With grid1
      For I = 0 To .Rows - 1
         .Row = I
         .Col = 0
         If .CellPicture = imgInc Then
            b = True
            
            .Col = 1
            sTemp = Trim(.Text)
            If sTemp <> "" Then
               If InStr(1, sPO, sTemp & "',") = 0 Then
                  sPO = sPO & "'" & sTemp & "',"
               End If
            End If
            
            .Col = 2
            sTemp = Trim(.Text)
            If sTemp <> "" Then
               If Not IsNumeric(sTemp) Then
                  sTemp = Right(sTemp, Len(sTemp) - 1)
               End If
               sTemp = CStr(CLng(sTemp))
               If InStr(1, sSO, sTemp & ",") = 0 Then
                  sSO = sSO & sTemp & ","
               End If
            End If
            
            .Col = 3
            sTemp = Trim(.Text)
            If sTemp <> "" Then
               If Not IsNumeric(Right(sTemp, 1)) Then
                  If InStr(1, sRev, Right(sTemp, 1) & "',") = 0 Then
                     sRev = sRev & "'" & Right(sTemp, 1) & "',"
                  End If
                  sTemp = Left(sTemp, Len(sTemp) - 1)
               End If
               If InStr(1, sItem, sTemp & ",") = 0 Then
                  sItem = sItem & sTemp & ","
               End If
            End If
            
            .Col = 4
            sTemp = Compress(.Text)
            If sTemp <> "" Then
               If InStr(1, sPart, sTemp & "',") = 0 Then
                  sPart = sPart & "'" & sTemp & "',"
               End If
            End If
         End If
      Next
   End With
   
   If b Then
      optPrn.Enabled = False
      optDis.Enabled = False
      
      SetMdiReportsize MdiSect
      
      MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
      MdiSect.crw.Formulas(1) = "RequestBy = 'Requested By: " _
                           & Secure.UserInitials & "'"
      
      MdiSect.crw.ReportFileName = sReportPath & "SkillsDemo.rpt"
      
      sSql = "{SohdTable.SOCUST} = '" & Compress(cmbCst) & "'"
      
      If Len(sPO) Then
         sSql = sSql & " AND {SohdTable.SOPO} IN [" & Left(sPO, Len(sPO) - 1) & "]"
      End If
      
      If Len(sSO) Then
         sSql = sSql & " AND {SohdTable.SONUMBER} IN [" & Left(sSO, Len(sSO) - 1) & "]"
      End If
      
      If Len(sItem) Then
         sSql = sSql & " AND {SoitTable.ITNUMBER} IN [" & Left(sItem, Len(sItem) - 1) & "]"
      End If
      
      'If Len(sRev) Then
      sRev = sRev & "''"
      sSql = sSql & " AND {SoitTable.ITREV} IN [" & sRev & "]"
      'End If
      
      If Len(sPart) Then
         sSql = sSql & " AND {SoitTable.ITPART} IN [" & Left(sPart, Len(sPart) - 1) & "]"
      End If
      
      'MsgBox sSql
      
      MdiSect.crw.SelectionFormula = sSql
      
      SetCrystalAction Me
      
      optPrn.Enabled = True
      optDis.Enabled = True
   Else
      MsgBox "No Items Selected.", vbInformation, Caption
      'grid1.SetFocus
   End If
   MouseCursor 0
   Exit Sub
   
   DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub FillCombo()
   FillCustomers Me
   bGoodCust = FindThisCustomer(Me)
   If bGoodCust Then
      sFillPO
      sRefreshSO
   End If
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
End Sub

Private Sub GetOptions()
   Dim sOptions As String
End Sub

Private Sub sRefreshGrid()
   Dim RdoItm As rdoResultset
   Dim sCust As String
   Dim sItem As String
   Dim sSO As String
   Dim I As Integer
   
   sCust = Compress(cmbCst)
   
   With grid1
      .Enabled = True
      .Clear
      .Rows = 0
   End With
   
   sSql = "SELECT SONUMBER,SOTYPE,SOPO,ITNUMBER,ITREV,PARTNUM,ITCUSTREQ,ITSCHED,ITACTUAL " _
          & "FROM SohdTable INNER JOIN SoitTable ON SohdTable.SONUMBER = SoitTable.ITSO " _
          & "INNER JOIN PartTable ON SoitTable.ITPART = PartTable.PARTREF WHERE SOCUST = '" & sCust & "' "
   If Trim(cmbPO) <> "" Then
      sSql = sSql & " AND SOPO = '" & Trim(cmbPO) & "' "
   End If
   
   If Trim(cmbSO) <> "" Then
      sSO = Trim(cmbSO)
      If Not IsNumeric(Left(sSO, 1)) Then
         sSO = Right(sSO, Len(sSO) - 1)
      End If
      sSql = sSql & " AND SONUMBER = " & CLng(sSO)
   End If
   
   If Trim(txtPrt) <> "" Then
      sSql = sSql & " AND ITPART = '" & Compress(txtPrt) & "'"
   End If
   
   bSqlRows = GetDataSet(RdoItm)
   If bSqlRows Then
      With grid1
         .Cols = 8
         .Rows = 1
         .ColWidth(0) = 500
         .ColWidth(1) = 1000
         .ColWidth(2) = 1000
         .ColWidth(3) = 500
         .ColWidth(4) = 2775
         .ColWidth(5) = 1000
         .ColWidth(6) = 1000
         .ColWidth(7) = 1000
      End With
      I = 1
      With RdoItm
         Do Until .EOF
            sItem = Chr(9) & " " & Trim(.rdoColumns(2)) _
                    & Chr(9) & " " & Trim(.rdoColumns(1)) _
                    & Format(.rdoColumns(0), "000000") _
                    & Chr(9) & " " & .rdoColumns(3) & Trim(.rdoColumns(4)) _
                    & Chr(9) & " " & Trim(.rdoColumns(5)) _
                    & Chr(9) & " " & Format("" & .rdoColumns(6), "m/d/yy") _
                    & Chr(9) & " " & Format("" & .rdoColumns(7), "m/d/yy") _
                    & Chr(9) & " " & Format("" & .rdoColumns(8), "m/d/yy")
            grid1.AddItem sItem
            grid1.Row = I
            grid1.Col = 0
            grid1.CellPictureAlignment = flexAlignCenterCenter
            Set grid1.CellPicture = imgdInc
            .MoveNext
            I = I + 1
         Loop
         .Cancel
      End With
      With grid1
         .FixedRows = 1
         .Row = 0
         .Col = 1
         .Text = "Customer PO"
         .Col = 2
         .Text = "Sales Order"
         .Col = 3
         .Text = "Item"
         .Col = 4
         .Text = "Part Number"
         .Col = 5
         .Text = "Requested"
         .Col = 6
         .Text = "Scheduled"
         .Col = 7
         .Text = "Shipped"
      End With
      
   Else
      With grid1
         .Cols = 1
         .Rows = 1
         .ColWidth(0) = 6500
         .CellAlignment = flexAlignCenterCenter
         .Text = "*** No Items Found ***"
      End With
   End If
   Set RdoItm = Nothing
   Exit Sub
   
   DiaErr1:
   sProcName = "sfreshgrid"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub sFillPO()
   Dim rdoPo As rdoResultset
   Dim sCust As String
   
   sCust = Compress(cmbCst)
   
   cmbPO.Clear
   sSql = "SELECT DISTINCT SOPO FROM SohdTable WHERE SOCUST = '" _
          & sCust & "'"
   bSqlRows = GetDataSet(rdoPo)
   If bSqlRows Then
      With rdoPo
         Do Until .EOF
            If "" & Trim(.rdoColumns(0)) <> "" Then
               AddComboStr cmbPO.hwnd, "" & Trim(.rdoColumns(0))
            End If
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoPo = Nothing
   
   Exit Sub
   
   DiaErr1:
   sProcName = "sfillpo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub Grid1_Click()
   With grid1
      If Left(.Text, 3) <> "***" Then
         .Col = 0
         .Row = .RowSel
         If .CellPicture = imgdInc Then
            Set .CellPicture = imgInc
         Else
            Set .CellPicture = imgdInc
         End If
      End If
   End With
End Sub

Private Sub txtPrt_LostFocus()
   If Not bCancel Then
      txtPrt = CheckLen(txtPrt, 30)
      If txtPrt <> "" Then
         FindPart Me, txtPrt
      Else
         lblDsc = "*** All Part Numbers Selected ***"
      End If
   End If
End Sub

Private Sub sRefreshSO()
   Dim rdoso As rdoResultset
   Dim sCust As String
   
   sCust = Compress(cmbCst)
   cmbSO.Clear
   sSql = "SELECT SOTYPE,SONUMBER FROM SohdTable WHERE SOCUST = '" _
          & sCust & "'"
   If Trim(cmbPO) <> "" Then
      sSql = sSql & " AND SOPO = '" & Trim(cmbPO) & "'"
   End If
   
   bSqlRows = GetDataSet(rdoso)
   If bSqlRows Then
      With rdoso
         Do Until .EOF
            AddComboStr cmbSO.hwnd, "" & Trim(.rdoColumns(0)) _
               & Format(CStr(.rdoColumns(1)), "000000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set rdoso = Nothing
End Sub
