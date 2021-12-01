VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaPin05
   BorderStyle = 3 'Fixed Dialog
   Caption = "Raw Material And Finished Goods Inventory"
   ClientHeight = 3525
   ClientLeft = 2115
   ClientTop = 1125
   ClientWidth = 7260
   ControlBox = 0 'False
   ForeColor = &H00C0C0C0&
   LinkTopic = "Form1"
   LockControls = -1 'True
   MDIChild = -1 'True
   MinButton = 0 'False
   PaletteMode = 1 'UseZOrder
   ScaleHeight = 3525
   ScaleWidth = 7260
   ShowInTaskbar = 0 'False
   Begin ComctlLib.ProgressBar Prg1
      Height = 255
      Left = 2040
      TabIndex = 24
      Top = 3120
      Visible = 0 'False
      Width = 5055
      _ExtentX = 8916
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.CheckBox OptCmt
      Caption = "____"
      ForeColor = &H00C0C0C0&
      Height = 285
      Left = 2040
      TabIndex = 11
      Top = 2760
      Width = 735
   End
   Begin VB.CheckBox optDsc
      Caption = "____"
      ForeColor = &H00C0C0C0&
      Height = 195
      Left = 2040
      TabIndex = 10
      Top = 2520
      Width = 735
   End
   Begin VB.OptionButton optFif
      Caption = "FIFO"
      Enabled = 0 'False
      Height = 195
      Left = 5400
      TabIndex = 9
      Top = 1920
      Width = 1215
   End
   Begin VB.OptionButton optLst
      Caption = "Last Cost"
      Height = 255
      Left = 4320
      TabIndex = 8
      Top = 1920
      Width = 1215
   End
   Begin VB.OptionButton optAve
      Caption = "Average"
      Height = 255
      Left = 3240
      TabIndex = 7
      Top = 1920
      Width = 1215
   End
   Begin VB.OptionButton optStd
      Caption = "Standard"
      Height = 195
      Left = 2040
      TabIndex = 6
      Top = 1920
      Value = -1 'True
      Width = 1215
   End
   Begin VB.CheckBox typ
      Caption = "1"
      Height = 255
      Index = 1
      Left = 2040
      TabIndex = 2
      Top = 1560
      Value = 1 'Checked
      Width = 495
   End
   Begin VB.CheckBox typ
      Caption = "2"
      Height = 255
      Index = 2
      Left = 2520
      TabIndex = 3
      Top = 1560
      Value = 1 'Checked
      Width = 495
   End
   Begin VB.CheckBox typ
      Caption = "3"
      Height = 255
      Index = 3
      Left = 3000
      TabIndex = 4
      Top = 1560
      Value = 1 'Checked
      Width = 495
   End
   Begin VB.CheckBox typ
      Caption = "4"
      Height = 255
      Index = 4
      Left = 3480
      TabIndex = 5
      Top = 1560
      Value = 1 'Checked
      Width = 495
   End
   Begin VB.TextBox txtMth
      Height = 285
      Left = 2040
      TabIndex = 0
      Tag = "4"
      Top = 720
      Width = 855
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 6120
      TabIndex = 17
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
      Begin VB.CommandButton optDis
         Height = 330
         Left = 0
         Picture = "diaPin05.frx":0000
         Style = 1 'Graphical
         TabIndex = 15
         ToolTipText = "Display The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optPrn
         Height = 330
         Left = 560
         Picture = "diaPin05.frx":017E
         Style = 1 'Graphical
         TabIndex = 16
         ToolTipText = "Print The Report"
         Top = 120
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.ComboBox txtDte
      Height = 315
      Left = 2040
      TabIndex = 1
      Tag = "4"
      Top = 1080
      Width = 1095
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
      PictureUp = "diaPin05.frx":0308
      PictureDn = "diaPin05.frx":044E
   End
   Begin ResizeLibCtl.ReSize ReSize1
      Left = 240
      Top = 3120
      _Version = 196615
      _ExtentX = 741
      _ExtentY = 741
      _StockProps = 0
      Enabled = -1 'True
      FormMinWidth = 0
      FormMinHeight = 0
      FormDesignHeight = 3525
      FormDesignWidth = 7260
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Descriptions?"
      Height = 285
      Index = 6
      Left = 240
      TabIndex = 23
      Top = 2520
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions?"
      Height = 285
      Index = 3
      Left = 240
      TabIndex = 22
      Top = 2760
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 5
      Left = 240
      TabIndex = 21
      Top = 2280
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Value Parts At"
      Height = 285
      Index = 2
      Left = 240
      TabIndex = 20
      Top = 1920
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Types?"
      Height = 285
      Index = 4
      Left = 240
      TabIndex = 19
      Top = 1560
      Width = 1305
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Month End"
      Height = 285
      Index = 1
      Left = 240
      TabIndex = 18
      Top = 720
      Width = 1185
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "As Of Date"
      Height = 285
      Index = 0
      Left = 240
      TabIndex = 12
      Top = 1080
      Width = 1185
   End
End
Attribute VB_Name = "diaPin05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOnload As Byte
Dim DbInv As Recordset 'Jet

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub CreateRmFgTable()
   Dim NewTb As TableDef
   Dim InvTb As TableDef
   Dim NewFld As Field
   Dim NewIdx As Index
   
   On Error Resume Next
   JetDb.Execute "DROP TABLE IrmfgTable"
   'Fields. Note that we allow empties
   Set NewTb = JetDb.CreateTableDef("IrmfgTable")
   With NewTb
      'PartRef
      .Fields.Append .CreateField("Inv00", dbText, 30)
      .Fields(0).AllowZeroLength = True
      'Part Number
      .Fields.Append .CreateField("Inv01", dbText, 30)
      .Fields(1).AllowZeroLength = True
      'Part Desc
      .Fields.Append .CreateField("Inv02", dbText, 30)
      .Fields(2).AllowZeroLength = True
      'Part Loc
      .Fields.Append .CreateField("Inv03", dbText, 4)
      .Fields(3).AllowZeroLength = True
      'Part Um
      .Fields.Append .CreateField("Inv04", dbText, 2)
      .Fields(4).AllowZeroLength = True
      'Part Qoh
      .Fields.Append .CreateField("Inv05", dbCurrency)
      .Fields(5).DefaultValue = 0
      'Cost
      .Fields.Append .CreateField("Inv06", dbCurrency)
      .Fields(6).DefaultValue = 0
      'Total Cost
      .Fields.Append .CreateField("Inv07", dbCurrency)
      .Fields(7).DefaultValue = 0
      'level
      .Fields.Append .CreateField("Inv08", dbInteger)
      .Fields(8).DefaultValue = 1
   End With
   
   'add the table and indexes to Jet.
   JetDb.TableDefs.Append NewTb
   Set InvTb = JetDb!IrmfgTable
   With InvTb
      Set NewIdx = .CreateIndex
      With NewIdx
         .Name = "PartRefIdx"
         .Fields.Append .CreateField("Inv00")
      End With
      .Indexes.Append NewIdx
   End With
   
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



Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   ReopenJet
   GetOptions
   GetMonth
   CreateRmFgTable
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   On Error Resume Next
   JetDb.Execute "DROP TABLE IrmfgTable"
   Set diaPin05 = Nothing
   
End Sub




Private Sub PrintReport()
   Dim sCostType As String
   Dim sIncludes As String
   Dim sWindows As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   Prg1.Visible = False
   If optStd.Value = True Then sCostType = "Standard Cost"
   If optAve.Value = True Then sCostType = "Average Cost"
   If optLst.Value = True Then sCostType = "Last Cost"
   If optFif.Value = True Then sCostType = "FIFO"
   
   If typ(1).Value = vbChecked Then sIncludes = ",1"
   If typ(2).Value = vbChecked Then sIncludes = sIncludes & ",2"
   If typ(3).Value = vbChecked Then sIncludes = sIncludes & ",3"
   If typ(4).Value = vbChecked Then sIncludes = sIncludes & ",4"
   SetMdiReportsize MdiSect
   sWindows = GetWindowsDir()
   MdiSect.crw.DataFiles(0) = sWindows & "\temp\esifina.mdb"
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "Includes='Cutoff Date " & txtDte & " And Types" _
                        & sIncludes & "'"
   MdiSect.crw.Formulas(2) = "UnitCost='" & sCostType & "'"
   MdiSect.crw.ReportFileName = sReportPath & "prdin05.rpt"
   If OptCmt.Value = vbChecked Then
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;T;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPFTR.0.1;T;;;"
   Else
      MdiSect.crw.SectionFormat(0) = "GROUPFTR.0.0;F;;;"
      MdiSect.crw.SectionFormat(1) = "GROUPFTR.0.1;F;;;"
   End If
   If optDsc = vbChecked Then
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.1.0;T;;;"
   Else
      MdiSect.crw.SectionFormat(2) = "GROUPFTR.1.0;F;;;"
   End If
   MdiSect.crw.SelectionFormula = ""
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
End Sub

Public Sub SaveOptions()
   Dim i As Integer
   Dim sOptions As String
   
   'Save by Menu Option
   On Error Resume Next
   For i = 1 To 3
      sOptions = sOptions & Trim(Str(typ(i).Value))
   Next
   sOptions = sOptions & Trim(Str(typ(i).Value))
   If optStd.Value = True Then
      sOptions = sOptions & "1"
   Else
      sOptions = sOptions & "0"
   End If
   If optAve.Value = True Then
      sOptions = sOptions & "1"
   Else
      sOptions = sOptions & "0"
   End If
   If optLst.Value = True Then
      sOptions = sOptions & "1"
   Else
      sOptions = sOptions & "0"
   End If
   If optFif.Value = True Then
      sOptions = sOptions & "1"
   Else
      sOptions = sOptions & "0"
   End If
   sOptions = sOptions & Trim(Str(optDsc.Value)) & Trim(Str(OptCmt.Value))
   SaveSetting "Esi2000", "EsiProd", "in05", Trim(sOptions)
   
End Sub

Public Sub GetOptions()
   Dim i As Integer
   Dim sOptions As String
   
   'Save by Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "in05", Trim(sOptions))
   If Len(sOptions) Then
      For i = 1 To 3
         typ(i).Value = Mid(sOptions, i, 1)
      Next
      typ(i).Value = Mid(sOptions, i, 1)
      If Mid(sOptions, i + 1, 1) = "1" Then optStd.Value = True
      If Mid(sOptions, i + 2, 1) = "1" Then optAve.Value = True
      If Mid(sOptions, i + 3, 1) = "1" Then optLst.Value = True
      If Mid(sOptions, i + 4, 1) = "1" Then optFif.Value = False
      optDsc.Value = Val(Mid(sOptions, i + 5, 1))
      OptCmt.Value = Val(Mid(sOptions, i + 6, 1))
   End If
   
End Sub

Private Sub optAve_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optCmt_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   BuildInventory
   
End Sub


Private Sub optDis_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDsc_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optFif_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optLst_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   BuildInventory
   
End Sub



Public Sub GetMonth()
   Dim a As Integer
   Dim n As Integer
   Dim i As Integer
   Dim sYear As String
   
   On Error GoTo DiaErr1
   If Len(Trim(txtMth)) = 0 Then
      i = Month(Now)
      i = i - 1
      n = Year(Now)
      If i = 0 Then
         i = 12
         n = n - 1
      End If
   Else
      i = Month(Format(txtMth, "mm/dd/yy"))
      n = Year(Format(txtMth, "mm/dd/yy"))
   End If
   Select Case i
      Case 1, 3, 5, 7, 8, 10, 12
         a = 31
      Case 2
         a = 28
      Case 4, 6, 9, 11
         a = 30
   End Select
   If a = 28 Then
      Select Case n
         Case 1992, 1996, 2000, 2004, 2008, 2012, 2016
            a = 29
         Case 2020, 2024, 2026, 2030, 2034, 2038, 2042
            a = 29
      End Select
   End If
   sYear = Trim(Str(n))
   txtMth = Format(i, "00") & "/" & Format(a, "00") & "/" _
            & Right(sYear, 2)
   txtDte = Format(Now, "mm/dd/yy")
   Exit Sub
   
   DiaErr1:
   sProcName = "getmonth"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Prg1.Visible = False
   DoModuleErrors Me
   
End Sub

Private Sub optStd_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub


Private Sub txtMth_LostFocus()
   txtMth = CheckDate(txtMth)
   GetMonth
   
End Sub


Private Sub typ_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub



Public Sub BuildInventory()
   Dim bByte As Boolean
   Dim a As Single
   Dim i As Integer
   Dim n As Single
   Dim RdoInv As rdoResultset
   bByte = False
   For i = 1 To 4
      If typ(i).Value = vbChecked Then bByte = True
   Next
   If Not bByte Then
      MsgBox "Must Select At Least One Type.", vbInformation, Caption
      Exit Sub
   End If
   MouseCursor 13
   JetDb.Execute "DELETE * FROM IrmfgTable"
   Prg1.Visible = True
   On Error GoTo DiaErr1
   Prg1.Value = 10
   sSql = "SELECT DISTINCT INPART,PARTREF,PARTNUM,PADESC," _
          & "PALOCATION,PAUNITS,PALEVEL FROM InvaTable,PartTable " _
          & "WHERE INPART = PARTREF AND INADATE<='" & txtDte & "' "
   If typ(1).Value = vbUnchecked Then sSql = sSql & "AND PALEVEL<>1 "
   If typ(2).Value = vbUnchecked Then sSql = sSql & "AND PALEVEL<>2 "
   If typ(3).Value = vbUnchecked Then sSql = sSql & "AND PALEVEL<>3 "
   If typ(4).Value = vbUnchecked Then sSql = sSql & "AND PALEVEL<>4 "
   sSql = sSql & "AND PALEVEL<>5 AND PALEVEL<>6 AND PALEVEL<>7 AND PALEVEL<>8"
   bSqlRows = GetDataSet(RdoInv)
   If bSqlRows Then
      Set DbInv = JetDb.OpenRecordset("IrmfgTable", dbOpenDynaset)
      With RdoInv
         Do Until .EOF
            i = i + 1
            DbInv.AddNew
            DbInv!Inv00 = "" & Trim(!PARTREF)
            DbInv!Inv01 = "" & Trim(!PARTNUM)
            DbInv!Inv02 = "" & Trim(!PADESC)
            DbInv!Inv03 = "" & Trim(!PALOCATION)
            DbInv!Inv04 = "" & Trim(!PAUNITS)
            DbInv!Inv08 = Format(!PALEVEL, "0")
            DbInv.Update
            .MoveNext
         Loop
         .Cancel
      End With
      n = 80 / i
      a = 20
      Prg1.Value = a
      On Error Resume Next
      If DbInv.RecordCount > 0 Then
         With DbInv
            .MoveFirst
            Do Until .EOF
               a = a + n
               If a > 95 Then a = 95
               Prg1.Value = a
               .Edit
               sSql = "SELECT SUM(INAQTY) FROM InvaTable WHERE INPART='" _
                      & Trim(!Inv00) & "' AND INADATE<='" & txtDte & "' "
               bSqlRows = GetDataSet(RdoInv)
               If bSqlRows Then
                  If Not IsNull(RdoInv.rdoColumns(0)) Then
                     !Inv05 = Format(RdoInv.rdoColumns(0), "######0.0000")
                  End If
                  RdoInv.Cancel
               End If
               If optAve.Value = True Then
                  sSql = "SELECT SUM(INAMT*Abs(INAQTY))/SUM(Abs(INAQTY)) " _
                         & "From InvaTable WHERE INAQTY<>0 AND " _
                         & "(INPART='" & !Inv00 & "' AND INADATE<='" & txtDte & "') "
                  Set RdoInv = RdoCon.OpenResultset(sSql, rdOpenForwardOnly, rdConcurReadOnly)
                  If Not IsNull(RdoInv.rdoColumns(0)) Then
                     !Inv06 = Format(RdoInv.rdoColumns(0), "######0.0000")
                  End If
               Else
                  If optStd.Value = True Then
                     sSql = "SELECT PARTREF,PASTDCOST FROM PartTable " _
                            & "WHERE PARTREF='" & !Inv00 & "'"
                     bSqlRows = GetDataSet(RdoInv)
                     If bSqlRows Then
                        If Not IsNull(RdoInv!PASTDCOST) Then
                           !Inv06 = Format(RdoInv!PASTDCOST, "######0.0000")
                        End If
                     End If
                  Else
                     sSql = "SELECT INPART,INAMT FROM InvaTable " _
                            & "WHERE INPART='" & !Inv00 & "' AND INADATE<='" _
                            & txtDte & "' ORDER BY INADATE DESC"
                     bSqlRows = GetDataSet(RdoInv)
                     If bSqlRows Then
                        !Inv06 = Format(RdoInv!INAMT, "######0.0000")
                     End If
                  End If
               End If
               RdoInv.Cancel
               !Inv07 = Format(!Inv06 * !Inv05, "######0.0000")
               .Update
               .MoveNext
            Loop
         End With
      End If
   End If
   On Error Resume Next
   Prg1.Value = 100
   DbInv.Close
   Set RdoInv = Nothing
   PrintReport
   Exit Sub
   
   DiaErr1:
   sProcName = "buildinv"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Prg1.Visible = False
   DoModuleErrors Me
   
End Sub
