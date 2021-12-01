VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form diaWip2
   BorderStyle = 3 'Fixed Dialog
   Caption = "Work In Process Report (2)"
   ClientHeight = 4020
   ClientLeft = 45
   ClientTop = 330
   ClientWidth = 6795
   LinkTopic = "Form1"
   MaxButton = 0 'False
   MDIChild = -1 'True
   MinButton = 0 'False
   ScaleHeight = 4020
   ScaleWidth = 6795
   ShowInTaskbar = 0 'False
   Begin VB.Frame fraPrn
      BorderStyle = 0 'None
      Height = 495
      Left = 5400
      TabIndex = 6
      Top = 360
      Width = 1335
      Begin VB.CommandButton optPrn
         Height = 390
         Left = 675
         Style = 1 'Graphical
         TabIndex = 8
         ToolTipText = "Print The Report"
         Top = 60
         UseMaskColor = -1 'True
         Width = 495
      End
      Begin VB.CommandButton optDis
         Height = 390
         Left = 120
         Style = 1 'Graphical
         TabIndex = 7
         ToolTipText = "Display The Report"
         Top = 60
         UseMaskColor = -1 'True
         Width = 495
      End
   End
   Begin VB.CommandButton cmdCan
      Cancel = -1 'True
      Caption = "Close"
      Height = 360
      Left = 5520
      TabIndex = 5
      TabStop = 0 'False
      Top = 0
      Width = 1065
   End
   Begin VB.CheckBox optExt
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 4
      Top = 2760
      Width = 855
   End
   Begin VB.CheckBox optDsc
      Caption = "___"
      ForeColor = &H8000000F&
      Height = 195
      Left = 2040
      TabIndex = 3
      Top = 2520
      Width = 855
   End
   Begin VB.ComboBox cboAsOf
      Height = 315
      Left = 2040
      TabIndex = 2
      Tag = "4"
      Top = 720
      Width = 1095
   End
   Begin VB.ComboBox cboClass
      Height = 315
      Left = 2040
      TabIndex = 1
      Top = 1200
      Width = 1095
   End
   Begin VB.ComboBox cboCode
      Height = 315
      Left = 2040
      TabIndex = 0
      Top = 1560
      Width = 1095
   End
   Begin Threed.SSRibbon cmdHlp
      Height = 225
      Left = 0
      TabIndex = 9
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
      PictureUp = "diaWip2.frx":0000
      PictureDn = "diaWip2.frx":0146
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
      FormDesignHeight = 4020
      FormDesignWidth = 6795
   End
   Begin Threed.SSRibbon ShowPrinters
      Height = 255
      Left = 360
      TabIndex = 10
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
      PictureUp = "diaWip2.frx":028C
      PictureDn = "diaWip2.frx":03D2
   End
   Begin ComctlLib.ProgressBar Prg1
      Height = 255
      Left = 120
      TabIndex = 11
      Top = 3600
      Visible = 0 'False
      Width = 6495
      _ExtentX = 11456
      _ExtentY = 450
      _Version = 327682
      Appearance = 1
   End
   Begin VB.Label Label2
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 4320
      TabIndex = 23
      Top = 3240
      Visible = 0 'False
      Width = 2295
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Descriptions"
      Height = 285
      Index = 6
      Left = 120
      TabIndex = 22
      Top = 2520
      Width = 1785
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Extended Descriptions"
      Height = 285
      Index = 3
      Left = 120
      TabIndex = 21
      Top = 2760
      Width = 1815
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Include:"
      Height = 285
      Index = 5
      Left = 120
      TabIndex = 20
      Top = 2280
      Width = 1785
   End
   Begin VB.Label lblPrinter
      Appearance = 0 'Flat
      BorderStyle = 1 'Fixed Single
      ForeColor = &H00800000&
      Height = 255
      Left = 720
      TabIndex = 19
      Top = 0
      Width = 2760
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Part Class"
      Height = 285
      Index = 2
      Left = 120
      TabIndex = 18
      Top = 1200
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "As Of "
      Height = 285
      Index = 0
      Left = 120
      TabIndex = 17
      Top = 720
      Width = 1065
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Product Code"
      Height = 285
      Index = 4
      Left = 120
      TabIndex = 16
      Top = 1560
      Width = 1065
   End
   Begin VB.Label z1
      Alignment = 2 'Center
      BackStyle = 0 'Transparent
      Caption = "Of"
      Height = 285
      Index = 7
      Left = 2040
      TabIndex = 15
      Top = 3240
      Visible = 0 'False
      Width = 945
   End
   Begin VB.Label z1
      BackStyle = 0 'Transparent
      Caption = "Record"
      Height = 285
      Index = 8
      Left = 120
      TabIndex = 14
      Top = 3240
      Visible = 0 'False
      Width = 945
   End
   Begin VB.Label lblRuns
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 1080
      TabIndex = 13
      Top = 3240
      Visible = 0 'False
      Width = 855
   End
   Begin VB.Label lblCount
      Alignment = 1 'Right Justify
      BackStyle = 0 'Transparent
      BorderStyle = 1 'Fixed Single
      Height = 255
      Left = 3000
      TabIndex = 12
      Top = 3240
      Visible = 0 'False
      Width = 855
   End
End
Attribute VB_Name = "diaWip2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1/7/05 CJS for WIP
Option Explicit
Dim RdoWip As rdoResultset
Dim bOnLoad As Byte
Dim sTempTable As String
Dim lTotalRuns As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   cboAsOf = Format(ES_SYSDATE, "mm/dd/yy")
   'cboAsOf = Left(cboAsOf, 3) & "01" & Right(cboAsOf, 3)
   Label2 = ""
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Function DataReady() As Boolean
   Dim b As Byte
   DataReady = False
   lTotalRuns = 0
   lblRuns = "0"
   sSql = "TRUNCATE TABLE EsReportWIP"
   RdoCon.Execute sSql, rdExecDirect
   If Trim(Label2) <> "" Then
      sSql = "TRUNCATE TABLE " & sTempTable
      RdoCon.Execute sSql, rdExecDirect
   Else
      BuildTempTable
   End If
   If sTempTable = "" Then
      MsgBox "Could Not Build The Temporary Table.", _
         vbInformation, Caption
   Else
      b = GetWipRuns()
      If b = 1 Then
         BuildReport
         DataReady = True
      Else
         MsgBox "No WIP Found With The Selected Parameters.", _
            vbInformation, Caption
      End If
   End If
   
End Function

Private Sub Form_Activate()
   If bOnLoad Then
      CheckReportTable
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   SetDiaPos Me
   FormatControls
   sCurrForm = Caption
   GetOptions
   optPrn.Picture = Resources.imgPrn.Picture
   optDis.Picture = Resources.imgDis.Picture
   bOnLoad = 1
   
   PopulateCombo cboClass, "PACLASS", "PartTable"
   PopulateCombo cboCode, "PAPRODCODE", "PartTable"
   
End Sub

Private Sub PopulateCombo(cbo As ComboBox, sColumn As String, sTable As String)
   'populate combobox from database table values of a specific column
   
   cbo.Clear
   cbo.AddItem "<ALL>"
   
   Dim rdo As rdoResultset
   sSql = "SELECT " & sColumn & " FROM " & sTable & " GROUP BY " & sColumn
   bSqlRows = GetDataSet(rdo, ES_FORWARD)
   If bSqlRows Then
      With rdo
         Do Until .EOF
            If Trim(.rdoColumns(0)) = "" Then
               cbo.AddItem "<BLANK>"
            Else
               cbo.AddItem Trim(.rdoColumns(0))
            End If
            .MoveNext
         Loop
      End With
   End If
   cbo.ListIndex = 0
End Sub

Private Function GetWipRuns() As Byte
   Dim RdoMos As rdoResultset
   Dim bLots As Byte
   Dim bShowLots As Byte
   
   bLots = CheckLotStatus()
   sSql = "SELECT RUNREF,RUNNO,RUNSTATUS,RUNCLOSED,PARTREF," _
          & "PALOTTRACK FROM RunsTable,PartTable WHERE RUNREF=PARTREF " _
          & "AND (RUNCLOSED IS NULL OR RUNCLOSED > '" & cboAsOf & "') " _
          & "AND RUNSTATUS<>'CA'"
   bSqlRows = GetDataSet(RdoMos, ES_FORWARD)
   If bSqlRows Then
      With RdoMos
         Do Until .EOF
            lTotalRuns = lTotalRuns + 1
            If bLots = 1 And !PALOTTRACK = 1 Then bShowLots = 1 _
                       Else bShowLots = 0
            sSql = "INSERT INTO " & sTempTable & " " _
                   & "(MOROWNUMBER,MOPARTREF,MORUNNO,MORUNSTATUS,MOLOTTRACK) " _
                   & "VALUES(" & lTotalRuns & ",'" & Trim(!RUNREF) & "'," & !RunNo _
                   & ",'" & Trim(!RUNSTATUS) & "'," & bShowLots & ")"
            RdoCon.Execute sSql, rdExecDirect
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If lTotalRuns Then GetWipRuns = 1 Else GetWipRuns = 0
   lblRuns = lTotalRuns
   lblRuns.Refresh
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   'Redundant because closing the RdoCon for the section drops these.
   'Reduces server clutter.
   sSql = "DROP TABLE " & sTempTable
   RdoCon.Execute sSql, rdExecDirect
   
   Set RdoWip = Nothing
   Set WipTest = Nothing
   
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub



Public Sub BuildTempTable()
   On Error Resume Next
   Err = 0
   On Error GoTo DiaErr1
   sTempTable = Compress(GetNextLotNumber())
   sTempTable = "##" & Right$(sTempTable, 8)
   sSql = "CREATE TABLE " & sTempTable & " (" _
          & "MOROWNUMBER INT NULL DEFAULT(1)," _
          & "MOPARTREF CHAR(30) NULL DEFAULT ('')," _
          & "MORUNNO INTEGER NULL DEFAULT(0)," _
          & "MORUNSTATUS CHAR(2) NULL DEFAULT('')," _
          & "MOLOTTRACK TINYINT NULL DEFAULT(0))"
   RdoCon.Execute sSql, rdExecDirect
   If Err = 0 Then
      sSql = "CREATE UNIQUE CLUSTERED INDEX MOCLOSE ON " & sTempTable & " " _
             & "(MOPARTREF,MORUNNO) WITH  FILLFACTOR = 80"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "CREATE INDEX MOROW ON " & sTempTable & " " _
             & "(MOROWNUMBER) WITH  FILLFACTOR = 80"
      RdoCon.Execute sSql, rdExecDirect
      Label2 = sTempTable
      Label2.Refresh
   End If
   Exit Sub
   
   DiaErr1:
   sTempTable = ""
   
End Sub


'Added some columns to the table

Private Sub CheckReportTable()
   On Error Resume Next
   Err = 0
   sSql = "SELECT WIPRUNSTATUS FROM EsReportWIP"
   RdoCon.Execute sSql, rdExecDirect
   If Err > 0 Then
      Err = 0
      sSql = "DROP TABLE EsReportWIP"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "CREATE TABLE EsReportWIP (" _
             & "WIPRUNREF CHAR(30) NULL DEFAULT('')," _
             & "WIPRUNNO INT NULL DEFAULT(0)," _
             & "WIPRUNSTATUS CHAR(2) NULL DEFAULT('')," _
             & "WIPCOSTTYPE CHAR(3) NULL DEFAULT('')," _
             & "WIPLABOR REAL NULL DEFAULT(0)," _
             & "WIPMISSTIME TINYINT NULL DEFAULT(0)," _
             & "WIPMATL REAL NULL DEFAULT(0)," _
             & "WIPMISSMATL TINYINT NULL DEFAULT(0)," _
             & "WIPOH REAL NULL DEFAULT(0)," _
             & "WIPEXP REAL NULL DEFAULT(0)," _
             & "WIPMISSEXP TINYINT NULL DEFAULT(0)," _
             & "WIPFREIGHT REAL NULL DEFAULT(0)," _
             & "WIPTAX REAL NULL DEFAULT(0)," _
             & "WIPUNCOSTED TINYINT DEFAULT(0))"
      RdoCon.Execute sSql, rdExecDirect
      
      sSql = "CREATE UNIQUE CLUSTERED INDEX WipReport ON EsReportWIP " _
             & "(WIPRUNREF,WIPRUNNO) WITH  FILLFACTOR = 80"
      RdoCon.Execute sSql, rdExecDirect
   End If
   
   
End Sub

Public Sub BuildReport()
   Dim a As Integer
   Dim cCounter As Currency
   Dim cValue As Currency
   Dim lList As Long
   
   sProcName = "collectruns"
   Prg1.Visible = True
   CollectRuns
   On Error GoTo DiaErr1
   cValue = 100 / lTotalRuns
   a = 5
   Prg1.Value = a
   MouseCursor 13
   For lList = 1 To lTotalRuns
      sSql = "SELECT MOPARTREF,MORUNNO,MOLOTTRACK FROM " & sTempTable & " " _
             & "WHERE MOROWNUMBER=" & lList & " "
      bSqlRows = GetDataSet(RdoWip, ES_FORWARD)
      If bSqlRows Then
         cCounter = cCounter + cValue
         If cCounter >= 5 Then
            a = a + cCounter
            If a > 95 Then a = 95
            Prg1.Value = a
            cCounter = 0
         End If
         sProcName = "getuninvoicedpo" '
         GetUnInvoicedPoItems Trim(RdoWip!MOPARTREF), RdoWip!MORUNNO
         
         sProcName = "getpicklist" '
         GetPickList Trim(RdoWip!MOPARTREF), RdoWip!MORUNNO
         
         sProcName = "getexpensecos" '
         GetExpenseCosts Trim(RdoWip!MOPARTREF), RdoWip!MORUNNO
         
         sProcName = "getlaborcosts" '
         GetLaborCosts Trim(RdoWip!MOPARTREF), RdoWip!MORUNNO
         
         sProcName = "getmaterialco"
         GetMaterialCosts Trim(RdoWip!MOPARTREF), RdoWip!MORUNNO, RdoWip!MOLOTTRACK
      End If
   Next
   Prg1.Value = 100
   MouseCursor 0
   'MsgBox "Completed.", _
   '    vbInformation, Caption
   Prg1.Visible = False
   Exit Sub
   
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.description = Err.description
   DoModuleErrors Me
   
End Sub

'Uninvoiced PO Items marked for expensed, but use as desired

Private Sub GetUnInvoicedPoItems(PartNumber As String, RunNumber As Long)
   Dim RdoInv As rdoResultset
   Dim bUninvoiced As Byte
   
   sSql = "SELECT PINUMBER,PITYPE,PIITEM,PIREV,PIPART,PIRUNPART,PIRUNNO,PIAQTY," _
          & " PARTREF,PARTNUM FROM PoitTable,PartTable " _
          & "WHERE (PIRUNPART='" & PartNumber & "' AND PIRUNNO=" _
          & RunNumber & " AND PIAQTY=0 AND PITYPE<>16) AND PARTREF=PIPART"
   bSqlRows = GetDataSet(RdoInv, ES_FORWARD)
   If bSqlRows = 1 Then
      bUninvoiced = 0
      With RdoInv
         Do Until .EOF
            If !PITYPE <> 17 Then bUninvoiced = 1
            .MoveNext
         Loop
         'Expenses?
         sSql = "UPDATE EsReportWIP SET WIPMISSEXP=" & bUninvoiced & " " _
                & "WHERE WIPRUNREF='" & PartNumber & "' AND WIPRUNNO=" _
                & RunNumber & " "
         RdoCon.Execute sSql, rdExecDirect
         .Cancel
      End With
   End If
   
End Sub

'Find Open Pick items

Private Sub GetPickList(PartNumber As String, RunNumber As Long)
   Dim RdoPck As rdoResultset
   Dim bUnpicked As Byte
   
   sSql = "SELECT PKPARTREF,PKTYPE,PKAQTY,PKMOPART,PKMORUN FROM MopkTable " _
          & "WHERE (PKTYPE<>12 AND PKMOPART='" & PartNumber & "' AND " _
          & "PKMORUN=" & RunNumber & ")"
   bSqlRows = GetDataSet(RdoPck, ES_FORWARD)
   If bSqlRows Then
      With RdoPck
         bUnpicked = 0
         Do Until .EOF
            If !PKTYPE = 9 Or !PKTYPE = 23 Then bUnpicked = 1
            If !PKAQTY = 0 Then bUnpicked = 1
            .MoveNext
         Loop
         'Incomplete Picks if any
         sSql = "UPDATE EsReportWIP SET WIPMISSMATL=" & bUnpicked & " " _
                & "WHERE WIPRUNREF='" & PartNumber & "' AND WIPRUNNO=" _
                & RunNumber & " "
         RdoCon.Execute sSql, rdExecDirect
         .Cancel
      End With
   Else
      
   End If
   
End Sub

'Get Costs, tax, freight and such

Private Sub GetExpenseCosts(PartNumber As String, RunNumber As Long)
   Dim RdoExp As rdoResultset
   Dim cMOEXPENSE As Currency
   Dim cFREIGHT As Currency
   Dim cTAXES As Currency
   
   'Purchased Expense Items
   sSql = "SELECT PINUMBER,PIPART,PIITEM,PIREV,PITYPE,PIAQTY,PIAMT," _
          & "PARTREF,PARTNUM FROM PoitTable,PartTable WHERE (PIRUNPART='" _
          & PartNumber & "' AND PIRUNNO=" & RunNumber & " AND " _
          & "PITYPE=17 AND PALEVEL=7) AND PIPART=PARTREF AND " _
          & "PIPDATE<= '" & cboAsOf & "'"
   bSqlRows = GetDataSet(RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         Do Until .EOF
            cMOEXPENSE = cMOEXPENSE + (!PIAQTY * !PIAMT)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   'Tax and freight
   sSql = "SELECT SUM(VIFREIGHT) AS FREIGHT,SUM(VITAX) AS TAX FROM VihdTable," _
          & "ViitTable WHERE VINO=VITNO AND (VITMO='" _
          & PartNumber & "' AND VITMORUN=" & RunNumber & ")"
   bSqlRows = GetDataSet(RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         If Not IsNull(!FREIGHT) Then
            cFREIGHT = cFREIGHT + !FREIGHT
         End If
         If Not IsNull(!tax) Then
            cTAXES = cTAXES + !tax
         End If
         .Cancel
      End With
   End If
   
   'Invoices without PO's
   sSql = "SELECT SUM(VITQTY*VITCOST) AS SUMCOST FROM ViitTable WHERE " _
          & "(VITPO=0 AND VITPOITEM=0) AND (VITMO='" & PartNumber _
          & "' AND VITMORUN=" & RunNumber & ")"
   bSqlRows = GetDataSet(RdoExp, ES_FORWARD)
   If bSqlRows Then
      With RdoExp
         If Not IsNull(!SUMCOST) Then _
                       cMOEXPENSE = cMOEXPENSE + !SUMCOST
         .Cancel
      End With
   End If
   'expense etc
   sSql = "UPDATE EsReportWIP SET WIPEXP=" & cMOEXPENSE & "," _
          & "WIPFREIGHT=" & cFREIGHT & ",WIPTAX=" & cTAXES & " " _
          & "WHERE WIPRUNREF='" & PartNumber & "' AND WIPRUNNO=" _
          & RunNumber & " "
   RdoCon.Execute sSql, rdExecDirect
   Set RdoExp = Nothing
   
End Sub

'Labor and Overhead

Public Sub GetLaborCosts(PartNumber As String, RunNumber As Long)
   Dim RdoLab As rdoResultset
   Dim cRunHours As Currency
   Dim cRunOvHd As Currency
   Dim cRunLabor As Currency
   
   sSql = "SELECT TCCARD,TCHOURS,TCTIME,TCRATE,TCOHRATE," _
          & "TCPARTREF,TCRUNNO,TMDATE FROM TcitTable,TchdTable WHERE " _
          & "(TCPARTREF='" & PartNumber & "' AND TCRUNNO=" _
          & RunNumber & ") AND TCCARD=TMCARD AND TMDATE<= '" & cboAsOf & "'"
   bSqlRows = GetDataSet(RdoLab, ES_FORWARD)
   If bSqlRows Then
      With RdoLab
         Do Until .EOF
            cRunHours = cRunHours + !TCHOURS
            cRunOvHd = cRunOvHd + (!TCOHRATE * !TCHOURS)
            cRunLabor = cRunLabor + (!TCRATE * !TCHOURS)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   'expense etc (no provison for hours, but maybe should be)
   sSql = "UPDATE EsReportWIP SET WIPLABOR=" & cRunLabor & "," _
          & "WIPOH=" & cRunOvHd & " WHERE WIPRUNREF='" _
          & PartNumber & "' AND WIPRUNNO=" & RunNumber & " "
   RdoCon.Execute sSql, rdExecDirect
   Set RdoLab = Nothing
   
End Sub

Private Sub GetMaterialCosts(PartNumber As String, RunNumber As Long, LOTTRACKED As Byte)
   Dim RdoMat As rdoResultset
   Dim bUncostedMat As Byte
   Dim bUncostedLot As Byte
   Dim cLotCost As Currency
   Dim cQuantity As Currency
   Dim cRunMatl As Currency
   Dim cStdCost As Currency
   
   bUncostedMat = 0
   cRunMatl = 0
   If LOTTRACKED Then
      'lot cost
      'Get uncosted Lots
      sSql = "SELECT LOINUMBER,LOITYPE,LOIMOPARTREF,LOIMORUNNO,LOTNUMBER," _
             & "LOTUNITCOST FROM LoitTable,LohdTable WHERE (LOINUMBER=LOTNUMBER " _
             & "AND LOIMOPARTREF='& PartNumber &' AND LOIMORUNNO=" & RunNumber & ") " _
             & "AND LOITYPE=10 AND LOTUNITCOST=0"
      bSqlRows = GetDataSet(RdoMat, ES_FORWARD)
      If bSqlRows Then
         bUncostedMat = 1
         bUncostedLot = 1
      End If
      RdoMat.Cancel
      
      'Get costed Lots (Picks)
      sSql = "SELECT LOINUMBER,LOITYPE,LOIQUANTITY,LOIMOPARTREF,LOIMORUNNO,LOTNUMBER," _
             & "LOTUNITCOST FROM LoitTable,LohdTable WHERE (LOINUMBER=LOTNUMBER " _
             & "AND LOIMOPARTREF='" & PartNumber & "' AND LOIMORUNNO=" & RunNumber & ") " _
             & "AND LOITYPE=10 AND LOTUNITCOST>0 AND LOIPDATE<= '" & cboAsOf & "'"
      bSqlRows = GetDataSet(RdoMat, ES_FORWARD)
      If bSqlRows Then
         With RdoMat
            Do Until .EOF
               cRunMatl = cRunMatl + (!LOTUNITCOST * !LOIQUANTITY)
               .MoveNext
            Loop
            .Cancel
         End With
      End If
      
      'Get costed Lots (Canceled Picks)
      sSql = "SELECT LOINUMBER,LOITYPE,LOIQUANTITY,LOIMOPARTREF,LOIMORUNNO,LOTNUMBER," _
             & "LOTUNITCOST FROM LoitTable,LohdTable WHERE (LOINUMBER=LOTNUMBER " _
             & "AND LOIMOPARTREF='" & PartNumber & "' AND LOIMORUNNO=" & RunNumber & ") " _
             & "AND LOITYPE=12 AND LOTUNITCOST>0 AND LOIPDATE<= '" & cboAsOf & "'"
      bSqlRows = GetDataSet(RdoMat, ES_FORWARD)
      If bSqlRows Then
         With RdoMat
            Do Until .EOF
               cRunMatl = cRunMatl - (!LOTUNITCOST * !LOIQUANTITY)
               .MoveNext
            Loop
            .Cancel
         End With
         'Could end up negative
         If cRunMatl < 0 Then cRunMatl = 0
      End If
   Else
      'Standard Costs
      sSql = "SELECT PKPARTREF,PKMOPART,PKMORUN,PKPDATE,PKAQTY,PKAMT,PARTREF," _
             & "PASTDCOST FROM MopkTable,PartTable WHERE (PKPARTREF=PARTREF AND " _
             & "PKAQTY>0 AND PKMOPART='" & PartNumber & "' AND PKMORUN=" _
             & RunNumber & ") AND PKPDATE<= '" & cboAsOf & "'"
      bSqlRows = GetDataSet(RdoMat, ES_FORWARD)
      If bSqlRows Then
         With RdoMat
            Do Until .EOF
               If !PKAMT > 0 Then cStdCost = !PKAMT _
                                             Else cStdCost = !PASTDCOST
               If cStdCost = 0 Then bUncostedMat = 1
               cRunMatl = cRunMatl + (cStdCost * !PKAQTY)
               .MoveNext
            Loop
            .Cancel
         End With
      End If
   End If
   sSql = "UPDATE EsReportWIP SET WIPMATL=" & cRunMatl & "," _
          & "WIPMISSMATL=" & bUncostedMat & "," _
          & "WIPUNCOSTED=" & bUncostedLot & " WHERE " _
          & "WIPRUNREF='" & PartNumber & "' AND WIPRUNNO=" _
          & RunNumber & " "
   RdoCon.Execute sSql, rdExecDirect
   Set RdoMat = Nothing
   
End Sub

Private Sub CollectRuns()
   Dim RdoMos As rdoResultset
   Dim sType As String
   
   sSql = "SELECT * FROM " & sTempTable & " "
   bSqlRows = GetDataSet(RdoMos, ES_FORWARD)
   If bSqlRows Then
      With RdoMos
         Do Until .EOF
            If !MOLOTTRACK = 1 Then sType = "LOT" _
                             Else sType = "STD"
            sSql = "INSERT INTO EsReportWIP (WIPRUNREF,WIPRUNNO," _
                   & "WIPRUNSTATUS,WIPCOSTTYPE) VALUES('" _
                   & Trim(!MOPARTREF) & "'," & !MORUNNO & ",'" _
                   & Trim(!MORUNSTATUS) & "','" & sType & "')"
            RdoCon.Execute sSql, rdExecDirect
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   Set RdoMos = Nothing
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   ' Get By Menu Option
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   If Len(Trim(sOptions)) > 0 Then
      optDsc.Value = Mid(sOptions, 1, 1)
      optExt.Value = Mid(sOptions, 2, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub


Private Sub optDis_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub optPrn_Click()
   If DataReady() Then
      PrintReport
   End If
End Sub

Private Sub PrintReport()
   Dim sCustomReport As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   'GenerateWIP cboAsOf
   
   sProcName = "printreport"
   SetMdiReportsize MdiSect, True
   sCustomReport = GetCustomReport("finwip.rpt")
   MdiSect.crw.ReportFileName = sReportPath & sCustomReport
   MdiSect.crw.Formulas(0) = "CompanyName='" & sFacility & "'"
   MdiSect.crw.Formulas(1) = "RequestBy='Requested By: " _
                        & sInitials & "'"
   MdiSect.crw.Formulas(2) = "Title1='Work In Process Inventory As Of " _
                        & cboAsOf & "'"
   MdiSect.crw.Formulas(3) = "Dsc=" & optDsc
   MdiSect.crw.Formulas(4) = "Ext=" & optExt
   'If Trim(cboClass) = "" Then cboClass = "ALL"
   'If Trim(cboCode) = "" Then cboCode = "ALL"
   MdiSect.crw.Formulas(5) = "Title3='Part Class " & cboClass _
                        & " And Product Code " & cboCode & "'"
   sSql = ""
   If UCase(cboClass) = "<BLANK>" Then
      sSql = "{PartTable.PACLASS} = ''"
   ElseIf UCase(cboClass) <> "<ALL>" Then
      sSql = "{PartTable.PACLASS} = '" & Compress(cboClass) & "'"
   End If
   
   If UCase(cboCode) <> "<ALL>" Then
      If Len(sSql) > 0 Then
         sSql = sSql & " AND "
      End If
      If cboCode = "<BLANK>" Then
         sSql = sSql & "{PartTable.PAPRODCODE} = ''"
      Else
         sSql = sSql & "{PartTable.PAPRODCODE} = '" & Compress(cboCode) & "'"
      End If
   End If
   MdiSect.crw.SelectionFormula = sSql
   SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   DiaErr1:
   CurrError.Number = Err.Number
   CurrError.description = Err.description
   DoModuleErrors Me
End Sub
