VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form PackPSp14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventory Available To Ship"
   ClientHeight    =   4530
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "RefreshExcel"
      Height          =   375
      Left            =   6120
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   5520
      TabIndex        =   25
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   3720
      Width           =   3975
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1680
      TabIndex        =   23
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1680
      TabIndex        =   19
      Top             =   1800
      Width           =   4215
      Begin VB.OptionButton optOrderBy 
         Caption         =   "Scheduled Date"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton optOrderBy 
         Caption         =   "Part Number"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PackPSp14a.frx":0000
      Height          =   315
      Left            =   4800
      Picture         =   "PackPSp14a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.CheckBox optDet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   6720
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox optIco 
      Caption         =   "___"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox txtEnd 
      Height          =   288
      Left            =   3600
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSp14a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Leading Chars Or Blank For All"
      Top             =   1080
      Visible         =   0   'False
      Width           =   3012
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "PackPSp14a.frx":0E32
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
         Picture         =   "PackPSp14a.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   1200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4530
      FormDesignWidth =   7260
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   5760
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   1305
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Report By"
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   22
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      Caption         =   "Part Description"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   288
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   4
      Left            =   5400
      TabIndex        =   14
      Top             =   1440
      Width           =   1428
   End
   Begin VB.Label z1 
      Caption         =   "Item Comments"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   288
      Index           =   5
      Left            =   2750
      TabIndex        =   12
      Top             =   1488
      Width           =   828
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipments From"
      Height          =   288
      Index           =   6
      Left            =   240
      TabIndex        =   11
      Top             =   1488
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   288
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   1080
      Width           =   1428
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Tag             =   " "
      Top             =   2400
      Width           =   1425
   End
End
Attribute VB_Name = "PackPSp14a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   '   Dim RdoCmb As ADODB.Recordset
   '   On Error GoTo DiaErr1
   '    sSql = "SELECT DISTINCT CUREF,CUNICKNAME,SOCUST FROM " _
   '        & "CustTable,SohdTable WHERE CUREF=SOCUST"
   '    bSqlRows = clsADOCon.GetDataSet(sSql,RdoCmb, ES_FORWARD)
   '        If bSqlRows Then
   '            With RdoCmb
   '                Do Until .EOF
   '                    AddComboStr cmbCst.hWnd, "" & Trim(!CUNICKNAME)
   '                    .MoveNext
   '                Loop
   '                .Cancel
   '            End With
   '        Else
   '            lblNme = "*** No Customers With SO's Found ***"
   '        End If
   '    Set RdoCmb = Nothing
   '    cmbCst = "ALL"
   '    GetCustomer
   '   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub cmbExport_Click()

   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
    GetSOParts (False)
    ExportPSNotShipped
   
   
End Sub


Private Function ExportPSNotShipped()

   Dim sParts As String
   Dim sCode As String
   Dim sClass As String
   Dim sBuyer As String
   Dim sMbe As String
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEndDate As String
   Dim sFileName As String
   
   On Error GoTo ExportError
   Dim RdoPO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(19) As String
   Dim sCust As String
   
   AddFieldsToExport sFieldsToExport
   
    sSql = "SELECT a.PARTNUM, a.STARTQOH, a.SALESORDERNO,  " & vbCrLf
    sSql = sSql & " a.ITEMNO, a.CUSTNICK, a.SCHEDDTE,  " & vbCrLf
    sSql = sSql & " a.ITCUSTREQ, a.QUANTITY, a.REMAINQOH, a.ITDOLLARS, (a.ITDOLLARS * a.QUANTITY) as EXTENDED_PRICE,a.RUNNO," & vbCrLf
    sSql = sSql & " PshdTable.PSSHIPPRINT, PshdTable.PSINVOICE, " & vbCrLf
    sSql = sSql & "  (CASE WHEN (PshdTable.PSSHIPPRINT = 1 AND PshdTable.PSINVOICE = 0) THEN " & vbCrLf
    sSql = sSql & " '*' ELSE '' END) as PREPACK " & vbCrLf
    sSql = sSql & "FROM PshdTable RIGHT OUTER JOIN " & vbCrLf
    sSql = sSql & "  PsitTable c ON PshdTable.PSNUMBER = c.PIPACKSLIP RIGHT OUTER JOIN " & vbCrLf
    sSql = sSql & "  EsReportPartsAvailable a ON c.PISONUMBER = a.SONUMBER AND c.PISOITEM = a.ITEMNO "


   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      SaveAsExcel RdoPO, sFieldsToExport, sFileName
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If

   Set RdoPO = Nothing
   Exit Function
   
ExportError:
   MouseCursor 0
   cmbExport.Enabled = True
   MsgBox Err.Description
   

End Function


Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "PARTNUM"
   sFieldsToExport(i + 1) = "STARTQOH"
   sFieldsToExport(i + 2) = "SALESORDERNO"
   sFieldsToExport(i + 3) = "ITEMNO"
   sFieldsToExport(i + 4) = "CUSTNICK"
   sFieldsToExport(i + 5) = "SCHEDDTE"
   sFieldsToExport(i + 6) = "ITCUSTREQ"
   sFieldsToExport(i + 7) = "QUANTITY"
   sFieldsToExport(i + 8) = "REMAINQOH"
   sFieldsToExport(i + 9) = "ITDOLLARS"
   sFieldsToExport(i + 10) = "EXTENDED_PRICE"
   sFieldsToExport(i + 11) = "RUNNO"
   sFieldsToExport(i + 12) = "PSSHIPPRINT"
   sFieldsToExport(i + 13) = "PSINVOICE"
   sFieldsToExport(i + 14) = "PREPACK"
   
   
   
End Function

Private Sub cmdRefresh_Click()
    GetSOParts (False)
    
     MsgBox "Refreshed the Data to view from Excel Sheet.", _
    vbInformation, Caption

End Sub

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.FileName
   End If

End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdFnd_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
   ViewParts.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      If cmdHlp Then
         MouseCursor 13
         OpenHelpContext 907
         MouseCursor 0
         cmdHlp = False
      End If
   End If
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then CreateReportTable
   bOnLoad = 0
   MouseCursor 0
   
   Dim bPartSearch As Boolean
   
   bPartSearch = GetPartSearchOption
   SetPartSearchOption (bPartSearch)
   
   If (Not bPartSearch) Then FillPartCombo cmbPrt

   
   cmbPrt = "ALL"
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
   Set PackPSp14a = Nothing
   
End Sub

Private Sub PrintReport()
    'Dim sBegDate As String
    'Dim sEnddate As String
    'Dim sPart As String
    Dim cCRViewer As EsCrystalRptViewer
    Dim sCustomReport As String
    Dim aRptPara As New Collection
    Dim aRptParaType As New Collection
    Dim aFormulaValue As New Collection
    Dim aFormulaName As New Collection
   
   'If IsDate(txtBeg) Then
   '   sBegDate = Format(txtBeg, "yyyy,mm,dd")
   'Else
   '   sBegDate = "1995,01,01"
   'End If
   '
   'If IsDate(txtEnd) Then
   '   sEnddate = Format(txtEnd, "yyyy,mm,dd")
   'Else
   '   sEnddate = "2024,12,31"
   'End If
   '
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowComments"
   aFormulaName.Add "GroupBy"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Part Number(s) " & CStr(cmbPrt & "..., " & txtBeg _
                        & " Through " & txtEnd) & "'")
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optIco.Value
   If Me.optOrderBy(0).Value = True Then aFormulaValue.Add 0 Else aFormulaValue.Add 1
   
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("sleps14")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   'sSql = "{SoitTable.ITSCHED} in Date(" & sBegDate _
   '       & ") to Date(" & sEnddate & ")" _
   '       & " and {SoitTable.ITCANCELED} = 0 and" _
   '       & " {SoitTable.ITPSITEM} = 0 and" _
   '       & " {SoitTable.ITINVOICE} = 0"
   '
   'cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
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

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtBeg = Format(Now, "mm/dd/yy")
   txtEnd = Format(Now + 30, "mm/dd/yy")
   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sOrderBy As String
   
   On Error Resume Next
   If optOrderBy(1).Value = True Then sOrderBy = "1" Else sOrderBy = "0"
   sOptions = Trim(str$(optDet.Value)) & Trim(str$(optIco.Value)) & sOrderBy
   SaveSetting "Esi2000", "EsiSale", "PackPsp14a", sOptions
   
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   sOptions = Trim(GetSetting("Esi2000", "EsiSale", "PackPsp14a", sOptions))
   If Len(Trim(sOptions)) > 0 Then
      optDet.Value = Val(Left(sOptions, 1))
      optIco.Value = Val(Right(sOptions, 1))
      If Len(Trim(sOptions)) > 2 Then optOrderBy(Val(Mid(sOptions, 3, 1))).Value = True Else optOrderBy(0).Value = True
   Else
      optDet.Value = vbChecked
      optIco.Value = vbChecked
      optOrderBy(0).Value = True
   End If
   On Error Resume Next
   
End Sub

Private Sub optDet_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optDis_Click()
   GetSOParts
   
End Sub


Private Sub optIco_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optPrn_Click()
   GetSOParts
   
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtBeg_LostFocus()
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If txtBeg <> "ALL" Then txtBeg = CheckDate(txtBeg)
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If txtEnd <> "ALL" Then txtEnd = CheckDate(txtEnd)
   
End Sub

Private Sub GetSOParts(Optional bPrint = True)
   Dim RdoRpt As ADODB.Recordset
   Dim RdoAvl As ADODB.Recordset
   Dim iCounter As Integer
   Dim sPartNumber As String
   Dim sCurrentPart As String
   Dim ht As New HashTable
   Dim cStrtQOH, cRemainQOH, cQuantity As Currency
   Dim strCmt As String
   
   Dim sBegDate As String
   Dim sEndDate As String
   
   If IsDate(txtBeg) Then
      sBegDate = Format(txtBeg, "mm/dd/yyyy")
   Else
      sBegDate = "01/01/1995"
   End If
   If IsDate(txtEnd) Then
      sEndDate = Format(txtEnd, "mm/dd/yyyy")
   Else
      sEndDate = "12/31/2024"
   End If
   
   If cmbPrt <> "ALL" Then sPartNumber = Compress(cmbPrt) Else sPartNumber = ""
   
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "TRUNCATE TABLE EsReportPartsAvailable"
   clsADOCon.ExecuteSql sSql 'rdExecDirect
   iCounter = 0
   sSql = "SELECT ITPART, PARTREF, PARTNUM, PADESC, PAQOH, SOTYPE, SOTEXT, SOTYPE, ITNUMBER, ITCUSTITEMNO, ITREV, " _
          & " CUNICKNAME, ITSCHED, ITCUSTREQ, ITQTY, ITCOMMENTS, PALOCATION, PARUN, ITSO,ITDOLLARS FROM SoitTable " _
          & " INNER JOIN SohdTable ON SoitTable.ITSO = SohdTable.SONUMBER " _
          & " INNER JOIN PartTable ON PartTable.PARTREF = SoitTable.ITPART " _
          & " INNER JOIN CustTable ON CustTable.CUREF = SohdTable.SOCUST " _
          & " WHERE " _
          & " ITSCHED BETWEEN '" & sBegDate & "' AND '" & sEndDate & "' AND " _
          & " PAQOH>0 AND (ITPSITEM + ITCANCELED + ITINVOICE) = 0 AND " _
          & " PARTREF LIKE '" & sPartNumber & "%' ORDER BY "
          If optOrderBy(0).Value = True Then sSql = sSql & "PARTREF, ITSCHED" Else sSql = sSql & "ITSCHED, PARTREF "
   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAvl, ES_FORWARD)
   clsADOCon.ADOErrNum = 0
   ht.SetSize 10000
   If bSqlRows Then
      With RdoAvl
'         sSql = "SELECT * FROM EsReportPartsAvailable"
'         bSqlRows = clsADOCon.GetDataSet(sSql,RdoRpt, ES_KEYSET)
         Do Until .EOF
            iCounter = iCounter + 1
            
'            RdoRpt.AddNew
'            RdoRpt!Counter = iCounter
             sCurrentPart = "" & Trim(!ITPART)
'            RdoRpt!PartRef = sCurrentPart
'            RdoRpt!PARTNUM = "" & Trim(!PARTNUM)
'            RdoRpt!PADESC = "" & Trim(!PADESC)
'            RdoRpt!PAQOH = "" & Trim(!PAQOH)
'            RdoRpt!SalesOrderNo = "" & Trim(!SOTYPE) & Trim(!SOTEXT)
'            RdoRpt!ItemNo = "" & Trim(!ITNUMBER)
'            RdoRpt!ItemRev = "" & Trim(!ITREV)
'            RdoRpt!CustNick = "" & Trim(!CUNICKNAME)
'            RdoRpt!SchedDte = "" & Trim(!ITSCHED)
            cQuantity = CCur("" & Trim(!ITQty))
'            RdoRpt!Quantity = cQuantity

            If ht.Exists(sCurrentPart) Then
                cStrtQOH = CCur(ht(sCurrentPart))
                cRemainQOH = cStrtQOH - cQuantity
                ht(sCurrentPart) = CVar(cRemainQOH)
            Else
                cStrtQOH = CCur("" & !PAQOH)
                cRemainQOH = cStrtQOH - cQuantity
                ht.Add sCurrentPart, CVar(cRemainQOH)
            End If
            
'            RdoRpt!QOH = cQOH
'            'RdoRpt!Comments = "" & !ITCOMMENTS
'            RdoRpt.Update

            strCmt = ReplaceSingleQuote(Trim(!ITCOMMENTS))

            sSql = "INSERT INTO EsReportPartsAvailable (COUNTER, PARTREF, PARTNUM, PADESC, STARTQOH, SALESORDERNO, ITEMNO, CUSTITEMNO, ITEMREV, CUSTNICK, SCHEDDTE, ITCUSTREQ, QUANTITY, REMAINQOH, LOCATION, RUNNO, SONUMBER, ITDOLLARS,COMMENTS) " & _
                   " VALUES (" & Trim(iCounter + 1) & ", '" & sCurrentPart & "','" & "" & Trim(!PartNum) & "','" & "" & Trim(!PADESC) & "'," & _
                   "" & Trim(cStrtQOH) & ",'" & "" & Trim(!SOTYPE) & Trim(!SOTEXT) & "','" & "" & Trim(!ITNUMBER) & "','" & "" & Trim(!ITCUSTITEMNO) & "','" & "" & Trim(!itrev) & "','" & _
                   "" & Trim(!CUNICKNAME) & "','" & "" & Trim(!itsched) & "','" & "" & Trim(!itcustreq) & _
                   "'," & cQuantity & "," & cRemainQOH & ",'" & "" & Trim(!PALOCATION) & "'," & "" & Trim(!PARUN) & "," & Trim(!itso) & "," & Trim(!ITDOLLARS) & ",'" & Trim(strCmt) & "')"
            'Debug.Print sSql
            clsADOCon.ExecuteSql sSql 'rdExecDirect
            
            .MoveNext
         Loop
      End With
      sPartNumber = "1"
   Else
      MouseCursor 0
      MsgBox "No Matching Data To Report.", _
         vbInformation, Caption
      sPartNumber = ""
   End If
   Set RdoRpt = Nothing
   Set RdoAvl = Nothing
   ht.RemoveAll
   Set ht = Nothing
   
   If (bPrint = True) Then
    If sPartNumber <> "" Then PrintReport
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "getsoparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CreateReportTable()
   On Error Resume Next
   'sSql = "SELECT PARTREF FROM EsReportPartsAvailable WHERE PARTREF='FUBAR'"
   'clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
   If TableExists("EsReportPartsAvailable") Then
        sSql = "DROP TABLE EsReportPartsAvailable"
        clsADOCon.ExecuteSql sSql 'rdExecDirect
   End If
   
   
   'If Err > 0 Then
      sSql = "CREATE TABLE EsReportPartsAvailable (" _
             & "COUNTER INT NOT NULL, " _
             & "PARTREF CHAR(30) NOT NULL, " _
             & "PARTNUM CHAR(30) NULL DEFAULT(''), " _
             & "PADESC CHAR(30) NULL DEFAULT(''), " _
             & "STARTQOH DEC(12,4) NULL DEFAULT(0), " _
             & "SALESORDERNO CHAR(7) NULL DEFAULT(''), " _
             & "SONUMBER INT NOT NULL, " _
             & "ITEMNO INT NOT NULL, " _
             & "CUSTITEMNO VARCHAR(15) NULL DEFAULT(''), " _
             & "ITEMREV CHAR(3) NULL DEFAULT(''), " _
             & "CUSTNICK CHAR(10) NULL DEFAULT(''), " _
             & "SCHEDDTE SMALLDATETIME NULL, " _
             & "ITCUSTREQ smalldatetime null, " _
             & "QUANTITY DEC(12,4) NOT NULL, " _
             & "REMAINQOH DEC(12,4) NOT NULL, " _
             & "LOCATION CHAR(4) NULL DEFAULT(''), " _
             & "RUNNO INT NOT NULL, " _
             & "ITDOLLARS DEC(12,4) NOT NULL, " _
             & "COMMENTS VARCHAR(3072) NULL DEFAULT('') )"
      clsADOCon.ExecuteSql sSql 'rdExecDirect
      
      'sSql = "ALTER TABLE EsReportPartsAvailable ADD Constraint PK_EsReportPartsAvailable_PARTREF PRIMARY KEY CLUSTERED (PARTREF) " _
      '       & "WITH FILLFACTOR=80 "
      'clsADOCon.ExecuteSQL sSql 'rdExecDirect
   'End If
   
End Sub

Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If txtPrt = "" Then txtPrt = "ALL"
   cmbPrt = txtPrt
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt = "" Then cmbPrt = "ALL"
   
End Sub

Private Sub txtPrt_Change()
   cmbPrt = txtPrt
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

