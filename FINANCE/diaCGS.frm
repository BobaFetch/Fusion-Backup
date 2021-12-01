VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaCGS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cost Of Goods Sold (Report)"
   ClientHeight    =   4665
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CheckBox optSORet 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox optSum 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox optSO 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.CheckBox optExt 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Tag             =   "4"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4800
      TabIndex        =   12
      Top             =   360
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaCGS.frx":0000
      PictureDn       =   "diaCGS.frx":0146
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   16
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
      PictureUp       =   "diaCGS.frx":028C
      PictureDn       =   "diaCGS.frx":03D2
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   5040
      Top             =   2400
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
      Index           =   8
      Left            =   240
      TabIndex        =   25
      Top             =   3720
      Width           =   1305
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show SO returns"
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   2640
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Summary Journal Only"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Numbers"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Descriptions"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descriptions"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   960
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1545
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaCGS"
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

'*********************************************************************************
' diaCGS - Cost Of Goods Sold
'
' Notes:
'
' Created: 02/18/04 (nth)
' Revisions:
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbExport_Click()
    If (txtFilePath <> "") Then
        ExportCGS
    Else
        MsgBox "Please Select The FileName", vbOKOnly
    End If
    
End Sub

Private Function ExportCGS()

   Dim sParts As String
   Dim sCode As String
   Dim sClass As String
   Dim sBuyer As String
   Dim sMbe As String
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEnddate As String
   Dim sFileName As String
   
   On Error GoTo ExportError

   Dim rdoPo As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(33) As String
   
   AddFieldsToExport sFieldsToExport
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBDate = "01/01/2000"
   Else
      sBDate = Format(txtBeg, "mm/dd/yyyy")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "12/31/2024"
   Else
      sEDate = Format(txtEnd, "mm/dd/yyyy")
   End If

    sSql = "SELECT CihdTable.INVNO, CihdTable.INVPRE, CihdTable.INVTYPE, CihdTable.INVTOTAL, " & vbCrLf
    sSql = sSql & " CihdTable.INVDATE, CihdTable.INVCUST, CihdTable.INVCANCELED,SoitTable.ITNUMBER, " & vbCrLf
    sSql = sSql & " SoitTable.ITREV, SoitTable.ITQTY, SoitTable.ITACTUAL, SoitTable.ITPSNUMBER, " & vbCrLf
    sSql = sSql & " SoitTable.ITPSITEM, SoitTable.ITCANCELED,InvaTable.INTYPE, InvaTable.INADATE, " & vbCrLf
    sSql = sSql & " InvaTable.INAQTY, InvaTable.INTOTMATL, InvaTable.INTOTLABOR, InvaTable.INTOTEXP, " & vbCrLf
    sSql = sSql & " InvaTable.INTOTOH, InvaTable.INPSNUMBER, InvaTable.INPSITEM,PartTable.PARTNUM, " & vbCrLf
    sSql = sSql & " PartTable.PADESC, ISNULL(PartTable.PAEXTDESC,'') AS PAEXTDESC, PartTable.PALEVEL, PartTable.PAUNITS, " & vbCrLf
    sSql = sSql & " PartTable.PALOTTRACK, PartTable.PAUSEACTUALCOST,SohdTable.SONUMBER, SohdTable.SOTYPE," & vbCrLf
    sSql = sSql & "    PsitTable.PIQTY,PshdTable.PSCANCELED" & vbCrLf
    sSql = sSql & "FROM" & vbCrLf
    sSql = sSql & "    (((((CihdTable CihdTable INNER JOIN SoitTable SoitTable ON" & vbCrLf
    sSql = sSql & "        CihdTable.INVNO = SoitTable.ITINVOICE)" & vbCrLf
    sSql = sSql & "     INNER JOIN PartTable PartTable ON" & vbCrLf
    sSql = sSql & "        SoitTable.ITPART = PartTable.PARTREF)" & vbCrLf
    sSql = sSql & "     INNER JOIN SohdTable SohdTable ON" & vbCrLf
    sSql = sSql & "        SoitTable.ITSO = SohdTable.SONUMBER)" & vbCrLf
    sSql = sSql & "     INNER JOIN InvaTable InvaTable ON" & vbCrLf
    sSql = sSql & "        SoitTable.ITSO = InvaTable.INSONUMBER AND" & vbCrLf
    sSql = sSql & "    SoitTable.ITNUMBER = InvaTable.INSOITEM AND" & vbCrLf
    sSql = sSql & "    SoitTable.ITREV = InvaTable.INSOREV)" & vbCrLf
    sSql = sSql & "     LEFT OUTER JOIN PsitTable PsitTable ON" & vbCrLf
    sSql = sSql & "        InvaTable.INPSNUMBER = PsitTable.PIPACKSLIP AND" & vbCrLf
    sSql = sSql & "    InvaTable.INPSITEM = PsitTable.PIITNO AND" & vbCrLf
    sSql = sSql & "    InvaTable.INPART = PsitTable.PIPART AND" & vbCrLf
    sSql = sSql & "    InvaTable.INSONUMBER = PsitTable.PISONUMBER AND" & vbCrLf
    sSql = sSql & "    InvaTable.INSOITEM = PsitTable.PISOITEM AND" & vbCrLf
    sSql = sSql & "    InvaTable.INSOREV = PsitTable.PISOREV)" & vbCrLf
    sSql = sSql & "     INNER JOIN PshdTable PshdTable ON" & vbCrLf
    sSql = sSql & "        PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf
    sSql = sSql & "WHERE" & vbCrLf
    sSql = sSql & "    (InvaTable.INTYPE = 4 OR" & vbCrLf
    sSql = sSql & "    InvaTable.INTYPE = 3 OR" & vbCrLf
    sSql = sSql & "    InvaTable.INTYPE = 26 OR" & vbCrLf
    sSql = sSql & "    InvaTable.INTYPE = 25 OR" & vbCrLf
    sSql = sSql & "    InvaTable.INTYPE = 24) AND" & vbCrLf
    sSql = sSql & "    PshdTable.PSCANCELED = 0 AND" & vbCrLf
    sSql = sSql & "    CihdTable.INVCANCELED = 0 AND" & vbCrLf
    sSql = sSql & "    SoitTable.ITCANCELED = 0" & vbCrLf
    sSql = sSql & "    AND invdate between '" & sBDate & "' and '" & sEDate & "'" & vbCrLf
    sSql = sSql & "    --and INVNO  like '%272611%'" & vbCrLf
    sSql = sSql & "ORDER BY" & vbCrLf
    sSql = sSql & "    CihdTable.INVNO ASC"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPo, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      'SaveAsExcel rdoPo, sFieldsToExport, sFileName
      SaveAsExcelSupDup rdoPo, sFieldsToExport, sFileName, True, 1, 4
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If

   Set rdoPo = Nothing
   Exit Function
   
ExportError:
   MouseCursor 0
   cmbExport.enabled = True
   MsgBox Err.Description
   

End Function

'*********************************************************************************

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   'fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.Filter = "Excel File (*.xlsx) | *.xlsx"
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.FileName
   End If

End Sub

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   'txtEnd = Format(GetServerDateTime(), "mm/dd/yy")
   'txtBeg = Format(txtEnd, "mm/01/yy")
   'optPrn.Picture = Resources.imgPrn.Picture
   'optDis.Picture = Resources.imgDis.Picture
   GetOptions
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaCGS = Nothing
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
   Dim b As Byte
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   sCustomReport = GetCustomReport("fincgs.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   aFormulaName.Add "Dsc"
   aFormulaName.Add "Ext"
   aFormulaName.Add "SO"
   aFormulaName.Add "SORet"
    
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'From " & CStr(txtBeg & " Through " & txtEnd) & "'")
   aFormulaValue.Add optDsc
   aFormulaValue.Add optExt
   aFormulaValue.Add optSO
   aFormulaValue.Add optSORet
   
   sSql = "{CihdTable.INVDATE} >= cdate('" & txtBeg _
          & "') and {CihdTable.INVDATE} <= cdate('" & txtEnd & "')"
        sSql = sSql & " and {InvaTable.INTYPE} in [24.00, 25.00, 26.00, 3.00, 4.00] and " _
                    & "{PshdTable.PSCANCELED} = 0 and " _
                    & "{CihdTable.INVCANCELED} = 0 and " _
                    & "{SoitTable.ITPSNUMBER} = {InvaTable.INPSNUMBER} and " _
                    & "{SoitTable.ITCANCELED} = 0.00"
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printrep"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   sOptions = Trim(txtBeg.Text) & Trim(txtEnd.Text) & Trim(optDsc.Value) & Trim(optExt.Value) _
              & Trim(optSO.Value) & Trim(optSum.Value)
   SaveSetting "Esi2000", "EsiFina", Me.Name, sOptions
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   Dim dToday As Integer
   On Error Resume Next
   dToday = CInt(Mid(Format(Now, "mm/dd/yy"), 4, 2))
   sOptions = GetSetting("Esi2000", "EsiFina", Me.Name, sOptions)
   
   If Len(Trim(sOptions)) > 0 Then
     If dToday < 21 Then
      txtBeg = Mid(sOptions, 1, 8)
      txtEnd = Mid(sOptions, 9, 8)
     Else
      txtBeg = Format(Now, "mm/01/yy")
      txtEnd = GetMonthEnd(txtBeg)
     End If
      optDsc.Value = Mid(sOptions, 17, 1)
      optExt.Value = Mid(sOptions, 18, 1)
      optSO.Value = Mid(sOptions, 19, 1)
      optSum.Value = Mid(sOptions, 20, 1)
   End If
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
End Sub

Private Sub txtBeg_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtBeg_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtBeg_LostFocus()
   txtBeg = CheckDate(txtBeg)
End Sub

Private Sub txtend_DropDown()
   ShowCalendar Me
End Sub

Private Sub txtEnd_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtEnd_LostFocus()
   txtEnd = CheckDate(txtEnd)
End Sub


Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "INVNO"
   sFieldsToExport(i + 1) = "INVPRE"
   sFieldsToExport(i + 2) = "INVTYPE"
   sFieldsToExport(i + 3) = "INVTOTAL"
   sFieldsToExport(i + 4) = "INVDATE"
   sFieldsToExport(i + 5) = "INVCUST"
   sFieldsToExport(i + 6) = "INVCANCELED"
   sFieldsToExport(i + 7) = "ITNUMBER"
   sFieldsToExport(i + 8) = "ITREV"
   sFieldsToExport(i + 9) = "ITQTY"
   sFieldsToExport(i + 10) = "ITACTUAL"
   sFieldsToExport(i + 11) = "ITPSNUMBER"
   sFieldsToExport(i + 12) = "ITPSITEM"
   sFieldsToExport(i + 13) = "ITCANCELED"
   sFieldsToExport(i + 14) = "INTYPE"
   sFieldsToExport(i + 15) = "INADATE"
   sFieldsToExport(i + 16) = "INAQTY"
   sFieldsToExport(i + 17) = "INTOTMATL"
   sFieldsToExport(i + 18) = "INTOTLABOR"
   sFieldsToExport(i + 19) = "INTOTEXP"
   sFieldsToExport(i + 20) = "INTOTOH"
   sFieldsToExport(i + 21) = "INPSNUMBER"
   sFieldsToExport(i + 22) = "INPSITEM"
   sFieldsToExport(i + 23) = "PARTNUM"
   sFieldsToExport(i + 24) = "PADESC"
   sFieldsToExport(i + 25) = "PAEXTDESC"
   sFieldsToExport(i + 26) = "PALEVEL"
   sFieldsToExport(i + 27) = "PAUNITS"
   sFieldsToExport(i + 28) = "PALOTTRACK"
   sFieldsToExport(i + 29) = "PAUSEACTUALCOST"
   sFieldsToExport(i + 30) = "SONUMBER"
   sFieldsToExport(i + 31) = "SOTYPE"
   sFieldsToExport(i + 32) = "PIQTY"
   sFieldsToExport(i + 33) = "PSCANCELED"

End Function

