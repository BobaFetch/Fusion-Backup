VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form MrplMRp08 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP Open Orders"
   ClientHeight    =   5205
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   42
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   4080
      Width           =   4695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   6480
      TabIndex        =   41
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   5280
      TabIndex        =   40
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CheckBox chkType 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   37
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   36
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   35
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   34
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   33
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   4560
      TabIndex        =   32
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "7"
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   31
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox chkType 
      Caption         =   "8"
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   30
      Top             =   2760
      Value           =   1  'Checked
      Width           =   435
   End
   Begin VB.CheckBox optVew 
      Caption         =   "View"
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbPart 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton ShowPrinters 
      Height          =   250
      Left            =   360
      Picture         =   "MrplMRp08.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Show Printers"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MrplMRp08.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame z2 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
      Begin VB.OptionButton optMbe 
         Caption         =   "ALL"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   9
         Top             =   200
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "E"
         Height          =   195
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "B"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton optMbe 
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Product Code From List"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cmbCls 
      Height          =   315
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select Product Class From List"
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   3240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6720
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "MrplMRp08.frx":0938
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "MrplMRp08.frx":0AB6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6840
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5205
      FormDesignWidth =   7935
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6840
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7560
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   43
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   288
      Index           =   3
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   1308
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Types"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   38
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   720
      TabIndex        =   28
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   13
      Left            =   5520
      TabIndex        =   25
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   11
      Left            =   5520
      TabIndex        =   24
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank for All)"
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Through"
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy, Either"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   2025
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Classes"
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   19
      Top             =   960
      Width           =   945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Codes"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Include:"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Comment"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   1785
   End
   Begin VB.Label p 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Numbers"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   525
      Width           =   1425
   End
End
Attribute VB_Name = "MrplMRp08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'5/19/06 Revised report and selections. Removed extra report.
Option Explicit
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Dim iOrder As Integer
Dim strBomRev As String
Dim sIns As String

'Least to greatest dates 10/12/01

Private Sub GetMRPDates()
   
   Dim RdoDte As ADODB.Recordset
    sSql = "SELECT MIN(MRP_PARTDATERQD) FROM MrplTable WHERE " _
           & "MRP_TYPE>" & MRPTYPE_BeginningBalance
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not ISNULL(.Fields(0)) Then
            txtBeg = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtBeg = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtBeg.ToolTipText = "Earliest Date By Default"
   
   sSql = "SELECT MAX(MRP_PARTDATERQD) FROM MrplTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDte, ES_FORWARD)
   If bSqlRows Then
      With RdoDte
         If Not ISNULL(.Fields(0)) Then
            txtEnd = Format(.Fields(0), "mm/dd/yyyy")
         Else
            txtEnd = Format(ES_SYSDATE, "mm/dd/yyyy")
         End If
         ClearResultSet RdoDte
      End With
   End If
   txtEnd.ToolTipText = "Latest Date By Default"
   Set RdoDte = Nothing
End Sub

Private Sub cmbCde_LostFocus()
    cmbCde = CheckLen(cmbCde, 6)
    If cmbCde = "" Then cmbCde = "ALL"

End Sub


Private Sub cmbCls_LostFocus()
    cmbCls = CheckLen(cmbCls, 6)
    If cmbCls = "" Then cmbCls = "ALL"

End Sub


Private Sub cmbExport_Click()

   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
   'CreateAllOpenOrders

   ExportOpenOrders
   
   
End Sub

Private Function ExportOpenOrders()

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
   GetMRPCreateDates sBegDate, sEndDate

   Dim RdoPO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(60) As String
   
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

   If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
   If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
   

    sSql = "SELECT dbo.MrplTable.MRP_PARTREF, dbo.MrplTable.MRP_PARTQTYRQD, dbo.MrplTable.MRP_PARTDATERQD, dbo.SohdTable.SOCUST, dbo.SoitTable.ITSO," & vbCrLf
    sSql = sSql & "  dbo.SoitTable.ITNUMBER , dbo.SoitTable.ITCUSTREQ, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, " & vbCrLf
    sSql = sSql & "  ISNULL(dbo.SoitTable.ITQTY * dbo.SoitTable.ITDOLLARS, 0) AS ExtendedCost, dbo.SoitTable.ITSCHED, dbo.SohdTable.SOPO, dbo.SohdTable.SOTYPE," & vbCrLf
    sSql = sSql & "  dbo.RunsTable.RUNCOMMENTS, PartTable.PAMAKEBUY, PartTable.PAQOH, PartTable.PALOCATION, dbo.SohdTable.SOSHIPDATE," & vbCrLf
    sSql = sSql & "  dbo.RunsTable.RUNREF, dbo.RunsTable.RUNNO, dbo.RunsTable.RUNSCHED, dbo.RunsTable.RUNQTY, dbo.RunsTable.RUNSTATUS," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO1, dbo.vw_RnopPivot.OPNO2, dbo.vw_RnopPivot.OPNO3, dbo.vw_RnopPivot.OPNO4, dbo.vw_RnopPivot.OPNO5," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO6, dbo.vw_RnopPivot.OPNO7, dbo.vw_RnopPivot.OPNO8, dbo.vw_RnopPivot.OPNO9, dbo.vw_RnopPivot.OPNO10," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO11, dbo.vw_RnopPivot.OPNO12, dbo.vw_RnopPivot.OPNO13, dbo.vw_RnopPivot.OPNO14, dbo.vw_RnopPivot.OPNO15," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO16, dbo.vw_RnopPivot.OPNO17, dbo.vw_RnopPivot.OPNO18, dbo.vw_RnopPivot.OPNO20, dbo.vw_RnopPivot.OPNO19," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO21, dbo.vw_RnopPivot.OPNO22, dbo.vw_RnopPivot.OPNO23, dbo.vw_RnopPivot.OPNO24, dbo.vw_RnopPivot.OPNO25," & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPNO26 , dbo.vw_RnopPivot.OPNO27" & vbCrLf
    sSql = sSql & "FROM dbo.vw_RnopPivot INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.PartTable AS PartTable INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.RunsTable INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.RnalTable ON dbo.RunsTable.RUNREF = dbo.RnalTable.RAREF AND dbo.RunsTable.RUNNO = dbo.RnalTable.RARUN INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.SoitTable ON dbo.RnalTable.RASO = dbo.SoitTable.ITSO AND dbo.RnalTable.RASOITEM = dbo.SoitTable.ITNUMBER AND" & vbCrLf
    sSql = sSql & "  dbo.RnalTable.RASOREV = dbo.SoitTable.ITREV INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.SohdTable ON dbo.SoitTable.ITSO = dbo.SohdTable.SONUMBER ON PartTable.PARTREF = dbo.SoitTable.ITPART AND" & vbCrLf
    sSql = sSql & "  PartTable.PARTREF = dbo.SoitTable.ITPART ON dbo.vw_RnopPivot.OPREF = dbo.RunsTable.RUNREF AND" & vbCrLf
    sSql = sSql & "  dbo.vw_RnopPivot.OPRUN = dbo.RunsTable.RUNNO INNER JOIN" & vbCrLf
    sSql = sSql & "  dbo.MrplTable ON dbo.RnalTable.RAREF = dbo.MrplTable.MRP_MOREF AND dbo.RnalTable.RARUN = dbo.MrplTable.MRP_MORUNNO" & vbCrLf
    sSql = sSql & "WHERE (dbo.RunsTable.RUNSTATUS <> 'CL') AND (dbo.RunsTable.RUNSTATUS <> 'CO')" & vbCrLf
    sSql = sSql & "    AND (dbo.SoitTable.ITCANCELED = 0) AND (dbo.SoitTable.ITPSNUMBER = '')" & vbCrLf
    sSql = sSql & "    AND (dbo.SoitTable.ITINVOICE = 0) AND (dbo.SoitTable.ITPSSHIPPED = 0)" & vbCrLf
    sSql = sSql & "    AND (dbo.SohdTable.SOSALESMAN LIKE '%%')" & vbCrLf
    sSql = sSql & "    AND (dbo.SoitTable.ITSCHED BETWEEN '" & sBDate & "' AND '" & sEDate & "')" & vbCrLf
    sSql = sSql & "    AND (PartTable.PACLASS LIKE '%" & sClass & "%')" & vbCrLf
    sSql = sSql & "    AND (PartTable.PAPRODCODE LIKE '%" & sCode & "%')"

    
'
'   sSql = "SELECT MrplTable.MRP_PARTREF, MrplTable.MRP_PARTQTYRQD, MrplTable.MRP_PARTDATERQD, dbo.SohdTable.SOCUST, dbo.SoitTable.ITSO," & vbCrLf
'   sSql = sSql & "  dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITCUSTREQ, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS," & vbCrLf
'   sSql = sSql & "  ISNULL(dbo.SoitTable.ITQTY * dbo.SoitTable.ITDOLLARS, 0) AS ExtendedCost, dbo.SoitTable.ITSCHED, dbo.SohdTable.SOPO, dbo.SohdTable.SOTYPE," & vbCrLf
'   sSql = sSql & "  RUNCOMMENTS, PartTable.PAMAKEBUY, PartTable.PAQOH, PartTable.PALOCATION, dbo.SohdTable.SOSHIPDATE," & vbCrLf
'   sSql = sSql & "  dbo.RunsTable.RUNREF, dbo.RunsTable.RUNNO, dbo.RunsTable.RUNSCHED, dbo.RunsTable.RUNQTY, dbo.RunsTable.RUNSTATUS," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO1, dbo.vw_RnopPivot.OPNO2, dbo.vw_RnopPivot.OPNO3, dbo.vw_RnopPivot.OPNO4, dbo.vw_RnopPivot.OPNO5," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO6, dbo.vw_RnopPivot.OPNO7, dbo.vw_RnopPivot.OPNO8, dbo.vw_RnopPivot.OPNO9, dbo.vw_RnopPivot.OPNO10," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO11, dbo.vw_RnopPivot.OPNO12, dbo.vw_RnopPivot.OPNO13, dbo.vw_RnopPivot.OPNO14, dbo.vw_RnopPivot.OPNO15," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO16, dbo.vw_RnopPivot.OPNO17, dbo.vw_RnopPivot.OPNO18, dbo.vw_RnopPivot.OPNO20, dbo.vw_RnopPivot.OPNO19," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO21, dbo.vw_RnopPivot.OPNO22, dbo.vw_RnopPivot.OPNO23, dbo.vw_RnopPivot.OPNO24, dbo.vw_RnopPivot.OPNO25," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO26 , dbo.vw_RnopPivot.OPNO27" & vbCrLf
'   sSql = sSql & "FROM  dbo.RunsTable INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.RnalTable ON dbo.RunsTable.RUNREF = dbo.RnalTable.RAREF AND dbo.RunsTable.RUNNO = dbo.RnalTable.RARUN INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.SoitTable ON dbo.RnalTable.RASO = dbo.SoitTable.ITSO AND dbo.RnalTable.RASOITEM = dbo.SoitTable.ITNUMBER AND" & vbCrLf
'   sSql = sSql & "  dbo.RnalTable.RASOREV = dbo.SoitTable.ITREV INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.SohdTable ON dbo.RnalTable.RASO = dbo.SohdTable.SONUMBER RIGHT OUTER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.MrpbaTable AS MrpbaTable INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.MrplTable AS MrplTable ON MrpbaTable.MRPBOM_PARTREF = MrplTable.MRP_PARTREF INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.PartTable AS PartTable ON MrplTable.MRP_PARTREF = PartTable.PARTREF INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot ON MrplTable.MRP_PARTREF = dbo.vw_RnopPivot.OPREF AND MrplTable.MRP_MORUNNO = dbo.vw_RnopPivot.OPRUN ON" & vbCrLf
'   sSql = sSql & "  dbo.RunsTable.Runno = MrplTable.MRP_MORUNNO And dbo.RunsTable.RUNREF = MrplTable.MRP_PARTREF" & vbCrLf
'   sSql = sSql & "WHERE (MrpbaTable.MRPBOM_LEVEL = 0) AND (MrplTable.MRP_TYPE = 3) AND (dbo.RunsTable.RUNSTATUS <> 'CL') AND" & vbCrLf
'   sSql = sSql & " (MrplTable.MRP_PARTDATERQD BETWEEN '" & sBDate & "' AND '" & sEDate & "') " & vbCrLf
'   sSql = sSql & "  AND (dbo.RunsTable.RUNSTATUS <> 'CO')"
   

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


'Private Function ExportOpenOrders()
'
'   Dim sParts As String
'   Dim sCode As String
'   Dim sClass As String
'   Dim sBuyer As String
'   Dim sMbe As String
'   Dim sBDate As String
'   Dim sEDate As String
'   Dim sBegDate As String
'   Dim sEndDate As String
'   Dim sFileName As String
'
'   On Error GoTo ExportError
'   GetMRPCreateDates sBegDate, sEndDate
'
'   Dim RdoPO As ADODB.Recordset
'   Dim i As Integer
'   Dim sFieldsToExport(60) As String
'
'   AddFieldsToExport sFieldsToExport
'
''   sSql = "SELECT MrplTable.MRP_PARTREF, MrplTable.MRP_PARTQTYRQD, MrplTable.MRP_PARTDATERQD, MrplTable.MRP_SOCUST, MrplTable.MRP_SONUM,MrplTable.MRP_SOITEM," _
''            & " MrplTable.MRP_COMMENT, PartTable.PAMAKEBUY, PartTable.PAQOH, SohdTable.SOSHIPDATE,dbo.RunsTable.RUNREF," _
''            & " dbo.RunsTable.RUNNO, dbo.RunsTable.RUNSCHED, dbo.RunsTable.RUNQTY, dbo.RunsTable.RUNSTATUS," _
''            & " dbo.vw_RnopPivot.OPNO1, dbo.vw_RnopPivot.OPNO2, dbo.vw_RnopPivot.OPNO3, dbo.vw_RnopPivot.OPNO4," _
''            & " dbo.vw_RnopPivot.OPNO5, dbo.vw_RnopPivot.OPNO6, dbo.vw_RnopPivot.OPNO7, dbo.vw_RnopPivot.OPNO8," _
''            & " dbo.vw_RnopPivot.OPNO9, dbo.vw_RnopPivot.OPNO10, dbo.vw_RnopPivot.OPNO11, dbo.vw_RnopPivot.OPNO12," _
''            & " dbo.vw_RnopPivot.OPNO13, dbo.vw_RnopPivot.OPNO14, dbo.vw_RnopPivot.OPNO15, dbo.vw_RnopPivot.OPNO16," _
''            & " dbo.vw_RnopPivot.OPNO17, dbo.vw_RnopPivot.OPNO18, dbo.vw_RnopPivot.OPNO20, dbo.vw_RnopPivot.OPNO19," _
''            & " dbo.vw_RnopPivot.OPNO21, dbo.vw_RnopPivot.OPNO22, dbo.vw_RnopPivot.OPNO23, dbo.vw_RnopPivot.OPNO24," _
''             & " dbo.vw_RnopPivot.OPNO25 , dbo.vw_RnopPivot.OPNO26, dbo.vw_RnopPivot.OPNO27" _
''         & " FROM dbo.vw_RnopPivot INNER JOIN" _
''            & " dbo.RunsTable ON dbo.vw_RnopPivot.OPREF = dbo.RunsTable.RUNREF" _
''            & " AND dbo.vw_RnopPivot.OPRUN = dbo.RunsTable.RUNNO RIGHT OUTER JOIN" _
''             & " dbo.MrpbaTable AS MrpbaTable INNER JOIN" _
''             & " dbo.MrplTable AS MrplTable ON MrpbaTable.MRPBOM_PARTREF = MrplTable.MRP_PARTREF INNER JOIN" _
''             & " dbo.PartTable AS PartTable ON MrplTable.MRP_PARTREF = PartTable.PARTREF ON" _
''             & " dbo.RunsTable.RUNREF = MrpbaTable.MRPBOM_ROOTPARTREF LEFT OUTER JOIN" _
''             & " dbo.SohdTable AS SohdTable ON MrplTable.MRP_SONUM = SohdTable.SONUMBER" _
''         & " where (MrpbaTable.MRPBOM_LEVEL = 0) And (MrplTable.MRP_TYPE = 11)" _
''            & " AND (dbo.RunsTable.RUNSTATUS <> 'CL') AND" _
''             & " (dbo.RunsTable.RUNSTATUS <> 'CO')"
'
'
''   sSql = "SELECT MrplTable.MRP_PARTREF, MrplTable.MRP_PARTQTYRQD, MrplTable.MRP_PARTDATERQD, MrplTable.MRP_SOCUST, MrplTable.MRP_SONUM," & vbCrLf
''   sSql = sSql & "  MrplTable.MRP_SOITEM , dbo.SoitTable.ITCUSTREQ, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS, " & vbCrLf
''   sSql = sSql & "  ISNULL((dbo.SoitTable.ITQTY * dbo.SoitTable.ITDOLLARS), 0) as ExtendedCost, dbo.SoitTable.ITSCHED, " & vbCrLf
''   sSql = sSql & "  SohdTable.SOPO, SohdTable.SOTYPE, MrplTable.MRP_COMMENT, PartTable.PAMAKEBUY, " & vbCrLf
''   sSql = sSql & "  PartTable.PAQOH , PartTable.PALOCATION, SohdTable.SOSHIPDATE, dbo.RunsTable.RUNREF, " & vbCrLf
''   sSql = sSql & "  dbo.RunsTable.RUNNO, dbo.RunsTable.RUNSCHED, dbo.RunsTable.RUNQTY, dbo.RunsTable.RUNSTATUS, dbo.vw_RnopPivot.OPNO1," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO2, dbo.vw_RnopPivot.OPNO3, dbo.vw_RnopPivot.OPNO4, dbo.vw_RnopPivot.OPNO5, dbo.vw_RnopPivot.OPNO6," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO7, dbo.vw_RnopPivot.OPNO8, dbo.vw_RnopPivot.OPNO9, dbo.vw_RnopPivot.OPNO10, dbo.vw_RnopPivot.OPNO11," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO12, dbo.vw_RnopPivot.OPNO13, dbo.vw_RnopPivot.OPNO14, dbo.vw_RnopPivot.OPNO15, dbo.vw_RnopPivot.OPNO16," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO17, dbo.vw_RnopPivot.OPNO18, dbo.vw_RnopPivot.OPNO20, dbo.vw_RnopPivot.OPNO19, dbo.vw_RnopPivot.OPNO21," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO22, dbo.vw_RnopPivot.OPNO23, dbo.vw_RnopPivot.OPNO24, dbo.vw_RnopPivot.OPNO25, dbo.vw_RnopPivot.OPNO26," & vbCrLf
''   sSql = sSql & "  dbo.vw_RnopPivot.OPNO27" & vbCrLf
''   sSql = sSql & " FROM dbo.SoitTable RIGHT OUTER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.MrpbaTable AS MrpbaTable INNER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.MrplTable AS MrplTable ON MrpbaTable.MRPBOM_PARTREF = MrplTable.MRP_PARTREF INNER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.PartTable AS PartTable ON MrplTable.MRP_PARTREF = PartTable.PARTREF" & vbCrLf
''   sSql = sSql & "      ON dbo.SoitTable.ITSO = MrplTable.MRP_SONUM AND" & vbCrLf
''   sSql = sSql & "      dbo.SoitTable.ITNUMBER = MrplTable.MRP_SOITEM LEFT OUTER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.vw_RnopPivot INNER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.RunsTable ON dbo.vw_RnopPivot.OPREF = dbo.RunsTable.RUNREF" & vbCrLf
''   sSql = sSql & "      AND dbo.vw_RnopPivot.OPRUN = dbo.RunsTable.RUNNO ON" & vbCrLf
''   sSql = sSql & "      MrpbaTable.MRPBOM_ROOTPARTREF = dbo.RunsTable.RUNREF LEFT OUTER JOIN" & vbCrLf
''   sSql = sSql & "      dbo.SohdTable AS SohdTable ON MrplTable.MRP_SONUM = SohdTable.SONUMBER" & vbCrLf
''   sSql = sSql & " where (MrpbaTable.MRPBOM_LEVEL = 0) And (MrplTable.MRP_TYPE = 11)" & vbCrLf
''   sSql = sSql & " AND (dbo.RunsTable.RUNSTATUS <> 'CL') AND" & vbCrLf
''   sSql = sSql & " (dbo.RunsTable.RUNSTATUS <> 'CO')"
'
''dbo.SoitTable.ITSO, dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITREV,
' '                     dbo.SohdTable.SOCUST
'
'
'   If Trim(txtBeg) = "" Then txtBeg = "ALL"
'   If Trim(txtEnd) = "" Then txtEnd = "ALL"
'   If Not IsDate(txtBeg) Then
'      sBDate = "01/01/2000"
'   Else
'      sBDate = Format(txtBeg, "mm/dd/yyyy")
'   End If
'   If Not IsDate(txtEnd) Then
'      sEDate = "12/31/2024"
'   Else
'      sEDate = Format(txtEnd, "mm/dd/yyyy")
'   End If
'
'
'   sSql = "SELECT MrplTable.MRP_PARTREF, MrplTable.MRP_PARTQTYRQD, MrplTable.MRP_PARTDATERQD, dbo.SohdTable.SOCUST, dbo.SoitTable.ITSO," & vbCrLf
'   sSql = sSql & "  dbo.SoitTable.ITNUMBER, dbo.SoitTable.ITCUSTREQ, dbo.SoitTable.ITQTY, dbo.SoitTable.ITDOLLARS," & vbCrLf
'   sSql = sSql & "  ISNULL(dbo.SoitTable.ITQTY * dbo.SoitTable.ITDOLLARS, 0) AS ExtendedCost, dbo.SoitTable.ITSCHED, dbo.SohdTable.SOPO, dbo.SohdTable.SOTYPE," & vbCrLf
'   sSql = sSql & "  RUNCOMMENTS, PartTable.PAMAKEBUY, PartTable.PAQOH, PartTable.PALOCATION, dbo.SohdTable.SOSHIPDATE," & vbCrLf
'   sSql = sSql & "  dbo.RunsTable.RUNREF, dbo.RunsTable.RUNNO, dbo.RunsTable.RUNSCHED, dbo.RunsTable.RUNQTY, dbo.RunsTable.RUNSTATUS," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO1, dbo.vw_RnopPivot.OPNO2, dbo.vw_RnopPivot.OPNO3, dbo.vw_RnopPivot.OPNO4, dbo.vw_RnopPivot.OPNO5," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO6, dbo.vw_RnopPivot.OPNO7, dbo.vw_RnopPivot.OPNO8, dbo.vw_RnopPivot.OPNO9, dbo.vw_RnopPivot.OPNO10," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO11, dbo.vw_RnopPivot.OPNO12, dbo.vw_RnopPivot.OPNO13, dbo.vw_RnopPivot.OPNO14, dbo.vw_RnopPivot.OPNO15," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO16, dbo.vw_RnopPivot.OPNO17, dbo.vw_RnopPivot.OPNO18, dbo.vw_RnopPivot.OPNO20, dbo.vw_RnopPivot.OPNO19," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO21, dbo.vw_RnopPivot.OPNO22, dbo.vw_RnopPivot.OPNO23, dbo.vw_RnopPivot.OPNO24, dbo.vw_RnopPivot.OPNO25," & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot.OPNO26 , dbo.vw_RnopPivot.OPNO27" & vbCrLf
'   sSql = sSql & "FROM  dbo.RunsTable INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.RnalTable ON dbo.RunsTable.RUNREF = dbo.RnalTable.RAREF AND dbo.RunsTable.RUNNO = dbo.RnalTable.RARUN INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.SoitTable ON dbo.RnalTable.RASO = dbo.SoitTable.ITSO AND dbo.RnalTable.RASOITEM = dbo.SoitTable.ITNUMBER AND" & vbCrLf
'   sSql = sSql & "  dbo.RnalTable.RASOREV = dbo.SoitTable.ITREV INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.SohdTable ON dbo.RnalTable.RASO = dbo.SohdTable.SONUMBER RIGHT OUTER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.MrpbaTable AS MrpbaTable INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.MrplTable AS MrplTable ON MrpbaTable.MRPBOM_PARTREF = MrplTable.MRP_PARTREF INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.PartTable AS PartTable ON MrplTable.MRP_PARTREF = PartTable.PARTREF INNER JOIN" & vbCrLf
'   sSql = sSql & "  dbo.vw_RnopPivot ON MrplTable.MRP_PARTREF = dbo.vw_RnopPivot.OPREF AND MrplTable.MRP_MORUNNO = dbo.vw_RnopPivot.OPRUN ON" & vbCrLf
'   sSql = sSql & "  dbo.RunsTable.Runno = MrplTable.MRP_MORUNNO And dbo.RunsTable.RUNREF = MrplTable.MRP_PARTREF" & vbCrLf
'   sSql = sSql & "WHERE (MrpbaTable.MRPBOM_LEVEL = 0) AND (MrplTable.MRP_TYPE = 3) AND (dbo.RunsTable.RUNSTATUS <> 'CL') AND" & vbCrLf
'   sSql = sSql & " (MrplTable.MRP_PARTDATERQD BETWEEN '" & sBDate & "' AND '" & sEDate & "') " & vbCrLf
'   sSql = sSql & "  AND (dbo.RunsTable.RUNSTATUS <> 'CO')"
'
'
'   Debug.Print sSql
'
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_STATIC)
'
'   If bSqlRows Then
'      sFileName = txtFilePath.Text
'      SaveAsExcel RdoPO, sFieldsToExport, sFileName
'   Else
'      MsgBox "No records found. Please try again.", vbOKOnly
'   End If
'
'   Set RdoPO = Nothing
'   Exit Function
'
'ExportError:
'   MouseCursor 0
'   cmbExport.Enabled = True
'   MsgBox Err.Description
'
'
'End Function

Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "MRP_PARTREF"
   sFieldsToExport(i + 1) = "MRP_PARTQTYRQD"
   sFieldsToExport(i + 2) = "MRP_PARTDATERQD"
   sFieldsToExport(i + 3) = "SOCUST"
   sFieldsToExport(i + 4) = "ITSO"
   sFieldsToExport(i + 5) = "ITNUMBER"
   sFieldsToExport(i + 6) = "ITCUSTREQ"
   sFieldsToExport(i + 7) = "ITQTY"
   sFieldsToExport(i + 8) = "ITDOLLARS"
   sFieldsToExport(i + 9) = "ExtendedCost"
   sFieldsToExport(i + 10) = "ITSCHED"
   sFieldsToExport(i + 11) = "SOPO"
   sFieldsToExport(i + 12) = "SOTYPE"
   sFieldsToExport(i + 13) = "RUNCOMMENTS"
   sFieldsToExport(i + 14) = "PAMAKEBUY"
   sFieldsToExport(i + 15) = "PAQOH"
   sFieldsToExport(i + 16) = "PALOCATION"
   sFieldsToExport(i + 17) = "SOSHIPDATE"
   sFieldsToExport(i + 18) = "RUNREF"
   sFieldsToExport(i + 19) = "RUNNO"
   sFieldsToExport(i + 20) = "RUNSCHED"
   sFieldsToExport(i + 21) = "RUNQTY"
   sFieldsToExport(i + 22) = "RUNSTATUS"
   sFieldsToExport(i + 23) = "OPNO1"
   sFieldsToExport(i + 24) = "OPNO2"
   sFieldsToExport(i + 25) = "OPNO3"
   sFieldsToExport(i + 26) = "OPNO4"
   sFieldsToExport(i + 27) = "OPNO5"
   sFieldsToExport(i + 28) = "OPNO6"
   sFieldsToExport(i + 29) = "OPNO7"
   sFieldsToExport(i + 30) = "OPNO8"
   sFieldsToExport(i + 31) = "OPNO9"
   sFieldsToExport(i + 32) = "OPNO10"
   sFieldsToExport(i + 33) = "OPNO11"
   sFieldsToExport(i + 34) = "OPNO12"
   sFieldsToExport(i + 35) = "OPNO13"
   sFieldsToExport(i + 36) = "OPNO14"
   sFieldsToExport(i + 37) = "OPNO15"
   sFieldsToExport(i + 38) = "OPNO16"
   sFieldsToExport(i + 39) = "OPNO17"
   sFieldsToExport(i + 40) = "OPNO18"
   sFieldsToExport(i + 41) = "OPNO19"
   sFieldsToExport(i + 42) = "OPNO20"
   sFieldsToExport(i + 43) = "OPNO21"
   sFieldsToExport(i + 44) = "OPNO22"
   sFieldsToExport(i + 45) = "OPNO23"
   sFieldsToExport(i + 46) = "OPNO24"
   sFieldsToExport(i + 47) = "OPNO25"
   sFieldsToExport(i + 48) = "OPNO26"
   sFieldsToExport(i + 49) = "OPNO27"


End Function
   
Private Sub cmdCan_Click()
    Unload Me

End Sub

'Private Sub cmdFnd_Click()
'   ViewParts.lblControl = "TXTPRT"
'   ViewParts.txtPrt = txtPrt
'   optVew.Value = vbChecked
'   ViewParts.Show
'
'End Sub

Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (907)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub


Private Sub FillCombos()
    On Error Resume Next
    sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
        & "FROM PartTable  " _
        & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
        & " WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 " _
        & "ORDER BY PARTREF"
    LoadComboBox cmbPart, 0
    cmbPart = "ALL"
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub


Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If

End Sub

Private Sub Form_Activate()
    On Error Resume Next
    MDISect.lblBotPanel = Caption
    If bOnLoad Then
        GetMRPDates
        GetOptions
        cmbCde.AddItem ("ALL")
        FillProductCodes
        If Trim(cmbCde) = "" Then cmbCde = cmbCde.List(0)
        cmbCls.AddItem ("ALL")
        FillProductClasses
        If Trim(cmbCls) = "" Then cmbCls = cmbCls.List(0)
        FillCombos
        bOnLoad = 0
    End If
    If optVew.Value = vbChecked Then
        optVew.Value = vbUnchecked
        Unload (ViewParts)
    End If
    MouseCursor (0)

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

'Private Sub Form_QueryUnload(ByVal Cancel As Integer, ByVal UnloadMode As Integer)
'    'SaveOptions
'End Sub

Private Sub Form_Resize()
    Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormUnload
    Set MrplMRp08 = Nothing

End Sub




Private Sub PrintReport()
   Dim sParts As String
   Dim sCode As String
   Dim sClass As String
   Dim sBuyer As String
   Dim sMbe As String
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEndDate As String
   
   Dim sCustomReport As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   Dim strIncludes As String
   Dim strDateDev As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   GetMRPCreateDates sBegDate, sEndDate
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBDate = "2000,01,01"
   Else
      sBDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024,12,31"
   Else
      sEDate = Format(txtEnd, "yyyy,mm,dd")
   End If
   
   If Trim(cmbPart) = "" Then cmbPart = "ALL"
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   If Trim(cmbPart) = "ALL" Then sParts = "" Else sParts = Compress(cmbPart)
   If Trim(cmbCde) = "ALL" Then sCode = "" Else sCode = Compress(cmbCde)
   If Trim(cmbCls) = "ALL" Then sClass = "" Else sClass = Compress(cmbCls)
   
   'get custom report name if one has been defined
   sCustomReport = GetCustomReport("prdmr08")
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.ShowGroupTree False
   
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "Includes"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "ShowComment"
   aFormulaName.Add "Mbe"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   
   strIncludes = Trim(cmbPart) & ", Prod Code(s) " & cmbCde & ", Class(es) " _
                           & cmbCls
   aFormulaValue.Add CStr("'" & CStr(strIncludes) & "...'")
   
   aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")
   aFormulaValue.Add optCom.Value
   
   sSql = "{MrplTable.MRP_PARTREF} LIKE '" & sParts & "*' " _
          & "AND {MrplTable.MRP_PARTPRODCODE} LIKE '" & sCode _
          & "*' AND {MrplTable.MRP_PARTCLASS} LIKE '" & sClass & "*' " _
          & " AND {MrpbaTable.MRPBOM_LEVEL} = 0 AND {MrplTable.MRP_TYPE} = 11"
   
   
   If optMbe(0).Value = True Then
      sMbe = "Make"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='M'"
   ElseIf optMbe(1).Value = True Then
      sMbe = "Buy"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='B'"
   ElseIf optMbe(2).Value = True Then
      sMbe = "Either"
      sSql = sSql & "AND {PartTable.PAMAKEBUY}='E'"
   Else
      sMbe = "Make, Buy And Either"
   End If
   
   aFormulaValue.Add CStr("'" & sMbe & "'")
   
   
   'select part types
   Dim types As String
   Dim includes As String
   Dim i As Integer
   For i = 1 To 8
     If Me.chkType(i).Value = vbChecked Then
        If types = "" Then
           types = " AND ("
        Else
           types = types & " OR "
        End If
        
        types = types & "{PartTable.PALEVEL} = " & i
        includes = includes & " " & i
     End If
   Next
   If types = "" Then
     MsgBox "No part types selected"
     Exit Sub
   
   Else
     sSql = sSql & types & ")"
   End If
   
   aFormulaName.Add "PartInc"
   aFormulaValue.Add CStr("'" & includes & "'")
      
      ' Set Formula values
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   ' set the report Selection
   cCRViewer.SetReportSelectionFormula (sSql)
   
   
   cCRViewer.CRViewerSize Me
   ' Set report parameter
   cCRViewer.SetDbTableConnection
   
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aRptParaType
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
'   txtPrt = "ALL"
   
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   Dim sCode As String * 6
   Dim sClass As String * 4
   sCode = cmbCde
   sClass = cmbCls
   sOptions = sCode & sClass & Trim(str(Val(optCom.Value)))
   SaveSetting "Esi2000", "EsiProd", "Prdmr08", sOptions
   SaveSetting "Esi2000", "EsiProd", "Pmr08", lblPrinter
   SaveSetting "Esi2000", "EsiProd", "MOExpFN", txtFilePath.Text
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   On Error Resume Next
   sOptions = GetSetting("Esi2000", "EsiProd", "Prdmr08", sOptions)
   If Len(Trim(sOptions)) > 0 Then
      cmbCde = Mid$(sOptions, 1, 6)
      cmbCls = Mid$(sOptions, 7, 4)
      optCom.Value = Val(Mid$(sOptions, 11, 1))
   End If
   lblPrinter = GetSetting("Esi2000", "EsiProd", "Pmr08", lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
   
   txtFilePath.Text = GetSetting("Esi2000", "EsiProd", "MOExpFN", txtFilePath.Text)
   
End Sub

Private Sub optDis_Click()
   
   CreateAllOpenOrders
   
   PrintReport
   
End Sub
Private Sub CreateAllOpenOrders()
   
   Dim strParts As String
   Dim strCode As String
   Dim strClass As String
   Dim strMbe As String
   Dim sBDate As String
   Dim sEDate As String
   Dim sBegDate As String
   Dim sEndDate As String
   
   MouseCursor 13
   On Error GoTo DiaErr1
   GetMRPCreateDates sBegDate, sEndDate
   
   If Trim(txtBeg) = "" Then txtBeg = "ALL"
   If Trim(txtEnd) = "" Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBDate = "2000/01/01"
   Else
      sBDate = Format(txtBeg, "yyyy/mm/dd")
   End If
   If Not IsDate(txtEnd) Then
      sEDate = "2024/12/31"
   Else
      sEDate = Format(txtEnd, "yyyy/mm/dd")
   End If
   
   If Trim(cmbPart) = "" Then cmbPart = "ALL"
   If Trim(cmbCde) = "" Then cmbCde = "ALL"
   If Trim(cmbCls) = "" Then cmbCls = "ALL"
   If Trim(cmbPart) = "ALL" Then strParts = "" Else strParts = Compress(cmbPart)
   If Trim(cmbCde) = "ALL" Then strCode = "" Else strCode = Compress(cmbCde)
   If Trim(cmbCls) = "ALL" Then strClass = "" Else strClass = Compress(cmbCls)
   
   sSql = "truncate table MrpbaTable"
   clsADOCon.ExecuteSQL sSql
   
   Dim RdoBom As ADODB.Recordset
   
   sSql = "SELECT DISTINCT PARTREF,PARTNUM " _
       & "FROM PartTable  " _
       & "INNER JOIN MrplTable ON MrplTable.MRP_PARTREF=PartTable.PARTREF " _
       & " AND MrplTable.MRP_PARTREF LIKE '" & strParts & "%' " _
       & " AND MrplTable.MRP_PARTPRODCODE LIKE '" & strCode & "%' " _
       & " AND MrplTable.MRP_PARTDATERQD between '" & sBDate & "' AND " _
       & " '" & sEDate & "' AND " _
       & " MrplTable.MRP_PARTCLASS LIKE '" & strClass & "%' " _
       & "ORDER BY PARTREF"
       
   Debug.Print sSql
   
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            sProcName = "getbomrev"
            GetBill Trim(!PartRef)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
   Exit Sub
DiaErr1:
   sProcName = "CreateAllOpenOrders"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

'Not using recursion here to keep the levels straight and
'make it easy to read

Private Sub GetBill(ByVal strRootPartRef As String)
   Dim RdoBom As ADODB.Recordset
   
   MouseCursor 13
   iOrder = 0
   On Error GoTo DiaErr1
   sSql = "INSERT INTO MrpbaTable (MRPBOM_ORDER, MRPBOM_ROOTPARTREF, MRPBOM_PARTREF," _
          & "MRPBOM_USEDON,MRPBOM_LEVEL) " _
          & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & strRootPartRef & "','" _
          & "',0)"
   clsADOCon.ExecuteSQL sSql
   
   strBomRev = GetBomRev(strRootPartRef)
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM BmplTable " _
          & "WHERE (BMASSYPART='" & strRootPartRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            sProcName = "getbomrev"
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF, MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',1)"
            clsADOCon.ExecuteSQL sSql
            GetNextBillLevel2 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "GetBill"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextBillLevel2(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel2"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF, MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',2)"
            clsADOCon.ExecuteSQL sSql
            GetNextBillLevel3 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel2"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetNextBillLevel3(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel3"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF,MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',3)"
            clsADOCon.ExecuteSQL sSql
            GetNextBillLevel4 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing
   
   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel3"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub GetNextBillLevel4(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel4"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF,MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',4)"
            clsADOCon.ExecuteSQL sSql
            GetNextBillLevel5 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing

   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel4"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetNextBillLevel5(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel5"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF,MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',5)"
            ' MM 6/19/2010 - Missing record update
            clsADOCon.ExecuteSQL sSql
            
            GetNextBillLevel6 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing

   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel5"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetNextBillLevel6(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel6"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF,MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',6)"
            clsADOCon.ExecuteSQL sSql
            GetNextBillLevel7 strRootPartRef, Trim(!BMASSYPART), Trim(!BMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing

   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel6"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetNextBillLevel7(ByVal strRootPartRef As String, sPartNumber As String, AssyRef As String)
   Dim RdoBom As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strBomRev = GetBomRev(sPartNumber)
   sProcName = "getnextbilllevel7"
   sSql = "SELECT BMASSYPART,BMPARTREF,BMREV,BMQTYREQD FROM " _
          & "BmplTable WHERE (BMASSYPART='" & AssyRef & "' " _
          & "AND BMREV='" & strBomRev & "') ORDER BY BMPARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      On Error Resume Next
      With RdoBom
         Do Until .EOF
            iOrder = iOrder + 1
            sIns = "INSERT INTO MrpbaTable (MRPBOM_ORDER,MRPBOM_ROOTPARTREF,MRPBOM_PARTREF," _
                   & "MRPBOM_USEDON,MRPBOM_QTYREQD, MRPBOM_LEVEL) " _
                   & "VALUES(" & iOrder & ",'" & strRootPartRef & "','" & Trim(!BMPARTREF) & "','" _
                   & Trim(!BMASSYPART) & "','" & Trim(!BMQTYREQD) & "',7)"
            .MoveNext
         Loop
         ClearResultSet RdoBom
      End With
   End If
   Set RdoBom = Nothing

   
   Exit Sub
   
DiaErr1:
   sProcName = "GetNextBillLevel7"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub




Private Sub optPrn_Click()
   PrintReport
   
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


'Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyF4 Then
'      ViewParts.lblControl = "TXTPRT"
'      ViewParts.txtPrt = txtPrt
'      optVew.Value = vbChecked
'      ViewParts.Show
'   End If
'
'End Sub

''Private Sub txtPrt_LostFocus()
 '  txtPrt = CheckLen(txtPrt, 30)
 '  If Trim(txtPrt) = "" Then txtPrt = "ALL"
 '
'End Sub



Private Sub cmbPart_LostFocus()
    cmbPart = CheckLen(cmbPart, 30)
    If Trim(cmbPart) = "" Then cmbPart = "ALL"
End Sub

Private Function GetBomRev(sPartNumber) As String
   Dim RdoRev As ADODB.Recordset
   
   sProcName = "getbomrev"
   sSql = "SELECT PARTREF,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRev, ES_FORWARD)
   If bSqlRows Then
      With RdoRev
         GetBomRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoRev
      End With
   Else
      GetBomRev = ""
   End If
   Set RdoRev = Nothing
End Function



