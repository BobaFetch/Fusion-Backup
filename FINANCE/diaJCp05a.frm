VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form diaJCp05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manufacturing Order Cost History"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "diaJCp05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtBeg 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox txtEnd 
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaJCp05a.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Display The Report"
         Top             =   110
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaJCp05a.frx":092C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print The Report"
         Top             =   110
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Manufacturing Orders"
      Top             =   960
      Width           =   3545
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2745
      FormDesignWidth =   7215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Costed Date"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   255
      Index           =   15
      Left            =   5520
      TabIndex        =   13
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblTyp 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6120
      TabIndex        =   12
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Blank For All)"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   9
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date:"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date:"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "diaJCp05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim bGoodPart As Byte
Dim bOnLoad As Byte

Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtEnd = Format(ES_SYSDATE, "mm/dd/yy")
   txtBeg = "01/01/" & Right(txtEnd, 2)
   
End Sub

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetPart()
   
End Sub
Private Sub PrintReport()
   Dim sBegDate As String
   Dim sEnddate As String
   
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   If Len(Trim(txtBeg)) = 0 Then txtBeg = "ALL"
   If Len(Trim(txtEnd)) = 0 Then txtEnd = "ALL"
   If Not IsDate(txtBeg) Then
      sBegDate = "1995,01,01"
   Else
      sBegDate = Format(txtBeg, "yyyy,mm,dd")
   End If
   If Not IsDate(txtEnd) Then
      sEnddate = "2024,12,31"
   Else
      sEnddate = Format(txtEnd, "yyyy,mm,dd")
   End If
   MouseCursor 13
   On Error GoTo Psh02
   'SetMdiReportsize MDISect
   sPartNumber = Compress(cmbPrt)
   
    aFormulaName.Add "CompanyName"
    aFormulaName.Add "Includes"
    aFormulaName.Add "RequestBy"
    
    aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
    aFormulaValue.Add CStr("'" & CStr(txtBeg) & "'")
    aFormulaValue.Add CStr("'Requested By:" & CStr(sInitials) & "'")

   sCustomReport = GetCustomReport("finMOCostHistory")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
   sSql = "{LOHDTABLE.LOTMOPARTREF} = '" & sPartNumber & "' AND {loitTable.LOITYPE} = 6 " _
          & " AND {LOHDTABLE.LOTADATE} in Date(" & Format(sBegDate, "yyyy,mm,dd") & ") " _
          & " to Date(" & Format(sEnddate, "yyyy,mm,dd") & ")"
   
   cCRViewer.SetReportSelectionFormula sSql
   
   sSql = "{LoitTable.LOIPARTREF} = {?Pm-LohdTable.LOTPARTREF} AND {loitTable.LOITYPE} = 25 " _
         & " AND {LOHDTABLE.LOTADATE} in Date(" & Format(sBegDate, "yyyy,mm,dd") & ") " _
          & " to Date(" & Format(sEnddate, "yyyy,mm,dd") & ")"
   
   ' set the sub sql variable pass the sub report name
   cCRViewer.SetSubRptSelFormula "subrptSalesPrices", sSql
      
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.CRViewerSize Me
   
   cCRViewer.OpenCrystalReportObject Me, aFormulaName

   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   'SetCrystalAction Me
   MouseCursor 0
   Exit Sub
   
Psh02:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Resume Psh02a
Psh02a:
   DoModuleErrors Me
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillAllRuns cmbPrt
      bGoodPart = GetPart()
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub




Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaJCp05a = Nothing
   
End Sub





Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   sPartNumber = Compress(cmbPrt)
   On Error GoTo DiaErr1
   If Len(sPartNumber) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL FROM PartTable WHERE PARTREF='" & sPartNumber & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
      If bSqlRows Then
         With RdoPrt
            cmbPrt = "" & Trim(!PARTNUM)
            lblDsc = "" & Trim(!PADESC)
            lblTyp = Format(0 + !PALEVEL, "#")
         End With
         GetPart = True
      Else
         MsgBox "Part Wasn't Found.", vbExclamation, Caption
         cmbPrt = ""
         lblDsc = ""
         GetPart = False
      End If
      On Error Resume Next
      RdoPrt.Close
   Else
      sPartNumber = ""
      cmbPrt = ""
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub optDis_Click()
   PrintReport
   
End Sub

Private Sub optPrn_Click()
   PrintReport
   
End Sub


Private Sub txtBeg_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtBeg_LostFocus()
   If Len(txtBeg) > 3 Then
      txtBeg = CheckDate(txtBeg)
   Else
      txtBeg = "ALL"
   End If
   
End Sub


Private Sub txtend_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtEnd_LostFocus()
   If Len(txtEnd) > 3 Then
      txtEnd = CheckDate(txtEnd)
   Else
      txtEnd = "ALL"
   End If
   
End Sub

