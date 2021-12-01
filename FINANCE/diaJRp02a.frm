VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaJRp02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Journals Status(Report)"
   ClientHeight    =   1740
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Tag             =   "4"
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select GL Type From List"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cmbFyr 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Fiscal Year"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   480
      Width           =   1095
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaJRp02a.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaJRp02a.frx":041D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      GroupAllowAllUp =   -1  'True
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaJRp02a.frx":0868
      PictureDn       =   "diaJRp02a.frx":09AE
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3600
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1740
      FormDesignWidth =   6555
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   10
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
      PictureUp       =   "diaJRp02a.frx":0AF4
      PictureDn       =   "diaJRp02a.frx":0C3A
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label lblkind 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   825
   End
   Begin VB.Label P 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal Year"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1425
   End
End
Attribute VB_Name = "diaJRp02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

'**************************************************************************************
' diaPjr01a - Display Journal Reports
'
' Notes: This form takes the places of all MCS journal viewing/reporting programs
'
' Created:  (cjs)
' Modified:
'   06/05/01 (nth) Redesigned window layout and included the SJ.
'   06/17/01 (nth) Added to INVCANCELED to sales journal selection formula.
'   11/11/02 (nth) Updated XC and PJ Journals.
'   11/14/02 (nth) Updated CR Journal.
'   12/11/02 (nth) Updated CC Journal.
'   04/04/03 (nth) Correctly display voided checks.
'   06/05/03 (nth) Removed IJ,TJ,and PR journal types.
'   09/18/03 (nth) Allow credit and debit memos to correctly display on SJ.
'   04/01/04 (nth) Canceled invoices on sales journal.
'   04/06/04 (nth) Add DCHEAD formula to CR journal.
'   08/16/04 (nth) Added printer saveoptions and getoptions
'
'**************************************************************************************

Option Explicit

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodId As Byte
Dim iFyear As Integer
Dim iJrnNo As Integer

Dim sKind(12) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbFyr_LostFocus()
   Dim i As Integer
   Dim b As Byte
   
   If Not bCancel = 0 Then
      cmbFyr = CheckLen(cmbFyr, 4)
      cmbFyr = Format(Abs(Val(cmbFyr)), "0000")
      For i = 0 To cmbFyr.ListCount - 1
         If cmbFyr = cmbFyr.List(i) Then b = 1
      Next
      If b = 0 Then
         Beep
         cmbFyr = Format(Now, "yyyy")
      End If
   End If
End Sub

Private Sub cmbTyp_Click()
   If cmbTyp.ListIndex > 0 Then lblkind = sKind(cmbTyp.ListIndex)
End Sub

Private Sub cmbTyp_LostFocus()
   If bCancel = 0 Then
      cmbTyp = CheckLen(cmbTyp, 3)
      If Trim(cmbTyp) = "" Then cmbTyp = "ALL"
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer _
                             , x As Single, y As Single)
   bCancel = 1
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
   Dim RdoPst As ADODB.Recordset
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT FYYEAR FROM GlfyTable order by FYYEAR desc"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPst, ES_FORWARD)
   
   If bSqlRows Then
      With RdoPst
         Do Until .EOF
            If Not IsNull(.Fields(0)) Then _
                          AddComboStr cmbFyr.hWnd, Format(.Fields(0), "0000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   
   AddComboStr cmbTyp.hWnd, "ALL"
   sKind(0) = "ALL"
   AddComboStr cmbTyp.hWnd, "SJ"
   sKind(1) = "Sales"
   AddComboStr cmbTyp.hWnd, "PJ"
   sKind(2) = "Purchases"
   AddComboStr cmbTyp.hWnd, "CR"
   sKind(3) = "Cash Receipts"
   AddComboStr cmbTyp.hWnd, "CC"
   sKind(4) = "Disp-Computer Checks"
   AddComboStr cmbTyp.hWnd, "XC"
   sKind(5) = "Disp-External Checks"
   AddComboStr cmbTyp.hWnd, "TJ"
   sKind(6) = "Time Charges"
   
   cmbTyp = "ALL"
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
   If bOnLoad = 1 Then
      MouseCursor 13
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
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
   Set diaJRp02a = Nothing
End Sub

Private Sub PrintReport()
   Dim sTemp As String
   Dim sType As String
   Dim sLaborAcct As String
   Dim sLaborDesc As String
   Dim sJournal As String
   Dim rptsql As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   
   If (Trim(cmbFyr.Text) = "") Then
      MsgBox ("Please select Journal Year")
      Exit Sub
   End If
   
   optPrn.enabled = False
   optDis.enabled = False
   
   On Error GoTo DiaErr1
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "JrnlYear"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'" & CStr(cmbFyr) & "'")
 
   rptsql = ""
   Dim sReportFile As String
   
   
   rptsql = "{JrhdTable.MJFY} = " & cmbFyr
   If cmbTyp = "ALL" Then
      rptsql = rptsql & " AND {JrhdTable.MJTYPE} like ['SJ','PJ','XC','CC','CR','TJ']"
   Else
      rptsql = rptsql & " AND {JrhdTable.MJTYPE} = '" & cmbTyp & "'"
   End If
   
   sReportFile = "finjr02"
   sCustomReport = GetCustomReport(sReportFile)
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   On Error GoTo DiaErr1
   '   SetCrystalAction Me
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
    
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    cCRViewer.SetReportSelectionFormula rptsql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aRptPara
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   optPrn.enabled = True
   optDis.enabled = True
   
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub SaveOptions()
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Public Sub GetOptions()
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then
      lblPrinter = TTDEFAULT
   End If
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

