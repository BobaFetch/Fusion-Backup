VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaForcastf01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forcast report"
   ClientHeight    =   2580
   ClientLeft      =   2115
   ClientTop       =   1125
   ClientWidth     =   7305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      ToolTipText     =   "Open Access file Name"
      Top             =   1080
      Width           =   255
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Access File to Import"
      Filter          =   "*.mdb"
   End
   Begin VB.CommandButton cmdImport 
      Cancel          =   -1  'True
      Caption         =   "Import Forecast data"
      Height          =   360
      Left            =   3840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2145
   End
   Begin VB.CheckBox optVew 
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtAccessFilePath 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Tag             =   "3"
      Text            =   "SAP_POPfcst.mdb"
      Top             =   1080
      Width           =   4695
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2580
      FormDesignWidth =   7305
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   360
      Left            =   6000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.Frame fraPrn 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      Begin VB.CommandButton optDis 
         Height          =   330
         Left            =   0
         Picture         =   "diaForcastf01.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Display The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton optPrn 
         Height          =   330
         Left            =   560
         Picture         =   "diaForcastf01.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Print The Report"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   4
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
      PictureUp       =   "diaForcastf01.frx":0308
      PictureDn       =   "diaForcastf01.frx":044E
   End
   Begin Threed.SSRibbon ShowPrinters 
      Height          =   255
      Left            =   360
      TabIndex        =   5
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
      PictureUp       =   "diaForcastf01.frx":0594
      PictureDn       =   "diaForcastf01.frx":06DA
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Forecast Access file"
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lblPrinter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "diaForcastf01"
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
' diaForcastf01 - Forcast report
'
' Notes:
'
' Created: 12/06/02 (nth)
' Revisions:
'   10/22/03 (nth) Added get custom report
'   01/12/03 (nth) Fix report runtime error
'
'*********************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private JetAccWkSpace As Workspace
Private JetAccDb As DAO.Database

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   On Error GoTo DiaErr1
   strpathFilename = txtAccessFilePath.Text
   
   If (Trim(strpathFilename) = "") Then
      MsgBox "Please select the Access file to import Forcast data.", _
            vbInformation, Caption
         Exit Sub
   End If
   'strWindows = GetWindowsDir()
   'strpathFilename = strWindows & "\temp\" & strAccFileName
   Set JetAccWkSpace = DBEngine.Workspaces(0)
   Set JetAccDb = JetAccWkSpace.OpenDatabase(strpathFilename)
   
   Dim RdoForecast As ADODB.Recordset
   Dim DbRawForecast As Recordset
   
   sSql = "SELECT DTE,STATUS_DT,SUPPL_FLDR,BUYER,MANAGER,SRMANAGER,SUPP_NAME,SUPP_ADDRESS, " _
         & " VENDOR_NO,MATL_NO,ASSY_NO,DESC,PLANNED_DELVY_DAYS,ZALLOY,ZGAGE, " _
         & " ZLENGTH,ZWIDTH,UNIT_MEASURE,ZSPECIFICATION,ZTEMPER,ORDER_UNIT,PRCH_DOC_NO, " _
         & " FRCST01_1,FRCST01_2,FRCST02_1,FRCST02_2,FRCST03_1,FRCST03_2,FRCST04_1, " _
         & " FRCST04_2,FRCST05_1,FRCST05_2,FRCST06_1,FRCST06_2,FRCST07_1,FRCST07_2, " _
         & " FRCST08_1,FRCST08_2,FRCST09_1,FRCST09_2,FRCST10_1,FRCST10_2,FRCST11_1, " _
         & " FRCST11_2,FRCST12_1,FRCST12_2,FRCST13_1,FRCST13_2,FRCST14_1,FRCST14_2, " _
         & " FRCST15_1,FRCST15_2,FRCST16_1,FRCST16_2,FRCST17_1,FRCST17_2,FRCST18_1, " _
         & " FRCST18_2,LAST2_DIGITS_PC,PC,DIGIT4_PC,DIGIT5_PC,MODEL,FRCST00_1, " _
         & " FRCST19_1,CONTRACT_EXPRY,PROFIT_CTR_ID FROM ForeCast"

   Set DbRawForecast = JetAccDb.OpenRecordset(sSql, dbOpenDynaset)

   ' Delete the Forecast table
   sSql = "DELETE FROM ForeCast"
   clsADOCon.ExecuteSQL sSql

   sSql = "DELETE FROM ForeCastSalesData"
   clsADOCon.ExecuteSQL sSql

   sSql = "SELECT DTE,STATUS_DT,SUPPL_FLDR,BUYER,MANAGER,SRMANAGER,SUPP_NAME,SUPP_ADDRESS, " _
         & " VENDOR_NO,MATL_NO,ASSY_NO,[DESC],PLANNED_DELVY_DAYS,ZALLOY,ZGAGE, " _
         & " ZLENGTH,ZWIDTH,UNIT_MEASURE,ZSPECIFICATION,ZTEMPER,ORDER_UNIT,PRCH_DOC_NO, " _
         & " FRCST01_1,FRCST01_2,FRCST02_1,FRCST02_2,FRCST03_1,FRCST03_2,FRCST04_1, " _
         & " FRCST04_2,FRCST05_1,FRCST05_2,FRCST06_1,FRCST06_2,FRCST07_1,FRCST07_2, " _
         & " FRCST08_1,FRCST08_2,FRCST09_1,FRCST09_2,FRCST10_1,FRCST10_2,FRCST11_1, " _
         & " FRCST11_2,FRCST12_1,FRCST12_2,FRCST13_1,FRCST13_2,FRCST14_1,FRCST14_2, " _
         & " FRCST15_1,FRCST15_2,FRCST16_1,FRCST16_2,FRCST17_1,FRCST17_2,FRCST18_1, " _
         & " FRCST18_2,LAST2_DIGITS_PC,PC,DIGIT4_PC,DIGIT5_PC,MODEL,FRCST00_1, " _
         & " FRCST19_1,CONTRACT_EXPRY,PROFIT_CTR_ID FROM ForeCast"
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoForecast, ES_DYNAMIC)
   With DbRawForecast
      While Not .EOF
         RdoForecast.AddNew
         RdoForecast!DTE = "" & Trim(!DTE)
         RdoForecast!STATUS_DT = !STATUS_DT
         RdoForecast!DTE = "" & Trim(!DTE)
         RdoForecast!STATUS_DT = "" & Trim(!STATUS_DT)
         RdoForecast!SUPPL_FLDR = "" & Trim(!SUPPL_FLDR)
         RdoForecast!BUYER = "" & Trim(!BUYER)
         RdoForecast!MANAGER = "" & Trim(!MANAGER)
         RdoForecast!SRMANAGER = "" & Trim(!SRMANAGER)
         RdoForecast!SUPP_NAME = "" & Trim(!SUPP_NAME)
         RdoForecast!SUPP_ADDRESS = "" & Trim(!SUPP_ADDRESS)
         RdoForecast!VENDOR_NO = "" & Trim(!VENDOR_NO)
         RdoForecast!MATL_NO = "" & Trim(!MATL_NO)
         RdoForecast!ASSY_NO = "" & Trim(!ASSY_NO)
         RdoForecast![Desc] = "" & Trim(![Desc])
         RdoForecast!PLANNED_DELVY_DAYS = "" & Trim(!PLANNED_DELVY_DAYS)
         RdoForecast!ZALLOY = "" & Trim(!ZALLOY)
         RdoForecast!ZGAGE = "" & Trim(!ZGAGE)
         RdoForecast!ZLENGTH = "" & Trim(!ZLENGTH)
         RdoForecast!ZWIDTH = "" & Trim(!ZWIDTH)
         RdoForecast!UNIT_MEASURE = "" & Trim(!UNIT_MEASURE)
         RdoForecast!ZSPECIFICATION = "" & Trim(!ZSPECIFICATION)
         RdoForecast!ZTEMPER = "" & Trim(!ZTEMPER)
         RdoForecast!ORDER_UNIT = "" & Trim(!ORDER_UNIT)
         RdoForecast!PRCH_DOC_NO = "" & Trim(!PRCH_DOC_NO)
         RdoForecast!FRCST01_1 = "" & Trim(!FRCST01_1)
         RdoForecast!FRCST01_2 = "" & Trim(!FRCST01_2)
         RdoForecast!FRCST02_1 = "" & Trim(!FRCST02_1)
         RdoForecast!FRCST02_2 = "" & Trim(!FRCST02_2)
         RdoForecast!FRCST03_1 = "" & Trim(!FRCST03_1)
         RdoForecast!FRCST03_2 = "" & Trim(!FRCST03_2)
         RdoForecast!FRCST04_1 = "" & Trim(!FRCST04_1)
         RdoForecast!FRCST04_2 = "" & Trim(!FRCST04_2)
         RdoForecast!FRCST05_1 = "" & Trim(!FRCST05_1)
         RdoForecast!FRCST05_2 = "" & Trim(!FRCST05_2)
         RdoForecast!FRCST06_1 = "" & Trim(!FRCST06_1)
         RdoForecast!FRCST06_2 = "" & Trim(!FRCST06_2)
         RdoForecast!FRCST07_1 = "" & Trim(!FRCST07_1)
         RdoForecast!FRCST07_2 = "" & Trim(!FRCST07_2)
         RdoForecast!FRCST08_1 = "" & Trim(!FRCST08_1)
         RdoForecast!FRCST08_2 = "" & Trim(!FRCST08_2)
         RdoForecast!FRCST09_1 = "" & Trim(!FRCST09_1)
         RdoForecast!FRCST09_2 = "" & Trim(!FRCST09_2)
         RdoForecast!FRCST10_1 = "" & Trim(!FRCST10_1)
         RdoForecast!FRCST10_2 = "" & Trim(!FRCST10_2)
         RdoForecast!FRCST11_1 = "" & Trim(!FRCST11_1)
         RdoForecast!FRCST11_2 = "" & Trim(!FRCST11_2)
         RdoForecast!FRCST12_1 = "" & Trim(!FRCST12_1)
         RdoForecast!FRCST12_2 = "" & Trim(!FRCST12_2)
         RdoForecast!FRCST13_1 = "" & Trim(!FRCST13_1)
         RdoForecast!FRCST13_2 = "" & Trim(!FRCST13_2)
         RdoForecast!FRCST14_1 = "" & Trim(!FRCST14_1)
         RdoForecast!FRCST14_2 = "" & Trim(!FRCST14_2)
         RdoForecast!FRCST15_1 = "" & Trim(!FRCST15_1)
         RdoForecast!FRCST15_2 = "" & Trim(!FRCST15_2)
         RdoForecast!FRCST16_1 = "" & Trim(!FRCST16_1)
         RdoForecast!FRCST16_2 = "" & Trim(!FRCST16_2)
         RdoForecast!FRCST17_1 = "" & Trim(!FRCST17_1)
         RdoForecast!FRCST17_2 = "" & Trim(!FRCST17_2)
         RdoForecast!FRCST18_1 = "" & Trim(!FRCST18_1)
         RdoForecast!FRCST18_2 = "" & Trim(!FRCST18_2)
         RdoForecast!LAST2_DIGITS_PC = "" & Trim(!LAST2_DIGITS_PC)
         RdoForecast!PC = "" & Trim(!PC)
         RdoForecast!DIGIT4_PC = "" & Trim(!DIGIT4_PC)
         RdoForecast!DIGIT5_PC = "" & Trim(!DIGIT5_PC)
         RdoForecast!MODEL = "" & Trim(!MODEL)
         RdoForecast!FRCST00_1 = "" & Trim(!FRCST00_1)
         RdoForecast!FRCST19_1 = "" & Trim(!FRCST19_1)
         ' We are not using the field
         'RdoForecast!CONTRACT_EXPRY = "" & Trim(!CONTRACT_EXPRY)
         RdoForecast!PROFIT_CTR_ID = "" & Trim(!PROFIT_CTR_ID)
         RdoForecast.Update
         .MoveNext
      Wend
   End With
   RdoForecast.Close
   Set RdoForecast = Nothing
   DbRawForecast.Close
   Set DbRawForecast = Nothing
   ' Normalize the data - Call storeprodecure
   Dim RdoNor As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "ForecastSales"
   clsADOCon.ExecuteSQL sSql
   
   ' Now show the report
   PrintReport
   
   Exit Sub
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   ' Open the table and read the data.
End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Access File (*.mdb) | *.mdb"
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtAccessFilePath.Text = ""
   Else
       txtAccessFilePath.Text = fileDlg.FileName
   End If
End Sub

'*********************************************************************************

Private Sub Form_Activate()
   On Error Resume Next
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      MouseCursor 13
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bGoodPart = 1 Then
      SaveCurrentSelections
   End If
   FormUnload
   Set diaForcastf01 = Nothing
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
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   
   sCustomReport = GetCustomReport("finforecast.rpt")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   aFormulaName.Add "CompanyName"
   aFormulaName.Add "RequestBy"
   aFormulaName.Add "Title1"
   
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   aFormulaValue.Add CStr("'Requested By: " & CStr(sInitials) & "'")
   aFormulaValue.Add CStr("'Forcast Report'")
   
    cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   
    cCRViewer.CRViewerSize Me
    ' Set report parameter
    cCRViewer.SetDbTableConnection
    ' print the copies
    'cCRViewer.SetReportSelectionFormula sSql
    cCRViewer.OpenCrystalReportObject Me, aFormulaName
    cCRViewer.ShowGroupTree False
    
    cCRViewer.ClearFieldCollection aFormulaName
    cCRViewer.ClearFieldCollection aFormulaValue
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   optPrn.enabled = True
   optDis.enabled = True
   sProcName = "PrintReport"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub SaveOptions()
   Dim sOptions As String
   SaveSetting "Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter
End Sub

Private Sub GetOptions()
   Dim sOptions As String
   lblPrinter = GetSetting("Esi2000", "EsiFina", Me.Name & TTSAVEPRN, lblPrinter)
   If lblPrinter = "" Then lblPrinter = "Default Printer"
End Sub

Private Sub cmdVew_Click()
   optVew.Value = vbChecked
   ViewParts.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveOptions
End Sub


