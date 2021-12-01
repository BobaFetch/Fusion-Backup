VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf15a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Sales Orders"
   ClientHeight    =   11505
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   15780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11505
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox OptOnlySel 
      Caption         =   "Consume Only Selected"
      Height          =   195
      Left            =   8760
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   10920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdExportSel 
      Caption         =   "Export Selected"
      Height          =   375
      Left            =   6960
      TabIndex        =   28
      Top             =   10800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdChkQOH 
      Caption         =   "Check QOH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13320
      TabIndex        =   27
      ToolTipText     =   "Check QOH"
      Top             =   3240
      Width           =   2280
   End
   Begin VB.CommandButton cmdPrintPS 
      Caption         =   "Print PackSLip PS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13320
      TabIndex        =   26
      ToolTipText     =   "Apply VOI Consumption"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton CmdSelAll 
      Caption         =   "Selection All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13320
      TabIndex        =   25
      ToolTipText     =   " Select All"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export All"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   10800
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   10320
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   10320
      Width           =   4695
   End
   Begin VB.Frame z2 
      Height          =   975
      Left            =   14520
      TabIndex        =   18
      Top             =   10080
      Visible         =   0   'False
      Width           =   3015
      Begin VB.OptionButton optExostar 
         Caption         =   "VOI Import"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optExostar 
         Caption         =   "VOI Import"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CheckBox OptSoXml 
      Caption         =   "FromXMLSO"
      Height          =   195
      Left            =   14160
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   10560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   360
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cmbPre 
      Height          =   315
      Left            =   13920
      TabIndex        =   11
      Tag             =   "3"
      Text            =   "S"
      ToolTipText     =   "Select or Enter Type A thru Z"
      Top             =   10440
      Visible         =   0   'False
      Width           =   520
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   14400
      TabIndex        =   10
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1555
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   14400
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Enter New Sales Order Number"
      Top             =   10440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdApplyVOI 
      Caption         =   "Apply VOI Consumption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13320
      TabIndex        =   8
      ToolTipText     =   "Apply VOI Consumption"
      Top             =   2400
      Width           =   2280
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13320
      TabIndex        =   7
      ToolTipText     =   " Close this Manufacturing Order"
      Top             =   6720
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.TextBox txtAccFilePath 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Cancel          =   -1  'True
      Caption         =   "Import VOI Sales data"
      Height          =   360
      Left            =   4080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf15a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   11505
      FormDesignWidth =   15780
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   14160
      Top             =   10320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSComDlg.CommonDialog ExpDlg 
      Left            =   8640
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar prg2 
      Height          =   300
      Left            =   9360
      TabIndex        =   31
      Top             =   9600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   7215
      Left            =   360
      TabIndex        =   32
      Top             =   2400
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12726
      _Version        =   393216
      Rows            =   3
      Cols            =   13
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   5
      Left            =   600
      TabIndex        =   24
      Top             =   10320
      Width           =   1305
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   13200
      Y1              =   9960
      Y2              =   9960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sales Order"
      Height          =   255
      Index           =   3
      Left            =   13800
      TabIndex        =   15
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   13920
      TabIndex        =   14
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   255
      Index           =   0
      Left            =   12360
      TabIndex        =   13
      Top             =   10560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9120
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15a.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15a.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select VOI db"
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1065
   End
End
Attribute VB_Name = "SaleSLf15a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE

Option Explicit

Dim cmdObj1 As ADODB.Command
Dim cmdObj2 As ADODB.Command

Dim bFIFO As Byte
Dim bGoodJrn As Boolean
Private Const PS_PACKSLIPNO = 0
Private Const PS_ITEMNO = 1
Private Const PS_QUANTITY = 2
'Private Const PS_PIPART = 3
Private Const PS_COST = 4
Private Const PS_LOTTRACKED = 5
Private Const PS_PARTNUM = 6

Dim vItems(800, 7) As Variant
Dim sPartGroup(800) As String '9/23/04 Compressed PartTable!PARTREF
Dim sSoItems(300, 3) As String 'Nathan 3/10/04
Dim sLots(50, 2) As String
Dim sCustomer As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Const SOITEM_SO = 0 ' string of PISONUMBER
Const SOITEM_ITEM = 1 ' string of PISOITEM
Const SOITEM_REV = 2 ' string of PISOREV


Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim strXML As String
Dim bNewImport As Boolean
Dim ExtName As String

Dim Fields(150) As String

Dim sCust As String
Dim cDiscount As Currency

Private txtKeyPress As New EsiKeyBd


Private Sub cmbExport_Click()
   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
   'CreateAllOpenOrders

   ExportVOISOs

End Sub

Private Function ExportVOISOs(Optional bSel = False)

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
   

   Dim RdoPO As ADODB.Recordset
   Dim i As Integer
   Dim sFieldsToExport(13) As String
   
   AddFieldsToExport sFieldsToExport
   
   Dim RdoPOIssued As ADODB.Recordset
   
   If (bSel = True) Then
      sSql = "SELECT DISTINCT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, a.WITHDRAWN_QTY, b.ISSUE_DTE, " _
               & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
               & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(b.Remarks,'') Remarks, ISNULL(a.REMARKPS,'') REMARKPS " _
               & " FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
               & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
               & "   INNER JOIN FusionSOVOIExport AS c ON " _
               & "   c.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO AND" _
               & "   C.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO"
      
   Else
   
      sSql = "SELECT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, a.WITHDRAWN_QTY, b.ISSUE_DTE, " _
               & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
               & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(b.Remarks,'') Remarks, ISNULL(a.REMARKPS,'') REMARKPS FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
               & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
               & " and a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   'Grd.Rows = 1
   Debug.Print sSql
   
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      SaveAsExcel RdoPOIssued, sFieldsToExport, sFileName
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If

   Set RdoPOIssued = Nothing
   Exit Function
   
ExportError:
   MouseCursor 0
   cmbExport.Enabled = True
   MsgBox Err.Description
   

End Function

Private Function AddFieldsToExport(ByRef sFieldsToExport() As String)
   
   Dim i As Integer
   i = 0
   sFieldsToExport(i) = "MATL_NO"
   sFieldsToExport(i + 1) = "PAYMENT_DOC_NO"
   sFieldsToExport(i + 2) = "PAYMENT_DOC_IT_NO"
   sFieldsToExport(i + 3) = "ISSUE_AMOUNT"
   sFieldsToExport(i + 4) = "WITHDRAWN_QTY"
   sFieldsToExport(i + 6) = "ISSUE_DTE"
   sFieldsToExport(i + 7) = "ITSO"
   sFieldsToExport(i + 8) = "ITNUMBER"
   sFieldsToExport(i + 9) = "ITREV"
   sFieldsToExport(i + 10) = "ITQty"
   sFieldsToExport(i + 11) = "ITSCHED"
   sFieldsToExport(i + 12) = "REMARKS"
   sFieldsToExport(i + 13) = "REMARKPS"
   
End Function


Private Sub cmdCan_Click()
   Unload Me

End Sub





Private Sub cmdExportSel_Click()

   Dim iList As Integer
   Dim strDocNo As String
   Dim strDocItNo As String
   
   If (txtFilePath.Text = "") Then
      MsgBox "Please Select Excel File.", vbExclamation
      Exit Sub
   End If
   
   sSql = "DELETE FROM FusionSOVOIExport"
   clsADOCon.ExecuteSql sSql ' rdExecDirect

   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 2
         strDocNo = Trim(Grd.Text)
         
         Grd.Col = 3
         strDocItNo = Trim(Grd.Text)
         
         sSql = "INSERT INTO FusionSOVOIExport (PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO)" _
                  & "VALUES('" & strDocNo & "','" & strDocItNo & "')"
   
         Debug.Print sSql
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
      End If
   Next
   
   ExportVOISOs True
   
End Sub

Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub FillGrid()

   Dim RdoPOIssued As ADODB.Recordset
   sSql = "SELECT DISTINCT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, b.WITHDRAWN_QTY, b.ISSUE_DTE, " _
            & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
            & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(b.Remarks,'') Remarks, ISNULL(a.REMARKPS,'') REMARKPS FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
            & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
            & " AND a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO  " _
            & " WHERE a.PAYMENT_DOC_IT_NO IS NULL"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   Grd.Rows = 1
   Debug.Print sSql
   
   If bSqlRows Then
   With RdoPOIssued
      While Not .EOF

         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         
         'Grd.Col = 0
         'Set Grd.CellPicture = Chkno.Picture
         Grd.Col = 0
         Grd.Text = Trim(!matl_no)
         
         Grd.Col = 1
         Grd.Text = Trim(!payment_doc_no)
         
         Grd.Col = 2
         Grd.Text = Trim(!payment_doc_it_no)
         
         Grd.Col = 3
         Grd.Text = Trim(!ISSUE_AMOUNT)
         
         Grd.Col = 4
         Grd.Text = Trim(!WITHDRAWN_QTY)
         
         Grd.Col = 5
         Grd.Text = Trim(!ISSUE_DTE)
         
         Grd.Col = 6
         Grd.Text = Trim(!itso)
         
         Grd.Col = 7
         Grd.Text = Trim(!ITNUMBER)
         
         Grd.Col = 8
         Grd.Text = Trim(!itrev)
         
         Grd.Col = 9
         Grd.Text = Trim(!ITQty)
   
         Grd.Col = 10
         Grd.Text = Trim(!itsched)
         
         Grd.Col = 11
         Grd.Text = Trim(!Remarks)
         
         Grd.Col = 12
         Grd.Text = Trim(!REMARKPS)

         .MoveNext
      Wend
      .Close
      ClearResultSet RdoPOIssued
      End With
   End If
   Set RdoPOIssued = Nothing

End Sub


Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   On Error GoTo DiaErr1
   strpathFilename = txtAccFilePath.Text
   MouseCursor 13
   
   If (Trim(strpathFilename) = "") Then
      MsgBox "Please select the Access file to import VOI data.", _
            vbInformation, Caption
         Exit Sub
   End If
   
   Dim conn As ADODB.Connection
   Dim RdoPOIssued As ADODB.Recordset
   Dim matNo As String
   Dim payDocNo As String
   Dim payDocItNo As Integer
   Dim issueAmt As Currency
   Dim WithDrawnQty As Currency
   Dim issueDate As String
   Dim CheckNo  As String
   Dim CheckAmt As Currency
   Dim CheckDate As String
   
   MouseCursor ccHourglass
   ' Delete the old table
   Debug.Print sSql
   clsADOCon.ExecuteSql "DELETE FROM VOIPmtIss"
   
   ' Delete the old table
   Debug.Print sSql
   ' MM clsADOCon.ExecuteSQL "DELETE FROM FusionSOVOI"

   clsADOCon.ExecuteSql "DELETE FROM tmpVOIPmtIss"

   sSql = "DELETE FROM FusionSOVOI WHERE (ITPSNUMBER IS NULL OR ITPSNUMBER = '')"
   
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql ' rdExecDirect


   Set conn = New ADODB.Connection
   Set RdoPOIssued = New ADODB.Recordset
   'conn.Open ("Provider=Microsoft.Jet.OLEDB 4.0;Data Source=C:\Program Files\Microsoft Office\Office\Samples\Northwind.mdb;Persist Security Info=False")
   'ASSIGNMENT_NO,CHECK_NO, CHECK_AMT, CHECK_DT

   conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strpathFilename + ";Persist Security Info=False"
   conn.open
   sSql = "SELECT DISTINCT MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, ISSUE_AMOUNT, WITHDRAWN_QTY," _
            & "ISSUE_DTE,CHECK_NO, CHECK_AMT, CHECK_DT  FROM Payment_Issues"

   RdoPOIssued.open sSql, conn, adOpenDynamic, adLockOptimistic

   Grd.Rows = 1
   Debug.Print sSql
   With RdoPOIssued
   While Not .EOF

      matNo = IIf(Not IsNull(!matl_no), Trim(!matl_no), "")
      payDocNo = IIf(Not IsNull(!payment_doc_no), Trim(!payment_doc_no), "")
      payDocItNo = IIf(Not IsNull(!payment_doc_it_no), Trim(!payment_doc_it_no), 0)
      issueAmt = IIf(Not IsNull(!ISSUE_AMOUNT), Trim(!ISSUE_AMOUNT), 0)
      WithDrawnQty = IIf(Not IsNull(!WITHDRAWN_QTY), Trim(!WITHDRAWN_QTY), 0)
      issueDate = IIf(Not IsNull(!ISSUE_DTE), Trim(!ISSUE_DTE), "")
      CheckNo = IIf(Not IsNull(!check_no), Trim(!check_no), "")
      CheckAmt = IIf(Not IsNull(!check_amt), Trim(!check_amt), 0)
      CheckDate = IIf(Not IsNull(!CHECK_DT), Trim(!CHECK_DT), "")

      sSql = "INSERT INTO tmpVOIPmtIss (MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, " _
            & " ISSUE_AMOUNT, WITHDRAWN_QTY," _
            & "ISSUE_DTE,CHECK_NO, CHECK_AMT, CHECK_DT)" _
      & "VALUES('" & Compress(matNo) & "','" & payDocNo & "','" _
         & payDocItNo & "','" & issueAmt & "','" & WithDrawnQty & "','" _
         & issueDate & "','" & CheckNo & "','" & CheckAmt & "','" _
         & CheckDate & "')"

      Debug.Print sSql
      clsADOCon.ExecuteSql sSql ' rdExecDirect

      .MoveNext
   Wend
   .Close
   End With
   Set RdoPOIssued = Nothing
   conn.Close
   
   
   
   Dim strDocNum As String
   ' get the doc number cut off
   strDocNum = GetDocNumCuffOff()
      
   If (strDocNum = "") Then
      MsgBox "The VOI cutoff Doc number is not populated. Contact Administrator.", vbExclamation, Caption
      MouseCursor 0
      Exit Sub
   End If
   
   
   sSql = "INSERT INTO VOIPmtIss (MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, " _
      & " ISSUE_AMOUNT, WITHDRAWN_QTY," _
      & "ISSUE_DTE,CHECK_NO, CHECK_AMT, CHECK_DT)" _
      & "     SELECT a.MATL_NO, a.PAYMENT_DOC_NO,a.PAYMENT_DOC_IT_NO," _
      & "          a.ISSUE_AMOUNT, a.WITHDRAWN_QTY," _
      & "         a.ISSUE_DTE,a.CHECK_NO, a.CHECK_AMT, a.CHECK_DT  FROM tmpVOIPmtIss a left outer join FusionSOVOI b " _
      & "      on a.payment_doc_no = b.payment_doc_no " _
      & "      and a.payment_doc_it_no = b.payment_doc_it_no " _
      & "   Where a.payment_doc_no > " & strDocNum _
      & "         and b.Payment_doc_it_no is null"
            
'
'& " SELECT MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, " _
'      & " ISSUE_AMOUNT, WITHDRAWN_QTY," _
'      & "ISSUE_DTE,CHECK_NO, CHECK_AMT, CHECK_DT FROM tmpVOIPmtIss " _
'      & " WHERE PAYMENT_DOC_NO IN (select distinct Payment_doc_no " _
'                        & " FROM dbo.tmpVOIPmtIss WHERE Payment_doc_no >= " & strDocNum & ")"
'
      Debug.Print sSql
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   
 'and check_no <> ''
 
      FillGrid
      
      MouseCursor 0
   Exit Sub

DiaErr1:
   MsgBox "Error DiaErr1"
   MouseCursor 0
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Function ConsumeOpenSOItems(strPartNum As String)

   Dim RdoSoit As ADODB.Recordset
   Dim itso As String
   Dim ITNum As Integer
   Dim itrev As String
   Dim ITQty As Currency
   Dim itsched As String

   Dim VOIRemQty As Currency
   Dim ConsumeQty As Currency
   Dim runningqty As Currency
   Dim VOIWithDrawnQty As Currency
   Dim strDocNo As String
   
   
   sSql = "DELETE FROM FusionSOVOI WHERE ITPART = '" & strPartNum & "'" _
      & " AND (ITPSNUMBER IS NULL OR ITPSNUMBER = '')"
   
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   
   ' copy to emp voi table
   '  delete any old data
   clsADOCon.ExecuteSql "DELETE FROM tempVOIPConsume" ' rdExecDirect
   
   sSql = "INSERT INTO tempVOIPConsume (MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, " _
            & " WITHDRAWN_QTY, REMQTY, ISSUE_DTE)" _
         & " SELECT MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO," _
            & "WITHDRAWN_QTY, WITHDRAWN_QTY, ISSUE_DTE " _
         & " FROM VOIPmtIss WHERE MATL_NO = '" & strPartNum & "'"
   
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   
   
   
'   sSql = "SELECT 1 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'            & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'            & " AND b.sonumber = a.itso and sotype = 'v'" _
'            & " AND itsched = '2/22/2022'" _
'            & " UNION " _
'      & " SELECT 2 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'            & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'            & " AND b.sonumber = a.itso and sotype = 'v'" _
'            & " AND itsched in ('12/25/2015', '12/31/2015', '12/25/2016', '12/31/2016')" _
'            & "  ORDER BY ordernum, itsched, itso, itnumber, itrev"
            
    ' revised to search most recent 12/25 for the current year and 12/25 for the next two years
    Dim yr As Integer
    Dim dates As String
    
    yr = year(Now())
    dates = "'12/25/" & CStr(yr) & "', '12/25/" & CStr(yr + 1) & "', '12/25/" & CStr(yr + 2) & "'"
'    sSql = "SELECT 1 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'          & " AND b.sonumber = a.itso and sotype = 'v'" _
'          & " AND itsched = '2/22/2022'" _
'          & " UNION " _
'    & " SELECT 2 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'          & " AND b.sonumber = a.itso and sotype = 'v'" _
'          & " AND itsched in (" & dates & ")" _
'          & "  ORDER BY ordernum, itsched, itso, itnumber, itrev"

' changed 4/3/2020
'    sSql = "SELECT 1 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'          & " AND b.sonumber = a.itso and sotype in ('N','V')" _
'          & " AND itsched = '2/22/2022'" _
'          & " UNION " _
'    & " SELECT 2 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
'          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
'          & " AND b.sonumber = a.itso and sotype in ('N','V')" _
'          & " AND itsched in (" & dates & ")" _
'          & "  ORDER BY ordernum, itsched, itso, itnumber, itrev"

   ' SO Type X added 4/3/2020 for holds and date changed from 2/22/2022 to 4/4/2044
    sSql = "SELECT 1 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
          & " AND b.sonumber = a.itso and sotype in ('N','V','X')" _
          & " AND itsched = '4/4/2044'" _
          & " UNION " _
    & " SELECT 2 as ordernum, itso, itnumber, itrev, itqty, itsched from soitTable a, sohdTable b " _
          & " WHERE itactual is null and ITCANCELDATE is null and itpart ='" & strPartNum & "' " _
          & " AND b.sonumber = a.itso and sotype in ('N','V','X')" _
          & " AND itsched in (" & dates & ")" _
          & "  ORDER BY ordernum, itsched, itso, itnumber, itrev"


   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_STATIC)
   If bSqlRows Then
      With RdoSoit
         While (Not .EOF)
            itso = "" & Trim(!itso)
            ITNum = "" & Trim(!ITNUMBER)
            itrev = "" & Trim(!itrev)
            ITQty = "" & Trim(!ITQty)
            itsched = "" & Trim(!itsched)
            
            runningqty = ITQty
            
            Dim RdoCon As ADODB.Recordset
            
            If (runningqty > 0) Then
               sSql = "SELECT MATL_NO, PAYMENT_DOC_NO,PAYMENT_DOC_IT_NO, " _
                        & " WITHDRAWN_QTY, REMQTY, ISSUE_DTE  FROM tempVOIPConsume " _
                        & " WHERE REMQTY > 0 " _
                        & " ORDER BY ISSUE_DTE"
                        
               bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_DYNAMIC)
               If bSqlRows Then
                  With RdoCon
                     While ((Not .EOF) And (runningqty <> 0))
                        
                        VOIRemQty = !REMQTY
                        VOIWithDrawnQty = !WITHDRAWN_QTY
                        
                        strDocNo = !payment_doc_no
                        
                        'If (strDocNo = "5103709935") Or (strDocNo = "5103707183") Then
                        '   strDocNo = !PAYMENT_DOC_NO
                        'End If
                        
                        If (VOIRemQty >= runningqty) Then
                        
                           ConsumeQty = runningqty
                           
                           sSql = "UPDATE tempVOIPConsume  SET REMQTY = REMQTY - " & runningqty _
                                    & " WHERE PAYMENT_DOC_NO = '" & !payment_doc_no & "' AND PAYMENT_DOC_IT_NO ='" _
                                    & !payment_doc_it_no & "'"
                           Debug.Print sSql
                           clsADOCon.ExecuteSql sSql ' rdExecDirect
                           
                           VOIRemQty = VOIRemQty - runningqty
                           
                           runningqty = 0
                        
                        Else
                           ConsumeQty = VOIRemQty
                           
                           sSql = "UPDATE tempVOIPConsume  SET REMQTY = 0" _
                                    & " WHERE PAYMENT_DOC_NO = '" & !payment_doc_no & "' AND PAYMENT_DOC_IT_NO ='" _
                                    & !payment_doc_it_no & "'"
                           Debug.Print sSql
                           clsADOCon.ExecuteSql sSql ' rdExecDirect
                           
                           runningqty = runningqty - VOIRemQty
                           
                        End If
                        
                         
                        sSql = "INSERT INTO FusionSOVOI (ITSO, ITNUMBER, ITREV, ITPART, ITQTY, WITHDRAWN_QTY, REMITQTY, " _
                                 & " COMPLFLG, PAYMENT_DOC_NO, PAYMENT_DOC_IT_NO, ITSCHED) " _
                                 & " VALUES ('" & itso & "'," & ITNum & ",'" & itrev & "','" _
                                 & strPartNum & "'," & ITQty & "," & ConsumeQty & "," & runningqty & ",1,'" _
                                 & !payment_doc_no & "','" & !payment_doc_it_no & "','" & itsched & "')"
                                 
                        Debug.Print sSql
                        clsADOCon.ExecuteSql sSql ' rdExecDirect
                     
                     .MoveNext
                     Wend
                  .Close
                  End With
                  Set RdoCon = Nothing
               End If
            End If
            
            .MoveNext ' next so Number
         Wend  ' SO Item
         .Close
      End With
      Set RdoSoit = Nothing
   End If
   'On Error Resume Next
   Set RdoSoit = Nothing

End Function


Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "Access Files (*.mdb) | *.mdb"
   
   fileDlg.ShowOpen
   If fileDlg.FileName = "" Then
       txtAccFilePath.Text = ""
   Else
       txtAccFilePath.Text = fileDlg.FileName
   End If
End Sub

Private Sub cmdPrintPS_Click()
   On Error GoTo DiaErr1
   ' Create the packslip
   Dim strDocNum As String
   Dim strPS As String
   Dim iList As Integer
   
   Dim strPrevDocNum As String
   strPrevDocNum = ""
   
   If (OptOnlySel = checked) Then
               ' Only if the part is checked
      For iList = 1 To Grd.Rows - 1
         Grd.Col = 0
         Grd.Row = iList
         If Grd.CellPicture = Chkyes.Picture Then
         
            Grd.Col = 2
            strDocNum = Grd.Text
            
            If (strPrevDocNum <> strDocNum) Then
               strPS = GetPSfromDoc(strDocNum)
               If (strPS <> "") Then
                  PrintVOIPS strPS, False
               End If
               strPrevDocNum = strDocNum
            End If
         End If
      Next
   Else
   
      Dim RdoDocNo As ADODB.Recordset
      sSql = "select distinct Payment_doc_no from FusionSOVOI WHERE REMARKPS IS NULL"
   
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoDocNo, ES_STATIC)
      If bSqlRows Then
         With RdoDocNo
            While Not .EOF
   
               strDocNum = "" & Trim(!payment_doc_no)
               strPS = GetPSfromDoc(strDocNum)
               If (strPS <> "") Then
                  PrintVOIPS strPS, False
               End If
               
               .MoveNext
            Wend
            .Close
            ClearResultSet RdoDocNo
         End With
      End If
      Set RdoDocNo = Nothing
   End If
DiaErr1:
   sProcName = "Print PS "
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub cmdChkQOH_Click()

   Dim RdoPrt As ADODB.Recordset
   Dim strPartNum As String
   
   Dim a As Integer
   Dim b As Currency
   Dim C As Currency
   
   MouseCursor 13
'   z1(1).Caption = "Step 1: Beginning Balance"
'   z1(1).Refresh
   

   sSql = "select distinct ITPART from FusionSOVOI" ' where ITPART = '183A914031'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
   If bSqlRows Then
      With RdoPrt
         While Not .EOF

            strPartNum = "" & Trim(!ITPART)
            CheckQOHForSO strPartNum
            DoEvents
            .MoveNext
         Wend
         .Close
         ClearResultSet RdoPrt
         DoEvents

      End With
   End If
   Set RdoPrt = Nothing
   
   Dim RdoPOIssued As ADODB.Recordset
   sSql = "SELECT DISTINCT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, b.WITHDRAWN_QTY, b.ISSUE_DTE, " _
            & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
            & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(b.Remarks,'') Remarks, " _
            & " ISNULL(a.RemarkPS,'') RemarkPS FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
            & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
            & " AND a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   Grd.Rows = 1
   Debug.Print sSql
   
   If bSqlRows Then
   With RdoPOIssued
      While Not .EOF

         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         
         Grd.Col = 0
         Grd.Text = Trim(!matl_no)
         
         Grd.Col = 1
         Grd.Text = Trim(!payment_doc_no)
         
         Grd.Col = 2
         Grd.Text = Trim(!payment_doc_it_no)
         
         Grd.Col = 3
         Grd.Text = Trim(!ISSUE_AMOUNT)
         
         Grd.Col = 4
         Grd.Text = Trim(!WITHDRAWN_QTY)
         
         Grd.Col = 5
         Grd.Text = Trim(!ISSUE_DTE)
         
         Grd.Col = 6
         Grd.Text = Trim(!itso)
         
         Grd.Col = 7
         Grd.Text = Trim(!ITNUMBER)
         
         Grd.Col = 8
         Grd.Text = Trim(!itrev)
         
         Grd.Col = 9
         Grd.Text = Trim(!ITQty)
   
         Grd.Col = 10
         Grd.Text = Trim(!itsched)
         
         Grd.Col = 11
         Grd.Text = Trim(!Remarks)
         
         Grd.Col = 12
         Grd.Text = Trim(!REMARKPS)

         .MoveNext
      Wend
      .Close
      ClearResultSet RdoPOIssued
      End With
   End If
   Set RdoPOIssued = Nothing
   
   
   
   MsgBox "Done Checking The QOH. Please check the status."
   
   
   
   MouseCursor 0
   
   
   Exit Sub
DiaErr1:
   sProcName = "Create PS"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   

End Sub

Private Sub cmdApplyVOI_Click()
   
   Dim RdoPrt As ADODB.Recordset
   Dim strPartNum As String
   
'   sSql = "SELECT DISTINCT itpart from soitTable a,sohdTable b" _
'               & " WHERE itactual is null " _
'               & " AND b.sonumber = a.itso and sotype = 'v'"
               '& " and itpart IN ('111A11911')"
   
      
   MouseCursor 13
   sSql = "SELECT DISTINCT matl_no from VOIPmtIss"
               
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
   If bSqlRows Then
      With RdoPrt
         While Not .EOF

            strPartNum = "" & Trim(!matl_no)
            ConsumeOpenSOItems Compress(strPartNum)

            ' Delete the old table
            Debug.Print sSql
            clsADOCon.ExecuteSql "DELETE FROM tempVOIPConsume"
            .MoveNext
         Wend
         .Close
         ClearResultSet RdoPrt
      End With
   End If
   Set RdoPrt = Nothing
   
   
   sSql = "UPDATE VOIPmtIss SET REMARKS = ''"
   
   Debug.Print sSql
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   
   'Update remark
   sSql = "UPDATE b SET b.REMARKS = 'SO Not found' from " _
            & " VOIPmtIss as b  left Outer join FusionSOVOI as a " _
            & "   on a.payment_doc_no = b.payment_doc_no " _
            & " AND a.Payment_doc_it_no = b.Payment_doc_it_no " _
            & " Where a.Payment_Doc_it_no Is Null "

   clsADOCon.ExecuteSql sSql ' rdExecDirect

   sSql = "UPDATE b SET b.REMARKS = 'Only Partial SO available :' + Convert(varchar(24), a.WITHDRAWN_QTY) FROM " _
            & "    (select ITPART, SUM(WITHDRAWN_QTY) as WITHDRAWN_QTY, " _
            & "            PAYMENT_DOC_NO , PAYMENT_DOC_IT_NO " _
            & "         from FusionSOVOI " _
            & "            GROUP BY ITPART, PAYMENT_DOC_NO, PAYMENT_DOC_IT_NO) as a, VOIPmtIss b " _
            & "   where a.itPart = b.matl_no " _
            & "   and a.Payment_Doc_it_no = b.Payment_Doc_it_no " _
            & "   and a.payment_doc_no = b.payment_doc_no " _
            & "   and a.withDrawn_qty <> b.withDrawn_qty "
   
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   
   sSql = "update VOIPmtIss set remarks = 'Credit Qty' where withDrawn_qty < 0"
   clsADOCon.ExecuteSql sSql ' rdExecDirect

   Dim RdoPOIssued As ADODB.Recordset
   sSql = "SELECT DISTINCT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, b.WITHDRAWN_QTY, b.ISSUE_DTE, " _
            & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
            & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(b.Remarks,'') Remarks, " _
            & " ISNULL(a.RemarkPS,'') RemarkPS FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
            & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
            & " AND a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   Grd.Rows = 1
   Debug.Print sSql
   
   If bSqlRows Then
   With RdoPOIssued
      While Not .EOF

         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         
         Grd.Col = 0
         Grd.Text = Trim(!matl_no)
         
         Grd.Col = 1
         Grd.Text = Trim(!payment_doc_no)
         
         Grd.Col = 2
         Grd.Text = Trim(!payment_doc_it_no)
         
         Grd.Col = 3
         Grd.Text = Trim(!ISSUE_AMOUNT)
         
         Grd.Col = 4
         Grd.Text = Trim(!WITHDRAWN_QTY)
         
         Grd.Col = 5
         Grd.Text = Trim(!ISSUE_DTE)
         
         Grd.Col = 6
         Grd.Text = Trim(!itso)
         
         Grd.Col = 7
         Grd.Text = Trim(!ITNUMBER)
         
         Grd.Col = 8
         Grd.Text = Trim(!itrev)
         
         Grd.Col = 9
         Grd.Text = Trim(!ITQty)
   
         Grd.Col = 10
         Grd.Text = Trim(!itsched)
         
         Grd.Col = 11
         Grd.Text = Trim(!Remarks)
         
         Grd.Col = 12
         Grd.Text = Trim(!REMARKPS)

         .MoveNext
      Wend
      .Close
      ClearResultSet RdoPOIssued
      End With
   End If
   Set RdoPOIssued = Nothing
   
   MouseCursor ccArrow
   
   Exit Sub
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   MouseCursor 0
   
End Sub
   
 
   
Public Function GetLotRemQty(strPartRef As String) As Currency

   Dim rdo As ADODB.Recordset
   sSql = "SELECT SUM(lotremainingqty) as LotRemQty from lohdTable where lotpartref = '" & strPartRef & "' AND LOTLOCATION = 'SPRT'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   
   If bSqlRows Then
      GetLotRemQty = Format(rdo!LotRemQty, ES_QuantityDataFormat)
   Else
      GetLotRemQty = Format(0, ES_QuantityDataFormat)
   End If
   Set rdo = Nothing

End Function

Private Function CheckQOHForSO(strPartNum As String)

   Dim RdoSoit As ADODB.Recordset
   Dim qoh As Currency
   Dim itso As String
   Dim ITNum As Integer
   Dim itrev As String
   Dim payDoc As String
   Dim payDocNo As String
   
   Dim WithDrawnQty As Currency
   Dim bDisplay As Boolean
   bDisplay = True
   qoh = GetLotRemQty(strPartNum)

   sSql = "Update FusionSOVOI set REMARKPS = '' FROM FusionSOVOI WHERE itpart = '" + strPartNum + "' AND ITPSNUMBER IS NULL"
   clsADOCon.ExecuteSql sSql ' rdExecDirect

   sSql = "select ITSO, ITNUMBER, ITREV, ITPART, WITHDRAWN_QTY,payment_doc_no, payment_doc_it_no FROM FusionSOVOI " _
            & "WHERE ITSO IS NOT NULL AND ITPART = '" & strPartNum & "' AND (ITPSNUMBER IS NULL OR ITPSNUMBER = '') ORDER BY ITSCHED"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_STATIC)
   If bSqlRows Then
      With RdoSoit
         While (Not .EOF)
            itso = "" & Trim(!itso)
            ITNum = "" & Trim(!ITNUMBER)
            itrev = "" & Trim(!itrev)
            WithDrawnQty = "" & Trim(!WITHDRAWN_QTY)
            payDoc = "" & Trim(!payment_doc_no)
            payDocNo = "" & Trim(!payment_doc_it_no)
            
            If (qoh < WithDrawnQty) Then
               If (bDisplay) Then
                  'MsgBox "Not sufficient QOH to cover the PartNumber:" & strPartNum
                  bDisplay = False
               End If
               
               sSql = "UPDATE FusionSOVOI  SET REMARKPS = 'Not sufficient QOH to cover the SO Number'" _
                        & " WHERE ITSO = " & itso & " AND ITNUMBER = " & ITNum & " AND ITREV = '" & itrev & "' " _
                        & " AND ITPART = '" & strPartNum & "' AND payment_doc_no = '" & payDoc & "' AND payment_doc_it_no = '" & payDocNo & "'"
               Debug.Print sSql
               clsADOCon.ExecuteSql sSql ' rdExecDirect
            Else
               Debug.Print "ITSO:" + CStr(itso) + ";ITNum:" + CStr(ITNum) + ";WithDrawnQty:" + CStr(WithDrawnQty)
               qoh = qoh - WithDrawnQty
            End If
            
            .MoveNext ' next so Number
         Wend  ' SO Item
         .Close
      End With
      Set RdoSoit = Nothing
   End If
   'On Error Resume Next
   Set RdoSoit = Nothing

End Function

Private Sub GetSoShipTo(ByVal strDocNum As String, ByRef sCust As String, ByRef strStnme As String, _
                        ByRef strSTAdr As String, ByRef strVia As String, ByRef strTerms As String)
   
   Dim lSoNum As Long
   Dim strITSO As String
   Dim Rdodis As ADODB.Recordset
   Dim RdoSto As ADODB.Recordset
   On Error GoTo DiaErr1
   
   lSoNum = 0
   ' get any one SO
   sSql = "select Top(1) ITSO from dbo.FusionSOVOI where Payment_doc_no = '" & strDocNum & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, Rdodis, ES_FORWARD)
   If bSqlRows Then
      With Rdodis
         strITSO = "" & Trim(!itso)
         ClearResultSet Rdodis
      End With
   Else
      MsgBox ("Primary SO number is Zero")
      
   End If
   Set Rdodis = Nothing

   If strITSO <> "" Then
      lSoNum = CLng(strITSO)
   End If
   
   ' Get the ship VIA information
   sSql = "SELECT SONUMBER,SOCUST, SOSTNAME,SOSTADR, SOVIA,SOSTERMS FROM SohdTable " _
          & "WHERE SONUMBER=" & lSoNum & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSto, ES_FORWARD)
   If bSqlRows Then
      With RdoSto
         strStnme = "" & Trim(!SOSTNAME)
         strSTAdr = "" & Trim(!SOSTADR)
         strVia = "" & Trim(!SOVIA)
         strTerms = "" & Trim(!SOSTERMS)
         sCust = "" & Trim(!SOCUST)
         ClearResultSet RdoSto
      End With
   Else
      lSoNum = 0
      strStnme = ""
      strSTAdr = ""
      strVia = ""
      strTerms = ""
      If (lSoNum = 0) Then
         MsgBox ("Primary SO number is Zero")
      End If
      
   End If
   Set RdoSto = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsoship "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Function GetValidITRev(strSoNum As String, ITNum As String, ByRef itrev As String _
                                    , ByRef ITQty As Currency) As String
   On Error GoTo DiaErr1
   
   Dim RdoRpt As ADODB.Recordset
   Dim RdoRptQ As ADODB.Recordset
   
   sSql = "SELECT MAX(ITREV) LstRev FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      
      If (Not IsNull(RdoRpt!LstRev)) Then
         itrev = Trim(RdoRpt!LstRev)
         ClearResultSet RdoRpt
      Else
         MsgBox "Couldn't find the Last Rev number - " & strSoNum & " Item :" & ITNum
      End If
   End If
   Set RdoRpt = Nothing
   
   sSql = "SELECT ITQTY FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITREV = '" & itrev & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRptQ, ES_FORWARD)
   If bSqlRows Then
      ITQty = Trim(RdoRptQ!ITQty)
      ClearResultSet RdoRptQ
   End If
   Set RdoRptQ = Nothing
   
   
   
   Exit Function

DiaErr1:
   sProcName = "GetValidITRev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Function GetSOITPrice(ByVal strSoNum As String, ByVal ITNum As Integer, _
                        ByVal rev As String, ByRef cPrice As Currency)
   On Error GoTo DiaErr1
   
   Dim RdoRpt As ADODB.Recordset
   
   sSql = "SELECT ITDOLLARS FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITREV = '" & rev & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      cPrice = Trim(RdoRpt!ITDOLLARS)
      ClearResultSet RdoRpt
   Else
      cPrice = 0
   End If
   Set RdoRpt = Nothing
   
   Exit Function

DiaErr1:
   sProcName = "GetSOITPrice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function GetPSfromDoc(strDocNum As String) As String
   
   Dim RdoDoc As ADODB.Recordset
   Dim strPS As String
   ' get any one SO
   sSql = "select DISTINCT PIPACKSLIP from psitTable where picomments like '%" & strDocNum & "%'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         strPS = "" & Trim(!PIPACKSLIP)
         GetPSfromDoc = strPS
         ClearResultSet RdoDoc
      End With
   Else
      MsgBox ("Document to Packslip is not Found")
      GetPSfromDoc = ""
   End If
   Set RdoDoc = Nothing

End Function

Private Sub PrintVOIPS(sPackSlip As String, Optional DontPrint As Boolean)
   
   Dim RdoPrint As ADODB.Recordset
   Dim bByte As Byte
   Dim iList As Integer
   Dim iLots As Integer
   Dim iRow As Integer
   
   Dim bInvType As Byte
   Dim bInvWritten As Byte
   Dim bLots As Byte
   Dim bLotsAct As Byte
   Dim bPrinted As Byte
   Dim bResponse As Byte
   Dim bMarkShipped As Byte
   
   Dim lSysCount As Long
   
   Dim cItmLot As Currency
   Dim cLotQty As Currency
   Dim cPartCost As Currency
   Dim cRemPqty As Currency
   Dim cPckQty As Currency
   Dim cQuantity As Currency
   
   'Costs
   Dim cMaterial As Currency
   Dim cLabor As Currency
   Dim cExpense As Currency
   Dim cOverhead As Currency
   Dim cHours As Currency
   
   Dim sMsg As String
   Dim sLot As String
   Dim sPart As String
   Dim cQtyLeft As Currency
   Dim cPsDtLotQty As Currency

   Dim vAdate As Variant
   Dim vPSdate As Variant
   Dim vCurrentdate As Variant
   Dim bPrevMonth As Boolean
   Dim bPrevPS As Boolean
   Dim iTotalItems As Integer
   
   On Error GoTo DiaErr2
   
   vAdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   vCurrentdate = vAdate
   vPSdate = Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   bPrevMonth = False
   bPrevPS = False
   
   bInvType = IATYPE_PackingSlip

   sJournalID = GetOpenJournal("IJ", Format(vAdate, "mm/dd/yy"))
   'sJournalID = GetOpenJournal("IJ", Format(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      bGoodJrn = 1
   Else
      If sJournalID = "" Then bGoodJrn = 0 Else bGoodJrn = 1
   End If
   If bGoodJrn = 0 Then
      MsgBox "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set The Pack Slip As Printed.", _
         vbExclamation, Caption
      Exit Sub
   End If
   
   
   
   'Was it Printed?
   bLotsAct = CheckLotStatus()
   'RdoQry1(0) = sPackSlip
   cmdObj1.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoPrint, cmdObj1, ES_FORWARD, True)
   
   If bSqlRows Then
      With RdoPrint
         If IsNull(!PSPRINTED) Then
            bPrinted = False
         Else
            bPrinted = True
         End If
         ClearResultSet RdoPrint
      End With
   End If
   Set RdoPrint = Nothing
   'If it was printed, then print again and bail out
   If bPrinted Then
      ' TODO: add logic to show/hide packslip
      If (False) Then
         MsgBox ("Packslip is already printed.")
      End If
      Exit Sub
   End If
   
   iTotalItems = GetItems(sPackSlip)
   If iTotalItems = 0 Then
      MsgBox "There Are No Unprinted Items On This Packing Slip.", vbInformation, Caption
      Exit Sub
   End If
   
   'quickly check that all lot-tracked items are available in sufficient quantity
   If bLotsAct Then
      For iRow = 1 To iTotalItems
         bLots = vItems(iRow, PS_LOTTRACKED)
         If bLots = 1 Then
            sPart = sPartGroup(iRow)
            cRemPqty = Val(vItems(iRow, PS_QUANTITY))
            'cLotQty = GetRemainingLotQty(sPart)
            cLotQty = GetLotRemainingQty(sPart)
            
            If cLotQty < cRemPqty Then
               If sMsg = "" Then
                  sMsg = "Insufficient Lot Quantity for the following parts:" & vbCrLf
               End If
               sMsg = sMsg & sPart & "    required=" & cRemPqty & " available=" & cLotQty & vbCrLf
            End If
         End If
      Next
      If sMsg <> "" Then
         
         ' TODO: add logic to show/hide packslip
         If (False) Then
            sMsg = sMsg & "The packing slip will not be printed."
            MsgBox sMsg, vbInformation, Caption
         End If
         Exit Sub
      End If
   End If
   
   'Packing slip hasn't been printed.  Confirm that printing is desired.
   ' TODO:
   If (False) Then
      sMsg = "Do You Want To Print This Pack Slip " & vbCrLf _
             & "And Adjust Inventory For The Parts?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      
      If bResponse = vbNo Then Exit Sub
   End If
   'Custom Transfer Option
   MouseCursor 13
   
   'determine lots from which items are drawn
   sSql = "delete from TempPsLots where PsNumber = '" & sPackSlip & "'" & vbCrLf _
      & "or DateDiff( hour, WhenCreated, getdate() ) > 24"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
   'if not using lots, start the transaction here
   
   For iRow = 1 To iTotalItems
      sPart = sPartGroup(iRow)
      bLots = vItems(iRow, PS_LOTTRACKED)
      
      'Lots
      'cRemPqty = Val(vItems(iRow, PS_QUANTITY))
      cQuantity = Format(Val(vItems(iRow, PS_QUANTITY)), ES_QuantityDataFormat)
      If bLotsAct = 1 And bLots = 1 Then
         '***** real lots
         'cLotQty = GetRemainingLotQty(sPart)
         cLotQty = GetLotRemainingQty(sPart)
         
         iLots = GetPartLots(sPart)
         cItmLot = 0
         cRemPqty = Format(Val(vItems(iRow, PS_QUANTITY)), ES_QuantityDataFormat)
         
         For iList = 1 To iLots
            If cRemPqty <= 0 Then
               Exit For
            End If
            cLotQty = Val(sLots(iList, 1))
            If cLotQty >= cRemPqty Then
               cPckQty = cRemPqty
               cLotQty = cLotQty - cRemPqty
               cRemPqty = 0
            Else
               cPckQty = cLotQty
               cRemPqty = cRemPqty - cLotQty
               cLotQty = 0
            End If
            If cPckQty > 0 Then
               cItmLot = cItmLot + cPckQty
               If cItmLot > Val(sLots(iList, 1)) Then cItmLot = Val(sLots(iList, 1))
               sLot = sLots(iList, 0)
               sSql = "INSERT INTO TempPsLots ( PsNumber, PsItem, LotID, LotQty , PartRef, LotItemID)" & vbCrLf _
                  & "Values ( '" & sPackSlip & "', " & vItems(iRow, PS_ITEMNO) & ", " _
                  & "'" & sLot & "', " & cPckQty _
                  & ", '" & sPart & "', '" & CStr(iList) & "') "
               clsADOCon.ExecuteSql sSql ', rdExecDirect
            End If
         Next
         ' If still we have remaining Qty we need to quit
         If (cRemPqty > 0) Then
         
            If (False) Then
            
               MsgBox "Not sufficient quantity for item " & vItems(iRow, PS_ITEMNO) _
                  & " part " & sPart & " available. " & vbCrLf _
                  & "It is short by (" & cRemPqty & ") quantity." & vbCrLf _
                  & "The packing slip will not be printed."
            End If
            GoTo NoCanDo
            
         End If
         
      End If
   Next
   
''''''''''''''''''''''''''''''''''''''

   'we have all the lots defined and there is no more user input,
   'so now ship the packing list in a single transaction
   
   'if using lots, start the transaction here
   If bLotsAct = 1 Then
      'clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
   End If
  
   'now that we're in the transaction, make sure that all selected lot quantities are available
   Dim rdo As ADODB.Recordset
   sSql = "select count(*) as ct" & vbCrLf _
      & "from TempPsLots tmp" & vbCrLf _
      & "join LohdTable lot on tmp.LotID = lot.LotNumber" & vbCrLf _
      & "where lot.LotRemainingQty < tmp.LotQty" & vbCrLf _
      & "and PSNUMBER = '" & sPackSlip & "'"
   If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
      If rdo!ct > 0 Then
         If bLotsAct = 1 Then
            clsADOCon.RollbackTrans
         End If
         MsgBox "Another user has allocated quantities from the lots selected.  Please try again."
         Exit Sub
      End If
   End If
   
   If bMarkShipped = 0 Then
      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
             & "PSSHIPPRINT=1,PSSHIPPED=0 WHERE " _
             & "PSNUMBER='" & sPackSlip & "' AND PSTYPE=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   Else
      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
             & "PSSHIPPRINT=1,PSSHIPPEDDATE='" & vAdate & "'," _
             & "PSSHIPPED=1 WHERE PSNUMBER='" & sPackSlip & "' " _
             & "AND PSTYPE=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   End If
   If clsADOCon.RowsAffected = 0 Then
      MouseCursor 0
      MsgBox "Could Not Update The Packing Slip. The Transaction " & vbCrLf _
         & "Has Been Aborted. Try Again In A Few Minutes.", _
         vbExclamation, Caption
      clsADOCon.RollbackTrans
      Exit Sub
   End If

   'Set date stamp for all items for this packing slip
   sSql = "UPDATE PsitTable SET PILOTNUMBER='" & Format(vAdate, "mm/dd/yy hh:mm") & "' " _
      & "WHERE PIPACKSLIP='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql ', rdExecDirect

   'Set all related SO items' ship dates
   sSql = "UPDATE SoitTable SET ITACTUAL='" & vAdate _
      & "',ITPSSHIPPED=" & bMarkShipped & " WHERE " _
      & "ITPSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql ', rdExecDirect

   'get next innumber to update wip accts later
   lSysCount = GetLastActivity + 1
  
   'loop through the packing slip items
   For iRow = 1 To iTotalItems
      'CurrentLotFailed = False
      'lCOUNTER = (GetLastActivity)
      'lSysCount = lCOUNTER + 1              'do above the loop
      cQuantity = Val(vItems(iRow, PS_QUANTITY))
      sPart = sPartGroup(iRow)
      bLots = vItems(iRow, PS_LOTTRACKED)
      
      ' set the cost as standard cost
      cPartCost = GetPartCost(sPart, ES_STANDARDCOST)
      
      vItems(iRow, PS_COST) = Format(cPartCost, ES_QuantityDataFormat)
      bByte = GetPartAccounts(sPart, sCreditAcct, sDebitAcct)
  
      Dim sSql1 As String
      Dim sSql2 As String
      Dim sSql3 As String

      'create inventory activities for lots for this packing slip item
      ' Fusion 5/15/2009
      sSql1 = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2, " & vbCrLf _
         & "INNUMBER,INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
         & "INPSNUMBER,INPSITEM,INLOTNUMBER,INSONUMBER,INSOITEM,INSOREV) " & vbCrLf _
         & "SELECT " & bInvType & ", '" & sPart & "', 'PACKING SLIP', "
         
    sSql2 = "tmp.PsNumber + '-' + " & "cast( tmp.PsItem as varchar(5) )," & vbCrLf _
         & "(SELECT MAX(INNUMBER) as num FROM INVATABLE) +  tmp.LotItemID," & vbCrLf _
         & "'" & vAdate & "', '" & vAdate & "',  -tmp.LotQty, " _
         & cPartCost & ", '" & sDebitAcct & "', '" & sCreditAcct & "', " & vbCrLf _
         & "'" & sPackSlip & "', " & Val(vItems(iRow, PS_ITEMNO)) & ", " _
         & "tmp.LotID, " & sSoItems(iRow, SOITEM_SO) & ", "

    sSql3 = sSoItems(iRow, SOITEM_ITEM) & ", '" & sSoItems(iRow, SOITEM_REV) & "'" & vbCrLf _
         & "FROM TempPsLots tmp" & vbCrLf _
         & "JOIN PartTable pt on tmp.PARTREF = pt.PartRef" & vbCrLf _
         & "WHERE tmp.PsNumber = '" & sPackSlip & "' AND tmp.PsItem = " & Trim(vItems(iRow, PS_ITEMNO))
         
      sSql = sSql1 & sSql2 & sSql3
         
      Debug.Print sSql
      
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      'insert lot items for this packing slip item
      sSql = "INSERT INTO LoitTable (LOINUMBER,LOIRECORD," _
         & "LOITYPE,LOIPARTREF,LOIADATE,LOIQUANTITY," & vbCrLf _
         & "LOIPSNUMBER,LOIPSITEM,LOICUST,LOIACTIVITY,LOICOMMENT) " & vbCrLf _
         & "SELECT tmp.LotID, dbo.fnGetNextLotItemNumber( tmp.LotID ), " _
         & bInvType & ", '" & sPart & "', '" & vAdate & "', " & vbCrLf _
         & "-tmp.LotQty, '" & sPackSlip & "', " _
         & Val(vItems(iRow, PS_ITEMNO)) & ", '" & sCustomer & "'," _
         & "ia.INNUMBER, 'Shipped Item'" & vbCrLf _
         & "FROM TempPsLots tmp" & vbCrLf _
         & "JOIN InvaTable ia ON ia.INPSNUMBER = tmp.PsNumber AND ia.INPSITEM = tmp.PsItem" & vbCrLf _
         & "and ia.INADATE = '" & vAdate & "' and ia.INLOTNUMBER = tmp.LotID" & vbCrLf _
         & "WHERE tmp.PsNumber = '" & sPackSlip & "' AND tmp.PsItem = " & Trim(vItems(iRow, PS_ITEMNO)) & vbCrLf _
         & "ORDER BY INNUMBER desc"
      
      Debug.Print sSql
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      'if there are quantities not covered by lots for automatic assignment, just create an ia record
      ' Fusion 5/15/2009
      
      sSql = "SELECT " & cQuantity & " + ( SELECT ISNULL( SUM( LOIQUANTITY ), 0 ) FROM LoitTable " & vbCrLf _
         & "WHERE LOIPSNUMBER = '" & sPackSlip & "' AND LOIPSITEM = " & Trim(vItems(iRow, PS_ITEMNO)) & " )"
      If clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD) Then
         Dim qtyLeft As Currency
         qtyLeft = rdo.Fields(0)
         If qtyLeft > 0 Then
            sSql1 = "INSERT INTO InvaTable (INTYPE,INPART,INREF1,INREF2, " & vbCrLf _
               & "INNUMBER,INPDATE,INADATE,INAQTY,INAMT,INCREDITACCT,INDEBITACCT," & vbCrLf _
               & "INPSNUMBER,INPSITEM,INLOTNUMBER,INSONUMBER,INSOITEM,INSOREV) " & vbCrLf _
               & "SELECT " & bInvType & ", '" & sPart & "', 'PACKING SLIP', "
               
            sSql2 = "'" & vItems(iRow, PS_PACKSLIPNO) & Trim(vItems(iRow, PS_ITEMNO)) & "', " & vbCrLf _
               & "(SELECT MAX(INNUMBER) as num FROM INVATABLE) + 1," & vbCrLf _
               & "'" & vAdate & "', '" & vAdate & "', " & -qtyLeft & ", " _
               & cPartCost & ", '" & sDebitAcct & "', '" & sCreditAcct & "', " & vbCrLf
               
               
            sSql3 = "'" & sPackSlip & "', " & Val(vItems(iRow, PS_ITEMNO)) & ", " _
               & "'No Lot Avail', " & sSoItems(iRow, SOITEM_SO) & ", " _
               & sSoItems(iRow, SOITEM_ITEM) & ", '" & sSoItems(iRow, SOITEM_REV) & "'"
               
            sSql = sSql1 & sSql2 & sSql3
            
            Debug.Print sSql
            clsADOCon.ExecuteSql sSql ', rdExecDirect
         End If
      End If
      rdo.Close
      
      'update quantities for part
      sSql = "UPDATE PartTable SET PAQOH=PAQOH - " & cQuantity & ", " _
             & "PALOTQTYREMAINING = PALOTQTYREMAINING - " & cQuantity & vbCrLf _
             & "WHERE PARTREF='" & sPart & "' "
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      AverageCost sPart
      
   
   Next
   
   'update remaining quantity in affected lots
   sSql = "UPDATE LohdTable" & vbCrLf _
      & "SET LOTREMAININGQTY = X.TOTAL" & vbCrLf _
      & "FROM LohdTable lt" & vbCrLf _
      & "JOIN (SELECT LOINUMBER, SUM(LOIQUANTITY) AS TOTAL FROM LOITTABLE GROUP BY LOINUMBER) AS X" & vbCrLf _
      & "ON X.LOINUMBER = LOTNUMBER" & vbCrLf _
      & "WHERE LOTNUMBER IN ( SELECT LotID from TempPsLots where PsNumber = '" & sPackSlip & "' )"
   clsADOCon.ExecuteSql sSql ', rdExecDirect
 
   'update ia costs from their associated lots
   Dim ia As New ClassInventoryActivity
   ia.UpdatePackingSlipCosts (sPackSlip)

   MouseCursor 0
   UpdateWipColumns lSysCount
   'clsADOCon.CommitTrans 'finally, commit the transaction
   SysMsg "Packing Slip Marked As Printed", True
   
   Exit Sub
   
DiaErr2:
   'On Error Resume Next
   
   MouseCursor 0
   sProcName = "PrintVOIPS"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   clsADOCon.RollbackTrans
   
   MsgBox "Couldn't complete inventory adjustments. " & vbCrLf _
      & "Packing list not printed.", vbExclamation, Caption
   Exit Sub
   
NoCanDo:
   MouseCursor 0
   Exit Sub
End Sub

Private Function GetItems(sPackSlip As String) As Integer
   Dim RdoItm As ADODB.Recordset
   Dim iRow As Integer
   Dim bLotsAct As Byte
   Dim iTotalItems As Integer
   
   Erase vItems
   Erase sSoItems
   Erase sPartGroup
   MouseCursor 13
   
   On Error GoTo DiaErr1
   iTotalItems = 0
   bLotsAct = CheckLotStatus()
   'RdoQry2(0) = sPackSlip
   cmdObj2.parameters(0).Value = sPackSlip
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, cmdObj2, ES_KEYSET, True)
   
   If bSqlRows Then
      'On Error Resume Next
      With RdoItm
         Do Until .EOF
            iRow = iRow + 1
            vItems(iRow, PS_PACKSLIPNO) = "" & Trim(!PIPACKSLIP) & "-"
            vItems(iRow, PS_ITEMNO) = Format(!PIITNO, "##0")
            vItems(iRow, PS_QUANTITY) = Format(!PIQTY, ES_QuantityDataFormat)
            'vItems(iRow, PS_PIPART) = "" & Trim(!PIPART)
            sPartGroup(iRow) = "" & Trim(!PIPART)
            vItems(iRow, PS_COST) = "0.000"
            If bLotsAct = 1 Then
               vItems(iRow, PS_LOTTRACKED) = !PALOTTRACK
            Else
               vItems(iRow, PS_LOTTRACKED) = 0
            End If
            vItems(iRow, PS_PARTNUM) = "" & Trim(!PartNum)
            sSoItems(iRow, SOITEM_SO) = str$(!PISONUMBER)
            sSoItems(iRow, SOITEM_ITEM) = str$(!PISOITEM)
            sSoItems(iRow, SOITEM_REV) = "" & Trim(!PISOREV)
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
   End If
   iTotalItems = iRow
   GetItems = iTotalItems
   Set RdoItm = Nothing
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetPartLots(sPartWithLot As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Erase sLots
   On Error GoTo DiaErr1
   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
          & "LOTREMAININGQTY>0 AND LOTAVAILABLE=1) "
   If bFIFO = 1 Then
      sSql = sSql & "ORDER BY LOTNUMBER ASC"
   Else
      sSql = sSql & "ORDER BY LOTNUMBER DESC"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            If (iRow >= 49) Then Exit Do
            iRow = iRow + 1
            sLots(iRow, 0) = "" & Trim(!lotNumber)
            sLots(iRow, 1) = Format$(!LOTREMAININGQTY, ES_QuantityDataFormat)
            .MoveNext
         Loop
         ClearResultSet RdoLots
      End With
      GetPartLots = iRow
   Else
      GetPartLots = 0
   End If
   Set RdoLots = Nothing
   Exit Function
   
DiaErr1:
   GetPartLots = 0
   
End Function


Private Function GetPartComm(ByVal strGetPart As String, _
            ByRef strPartNum As String, ByRef bComm As Boolean) As Byte
   Dim RdoPrt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   bComm = False
   strGetPart = Compress(strGetPart)
   If Len(strGetPart) > 0 Then
      sSql = "SELECT PARTNUM,PADESC,PAEXTDESC,PAPRICE,PAQOH," _
             & "PACOMMISSION FROM PartTable WHERE PARTREF='" & strGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_STATIC)
      If bSqlRows Then
         With RdoPrt
            strPartNum = "" & Trim(!PartNum)
            If !PACOMMISSION = 1 Then bComm = True _
                               Else bComm = False
            GetPartComm = 1
            ClearResultSet RdoPrt
         End With
      Else
         GetPartComm = 0
      End If
      'On Error Resume Next
      Set RdoPrt = Nothing
   Else
      GetPartComm = 0
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetPartComm"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetDocNumCuffOff() As String

   Dim RdoDoc As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT payment_doc_no FROM Voipdocnumcutoff"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc)
   If bSqlRows Then
      With RdoDoc
         GetDocNumCuffOff = Trim(!payment_doc_no)
         ClearResultSet RdoDoc
      End With
   Else
      GetDocNumCuffOff = ""
   End If
   Set RdoDoc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetDocNumCuffOff"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'GetCustomerRef = False
   DoModuleErrors Me
   

End Function



Private Sub GetCustomerRef(ByRef strCusFullName As String, ByRef strCusName As String)

   Dim RdoCus As ADODB.Recordset
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT CUREF FROM CustTable WHERE CUNAME = '" & strCusFullName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCus)
   If bSqlRows Then
      With RdoCus
         strCusName = Trim(!CUREF)
         ClearResultSet RdoCus
      End With
   Else
      strCusName = ""
   End If
   Set RdoCus = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetCustomerRef"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   'GetCustomerRef = False
   DoModuleErrors Me
   

End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   
End Sub

Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst, False
   
End Sub


Private Sub cmdSearch_Click()
   ExpDlg.Filter = "Excel File (*.xls) | *.xls"
   ExpDlg.ShowOpen
   If ExpDlg.FileName = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = ExpDlg.FileName
   End If

End Sub


Private Sub CmdSelAll_Click()
   Dim iList As Integer
   
   For iList = 1 To Grd.Rows - 1
       Grd.Col = 0
       Grd.Row = iList
       ' Only if the part is checked
       If Grd.CellPicture = Chkno.Picture Then
           Set Grd.CellPicture = Chkyes.Picture
       End If
   Next
End Sub



Private Sub Form_Activate()
Dim bSoAdded As Byte
   MdiSect.lblBotPanel = Caption
   
   ' Only if the import table is full
   FillGrid
  
   If bOnLoad Then
       bOnLoad = 0
   End If
    
    
   MouseCursor (0)

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hWnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   Call WheelUnHook(Me.hWnd)
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   Dim iChar As Integer
    
   With Grd
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1


      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "PartNumber"
      .Col = 1
      .Text = "Document Number"
      .Col = 2
      .Text = "Document Item"
      .Col = 3
      .Text = "Amount"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "Issued Date"
      .Col = 6
      .Text = "SO Number"
      .Col = 7
      .Text = "SO Item"
      .Col = 8
      .Text = "SO Rev"
      .Col = 9
      .Text = "SO Qty"
      .Col = 10
      .Text = "Sched Date"
      .Col = 11
      .Text = "SO Remarks"
      .Col = 12
      .Text = "PS Remarks"
      
      .ColWidth(0) = 2050
      .ColWidth(1) = 1800
      .ColWidth(2) = 1000
      .ColWidth(3) = 1000
      .ColWidth(4) = 800
      .ColWidth(5) = 1000
      .ColWidth(6) = 1050
      .ColWidth(7) = 850
      .ColWidth(8) = 500
      .ColWidth(9) = 800
      .ColWidth(10) = 800
      .ColWidth(11) = 1050
      .ColWidth(12) = 1050
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Function ReadAllFields(ByVal iIndex As Integer, ByRef ws As Worksheet)

    Dim iCols As Integer
        
    Fields(0) = ""
    While (iCols < 150)
        Fields(iCols) = ""
        iCols = iCols + 1
    Wend
    
    iCols = 0
    If (iIndex > 0 And Not ws Is Nothing) Then
        
        While (iCols < 150)
            Fields(iCols) = ws.Cells(iIndex, iCols + 1)
            iCols = iCols + 1
        Wend
    End If

End Function

Function RemoveCommas(sNextLine As String) As String
    
    Dim length As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim strStrip As String
    lngStart = 1
    lngStart = InStr(lngStart, sNextLine, """")
    'lngEnd = InStr(lngStart + 1, sNextLine, """")
    
    length = Len(sNextLine)

    'strStrip = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
    
    'RemoveCommas = strStrip
    
    While (lngStart > 0)
        lngEnd = InStr(lngStart + 1, sNextLine, """")
        If (lngEnd > 0) Then
            'ReplaceComma sNextLine, lngStart, lngEnd
            sNextLine = Left$(sNextLine, lngStart) & Right$(sNextLine, (length - lngEnd) + 1)
        End If
        lngStart = InStr(1, sNextLine, """")
    Wend
    
    

End Function

Function ReplaceComma(sNextLine As String, lngStart As Long, lngEnd As Long)
    Dim i As Long
    i = lngStart
    While ((i <= lngEnd) And i > 0)
        i = InStr(i, sNextLine, ",")
        If (i > 0 And i <= lngEnd) Then
            sNextLine = Replace(sNextLine, ",", "-", i, 1)
            i = i + 1
        End If
    Wend

End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
   If OptSoXml.Value = vbUnchecked Then FormUnload
   Set cmdObj1 = Nothing
   Set cmdObj2 = Nothing
    'FormUnload
    Set SaleSLf15a = Nothing
End Sub

'Private Sub grd_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
'      Grd.Col = 0
'      If Grd.Row >= 1 Then
'         If Grd.Row = 0 Then Grd.Row = 1
'         If Grd.CellPicture = Chkyes.Picture Then
'            Set Grd.CellPicture = Chkno.Picture
'         Else
'            Set Grd.CellPicture = Chkyes.Picture
'         End If
'      End If
'    End If
'
'
'End Sub


Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd.Rows - 1
        Grd.Col = 0
        Grd.Row = iList
        ' Only if the part is checked
        If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
        End If
    Next
End Sub

'
'Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Grd.Col = 0
'   If Grd.Row >= 1 Then
'      If Grd.Row = 0 Then Grd.Row = 1
'      If Grd.CellPicture = Chkyes.Picture Then
'         Set Grd.CellPicture = Chkno.Picture
'      Else
'         Set Grd.CellPicture = Chkyes.Picture
'      End If
'   End If
'End Sub


Private Function CheckForCustomerPO(ByVal strCustomer As String, ByVal strPONum As String) As Byte
   On Error GoTo modErr1
   Dim RdoCpo As ADODB.Recordset
   If Trim(strPONum) = "" Then
      CheckForCustomerPO = 0
   Else
      sSql = "Qry_GetCustomerPo '" & Compress(strCustomer) _
             & "','" & Trim(strPONum) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpo, ES_FORWARD)
      If bSqlRows Then
         With RdoCpo
            CheckForCustomerPO = 1
            ClearResultSet RdoCpo
         End With
      End If
   End If
   Set RdoCpo = Nothing
   Exit Function
   
modErr1:
   sProcName = "CheckForCustomerPO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   CheckForCustomerPO = 0
   DoModuleErrors MdiSect.ActiveForm
   
End Function




Private Function MakeAddress(strShipName2 As String, strStreet As String, strStreetSup1 As String, _
                  strStreetSup2 As String, strCity As String, strRegionCode As String, _
                  strPostalCode As String, ByRef strNewAddress As String)

   strNewAddress = ""
   
   ' MM not needed
   'If (strShipName2 <> "") Then strNewAddress = strNewAddress & strShipName2 & vbCrLf
   If (strStreet <> "") Then strNewAddress = strNewAddress & strStreet & vbCrLf
   If (strStreetSup1 <> "") Then strNewAddress = strNewAddress & strStreetSup1 & vbCrLf
   If (strStreetSup2 <> "") Then strNewAddress = strNewAddress & strStreetSup2 & vbCrLf
   
   ' moved Region ==> shiped
   'If (strRegionCode <> "") Then strNewAddress = strNewAddress & strRegionCode & vbCrLf
   If (strCity <> "") Then strNewAddress = strNewAddress & strCity
   
   If (strPostalCode <> "") Then
      If (strRegionCode <> "") Then
         strNewAddress = strNewAddress & ", " & IIf((strRegionCode <> ""), strRegionCode, "") & " - " & strPostalCode
      Else
         strNewAddress = strNewAddress & " - " & strPostalCode
      End If
   End If

End Function


