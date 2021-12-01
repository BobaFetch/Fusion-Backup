VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf15b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PackSlip & Print VOI SO"
   ClientHeight    =   12450
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   17115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12450
   ScaleWidth      =   17115
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPSDte 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdPostInvoice 
      Caption         =   "Post Invoice"
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
      Left            =   14520
      TabIndex        =   18
      ToolTipText     =   "Post the invoice"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   4680
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   1320
      Width           =   915
   End
   Begin VB.CommandButton cmdExportSel 
      Caption         =   "Export Selected"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   11520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdCreatePS 
      Caption         =   "Create PS"
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
      Left            =   14520
      TabIndex        =   13
      ToolTipText     =   "Apply VOI Consumption"
      Top             =   2280
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
      Left            =   14520
      TabIndex        =   12
      ToolTipText     =   "Apply VOI Consumption"
      Top             =   3120
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
      Left            =   14520
      TabIndex        =   11
      ToolTipText     =   " Select All"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton cmbExport 
      Caption         =   "Export All"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   11520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   11040
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   11040
      Width           =   4695
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   11640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cmbDoc 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1320
      Width           =   2760
   End
   Begin VB.CommandButton cmdInvoice 
      Caption         =   "Create Invoice"
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
      Left            =   14520
      TabIndex        =   3
      ToolTipText     =   "Create auto Invoice "
      Top             =   3840
      Width           =   2280
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf15b.frx":0000
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
      FormDesignHeight=   12450
      FormDesignWidth =   17115
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
      Left            =   14040
      Top             =   11040
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
   Begin MSComctlLib.ProgressBar prg2 
      Height          =   300
      Left            =   9240
      TabIndex        =   15
      Top             =   10800
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   8295
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   14631
      _Version        =   393216
      Rows            =   3
      Cols            =   18
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   315
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Caption         =   "Pack Slip Date"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Excel File"
      Height          =   285
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   11040
      Width           =   1305
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   13080
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "VOI Document"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15b.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf15b.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "SaleSLf15b"
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

Dim vItems(1000, 7) As Variant   ' 800 -> 1000
Dim sPartGroup(1000) As String '9/23/04 Compressed PartTable!PARTREF 800 -> 1000
Dim sSoItems(500, 3) As String 'Nathan 3/10/04 2017 300 -> 500
Dim sLots(60, 2) As String ' 50 -> 60
Dim sCustomer As String
Dim sCreditAcct As String
Dim sDebitAcct As String

Const SOITEM_SO = 0 ' string of PISONUMBER
Const SOITEM_ITEM = 1 ' string of PISOITEM
Const SOITEM_REV = 2 ' string of PISOREV

' Invoice global variables
Dim bGoodAct As Byte
Dim bGoodPs As Byte
Dim iRow As Integer
Dim lSo As Long
Dim lNewInv As Long
Dim lNextInv As Long
Dim lSalesOrder As Long
Dim iTotalItems As Integer
Dim iTotalChk As Integer
Dim sPsCust As String
Dim sPsStadr As String
Dim sPackSlip As String
Dim sAccount As String
Dim sMsg As String
Dim sDocNumber As String


Dim sTaxAccount As String
Dim sTaxState As String
Dim sTaxCode As String
Dim nTaxRate As Currency
Dim cFREIGHT As Currency

Dim cTax As Currency
Dim sType As String * 1
Dim currentCust As String

Dim sCrCashAcct As String
Dim sCrDiscAcct As String
Dim sCrExpAcct As String
Dim sSJARAcct As String
Dim sCrRevAcct As String
Dim sCrCommAcct As String


' Sales journal
Dim sCOSjARAcct As String
Dim sCOSjINVAcct As String
Dim sCOSjNFRTAcct As String
Dim sCOSjTFRTAcct As String
Dim sCOSjTaxAcct As String
Public lCurrInvoice As Long
Dim vpItems(1000, 12) As Variant


Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim strXML As String
Dim bNewImport As Boolean
Dim ExtName As String

Dim sJournalID As String

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
   Dim sFieldsToExport(16) As String
   
   Dim strDocNum As String
   strDocNum = cmbDoc
   
   AddFieldsToExport sFieldsToExport
   
   Dim RdoPOIssued As ADODB.Recordset
   

   
      sSql = "SELECT b.MATL_NO,b.PAYMENT_DOC_NO, b.PAYMENT_DOC_IT_NO, b.ISSUE_AMOUNT, a.WITHDRAWN_QTY, b.ISSUE_DTE, " _
               & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
               & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(a.ITPSNUMBER, '') ITPSNUMBER," _
               & " ISNULL(a.ITPSITEM,'') ITPSITEM, ISNULL(a.ITPSSHIPPED,0) ITPSSHIPPED," _
               & " ISNULL(b.Remarks,'') Remarks, ISNULL(a.REMARKPS,'') REMARKPS FROM VOIPmtIss AS b LEFT OUTER JOIN  FusionSOVOI AS a " _
               & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
               & " and a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO" _
               & " WHERE b.PAYMENT_DOC_NO = '" & strDocNum & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   'Grd.Rows = 1
   Debug.Print sSql
   
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_STATIC)
   
   If bSqlRows Then
      sFileName = txtFilePath.Text
      SaveAsExcel RdoPOIssued, sFieldsToExport, sFileName
      MsgBox "Saved the file.", vbOKOnly
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
   sFieldsToExport(i + 5) = "ISSUE_DTE"
   sFieldsToExport(i + 6) = "ITSO"
   sFieldsToExport(i + 7) = "ITNUMBER"
   sFieldsToExport(i + 8) = "ITREV"
   sFieldsToExport(i + 9) = "ITQty"
   sFieldsToExport(i + 10) = "ITSCHED"
   
   sFieldsToExport(i + 11) = "ITPSNUMBER"
   sFieldsToExport(i + 12) = "ITPSITEM"
   sFieldsToExport(i + 13) = "ITPSSHIPPED"
   
   sFieldsToExport(i + 14) = "REMARKS"
   sFieldsToExport(i + 15) = "REMARKPS"
   
End Function

Private Sub cmbPSDte_DropDown()
   ShowCalendarEx Me
End Sub

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

   For iList = 1 To Grd.rows - 1
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

   Dim DocNo As String
   
   DocNo = cmbDoc
   

   Dim RdoPOIssued As ADODB.Recordset
   sSql = "SELECT DISTINCT a.ITPART,a.PAYMENT_DOC_NO, a.PAYMENT_DOC_IT_NO, ISNULL(b.ISSUE_AMOUNT, 0) ISSUE_AMOUNT, a.WITHDRAWN_QTY, ISNULL(b.ISSUE_DTE, '') ISSUE_DTE, " _
            & " ISNULL(a.ITSO, '') ITSO, ISNULL(a.ITNUMBER, 0) ITNUMBER , ISNULL(a.ITREV, '') ITREV, " _
            & " ISNULL(a.ITQty, 0) ITQty, ISNULL(a.ITSCHED, '') ITSCHED, ISNULL(a.ITPSNUMBER, '') ITPSNUMBER," _
            & " ISNULL(a.ITPSITEM,'') ITPSITEM, ISNULL(a.ITPSSHIPPED,0) ITPSSHIPPED, " _
            & " ISNULL(a.INNO,'') INNO, ISNULL(a.CASHRECEIPT,0) CASHRECEIPT, " _
            & " ISNULL(b.Remarks,'') Remarks, ISNULL(a.REMARKPS,'') REMARKPS " _
         & " FROM FusionSOVOI AS a LEFT OUTER JOIN  VOIPmtIss AS b " _
            & "   ON  a.itPart = b.matl_no AND a.PAYMENT_DOC_NO = b.PAYMENT_DOC_NO " _
            & " AND a.PAYMENT_DOC_IT_NO = b.PAYMENT_DOC_IT_NO WHERE a.PAYMENT_DOC_NO = '" & DocNo & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPOIssued, ES_DYNAMIC)

   Grd.rows = 1
   Debug.Print sSql
   
   If bSqlRows Then
   With RdoPOIssued
      While Not .EOF

         Grd.rows = Grd.rows + 1
         Grd.Row = Grd.rows - 1
         
         Grd.Col = 0
         Grd.Text = Trim(!ITPART)
         
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
         Grd.Text = Trim(!ITPSNUMBER)
         
         Grd.Col = 12
         Grd.Text = Trim(!ITPSITEM)
         
         Grd.Col = 13
         Grd.Text = IIf(Trim(!ITPSSHIPPED) = 0, "No", "Yes")
         
         Grd.Col = 14
         Grd.Text = Trim(!INNO)
         
         Grd.Col = 15
         Grd.Text = Trim(!CASHRECEIPT)
         
         Grd.Col = 16
         Grd.Text = Trim(!Remarks)
         
         Grd.Col = 17
         Grd.Text = Trim(!REMARKPS)

         .MoveNext
      Wend
      .Close
      ClearResultSet RdoPOIssued
      End With
   End If
   Set RdoPOIssued = Nothing

End Sub


Private Sub cmdInvoice_Click()
   'make sure there is a journal open for the posting date
   'CurrentJournal "SJ", txtDte.Text, sJournalID
   
   Dim strDate As String
   strDate = cmbPSDte 'Format(Now, "mm/dd/yyyy")
   
   sJournalID = GetOpenJournal("SJ", strDate)
   If sJournalID = "" Then
      MsgBox "No Open Sales Journal Found For " _
         & strDate & " .", vbInformation, Caption
      Exit Sub
   End If
   
   'end date must be greater than or equal to start date
'   If DateDiff("d", CDate(txtstart.Text), CDate(txtEnd.Text)) < 0 Then
'      MsgBox "Ending date must be greater than or equal to start date", vbInformation, Caption
'      Exit Sub
'   End If
'
'   'make sure the posting date is greater than or equal to the end date
'   If DateDiff("d", CDate(txtEnd.Text), CDate(txtDte.Text)) < 0 Then
'      MsgBox "Posting date must be greater than or equal to end date", vbInformation, Caption
'      Exit Sub
'   End If
   
   
   CheckLoop
   
End Sub
Private Function GetPSfromDocWithOpenInv(strDocNum As String) As String
   
   Dim RdoDoc As ADODB.Recordset
   Dim strPS As String
   ' get any one SO
   sSql = "select DISTINCT ITPSNUMBER from FusionSOVOI where payment_doc_no like '%" & strDocNum & "%' " _
            & " AND ITPSSHIPPED = 1 AND (INNO IS NULL or INNO = '')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDoc, ES_FORWARD)
   If bSqlRows Then
      With RdoDoc
         strPS = "" & Trim(!ITPSNUMBER)
         GetPSfromDocWithOpenInv = strPS
         ClearResultSet RdoDoc
      End With
   Else
      MsgBox ("There are no open Invoices")
      GetPSfromDocWithOpenInv = ""
   End If
   Set RdoDoc = Nothing

End Function


Private Sub CheckLoop()
   Dim i As Integer
   Dim a As Integer
   Dim sMsg As String
   Dim bResponse As Byte
   Dim bErrOccur As Byte
   Dim lMinInv As Long
   Dim lMaxInv As Long
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   
   sMsg = "Create Invoices for the selected Packing Slips?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      MsgBox "Transaction Cancelled", vbExclamation, Caption
      Exit Sub
   End If
   GetSJAccounts
   sDocNumber = cmbDoc
   
   Grd.Row = 1
   Grd.Col = 11
   sPackSlip = GetPSfromDocWithOpenInv(sDocNumber)    'Grd.Text
   bGoodPs = GetPackslip(sPackSlip)
   Grd.Col = 6
   lSo = Grd.Text
   
   
   If bGoodPs Then
      lNewInv = AddInvoice(sPackSlip)
      If (lNewInv <> 0) Then
         MasterFunction
      End If

   End If

'   For i = 1 To Grd.Rows - 1
'      If bChecks(i) Then
'         Grd.Row = i
'         Grd.Col = 11
'         sPackSlip = Grd.Text
'         bGoodPs = GetPackslip(sPackSlip)
'         Grid1.Col = 6
'         lSo = Grid1.Text
'
'         ' Now the functions
'         If bGoodPs Then
'            If lMinInv = 0 Then
'               lMinInv = AddInvoice(sPackSlip)
'            Else
'               lMaxInv = AddInvoice(sPackSlip)
'            End If
'
'            If (lNewInv <> 0) Then
'               MasterFunction lSo
'            End If
'
'            If Err <> 0 Then
'               bErrOccur = True
'            End If
'         Else
'            bErrOccur = True
'         End If
'      End If
'   Next
'   If lMaxInv = 0 Then lMaxInv = lMinInv
'
'   If bErrOccur = True Then
'      sMsg = "Errors Occured Invoicing One or More Packing Slips"
'      MsgBox sMsg, vbExclamation, Caption
'   Else
'
'      If lMinInv <> lMaxInv Then
'         MsgBox "Invoices " & Format(lMinInv, "000000") & " Through " _
'            & Format(lMaxInv, "000000") _
'            & " Created From Packing Slips.", _
'            vbInformation, Caption
'      Else
'         MsgBox "Invoice " & Format(lMinInv, "000000") & " Created.", _
'            vbInformation, Caption
'      End If
'      iTotalChk = 0
'   End If
'   FillGrid
   Exit Sub
   
DiaErr1:
   sProcName = "CheckLoop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function AddInvoice(strPSNum As String) As Long
   AddInvoice = GetNextInvoice(strPSNum)
   On Error Resume Next
   Err.Clear
   lNewInv = AddInvoice
   'Use TM so that it won't show and can be safely deleted.
   sSql = "INSERT INTO CihdTable (INVNO,INVTYPE,INVSO,INVCANCELED) " _
          & "VALUES(" & lNewInv & ",'TM'," _
          & Val(sPackSlip) & ",0)"
   clsADOCon.ExecuteSql sSql
   Exit Function
   
DiaErr1:
   sProcName = "addinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub MasterFunction()
   Dim rdoTSo As ADODB.Recordset
   Dim sSO As String
   Dim sMsg As String
   
   On Error Resume Next
   Err.Clear
   If bGoodAct = False Then
      MsgBox "One Or More Journal Accounts Are Not Registered." & vbCr _
         & "Please Install All Accounts In the Company Setup.", _
         vbInformation, Caption
   Else
      sSql = "SELECT CUTAXCODE,SOTYPE,CUNICKNAME,SOTAXABLE " _
             & "FROM CustTable INNER JOIN " _
             & "SohdTable ON CustTable.CUREF = SohdTable.SOCUST " _
             & "WHERE SONUMBER = " & lSo & " "
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTSo)
      If Err > 0 Then
         MsgBox "This Packing Slip Cannot Be Invoiced.", _
            vbInformation, Caption
         Exit Sub
      End If
      currentCust = Compress(rdoTSo!CUNICKNAME)
      If rdoTSo!SOTAXABLE = 1 And ("" & Trim(rdoTSo!CUTAXCODE)) = "" Then
         sSO = Trim(rdoTSo!SOTYPE) & Format(lSo, "000000")
         sMsg = "Packslip " & sPackSlip & " Primary Sales Order " _
                & sSO & " Is Taxable." & vbCrLf & rdoTSo!CUNICKNAME & " Has No Tax Code Assigned. " _
                & "Packslip Cannot Be Invoiced."
         MsgBox sMsg, vbInformation, Caption
      Else
         GetPSItems
         If iTotalItems = 0 Then
            sMsg = "No Items On This Packing Slip " & sPackSlip & " To Invoice."
            MsgBox sMsg, vbInformation, Caption
         Else
            UpdateInvoice
         End If
      End If
      Set rdoTSo = Nothing
   End If
End Sub

Private Sub GetPSItems()
   Dim i As Integer
   Dim RdoPsi As ADODB.Recordset
   MouseCursor 13
   
   cTax = 0
   On Error GoTo DiaErr1
   
   Erase vItems
   Erase vpItems
   
   ' Update the packslip dollars from sales order
   sSql = "UPDATE PsitTable SET PISELLPRICE=" _
          & "ITDOLLARS FROM PsitTable,SoitTable WHERE " _
          & "(PISONUMBER=ITSO AND PISOITEM=ITNUMBER " _
          & "AND PISOREV=ITREV) AND PIPACKSLIP='" & Trim(sPackSlip) & "'"
   clsADOCon.ExecuteSql sSql
   
   ' Return packslip items
'   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM," _
'          & "PISOREV,PISELLPRICE,PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT," _
'          & "PASTDCOST FROM PsitTable,PartTable WHERE PIPART=" _
'          & "PARTREF AND PIPACKSLIP= '" & sPackSlip & "' ORDER BY PISONUMBER,PISOITEM"
   sSql = "SELECT PIPACKSLIP,PIQTY,PIPART,PISONUMBER,PISOITEM,PISOREV,PISELLPRICE," & vbCrLf _
          & "PARTREF,PARTNUM,PALEVEL,PAPRODCODE,PATAXEXEMPT,PASTDCOST,SOTAXABLE" & vbCrLf _
          & "FROM PsitTable" & vbCrLf _
          & "JOIN PartTable ON PIPART=PARTREF" & vbCrLf _
          & "JOIN  SohdTable on PISONUMBER = SONUMBER" & vbCrLf _
          & "WHERE PIPACKSLIP= '" & sPackSlip & "'" & vbCrLf _
          & "ORDER BY PISONUMBER,PISOITEM"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsi, ES_DYNAMIC)
   If bSqlRows Then
      With RdoPsi
         ' Load invoice items into vItems array for use in
         ' IpdateInvoice subroutine
         Do Until .EOF
            i = i + 1
            vpItems(i, 0) = Format(!PISONUMBER, SO_NUM_FORMAT)
            vpItems(i, 1) = Format(!PISOITEM, "##0")
            vpItems(i, 2) = Trim(!PISOREV)
            vpItems(i, 3) = "" & Trim(!PartRef)
            'vItems(i, 4) = Format(!PIQTY, "#####0.000")
            vpItems(i, 4) = !PIQTY
            vpItems(i, 5) = "" & sAccount
            vpItems(i, 6) = "" & !PIPART
            vpItems(i, 7) = "" & !PALEVEL
            vpItems(i, 8) = "" & Trim(!PAPRODCODE)
            'vItems(i, 9) = Format(!PASTDCOST, "#####0.000")
            vpItems(i, 9) = !PASTDCOST
            vpItems(i, 10) = 1
            'vItems(i, 11) = Format(!PISELLPRICE, "#####0.000")
            vpItems(i, 11) = !PISELLPRICE
            vpItems(i, 12) = !PATAXEXEMPT
            If vpItems(i, 12) = 0 And !SOTAXABLE = 1 Then
               cTax = cTax + ((vpItems(i, 11) * vpItems(i, 4)) * (nTaxRate / 100))
            End If
            .MoveNext
         Loop
      End With
   End If
   iTotalItems = i
   Set RdoPsi = Nothing
   MouseCursor 0
   Exit Sub
   ' Error handeling
DiaErr1:
   sProcName = "getpsitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub UpdateInvoice()
   Dim bByte As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim iTrans As Integer
   Dim iRef As Integer
   Dim iLevel As Integer
   Dim cCost As Currency
   Dim nLDollars As Single
   Dim nTdollars As Single
   Dim sPart As String
   Dim sProd As String
   Dim sMsg As String
   ' Accounts
   Dim sPartRevAcct As String
   Dim sPartCgsAcct As String
   Dim sREVAccount As String
   Dim sDisAccount As String
   Dim sCGSMaterialAccount As String
   Dim sCGSLaborAccount As String
   Dim sCGSExpAccount As String
   Dim sCGSOhAccount As String
   Dim sInvMaterialAccount As String
   Dim sInvLaborAccount As String
   Dim sInvExpAccount As String
   Dim sInvOhAccount As String
   Dim RdoTax As ADODB.Recordset
   
   ' BnO Taxes
   Dim nRate As Single
   Dim sType As String
   Dim sState As String
   Dim sCode As String
   Dim sPost As String
   Dim sTemp As String
   
   On Error GoTo DiaErr1
   
   ' Look For Accounts ?
   If sJournalID <> "" Then
      iTrans = GetNextTransaction(sJournalID)
   End If
   If iTrans > 0 Then
      bByte = True
      For i = 1 To iTotalItems
         If Val(vpItems(i, 8)) > 0 Then
            sPart = vpItems(i, 3)
            iLevel = Val(vpItems(i, 7))
            sProd = vpItems(i, 8)
            bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
                    sPartRevAcct, , sPartCgsAcct)
         End If
         If bByte = False Then Exit For
      Next
   End If
   
   lCurrInvoice = lNextInv
   sPost = cmbPSDte  'Format(Now, "mm/dd/yyyy")
   
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "UPDATE PshdTable SET PSINVOICE=" & lCurrInvoice & " " _
          & "WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   ' Update lot record
   sSql = "UPDATE LoitTable SET LOICUSTINVNO=" & lCurrInvoice _
          & ",LOICUST='" & sPsCust & "' WHERE LOIPSNUMBER = '" & sPackSlip & "'"
   clsADOCon.ExecuteSql sSql
   
   For i = 1 To iTotalItems
      ' Running invoice total
      nTdollars = nTdollars + (Val(vpItems(i, 4)) * Val(vpItems(i, 11)))
      ' Part accounts
      sPart = vpItems(i, 3)
      sProd = vpItems(i, 8)
      iLevel = Val(vpItems(i, 7))
      bByte = GetPartInvoiceAccounts(sPart, iLevel, sProd, _
              sREVAccount, sDisAccount, sCGSMaterialAccount, sCGSLaborAccount, _
              sCGSExpAccount, sCGSOhAccount, sInvMaterialAccount, sInvLaborAccount, _
              sInvExpAccount, sInvOhAccount)
      
      ' BnO tax
      sCode = ""
      nRate = 0
      sState = ""
      sType = ""
      
      'new functions (12/16/10)
      GetPartBnO vpItems(i, 3), nRate, sCode, sState, sType
      If sCode = "" Then
         GetCustBnO sPsCust, nRate, sCode, sState, sType
      End If
      
      
      ' Update the sales order items (revised 12/16/03)
      sSql = "UPDATE SoitTable SET " _
             & "ITINVOICE=" & lCurrInvoice & "," _
             & "ITREVACCT='" & sREVAccount & "'," _
             & "ITCGSACCT='" & sCOSjARAcct & "'," _
             & "ITBOSTATE='" & sState & "'," _
             & "ITBOCODE='" & sCode & "'," _
             & "ITSLSTXACCT='" & sTaxAccount & "'," _
             & "ITTAXCODE='" & sTaxCode & "'," _
             & "ITSTATE='" & sTaxState & "'," _
             & "ITTAXRATE=" & nTaxRate & "," _
             & "ITDISCAMOUNT=0.0," _
             & "ITTAXAMT=" & CCur((nTaxRate / 100) * (Val(vpItems(i, 4)) * Val(vpItems(i, 11)))) & " " _
             & "WHERE ITSO=" & Val(vpItems(i, 0)) & " AND " _
             & "ITNUMBER=" & Val(vpItems(i, 1)) & " AND " _
             & "ITREV='" & vpItems(i, 2) & "' "
      
      Debug.Print sSql
      clsADOCon.ExecuteSql sSql
      
      'Journal entries
      nLDollars = (Val(vpItems(i, 4)) * Val(vpItems(i, 11)))
      cCost = (Val(vpItems(i, 4)) * Val(vpItems(i, 9)))
      
      ' Debit A/R (+)
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
             & "DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & CCur(nLDollars) & ",'" _
             & sCOSjARAcct & "','" _
             & sPart & "'," _
             & vpItems(i, 0) & "," _
             & vpItems(i, 1) & ",'" _
             & vpItems(i, 2) & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      '& Format(Now, "mm/dd/yy") & "'," _

      ' Credit Revenue (-)
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCPARTNO,DCSONUMBER,DCSOITNUMBER,DCSOITREV," _
             & "DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & CCur(nLDollars) & ",'" _
             & sREVAccount & "','" _
             & sPart & "'," _
             & vpItems(i, 0) & "," _
             & vpItems(i, 1) & ",'" _
             & vpItems(i, 2) & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   Next
   
   ' Tax and freight
   
   If cFREIGHT > 0 Then
      
      ' Debit A/R Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjARAcct & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
      
      ' Credit Freight
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cFREIGHT & ",'" _
             & sCOSjNFRTAcct & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   'cTax = CCur(txtTax)
   If cTax > 0 Then
      ' Debit A/R Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCDEBIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjARAcct & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
       
     clsADOCon.ExecuteSql sSql
      ' Credit Taxes
      iRef = iRef + 1
      sSql = "INSERT INTO JritTable (DCHEAD,DCTRAN,DCREF,DCCREDIT," _
             & "DCACCTNO,DCCUST,DCDATE,DCINVNO) " _
             & "VALUES('" & Trim(sJournalID) & "'," _
             & iTrans & "," _
             & iRef & "," _
             & cTax & ",'" _
             & sCOSjTaxAcct & "','" _
             & Trim(sPsCust) & "','" _
             & cmbPSDte & "'," _
             & lCurrInvoice & ")"
      clsADOCon.ExecuteSql sSql
   End If
   
   Dim invoicePrefix As String
   
   ' Change TM invoice to PS
   sSql = "UPDATE CihdTable SET INVNO=" & lCurrInvoice & "," _
          & "INVPRE='" & invoicePrefix & "',INVSTADR='" & sPsStadr & "'," _
          & "INVTYPE='PS',INVSO=0," _
          & "INVCUST='" & Trim(sPsCust) & "' WHERE " _
          & "INVNO=" & lCurrInvoice & " AND INVTYPE='TM'"
          '& "INVNO=" & lNewInv & " AND INVTYPE='TM'"
   clsADOCon.ExecuteSql sSql
   
   ' MM commented - we are not saving the invoice number
   'MM Dim inv As New ClassARInvoice
   'MM inv.SaveLastInvoiceNumber lCurrInvoice
   
   ' Add freight and tax to invoice total
   nTdollars = nTdollars + (cTax + cFREIGHT)
   
   'MM added CANCELED flag to 0
   ' Then post the total to the invoice
   sSql = "UPDATE CihdTable SET INVTOTAL=" & nTdollars & "," _
          & "INVFREIGHT=" & cFREIGHT & "," _
          & "INVTAX=" & cTax & "," _
          & "INVSHIPDATE='" & sPost & "'," _
          & "INVDATE='" & sPost & "'," _
          & "INVCANCELED=0," _
          & "INVCOMMENTS=''," _
          & "INVPACKSLIP='" & sPackSlip & "' " _
          & "WHERE INVNO=" & lCurrInvoice & " "
   clsADOCon.ExecuteSql sSql
   
   sSql = "UPDATE PshdTable SET PSFREIGHT=" & cFREIGHT _
          & " WHERE PSNUMBER='" & sPackSlip & "' "
   clsADOCon.ExecuteSql sSql
   
   ' update the invoice number
   sSql = "UPDATE FusionSOVOI SET INNO =" & lCurrInvoice _
          & " WHERE ITPsnumber = '" & sPackSlip & "' AND payment_doc_no='" & sDocNumber & "' "
   clsADOCon.ExecuteSql sSql
   
   
   MouseCursor 0
   
   If (clsADOCon.ADOErrNum = 0) Then
      clsADOCon.CommitTrans
      sMsg = "Added new Invoice - " + CStr(lCurrInvoice)
      MsgBox sMsg, vbInformation, Caption
      FillGrid
   'TODO: may be we have to print later.
'      'print invoice if required
'      If (emailInvoice And chkPrintEmailedInv.Value = vbChecked) _
'         Or ((Not emailInvoice) And chkPrintNonEmailedInv.Value = vbChecked) Then
'         PrintInvoices lCurrInvoice, lCurrInvoice
'      End If
   Else
      sMsg = "Adding Invoice Number Failed - " + CStr(lCurrInvoice)
      MsgBox sMsg, vbInformation, Caption
      clsADOCon.RollbackTrans
   End If
   
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   clsADOCon.ADOErrNum = 0
   sProcName = "updateinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub



Private Function GetPackslip(sPack As String) As Boolean
   Dim RdoPsl As ADODB.Recordset
   nTaxRate = 0
   sTaxCode = ""
   sTaxState = ""
   sTaxAccount = ""
   
   On Error GoTo DiaErr1
   
   Erase vItems
   sSql = "SELECT DISTINCT PSNUMBER,PSCUST,PSTERMS,PSSTNAME,PSSTADR," _
          & "PSFREIGHT,CUREF,CUNICKNAME,CUNAME,PIPACKSLIP " & vbCrLf _
          & "FROM PshdTable" & vbCrLf _
          & "JOIN CustTable ON PshdTable.PSCUST = CustTable.CUREF" & vbCrLf _
          & "JOIN PsitTable ON PsitTable.PIPACKSLIP = PshdTable.PSNUMBER" & vbCrLf _
          & "AND (PSTYPE=1 AND PSSHIPPRINT=1 AND PSINVOICE=0 )" & vbCrLf _
          & "WHERE PSNUMBER = '" & sPack & "'"
   
   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_DYNAMIC)
   If bSqlRows Then
      With RdoPsl
         
         sPackSlip = "" & Trim(!PsNumber)
         sPsCust = "" & Trim(!CUNICKNAME)
         sPsCust = "" & Trim(!CUREF)
         sPsStadr = "" & Trim(!PSSTNAME) & vbCrLf _
                    & Trim(!PSSTADR)
         cFREIGHT = Format(!PSFREIGHT, "#####0.00")
         .Cancel
      End With
      GetSalesTaxInfo Compress(sPsCust), nTaxRate, sTaxCode, sTaxState, sTaxAccount
      sPsStadr = CheckComments(sPsStadr)
      GetPackslip = True
   Else
      cFREIGHT = 0
      GetPackslip = False
   End If
   Set RdoPsl = Nothing
   Exit Function
DiaErr1:
   sProcName = "getpackslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmdPostInvoice_Click()

   Dim strDate As String
   Dim strCheckNo As String
   Dim strCheckAmt As Currency
   Dim strDocNum As String
   Dim lInvoices As Long
   
   strDate = cmbPSDte   'Format(Now, "mm/dd/yyyy")
   sJournalID = GetOpenJournal("CR", strDate)
   If sJournalID <> "" Then
      'lTrans = GetNextTransaction(sJournalID)
   Else
      sMsg = "No Open Cash Recipts Journal Found For " & strDate & "."
      MsgBox sMsg, vbInformation, Caption
      Exit Sub
   End If
   
   strDocNum = cmbDoc
   ' get teh check number
   GetCheckNumVOI strDocNum, lInvoices, strCheckNo, strCheckAmt
   
   
   On Error Resume Next
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0

   Dim gl As GLTransaction
   Set gl = New GLTransaction
   gl.JournalID = sJournalID 'automatically sets next transaction
   gl.InvoiceDate = CDate(strDate)

   gl.InvoiceNumber = lInvoices
   
   'calculate amount applied.  if pif, amount = inv total
   'if pif, also calculate discount
   Dim amountPaid As Currency
   Dim debitCash As Currency
   Dim creditAR As Currency
   Dim creditOther As Currency
   
   Dim sCust As String
   
   sCust = "SPIAER"
   
   ' get Invoice amount
   GetInvoiceAmount lInvoices, amountPaid

   amountPaid = amountPaid
   creditAR = amountPaid
   debitCash = amountPaid
   
   sSql = "UPDATE CihdTable SET " _
          & "INVCHECKNO='" & strCheckNo & "'," _
          & "INVPAY=" & amountPaid & "," _
          & "INVPIF=1," _
          & "INVADJUST=0," _
          & "INVARDISC=0," _
          & "INVDAYS=0," _
          & "INVCHECKDATE='" & strDate & "'  " _
          & "WHERE INVNO=" & lInvoices & " "
   clsADOCon.ExecuteSql sSql
   
   
   gl.AddDebitCredit 0, creditAR, sSJARAcct, "", 0, 0, "", sCust, strCheckNo
   gl.AddDebitCredit debitCash, 0, sCrCashAcct, "", 0, 0, "", sCust, strCheckNo
   
End Sub

Private Function GetInvoiceAmount(lInvoices As Long, invAmount As Currency)
   Dim RdoInvAmt As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "SELECT INVTOTAL FROM CihdTable WHERE invno = " & lInvoices
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInvAmt)
   If bSqlRows Then
      With RdoInvAmt
         If .EOF Then
            invAmount = !INVTOTAL
         End If
         .Close
         ClearResultSet RdoInvAmt
         
      End With
   End If
   Set RdoInvAmt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetInvoiceAmount"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function
   
Private Function GetCashAccounts() As Byte
   Dim rdoCsh As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT COGLVERIFY,COCRCASHACCT,COCRDISCACCT,COSJARACCT," _
          & "COCRCOMMACCT,COCRREVACCT,COCREXPACCT,COTRANSFEEACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCsh, ES_FORWARD)
   sProcName = "getcashacct"
   If bSqlRows Then
      With rdoCsh
         For i = 1 To 7
            If "" & Trim(.Fields(i)) = "" Then
               b = 1
               Exit For
            End If
         Next
         sCrCashAcct = "" & Trim(!COCRCASHACCT)
         sCrDiscAcct = "" & Trim(!COCRDISCACCT)
         sSJARAcct = "" & Trim(!COSJARACCT)
         sCrCommAcct = "" & Trim(!COCRCOMMACCT)
         sCrRevAcct = "" & Trim(!COCRREVACCT)
         sCrExpAcct = "" & Trim(!COCREXPACCT)
         .Cancel
         If b = 1 Then GetCashAccounts = 3 Else GetCashAccounts = 2
      End With
   Else
      GetCashAccounts = 0
   End If
   Set rdoCsh = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetCashAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetCheckNumVOI(strDocNum As String, ByRef lInvNo As Long, ByRef strCheckNo As String, ByRef strCheckAmt As Currency)
   Dim RdoChk As ADODB.Recordset
   
   On Error GoTo DiaErr1
   strCheckNo = ""
   strCheckAmt = 0
   
   sSql = "select distinct voipmtIss.check_no as check_no, voipmtIss.check_amt as check_amt, FusionSovoi.inno as inno " & _
            " from voipmtIss join FusionSovoi " & _
            "    on voipmtIss.payment_doc_no = FusionSovoi.payment_doc_no " & _
             "    and voipmtIss.payment_doc_it_no = FusionSovoi.payment_doc_it_no " & _
             " Where voipmtIss.payment_doc_no = '" & strDocNum & "'" & _
             "    and inno is not null"
   
               
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
   If bSqlRows Then
      With RdoChk
         If .EOF Then
            strCheckNo = "" & Trim(!check_no)
            strCheckAmt = "" & Trim(!check_amt)
            lInvNo = !INNO
         End If
         .Close
         ClearResultSet RdoChk
         
      End With
   End If
   Set RdoChk = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetCheckNumVOI"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function


'Private Function CheckExists(sCheck As String) As Byte
'   Dim RdoChk As ADODB.Recordset
'
'   On Error GoTo DiaErr1
'   CheckExists = False
'   sSql = "SELECT COUNT(CACHECKNO) FROM CashTable WHERE CACHECKNO = '" _
'          & sCheck & "' AND CACUST = '" & Compress(cmbCst) & "'"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk)
'   If bSqlRows Then
'      With RdoChk
'         If .Fields(0) > 0 Then
'            CheckExists = True
'            sMsg = "Check # " & sCheck & " All Ready Exists " & vbCrLf _
'                   & "For Customer " & cmbCst
'            MsgBox sMsg, vbInformation, Caption
'         End If
'         .Cancel
'      End With
'   End If
'   Set RdoChk = Nothing
'   Exit Function
'
'DiaErr1:
'   sProcName = "CheckExists"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Function

Private Sub cmdPrintPS_Click()
   On Error GoTo DiaErr1
   ' Create the packslip
   Dim strDocNum As String
   Dim strPS As String
   Dim iList As Integer
   strDocNum = cmbDoc
   
   MouseCursor 13
   
   Dim RdoDocNo As ADODB.Recordset
   
   If (strDocNum = "ALL") Then
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
   Else
      Dim ret As String
      
      strPS = GetPSfromDoc(strDocNum)
      If (strPS <> "") Then
         ret = PrintVOIPS(strPS, False)
         
         If (ret = "") Then
            sSql = "UPDATE FusionSOVOI  SET ITPSSHIPPED = '1' WHERE ITPSNUMBER = '" & strPS & "' "
            clsADOCon.ExecuteSql sSql ' rdExecDirect
            
            MsgBox "printed Packslip :" & strPS
            FillGrid
         Else
            MsgBox ret
         End If
         
      End If
      
   End If
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "Print PS "
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   MouseCursor 0
End Sub

Private Sub cmdCreatePS_Click()

   On Error GoTo DiaErr1
   Dim iList As Integer
      
   MouseCursor 13
   
   ' Create the packslip
   Dim strDocNum As String
   Dim strPrevDocNum As String
   Dim bError As Boolean
   
   strPrevDocNum = cmbDoc
   
   bError = False
   clsADOCon.ADOErrNum = 0
   clsADOCon.BeginTrans
   Dim RdoDocNo As ADODB.Recordset
   
   sSql = "select distinct Payment_doc_no from FusionSOVOI WHERE Payment_doc_no = '" & strPrevDocNum & "' AND ( REMARKPS IS NULL OR  REMARKPS = '')"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoDocNo, ES_STATIC)
   If bSqlRows Then
      With RdoDocNo
         While Not .EOF

            strDocNum = "" & Trim(!payment_doc_no)
            bError = CreatePSForDoc(strDocNum)
            .MoveNext
         Wend
         .Close
         ClearResultSet RdoDocNo
      End With
   End If
   Set RdoDocNo = Nothing
   
   If bError = False Then
      clsADOCon.CommitTrans
   Else
      MsgBox "Error Updating the document number: "
      clsADOCon.RollbackTrans
   End If
   
   
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "Create PS"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   

End Sub

Private Function CreatePSForDoc(strDocNum As String) As Boolean
   
   
   ' get next packslip
   Dim packslipNum As String
   Dim sCust As String
   Dim strStnme As String
   Dim strSTAdr As String
   Dim strVia As String
   Dim strTerms As String
   Dim strDate As String
   Dim iItemNumber As Integer
   Dim ps As New ClassPackSlip
   
   Dim itso  As String
   Dim ITNum  As String
   Dim itrev  As String
   Dim WithDrawnQty As Currency
   Dim DocItemNum  As Integer
   Dim sPSComments  As String
   Dim sPartNumber As String
   
   
   strDate = cmbPSDte   'Format(Now, "mm/dd/yyyy")
   
   Dim RdoSOs As ADODB.Recordset
   Dim qoh As Currency
   Dim bFound As Boolean
   
   packslipNum = ps.GetNextPackSlipNumber
   
   bFound = CheckForPSinFusionVOI(packslipNum)
   If (bFound = True) Then
      MsgBox ("Packslip already found in the Fusion VIO table :" & packslipNum)
      Exit Function
   End If
   
   
   GetSoShipTo strDocNum, sCust, strStnme, strSTAdr, strVia, strTerms
   
   sSql = "INSERT INTO PshdTable (PSNUMBER,PSTYPE,PSDATE," _
          & "PSCUST,PSVIA,PSSTNAME,PSSTADR,PSTERMS) " _
          & "VALUES('" & packslipNum & "',1,'" & strDate & "','" & sCust & "','" _
          & strVia & "','" & Trim(strStnme) & "','" & Trim(strSTAdr) & "','" _
          & strTerms & "')"
   clsADOCon.ExecuteSql sSql ' rdExecDirect
   
   Dim PrevItso As String
   Dim PrevITNum As Integer
   Dim PrevItrev As String
   Dim CurItrev As String
   Dim NextItrev As String
   Dim RunTotItQty As Currency
   Dim ReqDt As String
   Dim ITQty As Currency
   Dim LastItrev As String
   
   
   PrevItso = ""
   PrevITNum = 0
   NextItrev = ""
   CurItrev = ""
   
   sSql = "select ITSO, ITNUMBER, ITREV, ITPART, ITQTY, ITSCHED, WITHDRAWN_QTY,Payment_doc_no, Payment_doc_it_no FROM FusionSOVOI " _
            & " WHERE Payment_doc_no = '" & strDocNum & "' and (ITPSNUMBER is null or ITPSNUMBER = '') ORDER by itso, itnumber, itrev"

   iItemNumber = 1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSOs, ES_STATIC)
   If bSqlRows Then
      With RdoSOs
         While (Not .EOF)
            '// read the data from each so item and create new rev
         
            itso = "" & Trim(!itso)
            ITNum = "" & Trim(!ITNUMBER)
            itrev = "" & Trim(!itrev)
            ITQty = "" & Trim(!ITQty)
            ReqDt = "" & Trim(!itsched)
            
            WithDrawnQty = "" & Trim(!WITHDRAWN_QTY)
            DocItemNum = "" & Trim(!payment_doc_it_no)
            sPartNumber = "" & Trim(!ITPART)
            sPSComments = "VOI Import:" & strDocNum & "; DOCItem:" & DocItemNum
            
            
            Dim newItPrev As String
            Dim cPrice As Currency
            GetSOITPrice itso, ITNum, itrev, cPrice
            
            
            GetValidITRev itso, ITNum, itrev, ITQty, WithDrawnQty
            CurItrev = itrev
            RunTotItQty = ITQty - WithDrawnQty
               
            If (ITQty <= 0) Then
               MsgBox "Error : SO: " & itso & " has insufficient parts."
               CreatePSForDoc = True
               Exit Function
            End If
            
'            If (CurItrev = "") Then
'               CurItrev = itrev
'            Else
'               CurItrev = NextItrev
'            End If
            
'            If (PrevItso = itso) And (PrevITNum = ITNum) And (PrevItrev = itrev) Then
'               ' get next Rev
'               'CurItrev = itrev
'               RunTotItQty = RunTotItQty - WithDrawnQty
'            Else
'               ' get the next Rev not taken
'
'               GetValidITRev itso, ITNum, itrev, ITQty
'               CurItrev = itrev
'               RunTotItQty = ITQty - WithDrawnQty
'            End If
            
            sSql = "UPDATE SoitTable SET ITPSNUMBER='" & packslipNum & "', ITPSITEM=" & iItemNumber _
                   & ", ItQty = " & WithDrawnQty _
                   & " WHERE ITSO=" & itso & " AND ITNUMBER=" _
                   & ITNum & " AND ITREV='" & CurItrev & "' "
            clsADOCon.ExecuteSql sSql ', rdExecDirect
               
            Dim strcustreq As String
            Dim strscheddel As String
            
            GetSOCustSchedDates itso, ITNum, strcustreq, strscheddel
            
            'Create PsitTable record
            sSql = "INSERT PsitTable (PIPACKSLIP,PIITNO,PITYPE,PIQTY,PIPART," _
                   & "PISONUMBER,PISOITEM,PISOREV,PISELLPRICE,PIBOX,PICOMMENTS) " _
                   & "VALUES('" & packslipNum & "'," & iItemNumber & ",1," & WithDrawnQty _
                   & ",'" & sPartNumber & "'," & itso & "," & ITNum _
                   & ",'" & CurItrev & "'," & cPrice & ",'','" & sPSComments & "')"
            clsADOCon.ExecuteSql sSql ', rdExecDirect
            
            sSql = "UPDATE FusionSOVOI  SET ITPSNUMBER = '" & packslipNum & "' ,ITPSITEM = '" & iItemNumber & "', " _
                  & " ITREV = '" & CurItrev & "' WHERE Payment_doc_no = '" & strDocNum & "' " _
                  & " AND Payment_doc_it_no = '" & DocItemNum & "' " _
                  & " AND ITSO = '" & itso & "' AND ITNUMBER = " & ITNum
            
            clsADOCon.ExecuteSql sSql ', rdExecDirect

            ' get next Max Revision number
            LastItrev = GetNextSOMaxRev(itso, ITNum)
            
            NextItrev = GetNextSORevision(LastItrev)
            
            ' get ITCOMMENTS and ITCUSTITEMNO from original SO Item (without a rev letter)
            Dim rs As ADODB.Recordset
            Dim comment As String, poitem As String
            sSql = "select ITCOMMENTS, ITCUSTITEMNO from SoitTable WHERE ITSO=" & itso & " AND ITNUMBER=" & ITNum & " AND ITREV=''"
            bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_STATIC)
            If bSqlRows Then
               comment = "" & Replace(Trim(rs!ITCOMMENTS), "'", "''")
               poitem = Replace(RTrim(rs!ITCUSTITEMNO), "'", "''")
            End If
            rs.Close
            Set rs = Nothing
            
            If NextItrev = "ZZX" Then
               MsgBox "SO = " & itso & " Item = " & ITNum & " - revision has reached max revision 'ZZX'"
               CreatePSForDoc = True
               Exit Function
               'clsADOCon.ADOErrNum = 0
            End If
            
            ' only if there is balance in the run total - split the SO item to create a new soitem.
            If (RunTotItQty > 0) Then
               sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITREV, ITCUSTITEMNO, ITPART,ITQTY,ITSCHED,ITBOOKDATE," _
                      & "ITCUSTREQ, ITSCHEDDEL, ITDOLLORIG, ITDOLLARS, ITUSER, ITCOMMENTS) " & vbCrLf _
                      & "VALUES(" & itso & "," & ITNum & ",'" & NextItrev & "','" & poitem & "','" _
                      & Compress(sPartNumber) & "'," & Val(RunTotItQty) & ",'" & ReqDt & "','" _
                      & Format(ES_SYSDATE, "mm/dd/yy") & "','" & strcustreq & "','" & strscheddel & "','" _
                      & CCur(cPrice) & "','" & CCur(cPrice) & "','" & sInitials & "','" & comment & "')"
               clsADOCon.ExecuteSql sSql ' rdExecDirect
            End If
            
            PrevItso = itso
            PrevITNum = ITNum
            PrevItrev = itrev
            
            If clsADOCon.ADOErrNum <> 0 Then
               MsgBox "Error Updating PS Number : SO: " & itso & " ITNum:" & ITNum & " ITREV:" & itrev
               CreatePSForDoc = True
               Exit Function
               'clsADOCon.ADOErrNum = 0
            End If
            
            iItemNumber = iItemNumber + 1
            .MoveNext ' next so Number
         Wend  ' SO Item
         .Close
      End With
      Set RdoSOs = Nothing
   End If
   'On Error Resume Next
   Set RdoSOs = Nothing
   
   'sCust = Compress("SPIAER")
   
   ps.SaveLastPSNumber packslipNum

   If clsADOCon.ADOErrNum <> 0 Then
      MsgBox "Error Updating PS Number : " & packslipNum
      CreatePSForDoc = True
      Exit Function
      
   End If
   
   FillGrid

   CreatePSForDoc = False

End Function

Private Function GetNextSOMaxRev(strSoNum As String, ITNum As String) As String
   On Error GoTo DiaErr1
   
   Dim rdo As ADODB.Recordset
   
   sSql = "SELECT MAX(ITREV) rev FROM soitTable WHERE ITSO ='" & strSoNum & "' AND ITNUMBER = " & ITNum _
            & " AND LEN(ITREV) = (SELECT MAX(LEN(ITREV)) FROM soitTable " _
                                 & " WHERE ITSO ='" & strSoNum & "' AND ITNUMBER = " & ITNum & ") "
                        
   bSqlRows = clsADOCon.GetDataSet(sSql, rdo, ES_FORWARD)
   If bSqlRows Then
      GetNextSOMaxRev = Trim(rdo!rev)
      ClearResultSet rdo
   Else
      GetNextSOMaxRev = ""
   End If
   Set rdo = Nothing
   
   Exit Function

DiaErr1:
   sProcName = "GetNextSOMaxRev"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

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

Function GetSOCustSchedDates(strSoNum As String, ITNum As String, _
                  ByRef itcustreq As String, ByRef itscheddel As String)
                  
   On Error GoTo DiaErr1
   
   Dim RdoRptQ As ADODB.Recordset
   
   sSql = "SELECT TOP 1 itcustreq,itscheddel FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRptQ, ES_FORWARD)
   If bSqlRows Then
      itcustreq = Format(Trim(RdoRptQ!itcustreq), "mm/dd/yyyy")
      itscheddel = Format(Trim(RdoRptQ!itscheddel), "mm/dd/yyyy")
      
      ClearResultSet RdoRptQ
   Else
      itcustreq = Format(GetServerDateTime(), "mm/dd/yyyy")
      itscheddel = Format(GetServerDateTime(), "mm/dd/yyyy")
   End If
   
   Set RdoRptQ = Nothing
   
   Exit Function

DiaErr1:
   sProcName = "GetSOCustSchedDates"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
                  
End Function
                  

Function GetValidITRev(strSoNum As String, ITNum As String, ByRef itrev As String _
                                    , ByRef ITQty As Currency, ByRef WithDrawQty As Currency) As String
   On Error GoTo DiaErr1
   
   Dim Rdolen As ADODB.Recordset
   Dim RdoRpt As ADODB.Recordset
   Dim RdoRptQ As ADODB.Recordset
   Dim lenRev As Integer
   
   sSql = "SELECT MAX(LEN(ITREV)) lenRev FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITACTUAL IS NULL and itpsnumber = '' " _
            & " and ITCANCELED <> 1 and itqty > 0 AND itqty >= " & WithDrawQty
   
   bSqlRows = clsADOCon.GetDataSet(sSql, Rdolen, ES_FORWARD)
   If bSqlRows Then
      If (Not IsNull(Rdolen!lenRev)) Then
         lenRev = Trim(Rdolen!lenRev)
         ClearResultSet Rdolen
      Else
         MsgBox "Couldn't find the Last Rev number - " & strSoNum & ":" & ITNum & ":" & WithDrawQty
         
      End If
   End If
   Set Rdolen = Nothing
   
   If (lenRev > 2) Then
      
      sSql = "SELECT MIN(ITREV) LstRev FROM soitTable WHERE ITSO ='" & strSoNum _
               & "' AND ITNUMBER = " & ITNum & " AND LEN(ITREV)  > 2  AND " _
               & " ITCANCELED <> 1 AND ITACTUAL IS NULL and itpsnumber = '' and itqty > 0  AND itqty >= " & WithDrawQty
   Else
      sSql = "SELECT MIN(ITREV) LstRev FROM soitTable WHERE ITSO ='" & strSoNum _
               & "' AND ITNUMBER = " & ITNum & "  AND ITACTUAL IS NULL and " _
               & " ITCANCELED <> 1 AND itpsnumber = '' and itqty > 0 AND itqty >= " & WithDrawQty
   End If
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRpt, ES_FORWARD)
   If bSqlRows Then
      itrev = Trim(RdoRpt!LstRev)
      ClearResultSet RdoRpt
   End If
   Set RdoRpt = Nothing
   
   
   sSql = "SELECT ITQTY FROM soitTable WHERE ITSO ='" & strSoNum _
            & "' AND ITNUMBER = " & ITNum & " AND ITREV = '" & itrev & "' AND ITACTUAL IS NULL " _
            & " AND ITCANCELED <> 1 and itpsnumber = '' and itqty > 0  AND itqty >= " & WithDrawQty
   
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
   sSql = "select DISTINCT PIPACKSLIP from psitTable where picomments like '%" & strDocNum & "%' and PILOTNUMBER = ''"
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

Private Function PrintVOIPS(sPackSlip As String, Optional DontPrint As Boolean) As String
   
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
   
   'MM - 11/3
   vAdate = cmbPSDte 'Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   vCurrentdate = vAdate
   vPSdate = cmbPSDte 'Format(GetServerDateTime(), "mm/dd/yyyy hh:mm")
   
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
      sMsg = "There Is No Open Inventory Journal For This" & vbCrLf _
         & "Period. Cannot Set The Pack Slip As Printed."
      
      PrintVOIPS = sMsg
      Exit Function
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
         sMsg = "Packslip is already printed."
         PrintVOIPS = sMsg
      End If
      Exit Function
   End If
   
   iTotalItems = GetItems(sPackSlip)
   If iTotalItems = 0 Then
      sMsg = "There Are No Unprinted Items On This Packing Slip."
      PrintVOIPS = sMsg
      Exit Function
   End If
   
   'quickly check that all lot-tracked items are available in sufficient quantity
   If bLotsAct Then
      For iRow = 1 To iTotalItems
         bLots = vItems(iRow, PS_LOTTRACKED)
         If bLots = 1 Then
            sPart = sPartGroup(iRow)
            cRemPqty = Val(vItems(iRow, PS_QUANTITY))
            'cLotQty = GetRemainingLotQty(sPart)
            cLotQty = GetSprintLotRemainingQty(sPart)

            
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
         If (True) Then
            sMsg = sMsg & "The packing slip will not be printed."
         End If
         PrintVOIPS = sMsg
         Exit Function
      End If
   End If
   
   'Packing slip hasn't been printed.  Confirm that printing is desired.
   ' TODO:
   If (False) Then
      sMsg = "Do You Want To Print This Pack Slip " & vbCrLf _
             & "And Adjust Inventory For The Parts?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      
      If bResponse = vbNo Then Exit Function
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
         cLotQty = GetSprintLotRemainingQty(sPart)
         
         iLots = GetPartLots(sPart, sPackSlip)
         cItmLot = 0
         cRemPqty = Format(Val(vItems(iRow, PS_QUANTITY)), ES_QuantityDataFormat)
         
         For iList = 1 To iLots
            If cRemPqty <= 0 Then
               Exit For
            End If
            cLotQty = Val(sLots(iList, 1))
            If (cLotQty > 0) Then
            
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
            End If
         Next
         ' If still we have remaining Qty we need to quit
         If (cRemPqty > 0) Then

            MsgBox "Not sufficient quantity for item " & vItems(iRow, PS_ITEMNO) _
               & " part " & sPart & " available. " & vbCrLf _
               & "It is short by (" & cRemPqty & ") quantity." & vbCrLf _
               & "The packing slip will not be printed."

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
         Exit Function
      End If
   End If
   
'   If bMarkShipped = 0 Then
'      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
'             & "PSSHIPPRINT=1,PSSHIPPED=0 WHERE " _
'             & "PSNUMBER='" & sPackSlip & "' AND PSTYPE=1"
'      clsADOCon.ExecuteSQL sSql ', rdExecDirect
'   Else
      sSql = "UPDATE PshdTable SET PSPRINTED='" & vAdate & "'," _
             & "PSSHIPPRINT=1,PSSHIPPEDDATE='" & vAdate & "'," _
             & "PSSHIPPED=1 WHERE PSNUMBER='" & sPackSlip & "' " _
             & "AND PSTYPE=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
'   End If
   If clsADOCon.RowsAffected = 0 Then
      MouseCursor 0
      MsgBox "Could Not Update The Packing Slip. The Transaction " & vbCrLf _
         & "Has Been Aborted. Try Again In A Few Minutes.", _
         vbExclamation, Caption
      clsADOCon.RollbackTrans
      Exit Function
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
         & "SELECT tmp.LotID, dbo.fnGetNextLotItemNumber( tmp.LotID )+ tmp.LotItemID - 1, " _
         & bInvType & ", '" & sPart & "', '" & vAdate & "', " & vbCrLf _
         & "-tmp.LotQty, '" & sPackSlip & "', " _
         & Val(vItems(iRow, PS_ITEMNO)) & ", 'SPIAER'," _
         & "ia.INNUMBER, 'Shipped Item'" & vbCrLf _
         & "FROM TempPsLots tmp" & vbCrLf _
         & "JOIN InvaTable ia ON ia.INPSNUMBER = tmp.PsNumber AND ia.INPSITEM = tmp.PsItem" & vbCrLf _
         & "and ia.INADATE = '" & vAdate & "' and ia.INLOTNUMBER = tmp.LotID" & vbCrLf _
         & "and ia.INAQty  = -tmp.LotQty " & vbCrLf _
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
   
   PrintVOIPS = ""
   Exit Function
   
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
   Exit Function
   
NoCanDo:
   MouseCursor 0
   PrintVOIPS = "No inventory"

   Exit Function
End Function

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

Private Function GetSprintLotRemainingQty(LotPart As String) As Currency

   'same as GetRemainingLotQty, except gets remaining qty from sum(LOTREMAININGQTY)
   'rather than SUM(LOIQUANTITY) to reduce a problem at LUMICOR
   
   Dim ADOQty As ADODB.Recordset
   
   GetSprintLotRemainingQty = 0
   sSql = "select isnull(sum(LOTREMAININGQTY),0)" & vbCrLf _
      & "from LohdTable" & vbCrLf _
      & "where LOTPARTREF='" & LotPart & "' AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT'"
   If clsADOCon.GetDataSet(sSql, ADOQty, ES_FORWARD) Then
      GetSprintLotRemainingQty = ADOQty.Fields(0)
   End If
   Set ADOQty = Nothing
   
End Function

Private Function GetPartLots(sPartWithLot As String, sPSNumber As String) As Integer
   Dim RdoLots As ADODB.Recordset
   Dim iRow As Integer
   Erase sLots
   On Error GoTo DiaErr1
   
   sSql = "select LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE,(LOTREMAININGQTY - tPsLots.TotQty) as consumeQty" _
         & " FROM  lohdTable join (SELECT PSNumber, LotID, Partref, SUM(LOTQTY) TotQty FROM TempPsLots where PSNumber = '" & sPSNumber & "'" _
         & "                           GROUP BY PSNumber, LotID, Partref ) as tPsLots on   lotNumber = lotID  where LOTREMAININGQTY > 0 AND LOTAVAILABLE=1" _
         & "   AND LOTLOCATION = 'SPRT'  and lotpartref = '" & sPartWithLot & "' " _
         & " Union " _
         & "select DISTINCT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE, LOTREMAININGQTY as consumeQty" _
         & "   FROM  lohdTable where LOTREMAININGQTY > 0 AND LOTAVAILABLE=1" _
         & "      AND LOTLOCATION = 'SPRT'  and lotpartref = '" & sPartWithLot & "' and lotNumber NOT IN" _
         & "            (SELECT DISTINCT lotID FROM TempPsLots WHERE PSNumber = '" & sPSNumber & "')" _
         & "      ORDER BY LOTNUMBER ASC"


   
'   sSql = "select LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE,(LOTREMAININGQTY - tPsLots.TotQty) as consumeQty " _
'          & "  FROM  lohdTable join (SELECT PSNumber, LotID, Partref, SUM(LOTQTY) TotQty FROM TempPsLots " _
'         & "                           GROUP BY PSNumber, LotID, Partref ) as tPsLots on   lotNumber = lotID " _
'         & " where LOTREMAININGQTY > 0 AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT' " _
'         & " and lotpartref = '" & sPartWithLot & "'" _
'         & " Union " _
'         & " select LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE, LOTREMAININGQTY as consumeQty " _
'         & "   FROM  lohdTable left outer join TempPsLots on   lotNumber = lotID " _
'         & " where LOTREMAININGQTY > 0 AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT' " _
'         & " and lotpartref = '" & sPartWithLot & "' and LotQty IS NULL " _
'         & " ORDER BY LOTNUMBER ASC "

Debug.Print sSql
'   sSql = "SELECT LOTNUMBER,LOTPARTREF,LOTREMAININGQTY,LOTAVAILABLE " _
'          & "FROM LohdTable WHERE (LOTPARTREF='" & sPartWithLot & "' AND " _
'          & "LOTREMAININGQTY > 0 AND LOTAVAILABLE=1 AND LOTLOCATION = 'SPRT') ORDER BY LOTNUMBER ASC"
'   If bFIFO = 1 Then
'      sSql = sSql & "ORDER BY LOTNUMBER ASC"
'   Else
'      sSql = sSql & "ORDER BY LOTNUMBER DESC"
'   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLots, ES_FORWARD)
   If bSqlRows Then
      With RdoLots
         Do Until .EOF
            If (iRow >= 49) Then Exit Do
            iRow = iRow + 1
            sLots(iRow, 0) = "" & Trim(!lotNumber)
            sLots(iRow, 1) = Format$(!ConsumeQty, ES_QuantityDataFormat)
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


Private Sub cmdSearch_Click()
   ExpDlg.Filter = "Excel File (*.xls) | *.xls"
   ExpDlg.ShowOpen
   If ExpDlg.FileName = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = ExpDlg.FileName
   End If

End Sub




Private Sub cmdSel_Click()
   FillGrid
End Sub
'
'Private Sub CmdSelAll_Click()
'   Dim iList As Integer
'
'   For iList = 1 To Grd.Rows - 1
'       Grd.Col = 0
'       Grd.Row = iList
'       ' Only if the part is checked
'       If Grd.CellPicture = Chkno.Picture Then
'           Set Grd.CellPicture = Chkyes.Picture
'       End If
'   Next
'End Sub




Private Sub Form_Activate()
   Dim bSoAdded As Byte
   MdiSect.lblBotPanel = Caption
   
   sSql = "select distinct payment_doc_no from FusionSOVOI where (itpsnumber is null OR ITPSSHIPPED IS NULL OR INNO IS NULL) order by payment_doc_no"
            
   LoadComboBox cmbDoc, -1
   'AddComboStr cmbDoc.hWnd, "" & Trim("ALL")
   'cmbDoc = "ALL"
   
   Set cmdObj1 = New ADODB.Command
   cmdObj1.CommandText = sSql
   'Set RdoQry1 = RdoCon.CreateQuery("", sSql)
   'RdoQry1.MaxRows = 1
   Dim prmObj1 As ADODB.Parameter
   Set prmObj1 = New ADODB.Parameter
   prmObj1.Type = adChar
   prmObj1.Size = 8
   cmdObj1.parameters.Append prmObj1
   
   sSql = "SELECT PIPACKSLIP,PIITNO,PIQTY,PIPART,PISONUMBER,PISOITEM," _
          & "PISOREV,PARTREF,PARTNUM,PALOTTRACK FROM " _
          & "PsitTable,PartTable WHERE (PIPART=PARTREF AND PIPACKSLIP = ?)" & vbCrLf _
          & "ORDER BY PIITNO"
   'Set RdoQry2 = RdoCon.CreateQuery("", sSql)
   Set cmdObj2 = New ADODB.Command
   cmdObj2.CommandText = sSql
   
   Dim prmObj2 As ADODB.Parameter
   
   Set prmObj2 = New ADODB.Parameter
   prmObj2.Type = adChar
   prmObj2.Size = 8
   cmdObj2.parameters.Append prmObj2
   
   ' Only if the import table is full
   'FillGrid
   cmbPSDte.Text = Format(Now, "mm/dd/yyyy")
   If bOnLoad Then
       bOnLoad = 0
   End If
    
    
   MouseCursor (0)

End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  
  For Each ctl In Me.Controls
    If TypeOf ctl Is MSFlexGrid Then
      If IsOver(ctl.hwnd, Xpos, Ypos) Then FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
    End If
  Next ctl
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' make sure that you release the Hook
   Call WheelUnHook(Me.hwnd)
   
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


      .rows = 1
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
      .Text = "Packslip"
      .Col = 12
      .Text = "PS Item"
      
      .Col = 13
      .Text = "PS Shipped"
      .Col = 14
      .Text = "Invoice"
      .Col = 15
      .Text = "Posted"
      .Col = 16
      .Text = "SO Remarks"
      .Col = 17
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
      .ColWidth(11) = 800
      .ColWidth(12) = 800
      .ColWidth(13) = 800
      .ColWidth(14) = 800
      .ColWidth(15) = 800
      .ColWidth(16) = 1050
      .ColWidth(17) = 1050
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
   
   Call WheelHook(Me.hwnd)
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
    'FormUnload
    Set SaleSLf15b = Nothing
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
'
'
'Private Sub cmdClear_Click()
'    Dim iList As Integer
'    For iList = 1 To Grd.Rows - 1
'        Grd.Col = 0
'        Grd.Row = iList
'        ' Only if the part is checked
'        If Grd.CellPicture = Chkyes.Picture Then
'            Set Grd.CellPicture = Chkno.Picture
'        End If
'    Next
'End Sub


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

Private Function CheckForPSinFusionVOI(ByVal strPONum As String) As Boolean
   On Error GoTo modErr1
   Dim RdoPO As ADODB.Recordset
   
   CheckForPSinFusionVOI = False
   If Trim(strPONum) = "" Then
      CheckForPSinFusionVOI = False
   Else
      sSql = "select distinct itpsnumber from FusionSOVOI WHERE itpsnumber = '" & strPONum & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPO, ES_FORWARD)
      If bSqlRows Then
         With RdoPO
            CheckForPSinFusionVOI = True
            ClearResultSet RdoPO
         End With
      End If
   End If
   Set RdoPO = Nothing
   Exit Function
   
modErr1:
   sProcName = "CheckForPSinFusionVOI"
   CheckForPSinFusionVOI = False
   
End Function



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


Private Sub GetSJAccounts()
   Dim rdoJrn As ADODB.Recordset
   Dim i As Integer
   Dim b As Byte
   If sJournalID = "" Then
      bGoodAct = True
      Exit Sub
   End If
   
   On Error GoTo DiaErr1
   sSql = "SELECT COREF,COSJARACCT,COSJNFRTACCT," _
          & "COSJTFRTACCT,COSJTAXACCT FROM ComnTable WHERE " _
          & "COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoJrn, ES_FORWARD)
   If bSqlRows Then
      With rdoJrn
         ' A/R
         sCOSjARAcct = "" & Trim(.Fields(1))
         If sCOSjARAcct = "" Then b = 1
         ' NonTaxable freight
         sCOSjNFRTAcct = "" & Trim(.Fields(2))
         If sCOSjNFRTAcct = "" Then b = 1
         ' Taxable freight
         sCOSjTFRTAcct = "" & Trim(.Fields(3))
         If sCOSjTFRTAcct = "" Then b = 1
         ' Sales tax
         sCOSjTaxAcct = "" & Trim(.Fields(4))
         If sCOSjTaxAcct = "" Then b = 1
         .Cancel
      End With
   End If
   If b = 1 Then
      bGoodAct = False
      '        lblJrn.Visible = True
   Else
      bGoodAct = True
      '        lblJrn.Visible = False
   End If
   Set rdoJrn = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsjacco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetNextTransaction(sJrnlId As String) As Long
   Dim RdoTrn As ADODB.Recordset
   On Error GoTo modErr1
   sSql = "SELECT MAX(DCTRAN) FROM JritTable WHERE DCHEAD='" _
          & Trim(sJrnlId) & "'"
bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrn, ES_FORWARD)
   If bSqlRows Then
      With RdoTrn
         If Not IsNull(.Fields(0)) Then
            GetNextTransaction = (.Fields(0)) + 1
         Else
            GetNextTransaction = 1
         End If
         .Cancel
      End With
   Else
      GetNextTransaction = 1
   End If
   Exit Function
modErr1:
   On Error Resume Next
   sProcName = "getnexttrans"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Function

Private Sub GetPartBnO(sPart, nRate, sCode, sState, sType)
   ' Get B&O tax codes from part
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
          & "PABORTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,PartTable WHERE " _
             & "PABOWTAX = TAXREF AND TAXTYPE = 0 AND PARTREF = '" & sPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getpartbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
End Sub

Private Sub GetSalesTaxInfo( _
                           sCust As String, _
                           nRate As Currency, _
                           sCode As String, _
                           sState As String, _
                           sAccount As String)
   
   On Error GoTo DiaErr1
   
   ' Load tax from customer.
   Dim RdoTax As ADODB.Recordset
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE,TAXACCT FROM CustTable INNER JOIN " _
          & "TxcdTable ON CustTable.CUTAXCODE = TxcdTable.TAXREF " _
          & "WHERE CUREF = '" & sCust & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTax)
   If bSqlRows Then
      With RdoTax
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sAccount = "" & Trim(!TAXACCT)
         .Cancel
      End With
   End If
   Set RdoTax = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getsaletaxinfo"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

Private Sub GetCustBnO(sCust, nRate, sCode, sState, sType)
   ' Get B&O tax codes from customer
   ' Retail takes precidence over wholesale
   
   Dim rdoTx1 As ADODB.Recordset
   Dim rdoTx2 As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
          & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
          & "' AND TAXTYPE = 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx1)
   If bSqlRows Then
      With rdoTx1
         nRate = !TAXRATE
         sCode = "" & Trim(!taxCode)
         sState = "" & Trim(!taxState)
         sType = "R"
         .Cancel
      End With
   Else
      sSql = "SELECT TAXCODE,TAXSTATE,TAXRATE FROM TxcdTable,CustTable " _
             & "WHERE CUBORTAXCODE = TAXREF AND CUREF = '" & sCust _
             & "' AND TAXTYPE = 0"
      bSqlRows = clsADOCon.GetDataSet(sSql, rdoTx2)
      If bSqlRows Then
         With rdoTx2
            nRate = !TAXRATE
            sCode = "" & Trim(!taxCode)
            sState = "" & Trim(!taxState)
            sType = "W"
            .Cancel
         End With
      End If
   End If
   
   Set rdoTx1 = Nothing
   Set rdoTx2 = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcustbno"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


Private Function GetNextInvoice(strFullPSNum As String) As Long
   
   Dim bDup As Boolean
   
   Dim inv As New ClassARInvoice
   ' Invoice numbers are created from packslip
   ' MM lNextInv = inv.GetNextInvoiceNumber

      
   Dim strPSNum As String
   strPSNum = Mid$(CStr(strFullPSNum), 3, Len(strFullPSNum))
   If (strPSNum <> "") Then
      lNextInv = Val(strPSNum)
   End If

   ' Validate the Invoice number
   If (Trim(lNextInv) <> "") Then
      Dim iCanceled As Integer
      
      iCanceled = 0
      bDup = inv.DuplicateInvNumber(CLng(lNextInv), iCanceled)
      
      If ((bDup = True) And (iCanceled = 0)) Then
         ' if the Inv PS is same then...get the next invoice from the invoice pool.
         lNextInv = inv.GetNextInvoiceNumber
         lNextInv = Format(lNextInv, "000000")
         MsgBox "Invoice number exists for PS Number " & strFullPSNum & ".Using the New Invoice number is " & lNextInv & ".", vbInformation, Caption
      End If
   End If
   
   GetNextInvoice = lNextInv

End Function


Private Function GetPartInvoiceAccounts(SPartRef As String, iLevel As Integer, sCode As String, _
                                       Optional sREVAccount As String, _
                                       Optional sDisAccount As String, _
                                       Optional sCGSMaterialAccount As String, _
                                       Optional sCGSLaborAccount As String, _
                                       Optional sCGSExpAccount As String, _
                                       Optional sCGSOhAccount As String, _
                                       Optional sInvMaterialAccount As String, _
                                       Optional sInvLaborAccount As String, _
                                       Optional sInvExpAccount As String, _
                                       Optional sInvOhAccount As String) As Boolean
   
   Dim rdoAct As ADODB.Recordset
   On Error GoTo modErr1
   
   'Part
   GetPartInvoiceAccounts = True
   SPartRef = Compress(SPartRef)
   
   sSql = "SELECT PACGSMATACCT,PACGSLABACCT,PACGSEXPACCT,PACGSOHDACCT," _
          & "PAINVMATACCT,PAINVLABACCT,PAINVEXPACCT,PAINVOHDACCT," _
          & "PAREVACCT,PADISACCT FROM PartTable WHERE PARTREF='" & SPartRef & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         sREVAccount = "" & Trim(!PAREVACCT)
         sDisAccount = "" & Trim(!PADISACCT)
         
         sCGSMaterialAccount = "" & Trim(!PACGSMATACCT)
         sCGSLaborAccount = "" & Trim(!PACGSLABACCT)
         sCGSExpAccount = "" & Trim(!PACGSEXPACCT)
         sCGSOhAccount = "" & Trim(!PACGSOHDACCT)
         
         sInvMaterialAccount = "" & Trim(!PAINVMATACCT)
         sInvLaborAccount = "" & Trim(!PAINVLABACCT)
         sInvExpAccount = "" & Trim(!PAINVEXPACCT)
         sInvOhAccount = "" & Trim(!PAINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Now check the accounts, if any are blank then fill then from the
   ' product code
   sCode = Compress(sCode)
   
   sSql = "SELECT PCCGSMATACCT,PCCGSLABACCT,PCCGSEXPACCT,PCCGSOHDACCT," _
          & "PCINVMATACCT,PCINVLABACCT,PCINVEXPACCT,PCINVOHDACCT," _
          & "PCREVACCT,PCDISCACCT FROM PcodTable WHERE PCREF='" & sCode & "'"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sREVAccount = "" Then sREVAccount = "" & Trim(!PCREVACCT)
         If sDisAccount = "" Then sDisAccount = "" & Trim(!PCDISCACCT)
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(!PCCGSMATACCT)
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(!PCCGSLABACCT)
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(!PCCGSEXPACCT)
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(!PCCGSOHDACCT)
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(!PCINVMATACCT)
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(!PCINVLABACCT)
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(!PCINVEXPACCT)
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(!PCINVOHDACCT)
         
         .Cancel
      End With
   End If
   
   ' Last check the company setup and fill any accounts that are still empty.
   sSql = "SELECT COREVACCT" & Trim(str(iLevel)) & "," _
          & "COAPDISCACCT," _
          & "COCGSMATACCT" & Trim(str(iLevel)) & "," _
          & "COCGSLABACCT" & Trim(str(iLevel)) & "," _
          & "COCGSEXPACCT" & Trim(str(iLevel)) & "," _
          & "COCGSOHDACCT" & Trim(str(iLevel)) & "," _
          & "COINVMATACCT" & Trim(str(iLevel)) & "," _
          & "COINVLABACCT" & Trim(str(iLevel)) & "," _
          & "COINVEXPACCT" & Trim(str(iLevel)) & "," _
          & "COINVOHDACCT" & Trim(str(iLevel)) & " FROM " _
          & "ComnTable WHERE COREF=1"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   
   If bSqlRows Then
      With rdoAct
         
         If sREVAccount = "" Then sREVAccount = "" & Trim(.Fields(0))
         If sDisAccount = "" Then sDisAccount = "" & Trim(.Fields(1))
         
         If sCGSMaterialAccount = "" Then sCGSMaterialAccount = "" & Trim(.Fields(2))
         If sCGSLaborAccount = "" Then sCGSLaborAccount = "" & Trim(.Fields(3))
         If sCGSExpAccount = "" Then sCGSExpAccount = "" & Trim(.Fields(4))
         If sCGSOhAccount = "" Then sCGSOhAccount = "" & Trim(.Fields(5))
         
         If sInvMaterialAccount = "" Then sInvMaterialAccount = "" & Trim(.Fields(6))
         If sInvLaborAccount = "" Then sInvLaborAccount = "" & Trim(.Fields(7))
         If sInvExpAccount = "" Then sInvExpAccount = "" & Trim(.Fields(8))
         If sInvOhAccount = "" Then sInvOhAccount = "" & Trim(.Fields(9))
         
         .Cancel
      End With
   End If
   
   
   Set rdoAct = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetPartInvoiceAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   MsgBox CurrError.Number & " " & CurrError.Description
End Function



