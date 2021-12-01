VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form SaleSLf14a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Sales Orders"
   ClientHeight    =   9705
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   15375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
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
      Left            =   13080
      TabIndex        =   28
      ToolTipText     =   " Select All"
      Top             =   5280
      Width           =   1920
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export to Excel"
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
      Left            =   10440
      TabIndex        =   27
      ToolTipText     =   " Export to Excel file"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame z2 
      Height          =   975
      Left            =   1920
      TabIndex        =   23
      Top             =   2280
      Width           =   2415
      Begin VB.OptionButton optEdiFile 
         Caption         =   "862's EDI file"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optEdiFile 
         Caption         =   "850's EDI file"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   200
         Width           =   1815
      End
   End
   Begin VB.CheckBox OptSoEDI 
      Caption         =   "FromEDI"
      Height          =   195
      Left            =   6240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   480
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ComboBox cmbPre 
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Tag             =   "3"
      Text            =   "S"
      ToolTipText     =   "Select or Enter Type A thru Z"
      Top             =   960
      Width           =   520
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1440
      Width           =   1555
   End
   Begin VB.TextBox txtSon 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Enter New Sales Order Number"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdSalesOrder 
      Caption         =   "Create SO"
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
      Left            =   13080
      TabIndex        =   9
      ToolTipText     =   " Create Sales Orders"
      Top             =   4560
      Width           =   1920
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
      Left            =   13080
      TabIndex        =   8
      ToolTipText     =   " Clear Selection"
      Top             =   6000
      Width           =   1920
   End
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   3360
      Width           =   4695
   End
   Begin VB.CommandButton cmdImport 
      Cancel          =   -1  'True
      Caption         =   "Import EDI Sales data"
      Height          =   360
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2145
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf14a.frx":0000
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
      FormDesignHeight=   9705
      FormDesignWidth =   15375
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   6960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd850 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4560
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grd862 
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4560
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grd830 
      Height          =   4935
      Left            =   120
      TabIndex        =   26
      ToolTipText     =   "Click To Select Or Scroll And Press Enter (Also Page Up And Page Down)"
      Top             =   4560
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   3
      Cols            =   9
      FixedRows       =   2
      FixedCols       =   0
      FocusRect       =   2
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label2 
      Caption         =   "** Part Not found in Fusion"
      Height          =   255
      Left            =   12960
      TabIndex        =   21
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "* Already added Sales Order Item"
      Height          =   375
      Left            =   12960
      TabIndex        =   20
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sales Order"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      ToolTipText     =   "Last Sales Order Entered"
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "SO Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf14a.frx":07AE
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7080
      Picture         =   "SaleSLf14a.frx":0B38
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select EDI File"
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Width           =   1305
   End
End
Attribute VB_Name = "SaleSLf14a"
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
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim bOptionSel As Boolean

Dim sLastPrefix As String
Dim sNewsonumber As String
Dim sCust As String
Dim sStName As String
Dim sStAdr As String
Dim sContact As String
Dim sConIntPhone As String
Dim sConPhone As String
Dim sConIntFax As String
Dim sConFax As String
Dim sConExt As String
Dim sDivision As String
Dim sOldSoNumber As String
Dim sRegion As String
Dim sSterms As String
Dim sVia As String
Dim sFob As String
Dim sSlsMan As String
Dim sTaxExempt As String
Dim strFilePath As String
Dim strEDIFormat  As String

Dim iDays As Integer
Dim iFrtDays As Integer
Dim iNetDays As Integer

Dim cDiscount As Currency

Dim arrValue() As Variant
Dim arrFieldName() As Variant

Dim strPartNum As String
Dim strPartCnt As String
Dim strPAUnit As String
Dim strPartInfo As String
Dim strPullNum As String
Dim strBinNum As String
Dim strECNum As String
Dim strBldStation As String

Dim strPartNumFld As String
Dim strPartCntFld As String
Dim strPAUnitFld As String
Dim strPartInfoFld As String
Dim strPullNumFld As String
Dim strBinNumFld As String
Dim strECNumFld As String
Dim strBldStationFld As String

Private txtKeyPress As New EsiKeyBd


Private Sub ClearArrays(iSize As Integer)
    Erase arrValue
    ReDim arrValue(0 To 96, iSize)
End Sub



Private Sub cmdCan_Click()
   sLastPrefix = cmbPre
   Unload Me

End Sub



Private Sub cmdExport_Click()
   Dim strExFileName As String
   
   ' Clear the data
   fileDlg.filename = ""
   fileDlg.Filter = "XLS File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
      strExFileName = ""
      MsgBox "Select file name to export EDI data.", vbOKOnly
      Exit Sub
   Else
       strExFileName = fileDlg.filename
   End If
   
'   strFilePath = txtEdiFilePath.Text
'   lPos = InStrRev(strFilePath, "\")
'   strFilePath = Mid(strFilePath, 1, lPos)
'   strFilePath = strFilePath & "Forcast"
   
   MouseCursor ccHourglass
   
   If (optEdiFile(0).Value = True) Then
      Exp850toExcel strExFileName
   ElseIf (optEdiFile(1).Value = True) Then
      Exp862toExcel strExFileName
   End If

   MouseCursor ccArrow
   Exit Sub
   
DiaErr1:
   sProcName = "cmdExport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub Exp850toExcel(strExFileName As String)
   
   Dim RdoHdEdi As ADODB.Recordset
   
   Dim sFieldsToExport(8) As String
   sFieldsToExport(0) = "EDISENDERCODE"
   sFieldsToExport(1) = "PARTNUM"
   sFieldsToExport(2) = "FSTQTY"
   sFieldsToExport(3) = "DUEDATE"
   sFieldsToExport(4) = "LASTSHPQTY"
   sFieldsToExport(5) = "LASTRECVEDDATE"
   sFieldsToExport(6) = "SHIPNAME"
   sFieldsToExport(7) = "SHIPNAME1"

      sSql = "SELECT a.EDISENDERCODE AS EDISENDERCODE, SHIPNAME, SHIPNAME1,PARTNUM, FSTQTY," _
            & " DUEDATE, LASTSHPQTY, LASTRECVEDDATE FROM Inhd830_850EDI a, Init830_850EDI b" _
              & " Where A.EDISENDERCODE = b.EDISENDERCODE"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdEdi, ES_STATIC)
   
   If bSqlRows Then
      SaveAsExcel RdoHdEdi, sFieldsToExport, strExFileName, False, True, True
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If
   
   Set RdoHdEdi = Nothing

End Sub

Private Sub Exp862toExcel(strExFileName As String)

   Dim RdoHdEdi As ADODB.Recordset
   
   Dim sFieldsToExport(10) As String
   sFieldsToExport(0) = "SHIPCODE"
   sFieldsToExport(1) = "PARTNUM"
   sFieldsToExport(2) = "PONUMBER"
   sFieldsToExport(3) = "PORELEASE"
   sFieldsToExport(4) = "SHIPNAME"
   sFieldsToExport(5) = "SHPQTY"
   sFieldsToExport(6) = "DUEDATE"
   sFieldsToExport(7) = "SHIPADDRESS1"
   sFieldsToExport(8) = "PULLNUM"
   sFieldsToExport(9) = "BINNUM"
   

   sSql = "SELECT DISTINCT a.SHIPCODE SHIPCODE, a.PARTNUM PARTNUM, b.PONUMBER PONUMBER," _
            & " PORELEASE , SHIPNAME, SHPQTY, DUEDATE, SHIPADDRESS1, PULLNUM, BINNUM " _
         & " FROM Init862_EDI a, Inhd862_EDI b" _
            & " WHERE a.ShipCode = b.ShipCode AND" _
            & " a.PONumber = b.PONumber AND EDI_ELEMENTTYPE <> 'D3'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdEdi, ES_STATIC)
   
   If bSqlRows Then
      SaveAsExcel RdoHdEdi, sFieldsToExport, strExFileName, False, True, True
   Else
      MsgBox "No records found. Please try again.", vbOKOnly
   End If
   
   Set RdoHdEdi = Nothing

End Sub

Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub cmdImport_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   On Error GoTo DiaErr1
   strFilePath = txtEdiFilePath.Text
   
   If (Trim(strFilePath) = "") Then
      MsgBox "Please select the EDI file to create Sales Order.", _
            vbInformation, Caption
         Exit Sub
   End If

   If (optEdiFile(0).Value = True) Then
      Grd862.Visible = False
      
      DeleteOldData ("Inhd830_EDI")
      DeleteOldData ("Init830_EDI")
      DeleteOldData ("Inhd830_850EDI")
      DeleteOldData ("Init830_850EDI")
      
      Dim strEDIDataType As String
      
      CheckFileFormat strFilePath, strEDIFormat
      
      If (strEDIFormat = "850_PO") Then
         Grd850.Visible = True
         Grd830.Visible = False
         cmdExport.Visible = False
         
         strEDIDataType = "850_EDI"
         ImportEDIFile strFilePath, strEDIDataType
         
         Fill850Grid (CStr(cmbCst))
      Else
         Grd850.Visible = False
         Grd830.Visible = True
         cmdExport.Visible = True
         
         strEDIDataType = "830_EDI"
         ImportEDIFile strFilePath, strEDIDataType
         ' Forecast Data
         Fill830Grid (CStr(cmbCst)) '"830_PlanSchedule"
      End If
      
   ElseIf (optEdiFile(1).Value = True) Then
      Grd850.Visible = False
      Grd830.Visible = False
      cmdExport.Visible = True
      Grd862.Visible = True
      
      DeleteOldData ("Inhd862_EDI")
      DeleteOldData ("Init862_EDI")
      
      strEDIDataType = "862_EDI"
      ImportEDIFile strFilePath, strEDIDataType
      Fill862Grid (CStr(cmbCst))
   Else
      MsgBox "Please select the EDI file type.", _
            vbInformation, Caption
      Exit Sub
   End If
   
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Sub

Private Sub cmdOpenDia_Click()
   fileDlg.Filter = "EDI File (*.edi) | *.edi"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtEdiFilePath.Text = ""
   Else
       txtEdiFilePath.Text = fileDlg.filename
   End If
End Sub

Private Sub cmdSalesOrder_Click()
   
   If (CStr(cmbCst) = "ALL") Then
      MsgBox "Please select customer.", _
            vbInformation, Caption
      Exit Sub
   End If
   
   If (optEdiFile(0).Value = True) Then
      Create850SalesOrder
      
   ElseIf (optEdiFile(1).Value = True) Then
      Create862SalesOrder
   Else
      MsgBox "Please select the EDI file type.", _
            vbInformation, Caption
      Exit Sub
   End If
   
   If optSORev.Value <> vbChecked Then
      GetLastSalesOrder sOldSoNumber, sNewsonumber, True
   End If
   
   
End Sub
   
Private Function Create850SalesOrder()
   Dim bByte As Byte
   Dim lNewSoNum As Long
   Dim iList As Long
   
   Dim strBuyerOrderNumber As String
   Dim strContactName As String
   Dim strContactNum As String
   Dim strContactType As String
   Dim strCusFullName As String
   Dim strShipName As String
   Dim strRefTypeCode As String
   Dim strNewAddress As String
   Dim strPONum As String
   Dim strPartID As String
   Dim strQty As String
   Dim strRefDesc As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strNewSO As String
   Dim strCusName As String
   
   For iList = 1 To Grd850.Rows - 1
      Grd850.Col = 0
      Grd850.Row = iList
      
      ' Only if the part is checked
      If Grd850.CellPicture = Chkyes.Picture Then
         
         Grd850.Col = 1
         strBuyerOrderNumber = Replace(Trim(Grd850.Text), Chr$(42), "")
         
         Grd850.Col = 3
         strPartID = Grd850.Text
         
         ' Get Customer Ref from Customer Name
         ' NOT needed now
         'GetCustomerRef strCusFullName, strCusName
         strCusName = cmbCst
         ' Get Customer P
' MM TODO: No need to warn user
'         bByte = CheckForCustomerPO(strCusName, strBuyerOrderNumber)
'         If bByte = 1 Then
'            bByte = MsgBox("The Customer PO Is In Use. Continue?", _
'                 ES_YESQUESTION, Caption)
'            If bByte = vbNo Then
'               Exit Function
'            End If
'         End If
         
         ' Get new Sales Order number
         Dim strSoType As String
         Dim strItem As String
         Dim strSoNum As String
         
         strSoType = cmbPre
         Dim bSoExists As Boolean

         ' if the SOPO exists Warn the users
         bSoExists = CheckOfExistingSO(strBuyerOrderNumber, strPartID, strSoNum)
         If (bSoExists = False) Then

            GetNewSO strNewSO, strSoType
            
            Dim strSenderCode As String
            strSenderCode = "097248199"
            CreateSOFromEDIData strSenderCode, strBuyerOrderNumber, strNewSO, strSoType, strCusName
            
            Grd850.Col = 1
            Grd850.Text = "*" & Trim(strBuyerOrderNumber)
            
            If optSORev.Value = vbChecked Then
               OptSoEDI = vbChecked
               SaleSLe02a.Show
               SaleSLe02a.OptSoEDI = vbChecked
               SaleSLe02a.SetFocus
               SaleSLe02a.cmbSon.SetFocus
            End If
         Else
            
            Dim sMsg As String
            sMsg = "The Purchase number '" & strBuyerOrderNumber & "' has Existing Sales Order '" & strSoNum & "'."
            MsgBox sMsg, vbOKOnly, Caption
         
         End If
      End If
   Next
   
End Function

Private Function Create862SalesOrder()
   Dim bByte As Byte
   Dim lNewSoNum As Long
   Dim iList As Long
   
   Dim strBuyerOrderNumber As String
   Dim strContactName As String
   Dim strContactNum As String
   Dim strContactType As String
   Dim strCusFullName As String
   Dim strShipName As String
   Dim strRefTypeCode As String
   Dim strNewAddress As String
   Dim strPONumber As String
   Dim strPORel As String
   Dim strPartID As String
   Dim strPart
   Dim strQty As String
   Dim strRefDesc As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim strSoNum As String
   Dim strCusName As String
   Dim strShipCode As String
   Dim strPrevPO As String
   Dim strBook As String
   Dim RdoEdi As ADODB.Recordset

    'look up book for customer
    sSql = "SELECT case when CUPRICEBOOK = '' then CUNICKNAME else CUPRICEBOOK end from CustTable where CUNICKNAME = '" & cmbCst & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
    If bSqlRows Then
        strBook = Trim(RdoEdi(0))
    End If
   
   strPrevPO = ""
   For iList = 1 To Grd862.Rows - 1
      Grd862.Col = 0
      Grd862.Row = iList
      
      ' Only if the part is checked
      If Grd862.CellPicture = Chkyes.Picture Then
         
         Grd862.Col = 1
         strShipCode = Trim(Grd862.Text)
         
         Grd862.Col = 3
         strPORel = Trim(Grd862.Text)
         
         
         Grd862.Col = 2
         strBuyerOrderNumber = Trim(Grd862.Text)
         
         Grd862.Col = 3
         strPORel = Trim(Grd862.Text)
         
         Grd862.Col = 4
         strPartID = Replace(Trim(Grd862.Text), Chr$(42), "")
         
         
         ' Get Customer Ref from Customer Name
         ' NOT needed now
         'GetCustomerRef strCusFullName, strCusName
         strCusName = cmbCst
         ' Get Customer P
         If (strPrevPO <> strBuyerOrderNumber) Then
' MM TODO: No need to warn user
'            bByte = CheckForCustomerPO(strCusName, strBuyerOrderNumber)
'            If bByte = 1 Then
'               bByte = MsgBox("The Customer PO Is In Use. Continue?", _
'                    ES_YESQUESTION, Caption)
'               If bByte = vbNo Then
'                  Exit Function
'               End If
'            End If
            strPrevPO = strBuyerOrderNumber
         End If
         
         ' Get new Sales Order number
         Dim strSoType As String
         Dim strItem As String
         strSoType = cmbPre
         
         Dim strSenderCode As String
         strSenderCode = "097248199"
         
         Dim strCustCont As String
         Dim strShpToName As String
         Dim strShpAddr1 As String
         Dim strShpAddr2 As String
         Dim strShipTo4 As String
         Dim strShipTo5 As String
         Dim strShpToAddress As String
         Dim strPOItem As String
         Dim strUOM As String
         Dim strDueDt As String
         Dim strShpDt As String
         Dim bPartFound, bIncRow As Boolean
         Dim strSORemark As String
         Dim bSoExists As Boolean
         Dim iItem As Integer
         Dim strBldStation As String
         Dim strECNum As String
         
         bSoExists = False
         
         sSql = "select DISTINCT a.PONUMBER, b.PARTNUM, b.PAUNITS, PORELEASE," _
                  & " SHPQTY, DUEDATE, SHPDATE,a.SHIPCODE," _
                  & " SHIPNAME, SHIPADDRESS1, SHIPADDRESS2," _
                  & " PULLNUM, BINNUM, BUILDSTATION, ECNUMBER " _
               & " FROM Inhd862_EDI a, Init862_EDI b" _
               & " WHERE a.PONUMBER = b.PONUMBER AND EDI_ELEMENTTYPE = 'D2'" _
               & " AND a.PONUMBER = '" & strBuyerOrderNumber & "' AND " _
               & " b.PORELEASE = '" & strPORel & "' " _
               & " AND a.SHIPCODE = '" & strShipCode & "' " _
               & " AND b.PARTNUM = '" & strPartID & "'"
          
         Debug.Print sSql
         
         bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
         
         If bSqlRows Then
            With RdoEdi
            While Not .EOF
               
               strPONumber = Trim(!PONumber)
               strPOItem = Trim(!PORELEASE)
               strPartID = Trim(!PartNum)
               strQty = Trim(!SHPQTY)
               strDueDt = ConvertToDate(Trim(!DUEDATE))
               strShpDt = ConvertToDate(Trim(!SHPDATE))
               strShpToName = Trim(!SHIPNAME)
               strShpAddr1 = Trim(!SHIPADDRESS1)
               strShpAddr2 = Trim(!SHIPADDRESS2)
               strUOM = Trim(!PAUNITS)
               
               strPullNum = Trim(!PULLNUM)
               strBinNum = Trim(!BINNUM)
               
               strBldStation = Trim(!BUILDSTATION)
               strECNum = Trim(!ECNUMBER)
               
               ' Commenting 9/5/2016
               'MM GetPartPrice strPartID, strUnitPrice
               'strBook = "" MUST LOOK UP PER CUSTOMER.  NOW DONE ABOVE
               GetBookPrice strPartID, strBook, strUnitPrice
               
               'If (strUnitPrice = "") Then
               If (CDec("0" & strUnitPrice) = 0) Then
                  GetPartPrice strPartID, strUnitPrice
               End If
               
               strCustCont = "" 'Trim(!SHIPPERSON)
               strContactNum = ""
               strSORemark = ""
               
               MakeAddress strShpToName, strShpAddr1, strShpAddr2, _
                        "", "", strShpToAddress
               
               ' if the SO header is alrady added don't add the PO again
               bSoExists = CheckOfExistingSO(strPONumber, strPartID, strSoNum)
               If (bSoExists = False) Then
                  GetNewSO strSoNum, strSoType
                  AddSalesOrder strSoNum, strPONumber, strCustCont, strContactNum, _
                                    strShpToName, strCusName, strShpToAddress, strSoType, strSORemark
               Else
                  ' Get customer profile
                  Dim bGoodCust As Byte
                  bGoodCust = GetCustomerData(strCusName)
                  If bCutOff = 1 Then
                     MsgBox "This Customer's Credit Is On Hold.", _
                        vbInformation, Caption
                     bGoodCust = 0
                  End If
                  If Not bGoodCust Then Exit Function
               
               End If
               
               ' Add So items
               AddSoItem strSoNum, CStr(strPOItem), strPONumber, strPOItem, _
                  strPartID, strQty, strUnitPrice, strDueDt, strPullNum, strBinNum, _
                  strBldStation, strECNum, strShpDt
      
               .MoveNext
            Wend
            .Close
            End With
         End If
         
         Set RdoEdi = Nothing
         
         Grd862.Col = 4
         Grd862.Text = "*" & Trim(strPartID)
         
         'MsgBox "Added Sales Order Items.", vbInformation, Caption
         
         
         If optSORev.Value = vbChecked Then
            txtSon = strSoNum
            OptSoEDI = vbChecked
            SaleSLe02a.Show
            SaleSLe02a.OptSoEDI = vbChecked
            SaleSLe02a.SetFocus
            SaleSLe02a.cmbSon.SetFocus
         End If
      
      End If
   Next
   

End Function


Private Sub AddSalesOrder(strNewSO As String, strBuyerOrderNumber As String, _
                     strContactName As String, strContactNum As String _
                     , strShipName As String, strCusName As String, strNewAddress As String, _
                     strSoType As String, strSORemark As String)
                     
   Dim sNewDate As Variant
   Dim bGoodCust As Byte
   
   bGoodCust = GetCustomerData(strCusName)
   If bCutOff = 1 Then
      MsgBox "This Customer's Credit Is On Hold.", _
         vbInformation, Caption
      bGoodCust = 0
   End If
   If Not bGoodCust Then Exit Sub
   On Error GoTo DiaErr1
   
   sNewDate = Format(ES_SYSDATE, "mm/dd/yy")
'   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
'          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
'          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
'          & "SOTEXT,SOTAXEXEMPT,SOPO, SOREMARKS) " _
'          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
'          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
'          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
'          & sRegion & "','" & sSterms & "','" & sVia & "','" _
'          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
'          & "," & iFrtDays & ",'" & strNewSO & "','" & sTaxExempt & "','" _
'          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
'
   sSql = "INSERT SohdTable (SONUMBER,SOTYPE,SOCUST,SODATE," _
          & "SOSALESMAN,SOSTNAME,SOSTADR,SODIVISION,SOREGION,SOSTERMS," _
          & "SOVIA,SOFOB,SOARDISC,SODAYS,SONETDAYS,SOFREIGHTDAYS," _
          & "SOTAXEXEMPT,SOPO, SOREMARKS) " _
          & "VALUES(" & Val(strNewSO) & ",'" & strSoType & "','" _
          & strCusName & "','" & sNewDate & "','" & sSlsMan & "','" _
          & strShipName & "','" & strNewAddress & "','" & sDivision & "','" _
          & sRegion & "','" & sSterms & "','" & sVia & "','" _
          & sFob & "'," & cDiscount & "," & iDays & "," & iNetDays _
          & "," & iFrtDays & ",'" & sTaxExempt & "','" _
          & Trim(strBuyerOrderNumber) & "','" & strSORemark & "')"
   
   'Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   If clsADOCon.RowsAffected Then
      On Error Resume Next
      MsgBox "Sales Order Added.", vbInformation, Caption
      sSql = "UPDATE SohdTable SET SOCCONTACT='" & strContactName & "'," _
             & "SOCPHONE='" & strContactNum & "',SOCINTFAX='" & sConIntFax _
             & "',SOCFAX='" & sConFax & "',SOCEXT=" & sConExt _
             & " WHERE SONUMBER=" & Val(strNewSO) & ""
      Debug.Print sSql
      
      clsADOCon.ExecuteSql sSql ', rdExecDirect
      
      sSql = "UPDATE ComnTable SET COLASTSALESORDER='" & Trim(strSoType) _
             & Trim(strNewSO) & "' WHERE COREF=1"
      clsADOCon.ExecuteSql sSql ', rdExecDirect
   
   Else
      MsgBox "Couldn't Add Sales Order.", vbExclamation, Caption
   End If
   Exit Sub
   
DiaErr1:
   MsgBox Err.Description
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub AddSoItem(strNewSO As String, strItem As String, _
                     strBuyerOrderNumber As String, strBuyerItemLine As String, _
                     strPartID As String, strQty As String, _
                     strUnitPrice As String, strReqDt As String, _
                     strPullNum As String, strBinNum As String, _
                     strBldStation As String, strECNum As String, _
                     Optional strShpDt As String = "")

   On Error GoTo DiaErr1
   
'   Dim strShedDt As String
'   ' Create the ShipDate
'   If (iFrtDays > 0) Then
'      strShedDt = Format(DateAdd("d", -iFrtDays, strReqDt), "mm/dd/yy")
'   Else
'      strShedDt = strReqDt
'   End If
   
   Dim strShedDt As String
   ' Create the ShipDate
   
   If (strShpDt <> "") Then
      strShedDt = Format(strShpDt, "mm/dd/yy")
   Else
      strShedDt = strReqDt
      
      If (iFrtDays > 0) Then
         strShedDt = Format(DateAdd("d", -iFrtDays, strReqDt), "mm/dd/yy")
      End If
   End If
   
   
   Dim RdoSoit As ADODB.Recordset
   Dim strShiped As String
   strShiped = ""
   sSql = "SELECT DISTINCT ITSO, ISNULL(ITPSSHIPPED, 0)  ITPSSHIPPED FROM SoitTable WHERE " _
             & " ITSO = '" & strNewSO & "'" _
             & "  AND ITNUMBER = '" & Val(strItem) & "'" ' AND ITPSSHIPPED <> 1"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_FORWARD)
   If bSqlRows Then
          strShiped = RdoSoit!ITPSSHIPPED
      ClearResultSet RdoSoit
      Set RdoSoit = Nothing
      
      If (CInt(strShiped) = 1) Then
         Exit Sub
      Else
         If (Val(strQty) = 0) Then
'            sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strShedDt & "'," _
'                     & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
'                     & "ITACTUAL=NULL, ITCANCELED=1, ITCANCELDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
'                     & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'"

            sSql = "UPDATE SoitTable SET ITQTY = " & Val(strQty) & ", ITSCHED = '" & strShedDt & "'," _
                     & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
                     & "ITACTUAL=NULL, ITCANCELED=1, ITCANCELDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "' " _
                     & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'"

                     
         Else
            ' get the latest revision of Item and them updated the part
            Dim strRev As String
            Dim lRemQty As Long
            
            GetLatesSoitRev strNewSO, strItem, Val(strQty), strRev, lRemQty
                                        
'            sSql = "UPDATE SoitTable SET ITQTY = ISNULL(ITQTY, 0) + " & lRemQty & ", ITSCHED = '" & strShedDt & "'," _
'                     & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
'                     & " BUILDSTATION = '" & strBldStation & "', ECNUMBER = '" & strECNum & "', " _
'                     & " ITCANCELDATE=NULL, ITCANCELED=0 " _
'                     & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'" _
'                               & " AND ITREV = '" & strRev & "'"

            sSql = "UPDATE SoitTable SET ITQTY = ISNULL(ITQTY, 0) + " & lRemQty & ", ITSCHED = '" & strShedDt & "'," _
                     & " ITCUSTREQ = '" & strReqDt & "', PULLNUM = '" & strPullNum & "', BINNUM = '" & strBinNum & "', " _
                     & " BUILDSTATION = '" & strBldStation & "', ECNUMBER = '" & strECNum & "', " _
                     & " ITCANCELDATE=NULL, ITCANCELED=0 " _
                     & " WHERE ITSO = '" & strNewSO & "' AND ITNUMBER = '" & Val(strItem) & "'" _
                           & " AND ITREV = '" & strRev & "'"

            End If
                         
            Debug.Print sSql
            clsADOCon.ExecuteSql sSql ', rdExecDirect
         End If
      
      'Not needed
      'MsgBox "Updated Sales Order '" & strNewSO & "' and Item '" & strItem & "'.", vbExclamation, Caption
      
      Exit Sub
   End If
      
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   
   sSql = "INSERT SoitTable (ITSO,ITNUMBER,ITCUSTITEMNO, ITPART,ITQTY,ITCUSTREQ, ITSCHED,ITBOOKDATE," _
          & "ITDOLLORIG, ITDOLLARS, ITUSER, PULLNUM, BINNUM, BUILDSTATION, ECNUMBER) " _
          & "VALUES(" & strNewSO & "," & strItem & ",'" & strBuyerItemLine & "','" _
          & Compress(strPartID) & "'," & Val(strQty) & ",'" & strReqDt & "','" & strShedDt & "','" _
          & Format(ES_SYSDATE, "mm/dd/yy") & "','" & CCur(strUnitPrice) & "','" _
          & CCur(strUnitPrice) & "','" & sInitials & "','" & strPullNum & "','" _
          & strBinNum & "','" & strBldStation & "','" & strECNum & "')"
   
   Debug.Print sSql
   
   clsADOCon.ExecuteSql sSql ', rdExecDirect
   
   'Add commission if applicable.
'   If cmdCom.Enabled Then
     Dim Item As New ClassSoItem
     Dim bUserMsg As Boolean
     bUserMsg = False
     Item.InsertCommission CLng(strNewSO), CLng(strItem), "", ""
     Item.UpdateCommissions CLng(strNewSO), CLng(strItem), "", bUserMsg
 '  End If
   
   clsADOCon.CommitTrans
   Exit Sub
   
DiaErr1:
   sProcName = "addsoitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetLatesSoitRev(ByVal strNewSO As String, ByVal strItem As String, _
            ByVal lEdiQty As Long, ByRef strRev As String, ByRef lRemQty As Long)
            
   Dim RdoSoit As ADODB.Recordset
   Dim strShiped As String
   Dim itemQty As Long
   
   lRemQty = 0
   strShiped = ""
   sSql = "SELECT DISTINCT ITSO, ITNUMBER, ITREV, ITQTY, ISNULL(ITPSSHIPPED, 0)  ITPSSHIPPED FROM SoitTable WHERE " _
             & " ITSO = '" & strNewSO & "'" _
             & "  AND ITNUMBER = '" & Val(strItem) & "'"

   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSoit, ES_FORWARD)
   If bSqlRows Then
   
      With RdoSoit
      While Not .EOF
            strRev = !itrev
            itemQty = !ITQty
            strShiped = !ITPSSHIPPED
            lEdiQty = lEdiQty - itemQty
            ' remainder
            lRemQty = lEdiQty
         .MoveNext
      Wend
      .Close
      End With
      Set RdoSoit = Nothing
   End If
            
End Function

Public Sub GetNewSO(ByRef sNewSo As String, ByVal sSoType As String)
   Dim RdoSon As ADODB.Recordset
   Dim lSales As Long
   On Error GoTo DiaErr1
   
   sSql = "SELECT (MAX(SONUMBER)+ 1)AS SalesOrder FROM SohdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         If Not IsNull(.Fields(0)) Then
            sNewSo = "" & Format$(!SalesOrder, SO_NUM_FORMAT)
         Else
            sNewSo = SO_NUM_FORMAT
         End If
         ClearResultSet RdoSon
      End With
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "GetNewSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Sub


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



Private Function GetCustomerData(strCusName As String) As Byte
   Dim RdoCst As ADODB.Recordset
   sCust = Compress(strCusName)
   On Error GoTo DiaErr1
   sSql = "SELECT CUREF,CUSTNAME,CUSTNAME,CUSTADR,CUARDISC," _
          & "CUDAYS,CUNETDAYS,CUDIVISION,CUREGION,CUSTERMS," _
          & "CUVIA,CUFOB,CUSALESMAN,CUDISCOUNT,CUSTSTATE," _
          & "CUSTCITY,CUSTZIP,CUCCONTACT,CUCPHONE,CUCEXT,CUCINTPHONE," _
          & "CUFRTDAYS,CUINTFAX,CUFAX,CUTAXEXEMPT,CUCUTOFF " _
          & "FROM CustTable WHERE CUREF='" & strCusName & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst)
   If bSqlRows Then
      With RdoCst
         bCutOff = !CUCUTOFF
         sStName = "" & Trim(!CUSTNAME)
         sStAdr = "" & Trim(!CUSTADR) & vbCrLf _
                  & "" & Trim(!CUSTCITY) & " " & Trim(!CUSTSTATE) _
                  & "  " & Trim(!CUSTZIP)
         sDivision = "" & Trim(!CUDIVISION)
         sRegion = "" & Trim(!CUREGION)
         sSterms = "" & Trim(!CUSTERMS)
         sVia = "" & Trim(!CUVIA)
         sFob = "" & Trim(!CUFOB)
         sSlsMan = "" & Trim(!CUSALESMAN)
         sContact = "" & Trim(!CUCCONTACT)
         sConIntPhone = "" & Trim(!CUCINTPHONE)
         sConPhone = "" & Trim(!CUCPHONE)
         sConIntFax = "" & Trim(!CUINTFAX)
         sConFax = "" & Trim(!CUFAX)
         sConExt = "" & Trim(str$(!CUCEXT))
         cDiscount = Format(0 + !CUARDISC, "##0.000")
         iDays = Format(!CUDAYS, "###0")
         iNetDays = Format(!CUNETDAYS, "###0")
         iFrtDays = Format(!CUFRTDAYS, "##0")
         sTaxExempt = "" & Trim(!CUTAXEXEMPT)
         ClearResultSet RdoCst
      End With
      GetCustomerData = True
   Else
      sStName = ""
      sStAdr = ""
      sDivision = ""
      sRegion = ""
      sSterms = ""
      sVia = ""
      sFob = ""
      sSlsMan = ""
      iFrtDays = 0
      MsgBox "Couldn't Retrieve Customer.", vbExclamation, Caption
      GetCustomerData = False
   End If
   Set RdoCst = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcustda"
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

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst, False
   
   If (cmbCst = "ALL") Then
      txtNme = "*** Selected All Customers ***"
   End If
   
   If bOnLoad = 0 And bOptionSel = False Then
      ' Filter the records if selected.
      FilterPOByCustomer (cmbCst)
   End If
End Sub

Private Sub cmbCst_LostFocus()
'   cmbCst = CheckLen(cmbCst, 10)
'   FindCustomer Me, cmbCst, False
'   If (cmbCst = "ALL") Then
'      txtNme = "*** All Customer selected ***"
'   End If
'   lblNotice.Visible = False
'
'   If bOnLoad = 0 Then
'      ' Filter the records if selected.
'      FilterPOByCustomer (cmbCst)
'   End If
End Sub


Private Sub cmbPre_LostFocus()
   Dim a As Integer
   cmbPre = CheckLen(cmbPre, 1)
   On Error Resume Next
   a = Asc(Left(cmbPre, 1))
   If a < 65 Or a > 90 Then
      MsgBox "Must Be Between A and Z..", vbInformation, Caption
      cmbPre = sLastPrefix
   End If
   If Len(Trim(cmbPre)) = 0 Then cmbPre = sLastPrefix
   
End Sub


Private Sub CmdSelAll_Click()
   Dim iList As Integer
   
   If (optEdiFile(0).Value = True) Then
      For iList = 1 To Grd850.Rows - 1
          Grd850.Col = 0
          Grd850.Row = iList
          ' Only if the part is checked
          If Grd850.CellPicture = Chkno.Picture Then
              Set Grd850.CellPicture = Chkyes.Picture
          End If
      Next
   ElseIf (optEdiFile(1).Value = True) Then
      For iList = 1 To Grd862.Rows - 1
          Grd862.Col = 0
          Grd862.Row = iList
          ' Only if the part is checked
          If Grd862.CellPicture = Chkno.Picture Then
              Set Grd862.CellPicture = Chkyes.Picture
          End If
      Next
   Else
      MsgBox "Please select the EDI file type.", _
            vbInformation, Caption
      Exit Sub
   End If

End Sub

Private Sub Form_Activate()
   
   MdiSect.lblBotPanel = Caption
   
   GetLastSalesOrder sOldSoNumber, sNewsonumber, True
   
   sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
            & " A.CUREF = b.CUREF"
            
   LoadComboBox cmbCst, -1
   AddComboStr cmbCst.hWnd, "" & Trim("ALL")
   cmbCst = "ALL"
   txtNme = "*** All Customer selected ***"

   'FillEDICust
   'FillCustomers
   'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
   
   'FindCustomer Me, cmbCst, False
   OptSoEDI.Value = vbUnchecked
     
   Grd830.Visible = True
   Grd862.Visible = False
   
   If bOnLoad Then
       bOnLoad = 0
   End If
    
   MouseCursor (0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiSale", "LastEDIPrefix", sLastPrefix
   
End Sub
Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
   
   Dim iChar As Integer
   sLastPrefix = GetSetting("Esi2000", "EsiSale", "LastEDIPrefix", sLastPrefix)
   If Len(sLastPrefix) = 0 Then sLastPrefix = "E"
   cmbPre = sLastPrefix
   For iChar = 65 To 90
      AddComboStr cmbPre.hWnd, Chr$(iChar)
   Next
    
    
   With Grd850
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
      .Text = "Apply"
      .Col = 1
      .Text = "PO Number"
      .Col = 2
      .Text = "Items"
      .Col = 3
      .Text = "PartNumber"
      .Col = 4
      .Text = "Qty"
      .Col = 5
      .Text = "UnitPrice"
      .Col = 6
      .Text = "Requestdate"
      .Col = 7
      .Text = "ShipTo Name"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1250
      .ColWidth(2) = 500
      .ColWidth(3) = 2500
      .ColWidth(4) = 1000
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 4500
'      .ColWidth(8) = 1000
'      .ColWidth(9) = 1000
'      .ColWidth(10) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
   With Grd862
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      .ColAlignment(9) = 1
   
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "ShipCode"
      .Col = 2
      .Text = "PO Number"
      .Col = 3
      .Text = "PO Rel"
      .Col = 4
      .Text = "PartNumber"
      .Col = 5
      .Text = "Qty"
      .Col = 6
      .Text = "Pull Num"
      .Col = 7
      .Text = "Bin Num"
      .Col = 8
      .Text = "ShipTo Name"
      .Col = 9
      .Text = "DueDate"
      .Col = 10
      .Text = "Ship Address"
      .Col = 11
      .Text = "Contact Person"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 1500
      .ColWidth(3) = 700
      .ColWidth(4) = 2000
      .ColWidth(5) = 700
      .ColWidth(6) = 1000
      .ColWidth(7) = 1000
      .ColWidth(8) = 1000
      .ColWidth(9) = 1000
      .ColWidth(10) = 2500
      .ColWidth(11) = 1500
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
   With Grd830
      .ColAlignment(0) = 0
      .ColAlignment(1) = 1
      .ColAlignment(2) = 1
      .ColAlignment(3) = 1
      .ColAlignment(4) = 1
      .ColAlignment(5) = 1
      .ColAlignment(6) = 1
      .ColAlignment(7) = 1
   
      .Rows = 1
      .Row = 0
      .Col = 0
      .Text = "Apply"
      .Col = 1
      .Text = "SenderCode"
      .Col = 2
      .Text = "Part Number"
      .Col = 3
      .Text = "ForcastQty"
      .Col = 4
      .Text = "DueDate"
      .Col = 5
      .Text = "LastShpQty"
      .Col = 6
      .Text = "DateLastQtyRcv"
      .Col = 7
      .Text = "ShipTo Name"
      .Col = 8
      .Text = "ShipTo Name1"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1000
      .ColWidth(2) = 2000
      .ColWidth(3) = 1000
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 3000
      .ColWidth(8) = 1700
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
    bOnLoad = 1

End Sub

Function ImportEDIFile(ByVal strFilePath As String, strEDIDataType As String) As Integer
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
   ' Read the content if the text file.
   Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
   Dim lngPos As Integer
   Dim bFound As Boolean
   Dim gstrSenderCode As String
' Get a free file number
   nFileNum = FreeFile
   
   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   bFound = False
   Do While Not EOF(nFileNum)
      Line Input #nFileNum, sNextLine
      Debug.Print sNextLine
      
      If (strEDIDataType = "850_EDI") Then
         Decode850EdiFormat sNextLine
      ElseIf (strEDIDataType = "830_EDI") Then
         Decode830EdiFormat sNextLine, gstrSenderCode
      ElseIf (strEDIDataType = "862_EDI") Then
         Decode862EdiFormat sNextLine
      End If
   Loop
   Close nFileNum

   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   If (nFileNum > 0) Then
      Close nFileNum
   End If
   sProcName = "ImportEDIFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Function Fill850Grid(strCust As String) As Integer
   
   Dim strSenderCode As String
   Dim RdoHdEdi As ADODB.Recordset
   
   MouseCursor ccHourglass
   Grd850.Rows = 1
   On Error GoTo DiaErr1
   
   Dim strCustCont, strShipName1, strShipName2 As String
   Dim strShipName3, strShipName4, strShipName5 As String
   Dim strPONumber As String
   Dim strPOItem As String
   Dim strPartID As String
   Dim strPrevPO, strQty, strUOM As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim bPartFound, bIncRow As Boolean
   Dim strTotQty As String
   Dim iItem As Integer
   
   
   strSenderCode = "097248199"

   sSql = "SELECT EDISENDERCODE, SHIPTO1, SHIPTO2, SHIPTO3," _
             & " SHIPTO4 , SHIPTO5, CUSTCONTACT, Inhd830_EDI.PONUMBER AS PONUMBER1," _
             & "POITEM,POPART ,POPAUNIT, POQTY, POREQDT, POAMT " _
            & " FROM Inhd830_EDI, Init830_EDI " _
            & "WHERE Inhd830_EDI.PONUMBER = Init830_EDI.PONUMBER AND " _
            & " EDISENDERCODE = '" & strSenderCode & "'"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdEdi, adOpenStatic)
   
   If bSqlRows Then
      With RdoHdEdi
      While Not .EOF
         
         strPONumber = Trim(!PONUMBER1)
         strPartID = Trim(!POPART)
         strSenderCode = Trim(!EDISENDERCODE)
         strShipName1 = Trim(!SHIPTO1)
         strShipName2 = Trim(!SHIPTO2)
         strShipName3 = Trim(!SHIPTO3)
         strShipName4 = Trim(!SHIPTO4)
         strShipName5 = Trim(!SHIPTO5)
         strCustCont = Trim(!CUSTCONTACT)
         strTotQty = Trim(!POQTY)
         strPOItem = Trim(!POITEM)
         strUOM = Trim(!POPAUNIT)
         strReqDt = ConvertToDate(Trim(!POREQDT))
         strUnitPrice = Trim(!POAMT)
         
         ' Filter the records if custrselected.
         If (CheckPOPrefix(strPONumber, strCust)) Then
            
            Grd850.Rows = Grd850.Rows + 1
            Grd850.Row = Grd850.Rows - 1
            bIncRow = False
            iItem = 1
            
            If (strPrevPO <> strPONumber) Then
               Grd850.Col = 0
               Set Grd850.CellPicture = Chkno.Picture
               Grd850.Col = 1
               Grd850.Text = Trim(strPONumber)
            End If
            
            Grd850.Col = 2
            Grd850.Text = Trim(strPOItem)
            
            sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartID) & "'"
            bPartFound = CheckRecordExits(sSql)
            Grd850.Col = 3
            If (bPartFound = False) Then
               Grd850.Text = "**" & Trim(strPartID)
            Else
               Grd850.Text = Trim(strPartID)
            End If
            
            Grd850.Col = 4
            Grd850.Text = Trim(strTotQty)
            Grd850.Col = 5
            Grd850.Text = Trim(strUnitPrice)
            Grd850.Col = 6
            Grd850.Text = Trim(strReqDt)
            
            Dim strShipAddr As String
            
            If (strShipName1 <> "") Then
               strShipAddr = strShipName1 & ", " & strShipName2
            Else
               strShipAddr = strShipName1 & ", " & strShipName2
            End If
            
            Grd850.Col = 7
            Grd850.Text = Trim(strShipAddr)
         
         End If
         
         strPrevPO = strPONumber
         .MoveNext
      Wend
      .Close
      End With
   End If

   Set RdoHdEdi = Nothing
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "Fill850Grid "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function Decode850EdiFormat(ByVal strEdiData As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As ADODB.Recordset
   Dim strValue As String
   Dim strType As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   Dim strTabName As String
   
   On Error GoTo DiaErr1
   
   If (strEdiData <> "") Then
      iIndex = 2
      iTotLen = Len(strEdiData)
      strType = Mid(strEdiData, 1, iIndex)
      iIndex = iIndex + 1
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'PO' ORDER BY FORATORDER"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
      ReDim arrValue(0 To RdoEdi.RecordCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            
            If (iNumChar > 0) Then
               strValue = Mid(strEdiData, iIndex, iNumChar)
            Else
               'strValue = Mid(strEdiData, iIndex, (iTotLen - iIndex))
               strValue = Mid(strEdiData, iIndex, (iTotLen - iIndex) + 1)
            End If
            
            arrValue(iTotalItems) = RemoveSQLString(Trim(strValue))
            arrFieldName(iTotalItems) = !FieldName
            iIndex = iIndex + iNumChar
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
      End If
      
      For j = 0 To iTotalItems - 1
         If (strFields = "") Then
            strFields = arrFieldName(j)
            strFldVal = "'" & arrValue(j) & "'"
         Else
            strFields = strFields + "," + arrFieldName(j)
            strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
         End If
      Next
      
      If (strFldVal <> "") Then
         If (strType = "H0") Then
            strTabName = "Inhd830_EDI"
         Else
            strTabName = "Init830_EDI"
         End If
         
         sSql = "INSERT INTO " & strTabName & " (" & strFields & ") " _
                & " VALUES (" & strFldVal & ")"
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSql sSql '
      End If
      
   End If

   Exit Function
DiaErr1:
   sProcName = "Decode850EdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Function Fill862Grid(strCust As String) As Integer
   
   Dim strSenderCode As String
   Dim RdoHdEdi As ADODB.Recordset
   Dim RdoItEdi As ADODB.Recordset
   
   
   MouseCursor ccHourglass
   Grd862.Rows = 1
   On Error GoTo DiaErr1
   
   Dim strShipCode As String
   Dim strPONumber As String
   Dim strPartNum As String
   Dim strPORel, strShipName, strShipQty As String
   Dim strShipDate, strShipAddr As String
   Dim strShipPer As String
   Dim bPartFound, bIncRow As Boolean
   Dim strTotQty As String
   Dim strBinNum As String
   Dim strPullNum As String
   Dim iItem As Integer

   sSql = "SELECT DISTINCT a.SHIPCODE SHIPCODE, a.PARTNUM PARTNUM, b.PONUMBER PONUMBER," _
            & " PORELEASE , SHIPNAME, SHPQTY, DUEDATE, SHIPADDRESS1, PULLNUM, BINNUM " _
         & " FROM Init862_EDI a, Inhd862_EDI b" _
            & " WHERE a.ShipCode = b.ShipCode AND" _
            & " a.PONumber = b.PONumber AND EDI_ELEMENTTYPE <> 'D3'"

   Debug.Print sSql

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdEdi, adOpenStatic)
   
   If bSqlRows Then
      With RdoHdEdi
      While Not .EOF
         
         strShipCode = Trim(!SHIPCODE)
         strPartNum = Trim(!PartNum)
         strPONumber = Trim(!PONumber)
         strPORel = Trim(!PORELEASE)
         strShipName = Trim(!SHIPNAME)
         strShipQty = Trim(!SHPQTY)
         strShipDate = ConvertToDate(Trim(!DUEDATE))
         strShipAddr = Trim(!SHIPADDRESS1)
         strShipPer = "" 'Trim(!SHIPPERSON)
         strPullNum = Trim(!PULLNUM)
         strBinNum = Trim(!BINNUM)
         
         If (CheckPOPrefix(strPONumber, strCust)) Then
            Grd862.Rows = Grd862.Rows + 1
            Grd862.Row = Grd862.Rows - 1
            bIncRow = False
            iItem = 1
            
            Grd862.Col = 0
            Set Grd862.CellPicture = Chkno.Picture
            Grd862.Col = 1
            Grd862.Text = Trim(strShipCode)
            
            Grd862.Col = 2
            Grd862.Text = Trim(strPONumber)
            
            Grd862.Col = 3
            Grd862.Text = Trim(strPORel)
            
            sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartNum) & "'"
            bPartFound = CheckRecordExits(sSql)
            Grd862.Col = 4
            If (bPartFound = False) Then
               Grd862.Text = "**" & Trim(strPartNum)
            Else
               Grd862.Text = Trim(strPartNum)
            End If
            
            Grd862.Col = 5
            Grd862.Text = Trim(strShipQty)
            Grd862.Col = 6
            Grd862.Text = Trim(strPullNum)
            Grd862.Col = 7
            Grd862.Text = Trim(strBinNum)
            Grd862.Col = 8
            Grd862.Text = Trim(strShipName)
            Grd862.Col = 9
            Grd862.Text = Trim(strShipDate)
            Grd862.Col = 10
            Grd862.Text = Trim(strShipAddr)
            Grd862.Col = 11
            Grd862.Text = Trim(strShipPer)
         End If
         .MoveNext
      Wend
      .Close
      End With
   End If

   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function Decode862EdiFormat(ByVal strEdiData As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As ADODB.Recordset
   Dim strValue As String
   Dim strType As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   Dim strTabName As String
   
   On Error GoTo DiaErr1
   
   If (strEdiData <> "") Then
      iIndex = 2
      iTotLen = Len(strEdiData)
      strType = Mid(strEdiData, 1, iIndex)
      iIndex = iIndex + 1
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'SHP' ORDER BY FORATORDER"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
      ReDim arrValue(0 To RdoEdi.RecordCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            
            If (iNumChar > 0) Then
               strValue = Mid(strEdiData, iIndex, iNumChar)
            Else
               strValue = Mid(strEdiData, iIndex, ((iTotLen - iIndex) + 1))
            End If
            
            arrValue(iTotalItems) = RemoveSQLString(Trim(strValue))
            arrFieldName(iTotalItems) = !FieldName
            iIndex = iIndex + iNumChar
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
      End If
      
      If (strType = "H1") Then
      
         For j = 0 To iTotalItems - 1
            If (strFields = "") Then
               strFields = arrFieldName(j)
               strFldVal = "'" & arrValue(j) & "'"
            Else
               strFields = strFields + "," + arrFieldName(j)
               strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
            End If
         
            strPartNum = ""
            strPartCnt = ""
            strPAUnit = ""
            strPartInfo = ""
            strBldStation = ""
            strECNum = ""
            
            strPartNumFld = ""
            strPartCntFld = ""
            strPAUnitFld = ""
            strPartInfoFld = ""
            strBldStationFld = ""
            strECNumFld = ""
         
         Next
         strTabName = "Inhd862_EDI"
         
         sSql = "INSERT INTO " & strTabName & " (" & strFields & ") " _
                & " VALUES (" & strFldVal & ")"
      
         Debug.Print sSql
         clsADOCon.ExecuteSql sSql '
      Else
         If (strType = "D1") Then
            ' Partnum
            strPartCntFld = arrFieldName(2)
            strPartCnt = arrValue(2)
            ' Partnum
            strPartNumFld = arrFieldName(3)
            strPartNum = arrValue(3)
            ' EC Number
            strECNumFld = arrFieldName(5)
            strECNum = arrValue(5)
            ' Pull#
            strPullNumFld = arrFieldName(6)
            strPullNum = arrValue(6)
            ' Partnum
            strPAUnitFld = arrFieldName(7)
            strPAUnit = arrValue(7)
            ' BinNum
            strBinNumFld = arrFieldName(8)
            strBinNum = arrValue(8)
            ' Build Station
            strBldStationFld = arrFieldName(9)
            strBldStation = arrValue(9)
            
         Else
            For j = 0 To iTotalItems - 1
               If (strFields = "") Then
                  strFields = arrFieldName(j)
                  strFldVal = "'" & arrValue(j) & "'"
               Else
                  strFields = strFields + "," + arrFieldName(j)
                  strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
               End If
            Next
         
            strTabName = "Init862_EDI"
            
            strPartInfoFld = "," + strPartNumFld + "," + strPartCntFld + "," + _
                           strPAUnitFld + "," + strPullNumFld + "," + _
                           strBinNumFld + "," + strECNumFld + "," + _
                           strBldStationFld + ","
            strPartInfo = ",'" + strPartNum + "','" + strPartCnt + "','" + _
                        strPAUnit + "','" + strPullNum + "','" + strBinNum + _
                           "','" + strECNum + "','" + strBldStation + "',"
            
            If (Trim(strPartNumFld) <> "") Then
               sSql = "INSERT INTO " & strTabName & " (EDI_ELEMENTTYPE" & strPartInfoFld & strFields & ") " _
                      & " VALUES ('" & strType & "'" & strPartInfo & strFldVal & ")"
               
               Debug.Print sSql
               clsADOCon.ExecuteSql sSql '
            End If
            
         End If
      End If
      
      
   End If

   Exit Function
DiaErr1:
   sProcName = "Decode862EdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function


Function Fill830Grid(strCust As String) As Integer
   
   Dim RdoHdEdi As ADODB.Recordset
   Dim RdoItEdi As ADODB.Recordset
   
   MouseCursor ccHourglass
   Grd830.Rows = 1
   On Error GoTo DiaErr1
       
   ' Read the content if the text file.
   Dim nFileNum As Integer, sText As String, sNextLine As String, lLineCount As Long
   Dim lngPos As Integer
   Dim bFound As Boolean
' Get a free file number
   nFileNum = FreeFile
      
   Dim strSenderCode As String
   Dim strPrevSenderCode As String
   Dim strCustCont, strShipName, strShipName1 As String
   Dim strPartNum As String, strFstQty As String, strLastShpQty As String
   Dim strDueDt As String, strLastShpDate As String
   Dim bPartFound, bIncRow As Boolean
   Dim strTotQty As String
   Dim iItem As Integer
   
   sSql = "SELECT a.EDISENDERCODE, SHIPNAME, SHIPNAME1,PARTNUM, FSTQTY," _
            & " DUEDATE, LASTSHPQTY, LASTRECVEDDATE FROM Inhd830_850EDI a, Init830_850EDI b" _
              & " Where A.EDISENDERCODE = b.EDISENDERCODE"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdEdi, adOpenStatic)
   
   If bSqlRows Then
      With RdoHdEdi
      While Not .EOF
         
         strSenderCode = Trim(!EDISENDERCODE)
         strPartNum = Trim(!PartNum)
         strFstQty = Trim(Val(!FSTQTY))
         strDueDt = ConvertToDate(Trim(!DUEDATE))
         strLastShpQty = Trim(Val(!LASTSHPQTY))
         strLastShpDate = ConvertToDate(Trim(!LASTRECVEDDATE))
         strShipName = Trim(!SHIPNAME)
         strShipName1 = Trim(!SHIPNAME1)
         
         Grd830.Rows = Grd830.Rows + 1
         Grd830.Row = Grd830.Rows - 1
         bIncRow = False
         iItem = 1
         
         Grd830.Col = 0
         Set Grd830.CellPicture = Chkno.Picture
         
         If (strPrevSenderCode <> strSenderCode) Then
            Grd830.Col = 1
            Grd830.Text = Trim(strSenderCode)
         End If
         
         
         sSql = "SELECT Partnum FROM partTable where partref = '" & Compress(strPartNum) & "'"
         bPartFound = CheckRecordExits(sSql)
         Grd830.Col = 2
         If (bPartFound = False) Then
            Grd830.Text = "**" & Trim(strPartNum)
         Else
            Grd830.Text = Trim(strPartNum)
         End If
         
         Grd830.Col = 3
         Grd830.Text = Trim(strFstQty)
         Grd830.Col = 4
         Grd830.Text = Trim(strDueDt)
         
         If (Val(strLastShpQty) <> 0) Then
            Grd830.Col = 5
            Grd830.Text = Trim(strLastShpQty)
         End If
         
         If (strLastShpDate <> "") Then
            Grd830.Col = 6
            Grd830.Text = Trim(strLastShpDate)
         End If
         
         Grd830.Col = 7
         Grd830.Text = Trim(strShipName)
         Grd830.Col = 8
         Grd830.Text = Trim(strShipName1)
         
         strPrevSenderCode = strSenderCode
         
         .MoveNext
      Wend
      .Close
      End With
   End If

   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "Fill830Grid "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Function Decode830EdiFormat(ByVal strEdiData As String, ByRef strSenderCode As String)
   Dim iIndex As Integer
   Dim j As Integer
   Dim RdoEdi As ADODB.Recordset
   Dim strValue As String
   Dim strType As String
   Dim iTotLen As Integer
   Dim iTotalItems As Integer
   Dim iNumChar As Integer
   Dim strFields As String
   Dim strFldVal As String
   Dim strTabName As String
   
   On Error GoTo DiaErr1
   
   If (strEdiData <> "") Then
      iIndex = 1
      iTotLen = Len(strEdiData)
      strType = Mid(strEdiData, 1, iIndex)
      iIndex = iIndex + 1
      sSql = "SELECT FIELDNAME,NUMCHARS FROM ProEdiFormat WHERE " _
             & "HEADER = '" & strType & "' AND IMPORTTYPE = 'PO' ORDER BY FORATORDER"
      
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
      ReDim arrValue(0 To RdoEdi.RecordCount + 1)
      ReDim arrFieldName(0 To RdoEdi.RecordCount + 1)
      If bSqlRows Then
         With RdoEdi
         iTotalItems = 0
         While Not .EOF
            iNumChar = !NUMCHARS
            
            If (iNumChar > 0) Then
               strValue = Mid(strEdiData, iIndex, iNumChar)
            Else
               strValue = Mid(strEdiData, iIndex, (iTotLen - iIndex))
            End If
            
            arrValue(iTotalItems) = RemoveSQLString(Trim(strValue))
            arrFieldName(iTotalItems) = !FieldName
            iIndex = iIndex + iNumChar
            iTotalItems = iTotalItems + 1
            .MoveNext
         Wend
         .Close
         End With
      End If
      
      For j = 0 To iTotalItems - 1
         If (strFields = "") Then
            strFields = arrFieldName(j)
            strFldVal = "'" & arrValue(j) & "'"
         Else
            strFields = strFields + "," + arrFieldName(j)
            strFldVal = strFldVal + "," + "'" + arrValue(j) + "'"
         End If
      Next
      
      If (strFldVal <> "") Then
         If (strType = "H") Then
            strTabName = "Inhd830_850EDI"
            strSenderCode = Trim(arrValue(0))
         
            sSql = "INSERT INTO " & strTabName & " (" & strFields & ") " _
                   & " VALUES (" & strFldVal & ")"
         
         
         Else
            strTabName = "Init830_850EDI"
         
            sSql = "INSERT INTO " & strTabName & " (EDISENDERCODE," & strFields & ") " _
                   & " VALUES ('" & strSenderCode & "'," & strFldVal & ")"
         End If
         
         Debug.Print sSql
         clsADOCon.ExecuteSql sSql '
      End If
      
   End If

   Exit Function
DiaErr1:
   sProcName = "Decode830EdiFormat"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function



Function CreateSOFromEDIData(ByVal strSenderCode As String, ByVal strInputPONum As String, _
               ByVal strNewSO As String, ByVal strSoType As String, ByVal strCusName As String) As Integer
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim RdoEdi As ADODB.Recordset
   Dim strCustCont As String
   Dim strShpToName As String
   Dim strShipTo2 As String
   Dim strShipTo3 As String
   Dim strShipTo4 As String
   Dim strShipTo5 As String
   Dim strShpToAddress As String
   Dim strPOItem As String
   Dim strPartID As String
   Dim strUOM As String
   Dim strUnitPrice As String
   Dim strReqDt As String
   Dim bPartFound, bIncRow As Boolean
   Dim strQty As String
   Dim strContactNum As String
   Dim strSORemark As String
   Dim bSOHdAdded As Boolean
   Dim iItem As Integer
   Dim strBook As String
   
   bSOHdAdded = False
   
   sSql = "SELECT EDISENDERCODE, SHIPTO1, SHIPTO2, SHIPTO3," _
             & " SHIPTO4 , SHIPTO5, CUSTCONTACT, Inhd830_EDI.PONUMBER AS PONUMBER1," _
             & "POITEM,POPART ,POPAUNIT, POQTY, POREQDT, POAMT " _
            & " FROM Inhd830_EDI, Init830_EDI " _
            & "WHERE Inhd830_EDI.PONUMBER = Init830_EDI.PONUMBER AND " _
            & " EDISENDERCODE = '" & strSenderCode & "' AND " _
            & "Inhd830_EDI.PONUMBER = '" & strInputPONum & "'"

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEdi, adOpenStatic)
   
   If bSqlRows Then
      With RdoEdi
      While Not .EOF
         
         'strPONumber = Trim(!PONUMBER1)
         strPartID = Trim(!POPART)
         strSenderCode = Trim(!EDISENDERCODE)
         strShpToName = Trim(!SHIPTO1)
         strShipTo2 = Trim(!SHIPTO2)
         strShipTo3 = Trim(!SHIPTO3)
         strShipTo4 = Trim(!SHIPTO4)
         strShipTo5 = Trim(!SHIPTO5)
         strCustCont = Trim(!CUSTCONTACT)
         strQty = Trim(!POQTY)
         strPOItem = Trim(!POITEM)
         strUOM = Trim(!POPAUNIT)
         strReqDt = ConvertToDate(Trim(!POREQDT))
         'Ignore the proice from EDI..get the price from ProceBook
         'strUnitPrice = Trim(!POAMT)
         
         strBook = "PACPARTS"
         GetBookPrice strPartID, strBook, strUnitPrice

         
         strContactNum = ""
         strSORemark = ""
         
         MakeAddress strShpToName, strShipTo2, strShipTo3, _
                  strShipTo4, strShipTo5, strShpToAddress
         
         ' if the SO header is alrady added don't add the PO again
         If (bSOHdAdded = False) Then
            AddSalesOrder strNewSO, strInputPONum, strCustCont, strContactNum, _
                              strShpToName, strCusName, strShpToAddress, strSoType, strSORemark
            bSOHdAdded = True
         Else
            
            ' Get the customer inforamtion
            Dim bGoodCust As Byte
            
            bGoodCust = GetCustomerData(strCusName)
            If bCutOff = 1 Then
               MsgBox "This Customer's Credit Is On Hold.", _
                  vbInformation, Caption
               bGoodCust = 0
            End If
            If Not bGoodCust Then Exit Function
            
         End If
         
         ' Add So items
         Dim strPullNum As String, strBinNum As String
         Dim strBldStation As String, strECNum As String
         strPullNum = ""
         strBinNum = ""
         strBldStation = ""
         strECNum = ""
         
         
         AddSoItem strNewSO, CStr(strPOItem), strInputPONum, strPOItem, _
            strPartID, strQty, strUnitPrice, strReqDt, strPullNum, _
            strBinNum, strBldStation, strECNum

         .MoveNext
      Wend
      .Close
      End With
   End If
   
   Set RdoEdi = Nothing
   MouseCursor ccArrow
   
   Exit Function
   
DiaErr1:
   sProcName = "CreateSOFromXMLData"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
   
   cUR.CurrentCustomer = cmbCst
   If OptSoEDI.Value = vbUnchecked Then FormUnload
    'FormUnload
    Set SaleSLf14a = Nothing
End Sub

Private Sub Grd850_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd850.Col = 0
      If Grd850.Row >= 1 Then
         If Grd850.Row = 0 Then Grd850.Row = 1
         If Grd850.CellPicture = Chkyes.Picture Then
            Set Grd850.CellPicture = Chkno.Picture
         Else
            Set Grd850.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub

Private Sub Grd862_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd862.Col = 0
      If Grd862.Row >= 1 Then
         If Grd862.Row = 0 Then Grd862.Row = 1
         If Grd862.CellPicture = Chkyes.Picture Then
            Set Grd862.CellPicture = Chkno.Picture
         Else
            Set Grd862.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub

Private Sub Grd830_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd830.Col = 0
      If Grd830.Row >= 1 Then
         If Grd830.Row = 0 Then Grd830.Row = 1
         If Grd830.CellPicture = Chkyes.Picture Then
            Set Grd830.CellPicture = Chkno.Picture
         Else
            Set Grd830.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub


Private Sub cmdClear_Click()
    Dim iList As Integer
    For iList = 1 To Grd850.Rows - 1
        Grd850.Col = 0
        Grd850.Row = iList
        ' Only if the part is checked
        If Grd850.CellPicture = Chkyes.Picture Then
            Set Grd850.CellPicture = Chkno.Picture
        End If
    Next
    For iList = 1 To Grd830.Rows - 1
        Grd830.Col = 0
        Grd830.Row = iList
        ' Only if the part is checked
        If Grd830.CellPicture = Chkyes.Picture Then
            Set Grd830.CellPicture = Chkno.Picture
        End If
    Next
    For iList = 1 To Grd862.Rows - 1
        Grd862.Col = 0
        Grd862.Row = iList
        ' Only if the part is checked
        If Grd862.CellPicture = Chkyes.Picture Then
            Set Grd862.CellPicture = Chkno.Picture
        End If
    Next
End Sub


Private Sub Grd850_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd850.Col = 0
   If Grd850.Row >= 1 Then
      If Grd850.Row = 0 Then Grd850.Row = 1
      If Grd850.CellPicture = Chkyes.Picture Then
         Set Grd850.CellPicture = Chkno.Picture
      Else
         Set Grd850.CellPicture = Chkyes.Picture
      End If
   End If
End Sub

Private Sub Grd862_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd862.Col = 0
   If Grd862.Row >= 1 Then
      If Grd862.Row = 0 Then Grd862.Row = 1
      If Grd862.CellPicture = Chkyes.Picture Then
         Set Grd862.CellPicture = Chkno.Picture
      Else
         Set Grd862.CellPicture = Chkyes.Picture
      End If
   End If
End Sub

Private Sub Grd830_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd830.Col = 0
   If Grd830.Row >= 1 Then
      If Grd830.Row = 0 Then Grd830.Row = 1
      If Grd830.CellPicture = Chkyes.Picture Then
         Set Grd830.CellPicture = Chkno.Picture
      Else
         Set Grd830.CellPicture = Chkyes.Picture
      End If
   End If
End Sub

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

Private Function MakeAddress(strShpTo1 As String, strShpTo2 As String, strStreet As String, _
                  strRegionCode As String, strPostalCode As String, ByRef strShpToAddress As String)

   Dim strNewAddress As String
   
   strShpToAddress = ""
   
   'If (strShpTo1 <> "") Then strNewAddress = strNewAddress & strShpTo1 & vbCrLf
   If (strShpTo2 <> "") Then strNewAddress = strNewAddress & strShpTo2 & vbCrLf
   If (strStreet <> "") Then strNewAddress = strNewAddress & strStreet & vbCrLf
   
   
   If (strPostalCode <> "") Then
      If (strRegionCode <> "") Then
         strNewAddress = strNewAddress & ", " & IIf((strRegionCode <> ""), strRegionCode, "") & " - " & strPostalCode
      Else
         strNewAddress = strNewAddress & " - " & strPostalCode
      End If
   End If
   
   strShpToAddress = strNewAddress

End Function

Private Function CheckRecordExits(sSql As String)
    
   Dim RdoCon As ADODB.Recordset
   
   On Error GoTo ERR1
      
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCon, ES_FORWARD)
   If bSqlRows Then
       CheckRecordExits = True
   Else
       CheckRecordExits = False
   End If
   Set RdoCon = Nothing
   Exit Function
   
ERR1:
    CheckRecordExits = False

End Function

Private Function DeleteOldData(strTableName As String)

   If (strTableName <> "") Then
      sSql = "DELETE FROM " & strTableName
      clsADOCon.ExecuteSql sSql '
   End If

End Function

Private Function CheckOfExistingSO(strPONumber As String, strPartID As String, ByRef strSoNum As String) As Boolean
   Dim RdoSO As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT DISTINCT ISNULL(MAX(SONUMBER),0) SONUMBER  FROM sohdTable,SoitTable WHERE " _
             & " SONUMBER = ITSO AND SOPO = '" & strPONumber & "'" _
             & "  AND ITPART = '" & Compress(strPartID) & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSO, ES_FORWARD)
   If bSqlRows Then
      With RdoSO
         If (Trim(!SoNumber) = 0) Then
            strSoNum = ""
            CheckOfExistingSO = False
         Else
            strSoNum = Trim(!SoNumber)
            CheckOfExistingSO = True
         End If
         ClearResultSet RdoSO
      End With
   Else
      strSoNum = ""
      CheckOfExistingSO = False
      
   End If
   
   Set RdoSO = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "CheckOfExistingSO"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   

End Function

Private Function ConvertToDate(strDate As String) As String
   
   Dim strDateConv As String
   If (Trim(strDate) <> "" And Not IsNull(strDate)) Then
      strDateConv = Mid(strDate, 3, 2) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 1, 2)
   Else
      strDateConv = ""
   End If
   
   ConvertToDate = strDateConv
End Function


Private Function GetBookPrice(strPart As String, strBook As String, ByRef strPrice As String)
   Dim RdoBok As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   If (strBook = "") Then
      sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
             & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) " _
             & "AND (PARTREF='" & Compress(strPart) & "')"
   Else
      sSql = "SELECT PARTREF,PARTNUM,PBIREF,PBIPARTREF,PBIPRICE " _
             & "FROM PartTable,PbitTable WHERE (PARTREF=PBIPARTREF) AND " _
             & "(PBIREF = '" & strBook & "') AND (PARTREF='" & Compress(strPart) & "')"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
   If bSqlRows Then
      With RdoBok
         strPrice = Format(!PBIPRICE, ES_SellingPriceFormat)
         ClearResultSet RdoBok
      End With
   Else
      strPrice = Format(0, ES_SellingPriceFormat)
   End If
   
   Set RdoBok = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
 
 Private Function GetPartPrice(strPart As String, ByRef strPrice As String)
   Dim RdoPrice As ADODB.Recordset
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT PAPRICE FROM PartTable WHERE PARTREF='" & Compress(strPart) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrice, ES_FORWARD)
   If bSqlRows Then
      With RdoPrice
         strPrice = Format(!PAPRICE, ES_SellingPriceFormat)
         ClearResultSet RdoPrice
      End With
   Else
      strPrice = Format(0, ES_SellingPriceFormat)
   End If
   
   Set RdoPrice = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function CheckFileFormat(ByVal strFilePath As String, ByRef strEDIFormat As String)

   On Error GoTo DiaErr1

   Dim strRecipientsCode As String
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   ' Read the content if the text file.
   Dim nFileNum As Integer
   Dim strLine As String
' Get a free file number
   nFileNum = FreeFile
   
   Open strFilePath For Input As nFileNum
   ' Read the contents of the file
   If Not EOF(nFileNum) Then
      Line Input #nFileNum, strLine
      Debug.Print strLine
      
      If (strLine <> "") Then
         strRecipientsCode = Mid(strLine, 16, 15)
         
         If (Trim(strRecipientsCode) = "11555AA") Then
            strEDIFormat = "850_PO"
         Else
            strEDIFormat = "830_PlanSchedule"
         End If
      End If
   End If
   Close nFileNum

   Exit Function

DiaErr1:
   sProcName = "getbookpr"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Private Function RemoveSQLString(varString As Variant) As String
   Dim PartNo As String
   Dim NewPart As String
   
   On Error GoTo modErr1
   PartNo = Trim$(varString)
   If Len(PartNo) > 0 Then
      NewPart = Replace(PartNo, Chr$(39), "")    'single quote
      NewPart = Replace(NewPart, Chr$(44), "")   ' comma
   End If
   RemoveSQLString = NewPart
   Exit Function
   
modErr1:
   Resume modErr2
modErr2:
   On Error Resume Next
   RemoveSQLString = varString
   
End Function

Private Function CheckPOPrefix(strPO As String, strCust As String) As Boolean
   
   Dim RdoCst As ADODB.Recordset
   Dim strPrefix As String
   Dim strPOPre1 As String
   Dim strPOPre2 As String
   Dim strPOPre3 As String
   Dim strPOPre4 As String
   
   MouseCursor 13
   On Error GoTo modErr1
   
   sSql = "SELECT DISTINCT POLETTERREF FROM " _
            & " ASNInfoTable WHERE CUREF = '" & strCust & "'"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   CheckPOPrefix = False
   If bSqlRows Then
      With RdoCst
         strPrefix = Trim(!POLETTERREF)
         ClearResultSet RdoCst
      End With
      
      strPOPre1 = "ST-" & strPrefix
      strPOPre2 = "EX-" & strPrefix
      strPOPre3 = "EA-" & strPrefix
      strPOPre4 = "SA-" & strPrefix
      ' ST-S or EX-S
      If ((Mid(strPO, 1, 4) = strPOPre1) Or _
          (Mid(strPO, 1, 4) = strPOPre2) Or _
          (Mid(strPO, 1, 4) = strPOPre3) Or _
          (Mid(strPO, 1, 4) = strPOPre4)) Then
         CheckPOPrefix = True
      ElseIf ((strPrefix = "#") And IsNumeric(strPO)) Then
          CheckPOPrefix = True
      ElseIf (strPrefix = Mid(strPO, 1, 4)) Then
          CheckPOPrefix = True
      ElseIf (strPrefix = Mid(strPO, 1, 3) And IsNumeric(Mid(strPO, 4, 1))) Then
          CheckPOPrefix = True
      Else
          CheckPOPrefix = False
      End If
   Else
      CheckPOPrefix = True
   End If
   
   Set RdoCst = Nothing
   Exit Function
modErr1:
   sProcName = "CheckPOLetter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set RdoCst = Nothing
   DoModuleErrors MdiSect.ActiveForm
   
End Function
Private Sub FillEDICust()

   MouseCursor 13
   Dim RdoCst As ADODB.Recordset
   On Error GoTo modErr1
   cmbCst.Clear
   
   'as CUREF, b.CUNAME as CUNAME
   sSql = "SELECT a.CUREF FROM " _
            & "ASNInfoTable a, custtable b WHERE A.CUREF = b.CUREF"

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCst, ES_FORWARD)
   If bSqlRows Then
      With RdoCst
         Do Until .EOF
            AddComboStr cmbCst.hWnd, "" & Trim(.Fields(0))
            .MoveNext
         Loop
         ClearResultSet RdoCst
      End With
   End If
   
   Set RdoCst = Nothing
   cmbCst = "ALL"
   
   MouseCursor 0
   Exit Sub
   
modErr1:
   sProcName = "fillcustomers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   Set RdoCst = Nothing
   DoModuleErrors MdiSect.ActiveForm
   
End Sub

   ' Filter the records if selected.
Private Sub FilterPOByCustomer(strCust As String)
      
   If (optEdiFile(0).Value = True) Then
      If (strEDIFormat = "850_PO") Then
         Fill850Grid (strCust)
         
         If (Grd850.Rows = 1) Then
            MsgBox "Purchase Order not found.", vbExclamation, Caption
         End If
      Else
         Fill830Grid (strCust) '"830_PlanSchedule"
      
         If (Grd830.Rows = 1) Then
            MsgBox "Purchase Order not found.", vbExclamation, Caption
         End If
      End If
   ElseIf (optEdiFile(1).Value = True) Then
      Fill862Grid (strCust)
   
      If (Grd862.Rows = 1) Then
         MsgBox "Purchase Order not found.", vbExclamation, Caption
      End If
   Else
      MsgBox "Please select the EDI file type.", _
            vbInformation, Caption
      Exit Sub
   End If


End Sub

Private Sub optEdiFile_Click(Index As Integer)
   
   If (optEdiFile(0).Value = True) Then
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
               & " A.CUREF = b.CUREF AND PACCARDPART = 1"
   ElseIf (optEdiFile(1).Value = True) Then
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
               & " A.CUREF = b.CUREF AND TRUCKPLANT = 1"
   
   End If
        
   bOptionSel = True
   LoadComboBox cmbCst, -1
   AddComboStr cmbCst.hWnd, "" & Trim("ALL")
   cmbCst = "ALL"
   txtNme = "*** All Customer selected ***"
   bOptionSel = False

End Sub
