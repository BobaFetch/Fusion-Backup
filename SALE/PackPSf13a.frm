VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PackPSf13a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export ASN Data from Manifest"
   ClientHeight    =   8175
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   14355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdSearch 
      Caption         =   "..."
      Height          =   255
      Left            =   10560
      TabIndex        =   22
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   1440
      Width           =   255
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   5760
      TabIndex        =   21
      Tag             =   "3"
      ToolTipText     =   "Select XML file to import"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   3600
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtASN 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "ASN Number"
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last ASN Number"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLastASN 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         ToolTipText     =   "Last Sales Order Entered"
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdASN 
      Caption         =   "Create ASN  file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   15
      ToolTipText     =   " Create PS from Sales Order"
      Top             =   1920
      Width           =   1920
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox txtStartDte 
      Height          =   315
      Left            =   1800
      TabIndex        =   11
      Tag             =   "4"
      Top             =   480
      Width           =   1095
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
      Left            =   12720
      TabIndex        =   10
      ToolTipText     =   " Select All"
      Top             =   2520
      Width           =   1560
   End
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Select import"
      Top             =   9480
      Width           =   4695
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   9480
      Width           =   255
   End
   Begin VB.CommandButton cmdGetPS 
      Caption         =   "Get PS detail"
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2145
   End
   Begin VB.ComboBox cmbMan 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1440
      Width           =   1555
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
      Left            =   12720
      TabIndex        =   5
      ToolTipText     =   " Clear the selection"
      Top             =   3240
      Width           =   1560
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf13a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
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
      FormDesignHeight=   8175
      FormDesignWidth =   14355
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   11520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9763
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
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   9720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select ASN File"
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   23
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7200
      Picture         =   "PackPSf13a.frx":07AE
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   9840
      Picture         =   "PackPSf13a.frx":0B38
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS End Date"
      Height          =   255
      Index           =   11
      Left            =   600
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Start Date"
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ASN File Name"
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   9
      Top             =   9480
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Manifest"
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "PackPSf13a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim bCutOff As Byte
Dim bOnLoad As Byte
Dim bUnload As Boolean

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
Dim strPOPrefix As String

Dim strSoNum As String
Dim strSOStnme As String
Dim strSOStadr As String
Dim strSOVia As String
Dim strSOTerms As String

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

Dim strPartNumFld As String
Dim strPartCntFld As String
Dim strPAUnitFld As String
Dim strPartInfoFld As String

Private txtKeyPress As New EsiKeyBd


Private Sub ClearArrays(iSize As Integer)
    Erase arrValue
    ReDim arrValue(0 To 96, iSize)
End Sub

Private Sub cmdASN_Click()
   CreateANS
End Sub

Private Sub cmdCan_Click()
   'sLastPrefix = cmbPre
   Unload Me

End Sub


Private Sub GetOptions()
   On Error Resume Next
   txtFilePath.Text = GetSetting("Esi2000", "EsiSale", "ASNFileName", txtFilePath.Text)
End Sub

Private Sub SaveOptions()
   On Error Resume Next
   SaveSetting "Esi2000", "EsiSale", "ASNFileName", txtFilePath.Text
   
End Sub



Private Sub cmdHlp_Click()
    If cmdHlp Then
        MouseCursor (13)
        OpenHelpContext (2150)
        MouseCursor (0)
        cmdHlp = False
    End If

End Sub

Private Sub cmdGetPS_Click()
   Dim strWindows As String
   Dim strAccFileName As String
   Dim strpathFilename As String
   
   On Error GoTo DiaErr1
   FillGrid
   
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
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
   End If
End Sub

Private Sub CreateANS()

   Dim strStartDate As String
   Dim strEndDate As String
   Dim strTmplFile As String
   Dim strManifest As String
   Dim strDestASNFile As String
   
   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
   strManifest = ""
   If (cmbMan <> "ALL") Then strManifest = cmbMan
   
   
   strTmplFile = App.Path & "\ASNUploadTemplateNew.xls"
   strDestASNFile = txtFilePath.Text

   If (strDestASNFile <> "") Then
      GenerateASNManifest strTmplFile, strDestASNFile
   Else
      MsgBox ("Please Select a FileName to Save.")
   End If
   
End Sub



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

Private Sub cmdSearch_Click()
   fileDlg.Filter = "Excel File (*.xls) | *.xls"
   fileDlg.ShowOpen
   If fileDlg.filename = "" Then
       txtFilePath.Text = ""
   Else
       txtFilePath.Text = fileDlg.filename
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
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      Dim strASN As String
      
      txtStartDte = Format(ES_SYSDATE, "mm/dd/yy")
      txtEndDte = Format(ES_SYSDATE, "mm/dd/yy")
      
      FillManifestNum
      strASN = GetLastASN
      If (strASN <> "") Then txtASN = CStr(Val(strASN) + 1) Else txtASN = ""
      
      bOnLoad = 0
   End If
   MouseCursor (0)

End Sub

Private Sub FillManifestNum()
         
   Dim strStartDate As String
   Dim strEndDate As String
   
   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
         
   sSql = "SELECT DISTINCT PSSHIPNO From PshdTable " _
         & " WHERE PshdTable.PSDATE BETWEEN '" & strStartDate _
         & "' AND '" & strEndDate & "' AND PSSHIPNO <> 0"
   
   LoadComboBox cmbMan, -1
   AddComboStr cmbMan.hWnd, "" & Trim("ALL")
   cmbMan = "ALL"
   
End Sub

Private Function GetLastASN() As String
   
   Dim RdoASN As ADODB.Recordset
   Dim strASN As String
   strASN = ""
   
   sSql = "SELECT DISTINCT LASTASNNUM, POLETTERREF FROM ASNInfoTable " _
            & " WHERE BOEINGPART = 1" _

   bSqlRows = clsADOCon.GetDataSet(sSql, RdoASN, ES_FORWARD)
   If bSqlRows Then
      With RdoASN
         lblLastASN = "" & Trim(!LASTASNNUM)
         'strPOPrefix = "" & Trim(!POLETTERREF)
         strASN = "" & Trim(!LASTASNNUM)
         
         ClearResultSet RdoASN
      End With
      
   End If
   GetLastASN = strASN
End Function


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
   SaveOptions
End Sub

Private Sub Form_Load()
    FormLoad Me, ES_DONTLIST
    
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
      .Text = "Apply"
      .Col = 1
      .Text = "PackSlip"
      .Col = 2
      .Text = "Manifest Num"
      .Col = 3
      .Text = "PO Number"
      .Col = 4
      .Text = "PO Date"
      .Col = 5
      .Text = "Ship Date"
      .Col = 6
      .Text = "Qty"
      .Col = 7
      .Text = "Boxes"
      .Col = 8
      .Text = "Weight"
      .Col = 9
      .Text = "Via"
      .Col = 10
      .Text = "POItem"
      .Col = 11
      .Text = "PartNumber"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1200
      .ColWidth(2) = 1200
      .ColWidth(3) = 1500
      .ColWidth(4) = 1000
      .ColWidth(5) = 1000
      .ColWidth(6) = 750
      .ColWidth(7) = 750
      .ColWidth(8) = 750
      .ColWidth(9) = 1200
      .ColWidth(10) = 750
      .ColWidth(11) = 1500
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
   GetOptions
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub


Function GenerateASNManifest(strTmplFile As String, strDestASNFile As String) As Integer
   
   Dim rdoPS As ADODB.Recordset
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
   
   Dim strSuppID As String
   Dim strBuyerID As String
   
   Dim strPONumber As String
   Dim strPSNum As String
   Dim strQty As String
   Dim strBoxes As String
   Dim strShipNo As String
   Dim strASNNum As String
   Dim strPSVia As String
   Dim strPurpose  As String
   Dim strWeight As String
   Dim strPreShipNo As String
   Dim strWtUOM As String
   
   Dim lASNNum As Integer
   Dim strLineItemID As String
   Dim strPOItem As String
   Dim strUOM As String
   Dim strPSdate As String
   Dim strShipDate As String
   Dim strShpNoPrefix As String
   Dim strShpID As String
   Dim strAction As String
   Dim strCarrier As String
   
   Dim iList As Integer
   Dim iExcelRow As Integer
   
   lASNNum = Val(txtASN.Text)
   
   Dim bRet As Boolean
   bRet = GetASNManifest(strSuppID, strBuyerID, strPOPrefix, strShpNoPrefix, strShpID)
   
   If (bRet = False) Then
      Exit Function
   End If
   
   
   'strSuppID = "0d5e8f78-78e1-1000-8e0d-0a1c0e080001"
   'strBuyerID = "a1d8e6d8-7802-1000-bfb4-ac16042a0001"
   'strPOPrefix = "VMS"
   strPurpose = "Original"
   strLineItemID = "1"
   strWtUOM = "Pounds"
   strUOM = "EACH"
   strAction = "InsertOrUpdate"
   strCarrier = "BOEING LICENSED TRANSPORTATION (BLT)"
   
   'strPOItem = "1"
   
   Err.Clear
   
   Dim xlApp As Excel.Application
   Dim ws As Worksheet
   
   Set xlApp = Nothing
'   Set xlApp = GetObject(, "Excel.Application")
   'Otherwise instantiate a new instance.
   If xlApp Is Nothing Then Set xlApp = New Excel.Application
   xlApp.Workbooks.open (strTmplFile)
   Set ws = xlApp.Worksheets(1) 'Specify your worksheet name
   
   iExcelRow = 3
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 1
         strPSNum = Grd.Text
         strASNNum = strPOPrefix & Right(Trim(strPSNum), (Len(Trim(strPSNum)) - 2))
         
         Grd.Col = 2
         strShipNo = Grd.Text
         strPreShipNo = Trim(strShpNoPrefix) & PadZeroString(strShipNo, 6, "0")

         Grd.Col = 3
         strPONumber = Grd.Text
         Grd.Col = 4
         strPSdate = Grd.Text
         Grd.Col = 5
         strShipDate = Grd.Text
         Grd.Col = 6
         strQty = Grd.Text
         Grd.Col = 7
         strBoxes = Grd.Text
         Grd.Col = 8
         strWeight = Grd.Text
         Grd.Col = 9
         strPSVia = Grd.Text
         Grd.Col = 10
         strPOItem = Grd.Text
         strPOItem = PadZeroString(CStr(Val(strPOItem)), 4, "0")
         Grd.Col = 11
         strPartNum = Grd.Text
      
         'ASN Number
         ws.Cells(iExcelRow, 1).Value = strASNNum
         'Buyer MPID
         ws.Cells(iExcelRow, 2).Value = strBuyerID
         'Supplier ID
         ws.Cells(iExcelRow, 3).Value = strSuppID
         'Supplier Code
         ws.Cells(iExcelRow, 4).Value = strShpID
         'Actual Ship Date
         ws.Cells(iExcelRow, 5).Value = strPSdate
         'Estimate Arrival
         ws.Cells(iExcelRow, 6).Value = strShipDate
         'Carrier ID
         ws.Cells(iExcelRow, 7).Value = strCarrier 'strPSVia
         'Manifest number / Bill Landing Tracking
         ws.Cells(iExcelRow, 9).Value = strPreShipNo
         'Pack slip
         strPSNum = Replace(strPSNum, "PS", "")    'PS001234
         ws.Cells(iExcelRow, 10).Value = strPSNum
         'Purpose
         ws.Cells(iExcelRow, 12).Value = strPurpose
         'Weight UOM
         ws.Cells(iExcelRow, 13).Value = strWtUOM
         'Gross Weight
         ws.Cells(iExcelRow, 14).Value = strWeight
         
         'Boxes / Total Package
         ws.Cells(iExcelRow, 16).Value = strBoxes
         
         'ASN Line ID
         ws.Cells(iExcelRow, 17).Value = PadZeroString(CStr(Val(strLineItemID)), 4, "0") 'strLineItemID
         
         ' Action
         ws.Cells(iExcelRow, 18).Value = strAction
         
         'TODO: Buyer Partnumber
         ws.Cells(iExcelRow, 19).Value = strPartNum
         
         'Ship Qty
         ws.Cells(iExcelRow, 20).Value = strQty
         'UOM
         ws.Cells(iExcelRow, 21).Value = strUOM
         'PO Item ID
         ws.Cells(iExcelRow, 24).Value = strPONumber  '// 23
         
         'Purchase Item
         ws.Cells(iExcelRow, 25).Value = strPOItem    '// 24
         
         iExcelRow = iExcelRow + 1
         
      End If
   Next
   MouseCursor ccArrow
         

   xlApp.Workbooks.Item(1).SaveAs (strDestASNFile)
   xlApp.Workbooks.Close

   If (Err.Number = 0) Then
      MsgBox "Generated ASN file to Export.", vbExclamation, Caption
   End If
   
   MouseCursor ccArrow
   Set rdoPS = Nothing
   Exit Function
   
DiaErr1:
   xlApp.Workbooks.Close
   sProcName = "GenerateASNFile"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


'Function GenerateASNManifest(strTmplFile As String, strDestASNFile As String) As Integer
'
'   Dim rdoPS As ADODB.Recordset
'
'   MouseCursor ccHourglass
'   On Error GoTo DiaErr1
'
'   Dim strSuppID As String
'   Dim strBuyerID As String
'
'   Dim strPONumber As String
'   Dim strPSNum As String
'   Dim strQty As String
'   Dim strBoxes As String
'   Dim strShipNo As String
'   Dim strASNNum As String
'   Dim strPSVia As String
'   Dim strPurpose  As String
'   Dim strWeight As String
'   Dim strPreShipNo As String
'
'   Dim lASNNum As Integer
'   Dim strLineItemID As String
'   Dim strPOItem As String
'   Dim strUOM As String
'   Dim strPSdate As String
'   Dim strShipDate As String
'   Dim strShpNoPrefix As String
'
'   Dim iList As Integer
'   Dim iExcelRow As Integer
'
'   lASNNum = Val(txtASN.Text)
'
'   Dim bRet As Boolean
'   bRet = GetASNManifest(strSuppID, strBuyerID, strPOPrefix, strShpNoPrefix)
'
'   If (bRet = False) Then
'      Exit Function
'   End If
'
'
'   'strSuppID = "0d5e8f78-78e1-1000-8e0d-0a1c0e080001"
'   'strBuyerID = "a1d8e6d8-7802-1000-bfb4-ac16042a0001"
'   'strPOPrefix = "VMS"
'   strPurpose = "Original"
'   strLineItemID = "1"
'   strUOM = "LBS"
'   'strPOItem = "1"
'
'   Err.Clear
'
'   Dim xlApp As Excel.Application
'   Dim ws As Worksheet
'
'   Set xlApp = Nothing
''   Set xlApp = GetObject(, "Excel.Application")
'   'Otherwise instantiate a new instance.
'   If xlApp Is Nothing Then Set xlApp = New Excel.Application
'   xlApp.Workbooks.open (strTmplFile)
'   Set ws = xlApp.Worksheets(1) 'Specify your worksheet name
'
'   iExcelRow = 3
'   For iList = 1 To Grd.Rows - 1
'      Grd.Col = 0
'      Grd.Row = iList
'
'      ' Only if the part is checked
'      If Grd.CellPicture = Chkyes.Picture Then
'
'         Grd.Col = 1
'         strPSNum = Grd.Text
'         strASNNum = strPOPrefix & Right(Trim(strPSNum), (Len(Trim(strPSNum)) - 2))
'
'         Grd.Col = 2
'         strShipNo = Grd.Text
'         strPreShipNo = Trim(strShpNoPrefix) & PadZeroString(strShipNo, 6, "0")
'
'         Grd.Col = 3
'         strPONumber = Grd.Text
'         Grd.Col = 4
'         strPSdate = Grd.Text
'         Grd.Col = 5
'         strShipDate = Grd.Text
'         Grd.Col = 6
'         strQty = Grd.Text
'         Grd.Col = 7
'         strBoxes = Grd.Text
'         Grd.Col = 8
'         strWeight = Grd.Text
'         Grd.Col = 9
'         strPSVia = Grd.Text
'         Grd.Col = 10
'         strPOItem = Grd.Text
'         strPOItem = PadZeroString(CStr(Val(strPOItem)), 4, "0")
'
'         'ASN Number
'         ws.Cells(iExcelRow, 1).Value = strASNNum
'         'Buyer MPID
'         ws.Cells(iExcelRow, 2).Value = strBuyerID
'         'Supplier ID
'         ws.Cells(iExcelRow, 3).Value = strSuppID
'         'Purpose
'         ws.Cells(iExcelRow, 4).Value = strPurpose
'         'Pack slip
'         strPSNum = Replace(strPSNum, "PS", "")    'PS001234
'         ws.Cells(iExcelRow, 5).Value = strPSNum
'         'Manifest number
'         ws.Cells(iExcelRow, 6).Value = strPreShipNo
'         'Estimate Dep
'         ws.Cells(iExcelRow, 7).Value = strPSdate
'         'Estimate Arrival
'         ws.Cells(iExcelRow, 8).Value = strShipDate
'         'Carrier ID
'         ws.Cells(iExcelRow, 9).Value = strPSVia
'         'Boxes
'         ws.Cells(iExcelRow, 10).Value = strBoxes
'         'Gross Weight
'         ws.Cells(iExcelRow, 11).Value = strWeight
'         'Gross Weight
'         ws.Cells(iExcelRow, 12).Value = strUOM
'
'         'Purchase Order num
'         ws.Cells(iExcelRow, 14).Value = strLineItemID
'
'         'PO Item ID
'         ws.Cells(iExcelRow, 16).Value = strPONumber
'
'
'         'Purchase Item
'         ws.Cells(iExcelRow, 17).Value = strPOItem
'         'Purchase Order num
'         ws.Cells(iExcelRow, 19).Value = strQty
'
'         iExcelRow = iExcelRow + 1
'
'      End If
'   Next
'   MouseCursor ccArrow
'
'
'   xlApp.Workbooks.Item(1).SaveAs (strDestASNFile)
'   xlApp.Workbooks.Close
'
'   If (Err.Number = 0) Then
'      MsgBox "Generated ASN file to Export.", vbExclamation, Caption
'   End If
'
'   MouseCursor ccArrow
'   Set rdoPS = Nothing
'   Exit Function
'
'DiaErr1:
'   xlApp.Workbooks.Close
'   sProcName = "GenerateASNFile"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Function
'
Private Function GetASNManifest(ByRef strSuppID As String, strBuyerID As String, _
         ByRef strPOPrefix As String, ByRef strShpNoPrefix As String, _
         ByRef strShpID As String) As Boolean
         
   Dim rdoMan As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT SUPPLIER_ID, BUYER_ID, PO_PREFIX, SHIPPING_PREFIX, SHIPPING_ID FROM ASNMfestTable"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoMan, ES_STATIC)
   
   If bSqlRows Then
      With rdoMan
         strSuppID = Trim(!SUPPLIER_ID)
         strBuyerID = Trim(!BUYER_ID)
         strPOPrefix = Trim(!PO_PREFIX)
         strShpNoPrefix = Trim(!SHIPPING_PREFIX)
         strShpID = Trim(!SHIPPING_ID)
         
      End With
      GetASNManifest = True
   Else
      MsgBox "Please Setup Manifest Details", vbCritical
      strSuppID = ""
      strBuyerID = ""
      strPOPrefix = ""
      strShpNoPrefix = ""
      strShpID = ""
      GetASNManifest = False
   End If
   MouseCursor ccArrow
   Set rdoMan = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GetASNManifest"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function


Function FillGrid() As Integer
   
   Dim strSenderCode As String
   Dim rdoPS As ADODB.Recordset
   
   MouseCursor ccHourglass
   Grd.Rows = 1
   On Error GoTo DiaErr1
       
       
   Dim strMan, strPONumber As String
   Dim strPSNum, strWeight As String
   Dim strStartDate, strEndDate, strQty As String
   Dim strBoxes As String
   Dim strLoadNo, strPSVia As String
   Dim strPSdate, strShipDate As String
   Dim strPOItem As String
   
   Dim bPartFound, bIncRow As Boolean
   Dim iItem As Integer

   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
   strMan = cmbMan.Text
   
   If (Trim(strMan) = "ALL") Then
      strMan = ""
   End If
   
          
   sSql = "SELECT PSNUMBER, PSSHIPNO, PSBOXES,PSGROSSLBS,  PSVIA, SOPO,PIQTY ,PSDATE, " _
         & " PSSHIPPEDDATE,ITCUSTITEMNO,Partnum  From PshdTable, psitTable, sohdTable, SoitTable, PartTable " _
         & " WHERE PshdTable.PSDATE BETWEEN '" & strStartDate & "' AND '" & strEndDate & "' " _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND PSSHIPNO <> 0 " _
          & " AND PSSHIPNO LIKE '" & strMan & "%'" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PsitTable.PIPART = PARTREF" _
          & " AND PshdTable.PSCUST IN " _
          & " (SELECT DISTINCT a.CUREF " _
          & "     FROM ASNInfoTable a, custtable b WHERE " _
          & "        A.CUREF = b.CUREF AND BOEINGPART = 1)"
      '& " GROUP BY PSNUMBER, PSSHIPNO, PSBOXES,PSGROSSLBS,  PSVIA, SOPO,PSDATE," _
            '& " PSSHIPPEDDATE"
          '& " ORDER BY PSSHIPNO"
          

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_STATIC)
   
   If bSqlRows Then
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strLoadNo = Trim(!PSSHIPNO)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strBoxes = Trim(!PSBOXES)
         strPSVia = Trim(!PSVIA)
         strWeight = Trim(!PSGROSSLBS)
         strPSdate = Format(!PSDATE, "mm/dd/yy")
         strShipDate = Format(!PSSHIPPEDDATE, "mm/dd/yy")
         strPOItem = PadZeroString(CStr(!ITCUSTITEMNO), 4, "0")
         strPartNum = Trim(!PartNum)
         
         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         bIncRow = False
         iItem = 1
         
         Grd.Col = 0
         Set Grd.CellPicture = Chkno.Picture
         Grd.Col = 1
         Grd.Text = Trim(strPSNum)
         Grd.Col = 2
         Grd.Text = Trim(strLoadNo)
         Grd.Col = 3
         Grd.Text = Trim(strPONumber)
         Grd.Col = 4
         Grd.Text = Trim(strPSdate)
         Grd.Col = 5
         Grd.Text = Trim(strShipDate)
         Grd.Col = 6
         Grd.Text = Trim(strQty)
         Grd.Col = 7
         Grd.Text = Trim(strBoxes)
         Grd.Col = 8
         Grd.Text = Trim(strWeight)
         Grd.Col = 9
         Grd.Text = Trim(strPSVia)
         Grd.Col = 10
         Grd.Text = Trim(strPOItem)
         Grd.Col = 11
         Grd.Text = Trim(strPartNum)
         .MoveNext
      Wend
      .Close
      End With
   End If

   MouseCursor ccArrow
   Set rdoPS = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub Form_Resize()
    Refresh

End Sub
Private Sub Form_Unload(Cancel As Integer)
    'FormUnload
    Set PackPSf13a = Nothing
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeySpace Then
      Grd.Col = 0
      If Grd.Row >= 1 Then
         If Grd.Row = 0 Then Grd.Row = 1
         If Grd.CellPicture = Chkyes.Picture Then
            Set Grd.CellPicture = Chkno.Picture
         Else
            Set Grd.CellPicture = Chkyes.Picture
         End If
      End If
    End If
   

End Sub

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


Private Sub Grd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Grd.Col = 0
   If Grd.Row >= 1 Then
      If Grd.Row = 0 Then Grd.Row = 1
      If Grd.CellPicture = Chkyes.Picture Then
         Set Grd.CellPicture = Chkno.Picture
      Else
         Set Grd.CellPicture = Chkyes.Picture
      End If
   End If
End Sub


Private Function GetBuyerInfo(ByVal strCust As String, _
                  ByRef strBusPartner As String, ByRef strBusDetail As String, _
                  ByRef strBuyerCode As String)

   On Error GoTo modErr1
   Dim RdoBuy As ADODB.Recordset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS, BUYERCODE FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strBusPartner = "" & Trim(!SHPREF)
            strBusDetail = "" & Trim(!SHPDETAIL)
            strBuyerCode = "" & Trim(!BUYERCODE)
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function


Private Function GetShipInfo(ByVal strCust As String, ByRef strFrmVnd As String, _
                  ByRef strFrmVndID As String, ByRef strFromAddrs As String, _
                  ByRef strToVnd As String, ByRef strToVndID As String, ByRef strToAddrs As String)

   On Error GoTo modErr1
   Dim RdoBuy As ADODB.Recordset
   If Trim(strCust) <> "" Then
      
      sSql = "SELECT SHPTOIDCODE, SHPTOCODEQUAL, SHPFRMIDCODE, " _
               & " SHPFRMCODEQUAL , SHPREF, SHPDETAIL, " _
               & " SHPADDRS FROM ASNInfoTable " _
               & "WHERE CUREF = '" & strCust & "'"

      bSqlRows = clsADOCon.GetDataSet(sSql, RdoBuy, ES_FORWARD)
      If bSqlRows Then
         With RdoBuy
            strFrmVnd = "" & Trim(!SHPFRMIDCODE)
            strFrmVndID = "" & Trim(!SHPFRMCODEQUAL)
            strFromAddrs = "U.S. CASTINGS LLC."
            strToVnd = "" & Trim(!SHPTOIDCODE)
            strToVndID = "" & Trim(!SHPTOCODEQUAL)
            strToAddrs = "" & Trim(!SHPADDRS)
            
            ClearResultSet RdoBuy
            
         End With
      End If
   End If
   Set RdoBuy = Nothing
   Exit Function
   
modErr1:
   sProcName = "GetBuyerInfo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors MdiSect.ActiveForm
   
End Function


Private Sub txtEndDte_LostFocus()
   FillManifestNum
End Sub

Private Sub txtStartDte_LostFocus()
   FillManifestNum
End Sub

Private Sub txtStartDte_DropDown()
   ShowCalendar Me
End Sub

Private Function strConverDate(strDate As String, ByRef strDateConv As String)
   strDateConv = Format(CDate(strDate), "yymmdd")
End Function

Private Function TotalPsSelected(strContainer As String) As Integer
   Dim iList As Integer
   Dim iTotCnt As Integer
   Dim strTCont As String
   
   iTotCnt = 0
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
         Grd.Col = 2
         strTCont = Grd.Text
         If (strTCont = strContainer) Then
            iTotCnt = iTotCnt + 1
         End If
      End If
   Next

   TotalPsSelected = iTotCnt
End Function

Private Function PadZeroString(strInput As String, iLen As Variant, strPad As String) As String
   
   If (iLen > 0) Then
      If (strPad = "0") Then
         strInput = Format(strInput, String(iLen, "0"))
      ElseIf (strPad = "@") Then
         strInput = Format(strInput, String(iLen, "@"))
      End If
   End If

   PadZeroString = strInput
   
End Function
   

Private Function CheckSelected(strPSNum As String, _
         strPONumber As String, strPartNum As String) As Boolean

   Dim bChecked As Boolean
   Dim strGrdPS As String
   Dim strGrdPO As String
   Dim strGrdPN As String
   Dim iList As Integer
   
   bChecked = False
   
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      If Grd.CellPicture = Chkyes.Picture Then
         
         Grd.Col = 1
         strGrdPS = Trim(Grd.Text)
         Grd.Col = 4
         strGrdPO = Trim(Grd.Text)
         Grd.Col = 5
         strGrdPN = Trim(Grd.Text)
               
         If ((strGrdPS = strPSNum) And (strGrdPO = strPONumber) And _
            (strGrdPN = strPartNum)) Then
               
            bChecked = True
            Exit For
         End If
      End If
   Next
   
   CheckSelected = bChecked
End Function


Private Sub txtEndDte_DropDown()
   ShowCalendar Me
End Sub
