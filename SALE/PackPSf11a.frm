VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form PackPSf11a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate Advance Shipping Manifest"
   ClientHeight    =   8445
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   15150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton optDis 
      Height          =   330
      Left            =   11520
      Picture         =   "PackPSf11a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Display The Report"
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton optPrn 
      Height          =   330
      Left            =   12120
      Picture         =   "PackPSf11a.frx":017E
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Print The Report"
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   490
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
      Height          =   600
      Left            =   13080
      TabIndex        =   29
      ToolTipText     =   " Create PS from Sales Order"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.ComboBox txtEndDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   26
      Tag             =   "4"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox txtStartDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   25
      Tag             =   "4"
      Top             =   840
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
      Left            =   12960
      TabIndex        =   24
      ToolTipText     =   " Select All"
      Top             =   3240
      Width           =   1920
   End
   Begin VB.TextBox txtEdiFilePath 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Select import"
      Top             =   9480
      Width           =   4695
   End
   Begin VB.CommandButton cmdOpenDia 
      Caption         =   "..."
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Browse XML file or Text file"
      Top             =   9480
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   5880
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdASNInfo 
         Caption         =   "Add Manifest Number"
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
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   " Add ASN number to PS"
         Top             =   1560
         Width           =   2160
      End
      Begin VB.TextBox txtMan 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Tag             =   "3"
         ToolTipText     =   "Select XML file to import"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblLastMan 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         ToolTipText     =   "Last Sales Order Entered"
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Manifest Number"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label z1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manifest Number"
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   960
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdGetPS 
      Caption         =   "Get PS detail"
      Height          =   360
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2145
   End
   Begin VB.ComboBox cmbCst 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Customer From List"
      Top             =   1800
      Width           =   1555
   End
   Begin VB.TextBox txtPsl 
      Height          =   285
      Left            =   9780
      MaxLength       =   8
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "New Pack Slip Number (6 char max)"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox OptSoXml 
      Caption         =   "FromXMLSO"
      Height          =   195
      Left            =   7080
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSORev 
      Caption         =   "Show Revise SO "
      Height          =   195
      Left            =   8280
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "The first PO will be created and Revise SO form is displayed"
      Top             =   120
      Width           =   1935
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
      Left            =   12960
      TabIndex        =   7
      ToolTipText     =   " Clear the selection"
      Top             =   4080
      Width           =   1920
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf11a.frx":0308
      Style           =   1  'Graphical
      TabIndex        =   10
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
      FormDesignHeight=   8445
      FormDesignWidth =   15150
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6360
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   11520
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1035
   End
   Begin MSComDlg.CommonDialog fileDlg 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open XML File for Import"
      Filter          =   "*.xml"
   End
   Begin MSFlexGridLib.MSFlexGrid Grd 
      Height          =   4935
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   12495
      _ExtentX        =   22040
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
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS End Date"
      Height          =   255
      Index           =   11
      Left            =   720
      TabIndex        =   28
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PS Start Date"
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ASN File Name"
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   23
      Top             =   9480
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer"
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblPrefix 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9480
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNotice 
      Caption         =   "Note: The Last Sales Order Number Has Changed"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Chkyes 
      Height          =   210
      Left            =   7680
      Picture         =   "PackPSf11a.frx":0AB6
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Chkno 
      Height          =   210
      Left            =   7680
      Picture         =   "PackPSf11a.frx":0E40
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "PackPSf11a"
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



Private Sub cmdASNInfo_Click()
   
   Dim strGrossWt As String
   Dim strCarrierNum As String
   Dim strPS As String
   
   Dim strCarton As String
   Dim strContainer As String
   Dim strLoadNum As String
   Dim strShipNum As String
   
   Dim strPONumber As String
   Dim strPartNum As String
   Dim strQty As String
   Dim strPullNum As String
   Dim strBinNum As String
   Dim strCust As String
   
   If (Trim(txtMan.Text) = "") Then
      MsgBox "Manifest number is empty.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   strCust = CStr(cmbCst)
   strContainer = PadZeroString(txtMan.Text, 6, "0")
   strShipNum = CStr(Val(strContainer))
   
   clsADOCon.BeginTrans
   
   Dim iList As Integer
   Dim iTotCnt As Integer
   iTotCnt = 0
   For iList = 1 To Grd.Rows - 1
      Grd.Col = 0
      Grd.Row = iList
      
      ' Only if the part is checked
      If Grd.CellPicture = Chkyes.Picture Then
                  
         Grd.Col = 1
         strPS = Trim(Grd.Text)
         
         sSql = "UPDATE PshdTable SET PSCONTAINER = '" & strContainer & "'," _
                  & " PSSHIPNO = " & strShipNum _
            & " FROM PshdTable " _
            & " WHERE PSNUMBER = '" & strPS & "'" _
                  & " AND PshdTable.PSINVOICE = 0"
                  '& " AND PshdTable.PSPRINTED IS NULL"
                  '& " AND PshdTable.PSSHIPPRINT = 0" _

         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
         sSql = "UPDATE ASNInfoTable SET LASTMANFSTNUM = '" & strShipNum & "' " _
                  & " WHERE BOEINGPART = 1"
                  
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
         Grd.Col = 2
         Grd.Text = Trim(strContainer)
         
         iTotCnt = iTotCnt + 1
      End If
   Next
   
   If clsADOCon.RowsAffected > 0 Then
      'clsADOCon.RollbackTrans
            
      MsgBox "Updated Manifest Number to the selected PackSlips.", _
         vbInformation, Caption
      
      Dim strMaxASN As String
      strMaxASN = GetLastManfest(CStr(cmbCst))
      
      If (Trim(strMaxASN) <> "") Then
         txtMan = Val(strMaxASN) + 1
      Else
         txtMan = ""
         
      End If
   
   End If
   clsADOCon.CommitTrans

End Sub

Private Sub cmdCan_Click()
   'sLastPrefix = cmbPre
   Unload Me

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

Private Sub CreateANS()

   Dim strStartDate As String
   Dim strEndDate As String
   Dim strASN As String
   
   Dim nFileNum As Integer, lLineCount As Long
   Dim strBlank As String
   
   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
   strASN = txtMan.Text

   GenerateASNManifest strStartDate, strEndDate, strASN
   PrintReport (Val(strASN) - 1)
   
End Sub

Private Sub PrintReport(ASN As Long)
   Dim sBeg As String
   Dim sEnd As String
   Dim sCust As String
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aRptPara As New Collection
   Dim aRptParaType As New Collection
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.Init
   sCustomReport = GetCustomReport("slesh05a")
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   
   sSql = "{EsReportASNManifest.PSSHIPNO} = " & Val(ASN)
   
   
   cCRViewer.SetReportSelectionFormula sSql
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.ShowGroupTree False
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
 
   cCRViewer.ClearFieldCollection aRptPara
   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue

   FillGrid
   
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "printreport"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCst_Click()
   
   Dim strMaxManfest  As String
   
   FindCustomer Me, cmbCst, False
   
   If bOnLoad = 0 Then
      ' Filter the records if selected.
      strMaxManfest = GetLastManfest(CStr(cmbCst))
      
      If (Trim(strMaxManfest) <> "") Then
         txtMan = Val(strMaxManfest) + 1
      Else
         txtMan = ""
         
      End If
   End If
   
End Sub

Private Sub cmbCst_LostFocus()
'   cmbCst = CheckLen(cmbCst, 10)
'   FindCustomer Me, cmbCst, False
'   lblNotice.Visible = False
   
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
   
      'FillCustomers
      sSql = "SELECT DISTINCT a.CUREF FROM ASNInfoTable a, custtable b WHERE " _
               & " A.CUREF = b.CUREF AND BOEINGPART = 1"
               
      LoadComboBox cmbCst, -1
      AddComboStr cmbCst.hWnd, "" & Trim("ALL")
      cmbCst = "ALL"
      txtNme = "*** All Customer selected ***"
      
      'If cUR.CurrentCustomer <> "" Then cmbCst = cUR.CurrentCustomer
      FindCustomer Me, cmbCst, False
   
      Dim ps As New ClassPackSlip
      Dim strMaxMan As String
      lblPrefix = ps.GetPackSlipPrefix
      txtPsl = ""
      txtPsl.MaxLength = 8 - Len(lblPrefix)
      strMaxMan = GetLastManfest(CStr(cmbCst))
      
      If (Trim(strMaxMan) <> "") Then
         txtMan = Val(strMaxMan) + 1
      End If
      txtStartDte = Format(ES_SYSDATE, "mm/dd/yy")
      txtEndDte = Format(ES_SYSDATE, "mm/dd/yy")

      'GetPackslip True
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
   
   SaveSetting "Esi2000", "EsiSale", "LastPrefix", sLastPrefix
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
      .ColAlignment(7) = 1
      .ColAlignment(8) = 1
      
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
      .Text = "PartNumber"
      .Col = 5
      .Text = "Qty"
      .Col = 6
      .Text = "Boxes"
      .Col = 7
      .Text = "Via"
      .Col = 8
      .Text = "ShipTo"
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1100
      .ColWidth(2) = 1100
      .ColWidth(3) = 1100
      .ColWidth(4) = 2500
      .ColWidth(5) = 750
      .ColWidth(6) = 750
      .ColWidth(7) = 1200
      .ColWidth(8) = 3500
      
      
      .ScrollBars = flexScrollBarBoth
      .AllowUserResizing = flexResizeColumns
      
   End With
    
   Call WheelHook(Me.hWnd)
   bOnLoad = 1

End Sub

Private Function GetASNManifestInfo(ByRef strShpNoPrefix As String, ByRef strShpID As String) As Boolean
         
   Dim rdoMan As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   sSql = "SELECT SHIPPING_PREFIX, SHIPPING_ID FROM ASNMfestTable"
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoMan, ES_STATIC)
   
   If bSqlRows Then
      With rdoMan
         strShpNoPrefix = Trim(!SHIPPING_PREFIX)
         strShpID = Trim(!SHIPPING_ID)
      End With
      GetASNManifestInfo = True
   Else
      MsgBox "Please setup Manifest Details", vbCritical
      strShpNoPrefix = ""
      strShpID = ""
      GetASNManifestInfo = False
   End If
   MouseCursor ccArrow
   Set rdoMan = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "fillgrid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Function GenerateASNManifest(strStartDate As String, strEndDate As String, strASN As String) As Integer
   
   Dim rdoPS As ADODB.Recordset
   
   MouseCursor ccHourglass
   On Error GoTo DiaErr1
       
   Dim strPONumber As String
   Dim strCust As String
   Dim strSTName As String
   Dim strSTAdr As String
   Dim strShpID As String
   
   Dim strPSNum As String
   Dim strPreShipNo As String
   Dim strGrossLbs As String
   
   Dim strQty As String
   Dim strCarton As String
   Dim strBoxes As String
   Dim bPartFound As Boolean
   Dim strShipNo As String
   Dim iItem As Integer
   Dim bSelected As Boolean
   
   Dim strShpNoPrefix As String
   
   bSelected = GetASNManifestInfo(strShpNoPrefix, strShpID)
   
   If (bSelected = False) Then
      Exit Function
   End If
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCUST, PSCONTAINER, PSSHIPNO, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
            & "ISNULL(PSBOXES, '0') PSBOXES,ISNULL(PSGROSSLBS, '0.00') PSGROSSLBS,ISNULL(PSCARRIERNUM, '') PSCARRIERNUM, " _
            & " PSLOADNO, PSVIA, PSSTNAME, PSSTADR,SOPO,PIQTY , PIPART, PARTNUM, " _
            & " ISNULL(PULLNUM, '') PULLNUM, ISNULL(BINNUM, '') BINNUM " _
         & " From PshdTable, psitTable, sohdTable, SoitTable, Parttable " _
         & " WHERE PshdTable.PSDATE BETWEEN '" & strStartDate & "' AND '" & strEndDate & "' " _
         & " AND PSSHIPNO = " & CStr(Val(strASN) - 1) _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PARTREF = PIPART" _
          & " AND PshdTable.PSCUST IN " _
          & " (SELECT DISTINCT a.CUREF " _
          & "     FROM ASNInfoTable a, custtable b WHERE " _
          & "        A.CUREF = b.CUREF AND BOEINGPART = 1)" _
          & " ORDER BY PSSHIPNO"
          ' PshdTable.PSCUST LIKE '" & strCust & "%' AND
          ' MM& " AND PshdTable.PSINVOICE = 0 " _

   Debug.Print sSql
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_STATIC)
   
   ' MM strShpID = "567904"
   bSelected = False
   If bSqlRows Then
      
      sSql = "DELETE FROM EsReportASNManifest WHERE PSSHIPNO = " & CStr(Val(strASN) - 1)
      clsADOCon.ExecuteSql sSql ' rdExecDirect
      
      With rdoPS
      While Not .EOF
         
         strShipNo = Trim(!PSSHIPNO)
         strPreShipNo = strShpNoPrefix & PadZeroString(strShipNo, 6, "0")
         strPSNum = Right(Trim(!PsNumber), (Len(Trim(!PsNumber)) - 2))
         
         strCust = Trim(!PSCUST)
         strGrossLbs = Trim(!PSGROSSLBS)
         strBoxes = Trim(!PSBOXES)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strSTName = Trim(!PSSTNAME)
         strSTAdr = Trim(!PSSTADR)
         
         
         sSql = "INSERT INTO EsReportASNManifest (PSSHIPNO,PSPRESHIPNO,PSCUST, SHPFRMIDCODE, " _
                  & "PSSTNAME, PSSTADR,SOPO, PSNUMBER, PSGROSSLBS, PSBOXES)" _
                  & " VALUES ('" & Val(strShipNo) & "','" & strPreShipNo & "','" _
                  & strCust & "','" & strShpID & "'," _
                  & "'" & strSTName & "','" & strSTAdr & "'," _
                  & "'" & strPONumber & "','" & CStr(Val(strPSNum)) & "'," _
                  & "'" & strGrossLbs & "','" & CStr(Val(strBoxes)) & "')"
         
         Debug.Print sSql
         
         clsADOCon.ExecuteSql sSql ' rdExecDirect
         
         .MoveNext
      Wend
      .Close
      End With
   
      If (bSelected = True) Then
         MsgBox "ASN File created.", vbExclamation, Caption
      End If
      
   End If

   MouseCursor ccArrow
   Set rdoPS = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "GenerateASNFile"
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
       
       
   Dim strCust, strPONumber, strPartNum As String
   Dim strPSNum As String
   Dim strStartDate, strEndDate, strQty As String
   Dim strCarton, strContainer As String
   Dim strBox As String
   Dim strLoadNo, strPSVia As String
   Dim bPartFound, bIncRow As Boolean
   Dim strSOAdr, strBinNum As String
   Dim iItem As Integer

   strStartDate = txtStartDte.Text
   strEndDate = txtEndDte.Text
   strCust = cmbCst.Text
   
   If (Trim(strCust) = "ALL") Then
      strCust = ""
   End If
   
   sSql = "SELECT DISTINCT PSNUMBER, PSCONTAINER, PSNUMBER, ISNULL(PSCARTON, '') PSCARTON," _
            & " PSLOADNO, PSVIA, SOPO,PIQTY , PIPART, PSBOXES, SOSTNAME " _
         & " From PshdTable, psitTable, sohdTable, SoitTable " _
         & " WHERE PshdTable.PSDATE BETWEEN '" & strStartDate & "' AND '" & strEndDate & "' " _
          & " AND PshdTable.PSCUST LIKE '" & strCust & "%'" _
          & " AND PshdTable.PSINVOICE = 0 " _
          & " AND PshdTable.PSSHIPNO = 0 " _
          & " AND PSNUMBER = PIPACKSLIP" _
          & " AND SONUMBER = ITSO" _
          & " AND ITPSNUMBER = ITPSNUMBER" _
          & " AND SoitTable.ITSO = PsitTable.PISONUMBER" _
          & " AND SoitTable.ITNUMBER = PsitTable.PISOITEM" _
          & " AND SoitTable.ITREV = PsitTable.PISOREV" _
          & " AND PshdTable.PSCUST IN (SELECT DISTINCT a.CUREF " _
          & "                FROM ASNInfoTable a, custtable b WHERE " _
          & "                A.CUREF = b.CUREF AND BOEINGPART = 1)"
          
          '" & strCust & "%'
          ' MM & " AND PshdTable.PSINVOICE = 0"
          '& " AND PshdTable.PSPRINTED IS NULL"
          '& " AND PshdTable.PSSHIPPRINT = 0" _

   Debug.Print sSql
   
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_STATIC)
   
   If bSqlRows Then
      With rdoPS
      While Not .EOF
         
         strPSNum = Trim(!PsNumber)
         strContainer = Trim(!PSCONTAINER)
         strPONumber = Trim(!PsNumber)
         strPartNum = Trim(!PIPART)
         strCarton = Trim(!PSCARTON)
         strLoadNo = Trim(!PSLOADNO)
         strPSVia = Trim(!PSVIA)
         strPONumber = Trim(!SOPO)
         strQty = Trim(!PIQTY)
         strBox = Trim(!PSBOXES)
         strSOAdr = Trim(!SOSTNAME)
         'strPullNum = Trim(!PULLNUM)
         'strBinNum = Trim(!BINNUM)
         
         Grd.Rows = Grd.Rows + 1
         Grd.Row = Grd.Rows - 1
         bIncRow = False
         iItem = 1
         
         Grd.Col = 0
         Set Grd.CellPicture = Chkno.Picture
         Grd.Col = 1
         Grd.Text = Trim(strPSNum)
         Grd.Col = 2
         Grd.Text = Trim(strContainer)
         
         Grd.Col = 3
         Grd.Text = Trim(strPONumber)
         
         Grd.Col = 4
         Grd.Text = Trim(strPartNum)
         
         Grd.Col = 5
         Grd.Text = Trim(strQty)
         
         Grd.Col = 6
         Grd.Text = Trim(strBox)
         
         Grd.Col = 7
         Grd.Text = Trim(strPSVia)
         
         Grd.Col = 8
         Grd.Text = Trim(strSOAdr)
         
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
   If OptSoXml.Value = vbUnchecked Then FormUnload
    'FormUnload
    Set PackPSf11a = Nothing
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

Private Function GetLastManfest(strCst As String) As String
   
   Dim rdoPS As ADODB.Recordset
   Dim strMan As String
   strMan = ""
   
   sSql = "SELECT DISTINCT LASTMANFSTNUM FROM ASNInfoTable " _
            & " WHERE BOEINGPART = 1 AND LASTMANFSTNUM IS NOT NULL" _

   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_FORWARD)
   If bSqlRows Then
      With rdoPS
         lblLastMan = "" & Trim(!LASTMANFSTNUM)
         strMan = "" & Trim(!LASTMANFSTNUM)
         ClearResultSet rdoPS
      End With
      
      Dim bRet As Boolean
      
      ' validate and make sure that the Ship Number is not duplicate
      bRet = ValidateManfest(strMan)
      
      If (bRet = False) Then
         strMan = ""
         lblLastMan = ""
      End If
      
   End If
   
   
   GetLastManfest = strMan
End Function


Private Function ValidateManfest(strShipNo As String) As Boolean
   
   Dim rdoPS As ADODB.Recordset
   
   ' Get the ship VIA information
   sSql = "SELECT PSSHIPNO FROM PshdTable,ASNInfoTable" _
            & " WHERE  PSCUST = CUREF AND BOEINGPART = 1" _
            & " AND PSSHIPNO = " & (Val(strShipNo) + 1)

   bSqlRows = clsADOCon.GetDataSet(sSql, rdoPS, ES_FORWARD)
   If bSqlRows Then
      With rdoPS
         MsgBox "Manifest number " & strShipNo & " exist in the System.", vbCritical
         ClearResultSet rdoPS
         ValidateManfest = False
      End With
   Else
      ValidateManfest = True
   End If

End Function


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


Private Sub GetPackslip(bFillText As Boolean)
   Dim ps As New ClassPackSlip
   'lblLastMan = ps.GetLastPackSlipNumber

   If bFillText Then
      txtPsl = Right(ps.GetNextPackSlipNumber, txtPsl.MaxLength)
   End If
End Sub


Private Sub optDis_Click()
   CreateANS
End Sub

Private Sub optPrn_Click()
   CreateANS
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
