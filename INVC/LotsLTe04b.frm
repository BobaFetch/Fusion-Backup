VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTe04b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Transfer Lots"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTe04b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtSplit 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Customer "
      Top             =   3000
      Width           =   1555
   End
   Begin VB.ComboBox cmbCpart 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Customer Part Number (Saves Entries)"
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "0.000"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ComboBox cmbLoc 
      Height          =   288
      Left            =   5400
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Enter Or Select A Location "
      Top             =   2640
      Width           =   860
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "LotsLTe04b.frx":07AE
      DownPicture     =   "LotsLTe04b.frx":1120
      Height          =   350
      Left            =   5520
      Picture         =   "LotsLTe04b.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Standard Comments"
      Top             =   3840
      Width           =   350
   End
   Begin VB.TextBox txtLot 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "User Produced Lot Number Click To Set User Lot The Same"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox lblNumber 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "System Produced Lot Number Click To Set User Lot The Same"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Change The User Lot ID"
      Top             =   1080
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      Height          =   1035
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "9"
      ToolTipText     =   "Comments (2048)"
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton optDis 
      Height          =   350
      Left            =   6000
      Picture         =   "LotsLTe04b.frx":2094
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Print or View Detail"
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   6700
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5220
      FormDesignWidth =   6930
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3480
      TabIndex        =   30
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Part"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   29
      Top             =   3360
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   27
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Comments"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Remaining"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   21
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   20
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Comments"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "LotsLTe04b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'8/24/05 New
Option Explicit
Dim RdoCur As ADODB.Recordset
Dim bGoodLot As Byte
Dim bGoodPart As Byte
Dim bOnLoad As Byte

Dim cOldCost As Currency
Dim sOldLot As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Public Sub GetCalledLot()
   Dim bByte As Byte
   bByte = GetLotPart(cmbPrt)
   
End Sub

Public Sub FillCustomerParts()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT LOTCUSTPART FROM LohdTable ORDER BY LOTCUSTPART"
   LoadComboBox cmbCpart, -1
   Exit Sub
   
DiaErr1:
   sProcName = "fillcustomerpar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillTransferCustomers()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT CUREF,CUNICKNAME,CUALLOWTRANSFERS FROM CustTable WHERE " _
          & "CUALLOWTRANSFERS =1 ORDER BY CUREF"
   LoadComboBox cmbCst
   If cmbCst.ListCount = 0 Then
      MsgBox "Please Mark Transfer Customers In Sales/Customers.", _
         vbInformation, Caption
   Else
      GetTransferCustomer
   End If
   'sSql = "SELECT SHIPREF FROM CshpTable "
   sSql = "SELECT DISTINCT LOTLOCATION FROM LohdTable WHERE LOTLOCATION<>'' ORDER BY LOTLOCATION"
   LoadComboBox cmbLoc, -1
   'If cmbLoc.ListCount > 0 Then
   '   cmbLoc = cmbLoc.List(0)
   'Else
   '   MsgBox "Locations Are Setup In Administration/Sales/Ship To..", _
   '      vbInformation, Caption
   'End If
   Exit Sub
   
DiaErr1:
   sProcName = "filltranscust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetTransferCustomer()
   On Error Resume Next
   Dim rdoCst As ADODB.Recordset
   sSql = "SELECT CUNICKNAME,CUNAME FROM CustTable WHERE CUREF='" _
          & Compress(cmbCst) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoCst, ES_FORWARD)
   If bSqlRows Then
      cmbCst = "" & Trim(rdoCst!CUNICKNAME)
      txtNme = "" & Trim(rdoCst!CUNAME)
   Else
      txtNme = ""
   End If
   Set rdoCst = Nothing
End Sub



Private Sub cmbCpart_LostFocus()
   cmbCpart = CheckLen(cmbCpart, 30)
   cmbCpart = StrCase(cmbCpart, ES_FIRSTWORD)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTCUSTPART = Trim(cmbCpart)
         .Update
      End With
   End If
   
End Sub


Private Sub cmbCst_Click()
   GetTransferCustomer
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 10)
   GetTransferCustomer
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTCUST = Compress(cmbCst)
         .Update
      End With
   End If
   
End Sub


Private Sub cmbLoc_LostFocus()
   cmbLoc = CheckLen(cmbLoc, 4)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTLOCATION = cmbLoc
         .Update
      End With
   End If
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetLotPart(Compress(cmbPrt))
   cmdChg.Enabled = False
   
End Sub


Private Sub cmbPrt_LostFocus()
   bGoodPart = GetLotPart(Compress(cmbPrt))
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
End Sub








Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   LotsLTe01b.txtLot = txtLot
   LotsLTe01b.Show
   
End Sub

Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 3
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5501"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub




Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillTransferCustomers
      FillCustomerParts
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   LotsLTe04a.Show
   Set LotsLTe04b = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   lblNumber.BackColor = Me.BackColor
   
End Sub




Private Sub optDis_Click()
   Dim cCRViewer As EsCrystalRptViewer
   Dim sCustomReport As String
   Dim aFormulaValue As New Collection
   Dim aFormulaName As New Collection
   
   'Dim sDate As String
   'Dim sVendor As String
   MouseCursor 13
   On Error GoTo DiaErr1
   'SetMdiReportsize MdiSect
   
   aFormulaName.Add "CompanyName"
   aFormulaValue.Add CStr("'" & CStr(sFacility) & "'")
   sCustomReport = GetCustomReport("lotdetail")
   Set cCRViewer = New EsCrystalRptViewer
   cCRViewer.SetReportFileName sCustomReport, sReportPath
   cCRViewer.SetReportTitle = sCustomReport
   cCRViewer.SetReportFormulaFields aFormulaName, aFormulaValue
   cCRViewer.ShowGroupTree False
   
   sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
   cCRViewer.SetReportSelectionFormula (sSql)
   cCRViewer.CRViewerSize Me
   cCRViewer.SetDbTableConnection
   cCRViewer.OpenCrystalReportObject Me, aFormulaName
   

   cCRViewer.ClearFieldCollection aFormulaName
   cCRViewer.ClearFieldCollection aFormulaValue
   
'   MdiSect.Crw.SelectionFormula = sSql
'   MdiSect.Crw.Destination = crptToWindow
'   MdiSect.Crw.Action = 1
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub





Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTCOMMENTS = txtCmt
         .Update
      End With
   End If
   
End Sub








'Leave Public - Called from elsewhere

Public Function GetThisLot() As Byte
   On Error GoTo DiaErr1
   ManageBoxes 0
   cmdChg.Enabled = False
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTADATE,LOTREMAININGQTY," _
          & "LOTLOCATION,LOTCOMMENTS,LOTSPLITCOMMENT,LOTCUST,LOTCUSTPART," _
          & "LOINUMBER,LOIRECORD,LOITYPE FROM LohdTable,LoitTable WHERE " _
          & "(LOTNUMBER='" & Trim(lblNumber) & " ' AND LOTNUMBER=LOINUMBER " _
          & "AND LOIRECORD=1)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_KEYSET)
   If bSqlRows Then
      ManageBoxes 1
      With RdoCur
         lblNumber = "" & Trim(!lotNumber)
         txtLot = "" & Trim(!LOTUSERLOTID)
         txtCmt = "" & Trim(!LOTCOMMENTS)
         cmbLoc = "" & Trim(!LOTLOCATION)
         txtCst = Format(!LotUnitCost, ES_QuantityDataFormat)
         cOldCost = !LotUnitCost
         If Not IsNull(.Fields(6)) Then
            lblDate = "" & Format(!LotADate, "mm/dd/yy")
         Else
            lblDate = Format(GetServerDateTime, "mm/dd/yy")
         End If
         lblRem = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
         lblType = GetLotType(!LOITYPE)
         txtSplit = "" & Trim(!LOTSPLITCOMMENT)
         cmbCst = "" & Trim(!LOTCUST)
         cmbCpart = "" & Trim(!LOTCUSTPART)
         sOldLot = lblNumber
         GetTransferCustomer
      End With
      GetThisLot = 1
      On Error Resume Next
   Else
      ManageBoxes 0, 1
      GetThisLot = 0
      MsgBox "The Request Lot Was Not Found Or Is Not Available.", _
         vbInformation, Caption
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthislot"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ManageBoxes(bOpen As Byte, Optional BlankNumber As Byte)
   'Temp
   On Error Resume Next
   z1(16).Enabled = False
   z1(17).Enabled = False
   
   z1(18).Enabled = False
   z1(19).Enabled = False
   
   'lblPart = ""
   'lblPart.ToolTipText = ""
   If BlankNumber = 1 Then lblNumber = ""
   lblType = ""
   lblRem = ""
   txtCmt = ""
   lblDate = ""
   cmbLoc = ""
   lblDate = ""
   On Error Resume Next
   
   'Open the bottom for use
   If bOpen = 1 Then
      txtLot.Enabled = True
      txtCst.Enabled = True
      txtCmt.Enabled = True
      lblDate.Enabled = True
      cmbLoc.Enabled = True
      
   Else
      'open the top for use
      cmdChg.Enabled = False
      txtCst.Enabled = False
      txtCmt.Enabled = False
      lblDate.Enabled = False
      cmbLoc.Enabled = False
   End If
   
End Sub

Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 9)
   txtCst = Format(Abs(Val(txtCst)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      If Val(txtCst) <> cOldCost Then
         With RdoCur
            !LotUnitCost = Format(Val(txtCst), ES_QuantityDataFormat)
            If Val(txtCst) > 0 Then
               !LOTDATECOSTED = Format(ES_SYSDATE, "mm/dd/yy")
            Else
               !LOTDATECOSTED = Null
            End If
            .Update
         End With
         cOldCost = Val(txtCst)
      End If
   End If
   
End Sub


Private Sub txtlot_LostFocus()
   txtLot = CheckLen(txtLot, 40)
   If Trim(txtLot) <> sOldLot Then
      If Len(Trim(txtLot)) < 5 Then
         Beep
         txtLot = sOldLot
         MsgBox "New User Lots Require At Least (5 chars).", _
            vbInformation
      Else
         With RdoCur
            !LOTUSERLOTID = txtLot
            .Update
         End With
      End If
   End If
   sOldLot = txtLot
   
End Sub











Private Function GetLotPart(sLotPart As String) As Byte
   Dim RdoPrt As ADODB.Recordset
   sSql = "Qry_GetPartsNotTools '" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         GetLotPart = 1
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "Part Number With Lot Wasn't Found."
      GetLotPart = 0
   End If
   Set RdoPrt = Nothing
   bGoodLot = GetThisLot()
   Exit Function
   
DiaErr1:
   sProcName = "getlotpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Function GetLotType(bType As Byte) As String
   Select Case bType
      Case 15
         GetLotType = "Purchase Order Receipt"
      Case 6
         GetLotType = "MO Completion"
      Case 19
         GetLotType = "Manual Adjustment"
      Case Else
         GetLotType = "Other Inventory Adustment"
   End Select
   
End Function

Private Sub txtSplit_LostFocus()
   txtSplit = CheckLen(txtSplit, 20)
   txtSplit = StrCase(txtSplit, ES_FIRSTWORD)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTSPLITCOMMENT = Trim(txtSplit)
         .Update
      End With
   End If
   
End Sub
