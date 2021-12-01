VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form LotsLTe03b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Split Lots"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "LotsLTe03b.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "LotsLTe03b.frx":07AE
      DownPicture     =   "LotsLTe03b.frx":1120
      Height          =   350
      Left            =   5520
      Picture         =   "LotsLTe03b.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Standard Comments"
      Top             =   2280
      Width           =   350
   End
   Begin VB.TextBox txtSplt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Splits Only 20 Char Alpha/Numeric"
      Top             =   3000
      Width           =   2150
   End
   Begin VB.TextBox txtLot 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "User Produced Lot Number Click To Set User Lot The Same"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox lblNumber 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   49
      ToolTipText     =   "System Produced Lot Number Click To Set User Lot The Same"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Change"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Change The User Lot ID"
      Top             =   1080
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      Height          =   675
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "9"
      ToolTipText     =   "Comments (2048)"
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton optDis 
      Height          =   350
      Left            =   6000
      Picture         =   "LotsLTe03b.frx":2094
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Print or View Detail"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   350
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Contains Part Numbers With Lots"
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Storage Location For This Lot"
      Top             =   3360
      Width           =   675
   End
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Tag             =   "1"
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      Tag             =   "3"
      ToolTipText     =   "Select Type From List (Or Blank)"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbItm 
      Height          =   315
      Left            =   4800
      TabIndex        =   15
      Tag             =   "1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbSon 
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Tag             =   "1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optSal 
      Caption         =   "SO"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      ToolTipText     =   "Select On Allocation Type"
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optMal 
      Caption         =   "MO"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      ToolTipText     =   "Select On Allocation Type"
      Top             =   5400
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtWum 
      Height          =   285
      Left            =   4920
      TabIndex        =   11
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   4440
      Width           =   435
   End
   Begin VB.TextBox txtLum 
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   4080
      Width           =   435
   End
   Begin VB.TextBox txtHum 
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Unit Of Measure (2)"
      Top             =   3720
      Width           =   435
   End
   Begin VB.TextBox txtWid 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Mat Width"
      Top             =   4440
      Width           =   915
   End
   Begin VB.TextBox txtLng 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Mat Length"
      Top             =   4080
      Width           =   915
   End
   Begin VB.TextBox txtHgt 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Mat Heght"
      Top             =   3720
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   50
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   6732
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5085
      FormDesignWidth =   6945
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Split Comments"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   50
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Lot Number"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   47
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   45
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Remaining"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   44
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblRem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   43
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   42
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   41
      Top             =   1920
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Cost"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   22
      Left            =   120
      TabIndex        =   39
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   38
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   21
      Left            =   120
      TabIndex        =   37
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Temporarily Disabled)"
      Height          =   255
      Index           =   20
      Left            =   3960
      TabIndex        =   36
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   19
      Left            =   4560
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO Number"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   34
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   315
      Index           =   17
      Left            =   4320
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSon 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2760
      TabIndex        =   32
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   31
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Allocate To"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   29
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
      Height          =   255
      Index           =   14
      Left            =   4080
      TabIndex        =   28
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   27
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   25
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Measure"
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   23
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   22
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot Comments"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Lot Number"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "LotsLTe03b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'7/27/05 new Called from LotsLTe03a (Splits)
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
   LotsLTe01b.txtlot = txtlot
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
   If bOnLoad Then bOnLoad = 0
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
   LotsLTe03a.Show
   Set RdoCur = Nothing
   Set LotsLTe03b = Nothing
   
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
   
   MouseCursor 13
   On Error GoTo DiaErr1
   
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
   MouseCursor 0
   Exit Sub
   
   
   
  ' Dim sDate As String
  ' Dim sVendor As String
  ' MouseCursor 13
  ' On Error GoTo DiaErr1
  ' 'SetMdiReportsize MdiSect
  ' MdiSect.Crw.Formulas(0) = "CompanyName='" & sFacility & "'"
  ' MdiSect.Crw.ReportFileName = sReportPath & "lotdetail.rpt"
  ' sSql = "{LohdTable.LOTNUMBER}='" & lblNumber & "'"
  ' MdiSect.Crw.SelectionFormula = sSql
  ' MdiSect.Crw.Destination = crptToWindow
  ' MdiSect.Crw.Action = 1
  ' MouseCursor 0
  ' Exit Sub
   
DiaErr1:
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   
End Sub


Private Sub optMal_Click()
   If optMal.Value = True Then
      z1(16).Visible = False
      z1(17).Visible = False
      lblSon.Visible = False
      cmbSon.Visible = False
      cmbItm.Visible = False
      
      z1(18).Visible = True
      z1(19).Visible = True
      cmbMon.Visible = True
      cmbRun.Visible = True
   Else
      z1(16).Visible = True
      z1(17).Visible = True
      lblSon.Visible = True
      cmbSon.Visible = True
      cmbItm.Visible = True
      
      z1(18).Visible = False
      z1(19).Visible = False
      cmbMon.Visible = False
      cmbRun.Visible = False
   End If
   
End Sub



Private Sub optSal_Click()
   If optMal.Value = True Then
      z1(16).Visible = False
      z1(17).Visible = False
      lblSon.Visible = False
      cmbSon.Visible = False
      cmbItm.Visible = False
      
      z1(18).Visible = True
      z1(19).Visible = True
      cmbMon.Visible = True
      cmbRun.Visible = True
   Else
      z1(16).Visible = True
      z1(17).Visible = True
      lblSon.Visible = True
      cmbSon.Visible = True
      cmbItm.Visible = True
      
      z1(18).Visible = False
      z1(19).Visible = False
      cmbMon.Visible = False
      cmbRun.Visible = False
   End If
   
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

Private Function GetThisLot() As Byte
   On Error GoTo DiaErr1
   ManageBoxes 0
   cmdChg.Enabled = False
   sSql = "SELECT LOTNUMBER,LOTUSERLOTID,LOTPARTREF," _
          & "LOTUNITCOST,LOTDATECOSTED,LOTADATE,LOTMATLENGTH," _
          & "LOTMATLENGTHUM,LOTMATHEIGHT,LOTMATHEIGHTHUM,LOTREMAININGQTY," _
          & "LOTMATWIDTH,LOTMATWIDTHHUM,LOTLOCATION,LOTCOMMENTS,LOTSPLITCOMMENT," _
          & "LOINUMBER,LOIRECORD,LOITYPE FROM LohdTable,LoitTable WHERE " _
          & "(LOTNUMBER='" & Trim(lblNumber) & " ' AND LOTNUMBER=LOINUMBER " _
          & "AND LOIRECORD=1)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCur, ES_KEYSET)
   If bSqlRows Then
      ManageBoxes 1
      With RdoCur
         lblNumber = "" & Trim(!lotNumber)
         txtlot = "" & Trim(!LOTUSERLOTID)
         txtCmt = "" & Trim(!LOTCOMMENTS)
         txtLoc = "" & Trim(!LOTLOCATION)
         txtCst = Format(!LotUnitCost, ES_QuantityDataFormat)
         cOldCost = !LotUnitCost
         If Not IsNull(.Fields(6)) Then
            lblDate = "" & Format(!LotADate, "mm/dd/yy")
         Else
            lblDate = Format(GetServerDateTime, "mm/dd/yy")
         End If
         txtHgt = Format(!LOTMATHEIGHT, ES_QuantityDataFormat)
         txtLng = Format(!LOTMATLENGTH, ES_QuantityDataFormat)
         txtWid = Format(!LOTMATWIDTH, ES_QuantityDataFormat)
         If Val(txtHgt) > 0 Then txtHum = "" & Trim(!LOTMATHEIGHTHUM)
         If Val(txtLng) > 0 Then txtLum = "" & Trim(!LOTMATLENGTHUM)
         If Val(txtWid) > 0 Then txtWum = "" & Trim(!LOTMATWIDTHHUM)
         lblRem = Format(!LOTREMAININGQTY, ES_QuantityDataFormat)
         lblType = GetLotType(!LOITYPE)
         txtSplt = "" & Trim(!LOTSPLITCOMMENT)
         sOldLot = lblNumber
      End With
      GetThisLot = 1
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
   lblSon.Enabled = False
   cmbSon.Enabled = False
   cmbItm.Enabled = False
   
   z1(18).Enabled = False
   z1(19).Enabled = False
   cmbMon.Enabled = False
   cmbRun.Enabled = False
   
   'lblPart = ""
   'lblPart.ToolTipText = ""
   If BlankNumber = 1 Then lblNumber = ""
   lblType = ""
   lblRem = ""
   txtCmt = ""
   lblDate = ""
   txtHgt = "0.000"
   txtLng = "0.000"
   txtWid = "0.000"
   txtHum = ""
   txtLum = ""
   txtWum = ""
   txtLoc = ""
   lblDate = ""
   On Error Resume Next
   
   'Open the bottom for use
   If bOpen = 1 Then
      txtlot.Enabled = True
      optMal.Enabled = True
      optSal.Enabled = True
      txtCst.Enabled = True
      txtCmt.Enabled = True
      lblDate.Enabled = True
      txtHgt.Enabled = True
      txtLng.Enabled = True
      txtWid.Enabled = True
      txtHum.Enabled = True
      txtLum.Enabled = True
      txtWum.Enabled = True
      txtLoc.Enabled = True
      
   Else
      'open the top for use
      cmdChg.Enabled = False
      optMal.Enabled = False
      optSal.Enabled = False
      txtCst.Enabled = False
      txtCmt.Enabled = False
      lblDate.Enabled = False
      txtHgt.Enabled = False
      txtLng.Enabled = False
      txtWid.Enabled = False
      txtHum.Enabled = False
      txtLum.Enabled = False
      txtWum.Enabled = False
      txtLoc.Enabled = False
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


Private Sub txtHgt_LostFocus()
   txtHgt = CheckLen(txtHgt, 8)
   txtHgt = Format(Abs(Val(txtHgt)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATHEIGHT = txtHgt
         .Update
      End With
   End If
   
End Sub


Private Sub txtHum_LostFocus()
   txtHum = CheckLen(txtHum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATHEIGHTHUM = txtHum
         .Update
      End With
   End If
   
End Sub


Private Sub txtLng_LostFocus()
   txtLng = CheckLen(txtLng, 8)
   txtLng = Format(Abs(Val(txtLng)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATLENGTH = txtLng
         .Update
      End With
   End If
   
End Sub


Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 4)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTLOCATION = txtLoc
         .Update
      End With
   End If
   
End Sub



Private Sub txtlot_LostFocus()
   txtlot = CheckLen(txtlot, 40)
   If Trim(txtlot) <> sOldLot Then
      If Len(Trim(txtlot)) < 5 Then
         Beep
         txtlot = sOldLot
         MsgBox "New User Lots Require At Least (5 chars).", _
            vbInformation
      Else
         If bGoodLot = 1 Then
            '                With RdoCur

            '                    !LOTUSERLOTID = txtLot
            '                    .Update
            '                End With
         End If
      End If
   End If
   sOldLot = txtlot
   
End Sub


Private Sub txtLum_LostFocus()
   txtLum = CheckLen(txtLum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATLENGTHUM = txtLum
         .Update
      End With
   End If
   
End Sub


Private Sub txtSplt_LostFocus()
   txtSplt = CheckLen(txtSplt, 20)
   txtSplt = StrCase(txtSplt, ES_FIRSTWORD)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTSPLITCOMMENT = Trim(txtSplt)
         .Update
      End With
   End If
   
End Sub


Private Sub txtWid_LostFocus()
   txtWid = CheckLen(txtWid, 8)
   txtWid = Format(Abs(Val(txtWid)), ES_QuantityDataFormat)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATWIDTH = txtWid
         .Update
      End With
   End If
   
End Sub


Private Sub txtWum_LostFocus()
   txtWum = CheckLen(txtWum, 2)
   If bGoodLot = 1 Then
      On Error Resume Next
      With RdoCur
         !LOTMATWIDTHHUM = txtWum
         .Update
      End With
   End If
   
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

Public Sub GetCalledLot()
   Dim bByte As Byte
   bByte = GetLotPart(cmbPrt)
   
End Sub
