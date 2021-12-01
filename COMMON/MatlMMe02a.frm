VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form MatlMMe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Cost"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "MatlMMe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optAbc 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3000
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbCde 
      Height          =   288
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   37
      Tag             =   "3"
      ToolTipText     =   "Product Code (Leading Characters Or Blank For All)"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox optBom 
      Alignment       =   1  'Right Justify
      Caption         =   "Update BOM Level Costs"
      Height          =   252
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Updates The Bill Of Material Costs At This Part Level"
      Top             =   4320
      Width           =   2412
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Inventory Location"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Height          =   315
      Left            =   6000
      TabIndex        =   12
      ToolTipText     =   "Next Part Number"
      Top             =   4340
      Width           =   875
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Height          =   315
      Left            =   5100
      TabIndex        =   11
      ToolTipText     =   "Last Part Number"
      Top             =   4340
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   10
      ToolTipText     =   "Update Standard Cost To Calculated Total"
      Top             =   3960
      Width           =   875
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Current Standard Cost"
      Top             =   3600
      Width           =   1000
   End
   Begin VB.TextBox txtHrs 
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Standard Hours To Manufacture"
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox txtOhd 
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      ToolTipText     =   "Manufacturing (Factory) Overhead"
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Cost Of Materials Used"
      Top             =   2400
      Width           =   1000
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Manufacturing Expenses (Outside Services)"
      Top             =   2040
      Width           =   1000
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Will Contain A Maximum Of 300 Part Numbers (Enter Leading Characters To Refine)"
      Top             =   720
      Width           =   3255
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "S&elect"
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      ToolTipText     =   "Select Series Of Parts (Equal Or Greater Than Search) Up To 300 Entries"
      Top             =   720
      Width           =   875
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Standard Labor Cost (Hours * Rate)"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   240
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4830
      FormDesignWidth =   7005
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   312
      Left            =   1320
      TabIndex        =   42
      ToolTipText     =   "Part Unit Of Measure"
      Top             =   3960
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Of Meas"
      Height          =   288
      Index           =   17
      Left            =   120
      TabIndex        =   41
      Top             =   3960
      Width           =   1512
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Classes Initialized "
      Height          =   288
      Index           =   14
      Left            =   600
      TabIndex        =   40
      ToolTipText     =   "Checked If Setup"
      Top             =   60
      Visible         =   0   'False
      Width           =   2232
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Part Numbers And Product Codes For A Refined Search"
      Height          =   288
      Index           =   16
      Left            =   120
      TabIndex        =   38
      Top             =   360
      Width           =   4872
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   285
      Index           =   15
      Left            =   2520
      TabIndex        =   36
      Top             =   3960
      Width           =   1515
   End
   Begin VB.Label lblABC 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   35
      ToolTipText     =   "ABC Class (If Initialized)"
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Class"
      Height          =   285
      Index           =   12
      Left            =   2520
      TabIndex        =   34
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   6840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   33
      ToolTipText     =   "Last Revised Date"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Revised"
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label lblMbe 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   31
      ToolTipText     =   "Responsibility"
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   30
      ToolTipText     =   "Part Type (Level 1 - 8)"
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make/Buy/Either"
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   29
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      Height          =   285
      Index           =   9
      Left            =   4560
      TabIndex        =   27
      Top             =   3600
      Width           =   1400
   End
   Begin VB.Label lblTot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   26
      ToolTipText     =   "Total Of Columns -  Cost To Update To"
      Top             =   3240
      Width           =   1000
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   285
      Index           =   8
      Left            =   4560
      TabIndex        =   25
      Top             =   3240
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Hours"
      Height          =   285
      Index           =   7
      Left            =   4560
      TabIndex        =   24
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   23
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   22
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Cost"
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   21
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   20
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   1320
      TabIndex        =   18
      ToolTipText     =   "Extended Description"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblCnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   315
      Left            =   1320
      TabIndex        =   17
      ToolTipText     =   "Part Numbers Selected In The Group"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      ToolTipText     =   "Part Description"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1275
   End
End
Attribute VB_Name = "MatlMMe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/16/03 Revamped and moved to Inventory Management
'5/4/04   Edited and revised scrolling for ABC CC
'11/9/05  Changed LEV Costs to TOT Costs
Option Explicit
Dim AdoPrt As ADODB.Recordset
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bFilling As Byte
Dim bOnLoad As Byte
Dim bGoodPart As Byte
Dim bSelect As Byte

Dim iCurrIdx As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetSetup()
   Dim RdoSet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CycleCountInitialized FROM Preferences"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSet, ES_FORWARD)
   If bSqlRows Then
      With RdoSet
         If Not IsNull(.Fields(0)) Then
            optAbc.Value = .Fields(0)
         Else
            optAbc.Value = vbUnchecked
         End If
         ClearResultSet RdoSet
      End With
   End If
   Set RdoSet = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getsetup"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbCde_LostFocus()
   If cmbCde = "" Then cmbCde = "ALL"
   
End Sub


Private Sub cmbPrt_Click()
   On Error Resume Next
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   If cmbPrt.ListIndex >= 0 Then iCurrIdx = cmbPrt.ListIndex
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Trim(cmbPrt) <> "" Then
      If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   Else
      ClearBoxes
   End If
   
End Sub


Private Sub cmdCan_Click()
   UpdateKeySet
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5401"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdLst_Click()
   UpdateTotals
   UpdateKeySet
   iCurrIdx = iCurrIdx - 1
   If iCurrIdx < 0 Then iCurrIdx = 0
   cmbPrt = cmbPrt.List(iCurrIdx)
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   
End Sub

Private Sub cmdNxt_Click()
   UpdateTotals
   UpdateKeySet
   iCurrIdx = iCurrIdx + 1
   If iCurrIdx > cmbPrt.ListCount - 1 Then iCurrIdx = cmbPrt.ListCount - 1
   cmbPrt = cmbPrt.List(iCurrIdx)
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   
End Sub

Private Sub cmdSel_Click()
   FillCombo
   
End Sub

Private Sub cmdSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bSelect = 1
   
End Sub


Private Sub cmdUpd_Click()
   Dim bResponse As Byte
   If bGoodPart Then
      If Val(lblTot) = 0 Then
         bResponse = MsgBox("The Value To Be Updated To Is Zero. Are You" & vbCr _
                     & "Certain That You Want The Standard Cost Set To Zero?", _
                     ES_NOQUESTION, Caption)
         If bResponse = vbNo Then
            If optAbc.Value = vbChecked Then
               GetClassCode
               bResponse = MsgBox("Update The ABC Class For The Current Cost?", _
                           ES_NOQUESTION, Caption)
               If bResponse = vbYes Then
                  AdoPrt!PASTDCOST = Val(txtCst)
                  AdoPrt!PAABC = Trim(lblABC)
                  AdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
                  AdoPrt.Update
                  MsgBox "Standard Cost And Class Updated For The Part.", _
                     vbInformation, Caption
                  lblRev = Format(ES_SYSDATE, "mm/dd/yyyy")
                  Exit Sub
               End If
               CancelTrans
               Exit Sub
            End If
            bResponse = MsgBox("Update The Standard Cost To The New Current Cost?", _
                        ES_NOQUESTION, Caption)
            If bResponse = vbYes Then
               AdoPrt!PASTDCOST = Val(txtCst)
               AdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
               AdoPrt.Update
               MsgBox "Standard Cost Updated For The Part.", _
                  vbInformation, Caption
               lblRev = Format(ES_SYSDATE, "mm/dd/yyyy")
               Exit Sub
            Else
               CancelTrans
               Exit Sub
            End If
         End If
      End If
      
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      GetClassCode
      AdoPrt!PASTDCOST = Val(lblTot)
      AdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
      If optAbc.Value = vbChecked Then AdoPrt!PAABC = Trim(lblABC)
      AdoPrt.Update
      
      If clsADOCon.ADOErrNum > 0 Then ValidateEdit
      If clsADOCon.ADOErrNum = 0 Then
         txtCst = lblTot
         MsgBox "Standard Cost Updated For The Part.", _
            vbInformation, Caption
         lblRev = Format(ES_SYSDATE, "mm/dd/yyyy")
      Else
         MsgBox "Couldn't Update Standard Cost For The Part.", _
            vbInformation, Caption
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetSetup
      FillProductCodes
      cmbCde = "ALL"
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   GetOptions
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PALOCATION,PAREVDATE," _
          & "PAMAKEBUY,PAUNITS,PAABC,PATOTLABOR,PATOTEXP,PATOTMATL,PATOTOH," _
          & "PATOTHRS,PASTDCOST,PAEXTDESC FROM PartTable WHERE PARTREF= ? "
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.Size = 30
   
   AdoQry.Parameters.Append AdoParameter
   
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
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set AdoPrt = Nothing
   Set MatlMMe02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtLab = "0.000"
   txtMat = "0.000"
   txtExp = "0.000"
   txtOhd = "0.000"
   txtHrs = "0.000"
   lblTot = "0.000"
   txtCst = "0.000"
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim bLen As Byte
   Dim iRows As Integer
   Dim sCode As String
   Dim sPartNumber As String
   
   On Error GoTo DiaErr1
   iCurrIdx = 0
   bFilling = 1
   sPartNumber = Compress(cmbPrt)
   bLen = Len(sPartNumber)
   If bLen = 0 Then bLen = 1
   cmdLst.Enabled = False
   cmdNxt.Enabled = False
   If cmbCde = "ALL" Then sCode = "" Else sCode = cmbCde
   cmbPrt.Clear
   ClearBoxes
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE (LEFT(PARTREF," _
          & str$(bLen) & ")> ='" & sPartNumber & "' AND PATOOL=0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 "
   If sCode <> "" Then sSql = sSql & " AND PAPRODCODE LIKE '" & sCode & "%' "
   sSql = sSql & ") ORDER BY PARTREF "
   
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            If iRows > 299 Then Exit Do
            iRows = iRows + 1
            AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
      lblCnt = iRows
      bFilling = 0
      If cmbPrt.ListCount > 0 Then
         cmdNxt.Enabled = True
         cmbPrt = cmbPrt.List(0)
         bGoodPart = GetPart()
      End If
   Else
      MsgBox "No Matching Parts Were Found.", vbInformation, _
         Caption
   End If
   bSelect = 0
   bFilling = 0
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart() As Byte
   Dim sPartNumber As String
   sPartNumber = Compress(cmbPrt)
   'RdoQry(0) = sPartNumber
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   On Error GoTo DiaErr1
   If iCurrIdx > 0 Then cmdLst.Enabled = True _
                                         Else cmdLst.Enabled = False
   If iCurrIdx = cmbPrt.ListCount - 1 Then cmdNxt.Enabled = False _
                 Else cmdNxt.Enabled = True
   bSqlRows = clsADOCon.GetQuerySet(AdoPrt, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With AdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblLvl = Format(!PALEVEL, "0")
         lblMbe = "" & Trim(!PAMAKEBUY)
         lblABC = "" & Trim(!PAABC)
         lblRev = Format(!PAREVDATE, "mm/dd/yyyy")
         txtLab = Format(!PATOTLABOR, ES_QuantityDataFormat)
         txtExp = Format(!PATOTEXP, ES_QuantityDataFormat)
         txtMat = Format(!PATOTMATL, ES_QuantityDataFormat)
         txtOhd = Format(!PATOTOH, ES_QuantityDataFormat)
         txtHrs = Format(!PATOTHRS, ES_QuantityDataFormat)
         txtCst = Format(!PASTDCOST, ES_QuantityDataFormat)
         txtLoc = "" & Trim(!PALOCATION)
         lblUom = "" & Trim(!PAUNITS)
      End With
      UpdateTotals
      GetPart = 1
      cmdUpd.Enabled = True
   Else
      ClearBoxes
      cmdUpd.Enabled = False
      GetPart = 0
      If bSelect = 0 Then lblDsc = "*** No Current Part ***"
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If bFilling = 0 Then
      If lblDsc = "*** No Current Part ***" Then
         lblDsc.ForeColor = ES_RED
      Else
         lblDsc.ForeColor = vbBlack
      End If
   End If
   
End Sub


Private Sub UpdateTotals()
   Dim cTotal As Currency
   cTotal = cTotal + Val(txtLab)
   cTotal = cTotal + Val(txtMat)
   cTotal = cTotal + Val(txtExp)
   cTotal = cTotal + Val(txtOhd)
   lblTot = Format(cTotal, ES_QuantityDataFormat)
   
End Sub

Private Sub txtCst_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtCst_LostFocus()
   txtCst = CheckLen(txtCst, 10)
   txtCst = Format(Abs(Val(txtCst)), ES_QuantityDataFormat)
   UpdateTotals
   UpdateKeySet
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PASTDCOST = Format(Val(txtCst), ES_QuantityDataFormat)
      AdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
      lblRev = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
End Sub


Private Sub txtExp_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtExp_LostFocus()
   txtExp = CheckLen(txtExp, 10)
   txtExp = Format(Abs(Val(txtExp)), ES_QuantityDataFormat)
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PATOTEXP = Format(Val(txtExp), ES_QuantityDataFormat)
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtHrs_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 10)
   txtHrs = Format(Abs(Val(txtHrs)), ES_QuantityDataFormat)
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PATOTHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLab_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 10)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PATOTLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtLoc_LostFocus()
   txtLoc = Compress(txtLoc)
   txtLoc = CheckLen(txtLoc, 4)
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PALOCATION = txtLoc
      AdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yyyy")
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
      lblRev = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If

End Sub

Private Sub txtMat_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 10)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PATOTMATL = Format(Val(txtMat), ES_QuantityDataFormat)
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtOhd_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtOhd_LostFocus()
   txtOhd = CheckLen(txtOhd, 10)
   txtOhd = Format(Abs(Val(txtOhd)), ES_QuantityDataFormat)
   UpdateTotals
   If bGoodPart Then
      On Error Resume Next
      'AdoPrt.Edit
      AdoPrt!PATOTOH = Format(Val(txtOhd), ES_QuantityDataFormat)
      AdoPrt.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



'11/16/03

Private Sub GetClassCode()
   Dim RdoCde As ADODB.Recordset
   Dim cCost As Currency
   
   On Error GoTo DiaErr1
   cCost = Val(txtCst)
   sSql = "SELECT COABCCODE,COABCLOWCOST,COABCHIGHCOST FROM CabcTable " _
          & "WHERE " & Format(cCost, "######0.000") & " BETWEEN COABCLOWCOST AND COABCHIGHCOST " _
          & "AND COABCUSED=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         'if exists, else leave it alone
         If Not IsNull(.Fields(0)) Then _
                       lblABC = "" & Trim(.Fields(0))
         ClearResultSet RdoCde
      End With
   End If
   Set RdoCde = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getclassco"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub UpdateKeySet()
   On Error Resume Next
   With AdoPrt
      '.Edit
      !PATOTHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
      !PATOTLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
      !PATOTEXP = Format(Val(txtExp), ES_QuantityDataFormat)
      !PATOTMATL = Format(Val(txtMat), ES_QuantityDataFormat)
      !PATOTOH = Format(Val(txtOhd), ES_QuantityDataFormat)
      !PALOCATION = txtLoc
      If optBom.Value = vbChecked Then
         !PALEVHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
         !PALEVLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
         !PALEVEXP = Format(Val(txtExp), ES_QuantityDataFormat)
         !PALEVMATL = Format(Val(txtMat), ES_QuantityDataFormat)
         !PALEVOH = Format(Val(txtOhd), ES_QuantityDataFormat)
      End If
      .Update
   End With
End Sub


Private Sub SaveOptions()
   SaveSetting "Esi2000", "EsiInvc", "isstd", optBom.Value
   
End Sub

Private Sub GetOptions()
   On Error Resume Next
   optBom = GetSetting("Esi2000", "EsiInvc", "isstd", optBom.Value)
   
End Sub


Private Sub ClearBoxes()
   lblDsc = ""
   lblExt = ""
   lblLvl = ""
   lblMbe = ""
   lblRev = ""
   txtLab = "0.000"
   txtMat = "0.000"
   txtExp = "0.000"
   txtOhd = "0.000"
   txtHrs = "0.000"
   lblTot = "0.000"
   txtCst = "0.000"
   lblUom = ""
   lblABC = ""
   lblCnt = "0"
   
End Sub

