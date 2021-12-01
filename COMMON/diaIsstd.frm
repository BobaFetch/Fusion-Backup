VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaIsstd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Cost"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSel 
      Cancel          =   -1  'True
      Caption         =   "Select"
      Height          =   375
      Left            =   4860
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   720
      Width           =   875
   End
   Begin VB.CheckBox chkCurrent 
      Caption         =   "Set Current"
      Height          =   255
      Left            =   4380
      TabIndex        =   2
      Top             =   180
      Width           =   1275
   End
   Begin VB.CheckBox chkProposed 
      Caption         =   "Set Proposed"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   180
      Width           =   1335
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   3720
      TabIndex        =   10
      Tag             =   "3"
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox chkAbc 
      Caption         =   "ABC Classes Initialized"
      Enabled         =   0   'False
      Height          =   255
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   1995
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Height          =   315
      Left            =   6000
      TabIndex        =   13
      ToolTipText     =   "Next Part Number"
      Top             =   4320
      Width           =   875
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Height          =   315
      Left            =   5100
      TabIndex        =   12
      ToolTipText     =   "Last Part Number"
      Top             =   4320
      Width           =   875
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6000
      TabIndex        =   11
      ToolTipText     =   "Update Standard Cost To Calculated Total"
      Top             =   3960
      Width           =   875
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      ToolTipText     =   "Current Standard Cost"
      Top             =   3600
      Width           =   1000
   End
   Begin VB.TextBox txtHrs 
      Height          =   285
      Left            =   5880
      TabIndex        =   4
      ToolTipText     =   "Standard Hours To Manufacture"
      Top             =   1320
      Width           =   1000
   End
   Begin VB.TextBox txtOhd 
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      ToolTipText     =   "Manufacturing (Factory) Overhead"
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      ToolTipText     =   "Cost Of Materials Used"
      Top             =   2040
      Width           =   1000
   End
   Begin VB.TextBox txtExp 
      Height          =   285
      Left            =   5880
      TabIndex        =   7
      ToolTipText     =   "Manufacturing Expenses (Outside Services)"
      Top             =   2400
      Width           =   1000
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Will Contain A Maximum Of 300 Part Numbers (Enter Leading Characters To Refine)"
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Standard Labor Cost (Hours * Rate)"
      Top             =   1680
      Width           =   1000
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaIsstd.frx":0000
      PictureDn       =   "diaIsstd.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4770
      FormDesignWidth =   7125
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extended Description"
      Height          =   465
      Index           =   14
      Left            =   120
      TabIndex        =   40
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   285
      Index           =   15
      Left            =   2520
      TabIndex        =   38
      Top             =   3960
      Width           =   1515
   End
   Begin VB.Label lblABC 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   37
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC Class"
      Height          =   285
      Index           =   12
      Left            =   2520
      TabIndex        =   36
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
      TabIndex        =   35
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Revised"
      Height          =   285
      Index           =   13
      Left            =   120
      TabIndex        =   34
      Top             =   3600
      Width           =   1275
   End
   Begin VB.Label lblMbe 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   33
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   32
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make/Buy/Either"
      Height          =   285
      Index           =   11
      Left            =   2520
      TabIndex        =   31
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      Height          =   285
      Index           =   9
      Left            =   4560
      TabIndex        =   29
      Top             =   3600
      Width           =   1400
   End
   Begin VB.Label lblTot 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5880
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   3240
      Width           =   1400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Hours"
      Height          =   285
      Index           =   7
      Left            =   4560
      TabIndex        =   26
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      Height          =   285
      Index           =   6
      Left            =   4560
      TabIndex        =   25
      Top             =   2760
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   24
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Cost"
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   23
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   285
      Index           =   3
      Left            =   4560
      TabIndex        =   22
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   795
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   1320
      TabIndex        =   20
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
      TabIndex        =   19
      ToolTipText     =   "Part Numbers Selected In The Group"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1275
   End
End
Attribute VB_Name = "diaIsstd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/16/03 Revamped and moved to Inventory Management
'5/4/04   Edited and revised scrolling for ABC CC
'
' diaIsstd - Standard Cost

Option Explicit
Dim RdoPrt As ADODB.Recordset
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
            chkAbc.Value = .Fields(0)
         Else
            chkAbc.Value = vbUnchecked
         End If
         .Cancel
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

Private Sub chkCurrent_Click()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Current", chkCurrent.Value
End Sub

Private Sub chkProposed_Click()
   SaveSetting "Esi2000", "EsiFina", Me.Name & "_Proposed", chkProposed.Value
End Sub

Private Sub cmbPrt_Click()
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs5401"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdLst_Click()
   iCurrIdx = iCurrIdx - 1
   If iCurrIdx < 0 Then iCurrIdx = 0
   cmbPrt = cmbPrt.List(iCurrIdx)
   If cmbPrt.ListCount > 0 Then bGoodPart = GetPart()
   
End Sub

Private Sub cmdNxt_Click()
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
            If chkAbc.Value = vbChecked Then
               'GetClassCode
               bResponse = MsgBox("Update The ABC Class For The Current Cost?", _
                           ES_NOQUESTION, Caption)
               If bResponse = vbYes Then
                  RdoPrt!PASTDCOST = Val(txtCst)
                  RdoPrt!PAABC = Trim(lblABC)
                  RdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
                  RdoPrt.Update
                  MsgBox "Standard Cost And Class Updated For The Part.", _
                     vbInformation, Caption
                  lblRev = Format(ES_SYSDATE, "mm/dd/yy")
                  Exit Sub
               End If
               CancelTrans
               Exit Sub
            End If
            bResponse = MsgBox("Update The Standard Cost To The New Current Cost?", _
                        ES_NOQUESTION, Caption)
            If bResponse = vbYes Then
               RdoPrt!PASTDCOST = Val(txtCst)
               RdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
               RdoPrt.Update
               MsgBox "Standard Cost Updated For The Part.", _
                  vbInformation, Caption
               lblRev = Format(ES_SYSDATE, "mm/dd/yy")
               Exit Sub
            Else
               CancelTrans
               Exit Sub
            End If
         End If
      End If
      
      UpdatePartRecord
      
      '        On Error Resume Next
      '        GetClassCode
      '        RdoPrt!PASTDCOST = Val(lblTot)
      '        RdoPrt!PATOTCOST = Val(lblTot)
      '        RdoPrt!PABOMOH = 0
      '        RdoPrt!PABOMLABOR = 0
      '        RdoPrt!PABOMMATL = 0
      '        RdoPrt!PABOMEXP = 0
      '        RdoPrt!PABOMHRS = 0
      '        RdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
      '        If chkAbc.Value = vbChecked Then RdoPrt!PAABC = Trim(lblABC)
      '        RdoPrt.Update
      '
      '        If Err > 0 Then ValidateEdit Me
      '        If Err = 0 Then
      '            txtCst = lblTot
      '            MsgBox "Standard Cost Updated For The Part.", _
      '                vbInformation, Caption
      '            lblRev = Format(ES_SYSDATE, "mm/dd/yy")
      '        Else
      '            MsgBox "Couldn't Update Standard Cost For The Part.", _
      '                vbInformation, Caption
      '        End If
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetSetup
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PALOCATION,PAREVDATE," & vbCrLf _
          & "PAMAKEBUY,PAABC,PALEVLABOR,PALEVEXP,PALEVMATL,PALEVOH," & vbCrLf _
          & "PALEVHRS,PASTDCOST,PAEXTDESC,PATOTCOST," & vbCrLf _
          & "PATOTOH,PATOTLABOR,PATOTMATL,PATOTEXP,PATOTHRS," & vbCrLf _
          & "PABOMOH,PABOMLABOR,PABOMMATL,PABOMEXP,PABOMHRS" & vbCrLf _
          & "FROM PartTable WHERE PARTREF= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   AdoQry.parameters.Append AdoParameter
   
   Me.chkCurrent.Value = GetSetting("Esi2000", "EsiFina", Me.Name & "_Current", 1)
   Me.chkProposed.Value = GetSetting("Esi2000", "EsiFina", Me.Name & "_Proposed", 1)
   
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set RdoPrt = Nothing
   Set diaIsstd = Nothing
   
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
   Dim sPartNumber As String
   
   On Error GoTo DiaErr1
   iCurrIdx = 0
   bFilling = 1
   sPartNumber = Compress(cmbPrt)
   bLen = Len(sPartNumber)
   If bLen = 0 Then bLen = 1
   cmdLst.enabled = False
   cmdNxt.enabled = False
   cmbPrt.Clear
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE (LEFT(PARTREF," _
          & str$(bLen) & ")> ='" & sPartNumber & "' AND PATOOL=0) " _
          & "ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            If iRows > 299 Then Exit Do
            iRows = iRows + 1
            AddComboStr cmbPrt.hWnd, "" & Trim(!PARTNUM)
            .MoveNext
         Loop
         .Cancel
      End With
      lblCnt = iRows
      bFilling = 0
      If cmbPrt.ListCount > 0 Then
         cmdNxt.enabled = True
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
   AdoQry.parameters(0).Value = sPartNumber
   On Error GoTo DiaErr1
   If iCurrIdx > 0 Then cmdLst.enabled = True _
                                         Else cmdLst.enabled = False
   If iCurrIdx = cmbPrt.ListCount - 1 Then cmdNxt.enabled = False _
                 Else cmdNxt.enabled = True
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry, ES_KEYSET, True)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PARTNUM)
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblLvl = Format(!PALEVEL, "0")
         lblMbe = "" & Trim(!PAMAKEBUY)
         lblABC = "" & Trim(!PAABC)
         lblRev = Format(!PAREVDATE, "mm/dd/yy")
         
         txtHrs = Format(!PALEVHRS + !PABOMHRS, ES_QuantityDataFormat)
         txtLab = Format(!PALEVLABOR + !PABOMLABOR, ES_QuantityDataFormat)
         txtMat = Format(!PALEVMATL + !PABOMMATL, ES_QuantityDataFormat)
         txtExp = Format(!PALEVEXP + !PABOMEXP, ES_QuantityDataFormat)
         txtOhd = Format(!PALEVOH + !PABOMOH, ES_QuantityDataFormat)
         
         txtCst = Format(!PASTDCOST, ES_QuantityDataFormat)
         txtLoc = "" & Trim(!PALOCATION)
         .Cancel
      End With
      UpdateTotals
      GetPart = 1
      cmdUpd.enabled = True
   Else
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
      cmdUpd.enabled = False
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
   If bGoodPart Then
      On Error Resume Next
      RdoPrt!PASTDCOST = Format(Val(txtCst), ES_QuantityDataFormat)
      RdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
      RdoPrt.Update
      If Err > 0 Then ValidateEdit Me
      lblRev = Format(ES_SYSDATE, "mm/dd/yy")
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
   '    If bGoodPart Then
   '        On Error Resume Next
   '        RdoPrt!PALEVEXP = Format(Val(txtExp), ES_QuantityDataFormat)
   '        RdoPrt!PATOTEXP = Format(Val(txtExp), ES_QuantityDataFormat)
   '        RdoPrt.Update
   '        If Err > 0 Then ValidateEdit Me
   '    End If
   
End Sub


Private Sub txtHrs_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtHrs_LostFocus()
   txtHrs = CheckLen(txtHrs, 10)
   txtHrs = Format(Abs(Val(txtHrs)), ES_QuantityDataFormat)
   UpdateTotals
   '        If bGoodPart Then
   '            On Error Resume Next
   '            RdoPrt!PALEVHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
   '            RdoPrt!PATOTHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
   '            RdoPrt.Update
   '            If Err > 0 Then ValidateEdit Me
   '        End If
   
End Sub


Private Sub txtLab_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 10)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   UpdateTotals
   '    If bGoodPart Then
   '        On Error Resume Next
   '        RdoPrt!PALEVLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
   '        RdoPrt!PATOTLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
   '        RdoPrt.Update
   '        If Err > 0 Then ValidateEdit Me
   '    End If
   
End Sub


Private Sub txtLoc_LostFocus()
   txtLoc = Compress(txtLoc)
   txtLoc = CheckLen(txtLoc, 4)
   If bGoodPart Then
      On Error Resume Next
      RdoPrt!PALOCATION = txtLoc
      RdoPrt!PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
      RdoPrt.Update
      If Err > 0 Then ValidateEdit Me
      lblRev = Format(ES_SYSDATE, "mm/dd/yy")
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
   '    If bGoodPart Then
   '        On Error Resume Next
   '        RdoPrt!PALEVMATL = Format(Val(txtMat), ES_QuantityDataFormat)
   '        RdoPrt!PATOTMATL = Format(Val(txtMat), ES_QuantityDataFormat)
   '        RdoPrt.Update
   '        If Err > 0 Then ValidateEdit Me
   '    End If
   
End Sub


Private Sub txtOhd_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtOhd_LostFocus()
   txtOhd = CheckLen(txtOhd, 10)
   txtOhd = Format(Abs(Val(txtOhd)), ES_QuantityDataFormat)
   UpdateTotals
   '    If bGoodPart Then
   '        On Error Resume Next
   '        RdoPrt!PALEVOH = Format(Val(txtOhd), ES_QuantityDataFormat)
   '        RdoPrt!PATOTOH = Format(Val(txtOhd), ES_QuantityDataFormat)
   '        RdoPrt.Update
   '        If Err > 0 Then ValidateEdit Me
   '    End If
   
End Sub



'11/16/03
'5/19/2021 This flawed logic returned all codes for CASGAS and typically set the ABC code to A (the first one).
'Now, the current code is retained.
'Private Sub GetClassCode()
'   Dim RdoCde As ADODB.Recordset
'   Dim cCost As Currency
'
'   On Error GoTo DiaErr1
'   cCost = Val(txtCst)
'   sSql = "SELECT COABCCODE,COABCLOWCOST,COABCHIGHCOST FROM CabcTable " _
'          & "WHERE " & Format(cCost, "######0.000") & " BETWEEN COABCLOWCOST AND COABCHIGHCOST " _
'          & "AND COABCUSED=1"
'   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
'   If bSqlRows Then
'      With RdoCde
'         'if exists, else leave it alone
'         If Not IsNull(.Fields(0)) Then _
'                       lblABC = "" & Trim(.Fields(0))
'         .Cancel
'      End With
'   End If
'   Set RdoCde = Nothing
'   Exit Sub
'
'DiaErr1:
'   sProcName = "getclassco"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'End Sub
'
Private Sub UpdatePartRecord()
   
   On Error Resume Next
   'GetClassCode 'set ABC class     ' leave the code as is 5/19/2021
   
   UpdateTotals
   With RdoPrt
      If chkProposed = vbChecked Then
         !PALEVHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
         !PABOMHRS = 0
         
         !PALEVLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
         !PABOMLABOR = 0
         
         !PALEVMATL = Format(Val(txtMat), ES_QuantityDataFormat)
         !PABOMMATL = 0
         
         !PALEVEXP = Format(Val(txtExp), ES_QuantityDataFormat)
         !PABOMEXP = 0
         
         !PALEVOH = Format(Val(txtOhd), ES_QuantityDataFormat)
         !PABOMOH = 0
      End If
      
      If chkCurrent = vbChecked Then
         !PATOTHRS = Format(Val(txtHrs), ES_QuantityDataFormat)
         !PATOTLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
         !PATOTMATL = Format(Val(txtMat), ES_QuantityDataFormat)
         !PATOTEXP = Format(Val(txtExp), ES_QuantityDataFormat)
         !PATOTOH = Format(Val(txtOhd), ES_QuantityDataFormat)
         
         !PASTDCOST = Val(lblTot)
         !PATOTCOST = Val(lblTot)
         !PAREVDATE = Format(ES_SYSDATE, "mm/dd/yy")
      End If
      
'5/19/2021 No longer update PAABC in this transaction
'      If chkAbc.Value = vbChecked Then
'         !PAABC = Trim(lblABC)
'      End If
      
      .Update
      
   End With
   
   If Err > 0 Then ValidateEdit Me
   If Err = 0 Then
      txtCst = lblTot
      
      Dim sMsg As String
      If chkProposed.Value = 1 And chkCurrent = 1 Then
         sMsg = "Proposed costs and current costs"
      ElseIf chkProposed.Value = 1 Then
         sMsg = "Proposed costs"
      ElseIf chkCurrent.Value = 1 Then
         sMsg = "Current costs"
      Else
         sMsg = "No costs"
      End If
      
      MsgBox sMsg & " updated for the part", _
         vbInformation, Caption
      lblRev = Format(ES_SYSDATE, "mm/dd/yy")
   Else
      MsgBox "Couldn't Update Standard Cost For The Part.", _
         vbInformation, Caption
   End If
End Sub
