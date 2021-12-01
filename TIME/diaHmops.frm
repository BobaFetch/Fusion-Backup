VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaHmops 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operation Completions"
   ClientHeight    =   2340
   ClientLeft      =   1740
   ClientTop       =   1065
   ClientWidth     =   6435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "Press To Update This Operation"
      Top             =   1560
      Width           =   875
   End
   Begin VB.TextBox txtScr 
      Height          =   315
      Left            =   3960
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Scrap Quantity (Rejected)"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtRwk 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Quantity For Rework (Rejected)"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox txtAcd 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Accepted Quantity"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.ComboBox cmbShp 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Shop"
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Work Center"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox optCom 
      Alignment       =   1  'Right Justify
      Caption         =   "Complete"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Mark Operation Complete"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtOpn 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "1"
      ToolTipText     =   "Operation "
      Top             =   840
      Width           =   495
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2640
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2340
      FormDesignWidth =   6435
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   19
      ToolTipText     =   "Subject Help"
      Top             =   5520
      Visible         =   0   'False
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
      PictureUp       =   "diaHmops.frx":0000
      PictureDn       =   "diaHmops.frx":0146
   End
   Begin VB.Label lblOrig 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5160
      TabIndex        =   23
      ToolTipText     =   "Beginning Run Quantity"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Order"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Scrap"
      Height          =   285
      Index           =   9
      Left            =   2760
      TabIndex        =   21
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rework"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label lblRun 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Org Quantity    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Req/Acc Qty    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   14
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sch/Act Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop/Work Center          "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Op No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblReq 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Most Current Quantity Remaining"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblSch 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   1035
   End
End
Attribute VB_Name = "diaHmops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/7/03 New for Time Cards
Option Explicit
Dim AdoOps As ADODB.Command
Dim ADOParameter1 As ADODB.Parameter
Dim ADOParameter2 As ADODB.Parameter
Dim ADOParameter3 As ADODB.Parameter

Dim bChanged As Byte
Dim bOnLoad As Byte
Dim bOpComplete As Byte
Dim iIndex As Integer
Dim iTotalOps As Integer
Dim iOpCur As Integer

Dim cOrigQty As Currency
Dim cRunqty As Currency

Dim sShop As String
Dim sCenter As String

Dim sOldShop As String
Dim sOldCenter As String
Dim sComments As String

Dim iOpno(300) As Integer

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtAcd = Format(ES_SYSDATE, "mm/dd/yy")
   
End Sub


Private Sub cmbShp_Click()
   FillWorkCenters
   
End Sub


Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   FindShop Me
   sShop = Compress(cmbShp)
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   sCenter = Compress(cmbWcn)
   
End Sub

Private Sub cmdCan_Click()
   Dim b As Byte
   If bChanged Then
      b = MsgBox("There Are Changes To The Data. " & vbCr _
          & "Do You Want To Exit Without Udating?", _
          ES_NOQUESTION, Caption)
      If b = vbNo Then Exit Sub
   End If
   SetMoCurrentOp
   Unload Me
   
End Sub




Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs4103.htm"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub



Private Sub cmdUpd_Click()
   Dim b As Byte
   b = CheckOpQuantity()
   If b = 0 Then
      MsgBox "The Rejected And Accepted Quantities Are" & vbCr _
         & "Greater Than The Available Quantity.", _
         vbInformation, Caption
      Exit Sub
   End If
   If optCom.Value = vbUnchecked Then txtAcd = ""
   UpdateOp
   
End Sub

Private Sub Form_Activate()
   If bOnLoad Then
      GetMoreRunInfo
      GetThisOp
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub

Private Sub Form_Load()
   SetFormSize Me
   Move Left + 500, Top + 1000
   FormatControls
   sSql = "SELECT TOP 1 OPREF,OPRUN,OPNO,OPSHOP,OPCENTER,OPCOMT," _
          & "OPSCHEDDATE,OPNOTES,OPCOMPDATE,OPINSP,OPYIELD," _
          & "OPCOMPLETE,OPACCEPT,OPREWORK,OPSCRAP " _
          & "FROM RnopTable WHERE OPREF= ? AND OPRUN= ? AND OPNO= ?"
    Set AdoOps = New ADODB.Command
    AdoOps.CommandText = sSql
    
    Set ADOParameter1 = New ADODB.Parameter
    ADOParameter1.Type = adChar
    ADOParameter1.Size = 30
    
    Set ADOParameter2 = New ADODB.Parameter
    ADOParameter2.Type = adInteger
    
    Set ADOParameter3 = New ADODB.Parameter
    ADOParameter3.Type = adInteger
    
    AdoOps.Parameters.Append ADOParameter1
    AdoOps.Parameters.Append ADOParameter2
    AdoOps.Parameters.Append ADOParameter3

   
   'Set RdoOps = RdoCon.CreateQuery("", sSql)
   'RdoOps.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub UpdateOp()
   Dim cYield As Currency
   Dim sCompDate As String
   MouseCursor 13
   If Len(Trim(cmbShp)) = "" Then
      MsgBox "Requires A Valid Shop.", vbExclamation, Caption
      Exit Sub
   End If
   If optCom Then
      cYield = Val(txtQty)
      sCompDate = "'" & txtAcd & "'"
   Else
      cYield = 0
      sCompDate = "Null"
   End If
   
   sShop = Compress(cmbShp)
   sCenter = Compress(cmbWcn)
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "UPDATE RnopTable SET " _
          & "OPSHOP='" & sShop & "'," _
          & "OPCENTER='" & sCenter & "'," _
          & "OPCOMPDATE=" & sCompDate & "," _
          & "OPYIELD=" & cYield & "," _
          & "OPACCEPT=" & Val(cYield) & "," _
          & "OPREWORK=" & Val(txtRwk) & "," _
          & "OPSCRAP=" & Val(txtScr) & "," _
          & "OPCOMPLETE=" & str(optCom.Value) & " " _
          & "WHERE OPREF='" & Compress(lblPrt) & "' AND " _
          & "OPRUN=" & Val(lblRun) & " AND OPNO=" & Val(txtOpn)
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.ADOErrNum = 0 Then UpdateMo
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "updateop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'diaScoop.optLoaded = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set ADOParameter1 = Nothing
   Set ADOParameter2 = Nothing
   Set ADOParameter3 = Nothing
   Set AdoOps = Nothing
   'RdoRes.Close
   'RdoOps.Close
   Set diaHmops = Nothing
   
End Sub


Private Sub GetThisOp()
   Dim RdoRps As ADODB.Recordset
   On Error GoTo DiaErr1
'   RdoOps.RowsetSize = 1
'   RdoOps(0) = Compress(lblPrt)
'   RdoOps(1) = Val(lblRun)
'   RdoOps(2) = Val(txtOpn)
   ADOParameter1.Value = Compress(lblPrt)
   ADOParameter2.Value = Val(lblRun)
   ADOParameter3.Value = Val(txtOpn)
   
   bSqlRows = clsADOCon.GetQuerySet(RdoRps, AdoOps, ES_KEYSET)
   If bSqlRows Then
      With RdoRps
         cmbShp = "" & Trim(!OPSHOP)
         cmbWcn = "" & Trim(!OPCENTER)
         sOldCenter = cmbWcn
         cmbWcn = GetCenter(cmbWcn)
         lblSch = "" & Format(!OPSCHEDDATE, "mm/dd/yy")
         txtRwk = Format(!OPREWORK, ES_QuantityDataFormat)
         txtScr = Format(!OPSCRAP, ES_QuantityDataFormat)
         lblOrig = Format(cOrigQty, ES_QuantityDataFormat)
         If !OPCOMPLETE = 0 Then
            cmdUpd.Enabled = True
            lblReq = Format(cRunqty, ES_QuantityDataFormat)
            cmbShp.Enabled = True
            cmbWcn.Enabled = True
            cmbShp.ForeColor = ES_BLUE
            cmbWcn.ForeColor = ES_BLUE
            txtRwk.Enabled = True
            txtScr.Enabled = True
            txtAcd.Enabled = True
            txtQty.Enabled = True
            optCom.Value = vbUnchecked
            txtQty = Format(cRunqty, ES_QuantityDataFormat)
         Else
            cmdUpd.Enabled = False
            lblReq = Format(!OPACCEPT, ES_QuantityDataFormat)
            cmbShp.Enabled = False
            cmbWcn.Enabled = False
            cmbShp.ForeColor = vbGrayText
            cmbWcn.ForeColor = vbGrayText
            txtRwk.Enabled = False
            txtScr.Enabled = False
            txtAcd.Enabled = False
            txtQty.Enabled = False
            optCom.Value = vbChecked
            txtAcd = "" & Format(!OPCOMPDATE, "mm/dd/yy")
            txtQty = Format(0 + !OPYIELD, ES_QuantityDataFormat)
         End If
         bOpComplete = !OPCOMPLETE
         .Cancel
      End With
      bChanged = 0
   End If
   On Error Resume Next
   Set RdoRps = Nothing
   sOldShop = cmbShp
   FindShop Me
   FillWorkCenters
   Exit Sub
   
DiaErr1:
   sProcName = "getthisop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillOps()
   Dim RdoFil As ADODB.Recordset
   Erase iOpno
   iTotalOps = 0
   iIndex = 1
   On Error GoTo DiaErr1
   sSql = "SELECT OPREF,OPNO,OPRUN,OPCOMPLETE,OPCOMT FROM RnopTable WHERE " _
          & "OPREF='" & Compress(lblPrt) & "' AND OPRUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFil, ES_KEYSET)
   If bSqlRows Then
      With RdoFil
         txtOpn = Format(!opNo, "000")
         Do Until .EOF
            iTotalOps = iTotalOps + 1
            iOpno(iTotalOps) = !opNo
            sComments = "" & Trim(!OPCOMT)
            .MoveNext
         Loop
         .Cancel
      End With
   Else
      MsgBox "There Are No Incomplete Operations For This MO.", _
         vbInformation, Caption
      Unload Me
      Exit Sub
   End If
   Set RdoFil = Nothing
   sSql = "Qry_FillShops "
   LoadComboBox cmbShp
   GetThisOp
   Exit Sub
   
DiaErr1:
   sProcName = "fillops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillWorkCenters()
   Dim sShop As String
   cmbWcn.Clear
   If cmbShp = "" Then cmbShp = sOldShop
   sShop = Compress(cmbShp)
   
   On Error GoTo DiaErr1
   '1/6/04
   sSql = "Qry_FillWorkCenters '" & sShop & "'"
   LoadComboBox cmbWcn
   If Trim(sOldCenter) = "" Then
      If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   Else
      cmbWcn = sOldCenter
   End If
   cmbWcn = GetCenter(sOldCenter)
   Exit Sub
   
DiaErr1:
   sProcName = "fillworkcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetMoreRunInfo()
   Dim RdoInf As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RUNREF,RUNNO,RUNQTY,RUNOPCUR,RUNREMAININGQTY " _
          & "FROM RunsTable WHERE RUNREF='" & Compress(lblPrt) & "'" _
          & "AND RUNNO=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInf, ES_FORWARD)
   If bSqlRows Then
      With RdoInf
         cOrigQty = !RUNQTY
         cRunqty = !RUNREMAININGQTY
         iOpCur = !RUNOPCUR
         .Cancel
      End With
   End If
   Set RdoInf = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getmorerun"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub optCom_Click()
   If Not bOnLoad Then
      CheckOps
   Else
      bOnLoad = 0
   End If
   
End Sub

Private Sub txtAcd_Change()
   bChanged = 1
   
End Sub

Private Sub txtAcd_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtAcd_LostFocus()
   If Len(Trim(txtAcd)) Then txtAcd = CheckDate(txtAcd)
   If bChanged Then cmdUpd.Enabled = True
   
End Sub


Private Sub txtQty_Change()
   bChanged = 1
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If bChanged Then cmdUpd.Enabled = True
   
End Sub



Private Sub CheckOps()
   If optCom.Value = vbChecked Then
      cmbShp.Enabled = False
      cmbWcn.Enabled = False
      cmbShp.ForeColor = vbGrayText
      cmbWcn.ForeColor = vbGrayText
      txtRwk.Enabled = False
      txtScr.Enabled = False
      txtAcd.Enabled = False
      txtQty.Enabled = False
   Else
      cmbShp.Enabled = True
      cmbWcn.Enabled = True
      cmbShp.ForeColor = ES_BLUE
      cmbWcn.ForeColor = ES_BLUE
      txtRwk.Enabled = True
      txtScr.Enabled = True
      txtAcd.Enabled = True
      If Trim(txtAcd) = "" Then txtAcd = Format(ES_SYSDATE, "mm/dd/yy")
      txtQty.Enabled = True
   End If
   bChanged = 1
   
End Sub

Private Function GetCenter(sNewCenter As String) As String
   Dim RdoCnt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT WCNREF,WCNNUM FROM WcntTable " _
          & "WHERE WCNREF='" & Compress(sNewCenter) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_FORWARD)
   If bSqlRows Then
      With RdoCnt
         GetCenter = Trim(!WCNNUM)
         .Cancel
      End With
   Else
      GetCenter = ""
   End If
   Set RdoCnt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

'Set the current op

Private Sub SetMoCurrentOp()
   On Error GoTo DiaErr1
   Dim RdoLst As ADODB.Recordset
   sSql = "SELECT OPREF,OPRUN,OPNO,OPCOMPLETE FROM RnopTable " _
          & "WHERE (OPREF='" & Compress(lblPrt) & "' AND OPRUN=" & Val(lblRun) & " " _
          & "AND OPCOMPLETE=0) ORDER BY OPNO "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         iOpCur = !opNo
         .Cancel
      End With
   Else
      sSql = "SELECT MAX(OPNO) FROM RnopTable " _
             & "WHERE (OPREF='" & Compress(lblPrt) & "' AND OPRUN=" & Val(lblRun) & ")"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
      If bSqlRows Then
         With RdoLst
            If Not IsNull(.Fields(0)) Then _
                          iOpCur = .Fields(0)
            .Cancel
         End With
      End If
   End If
   Set RdoLst = Nothing
   sSql = "UPDATE RunsTable SET RUNOPCUR=" & Trim(str(iOpCur)) & " " _
          & "WHERE RUNREF='" & Compress(lblPrt) & "' AND RUNNO=" & Val(lblRun) & " "
   clsADOCon.ExecuteSQL sSql
   Exit Sub
   
DiaErr1:
   sProcName = "SetMoCurrent"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtRwk_Change()
   bChanged = 1
   
End Sub

Private Sub txtRwk_LostFocus()
   txtRwk = CheckLen(txtRwk, 9)
   txtRwk = Format(Abs(Val(txtRwk)), ES_QuantityDataFormat)
   If bChanged Then cmdUpd.Enabled = True
   
End Sub


'10/6/03

Private Function CheckOpQuantity() As Byte
   Dim sAvail As Currency
   Dim sAccept As Currency
   Dim sReject As Currency
   
   sAvail = Val(lblReq)
   sAccept = Val(txtQty)
   sReject = Val(txtRwk) + Val(txtScr) + sAccept
   If sReject > sAvail Then CheckOpQuantity = 0 Else _
                                              CheckOpQuantity = 1
   
End Function

Private Sub txtScr_Change()
   bChanged = 1
   
End Sub

Private Sub txtScr_LostFocus()
   txtScr = CheckLen(txtScr, 9)
   txtScr = Format(Abs(Val(txtScr)), ES_QuantityDataFormat)
   If bChanged Then cmdUpd.Enabled = True
   
End Sub



'10/6/03

Private Sub UpdateMo()
   Dim RdoQty As ADODB.Recordset
   Dim cRework As Currency
   Dim cReject As Currency
   Dim cScrap As Currency
   
   sSql = "SELECT SUM(OPREWORK),SUM(OPSCRAP) FROM RnopTable " _
          & "WHERE OPREF='" & Compress(lblPrt) & "' AND OPRUN=" & Val(lblRun) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoQty, ES_FORWARD)
   If bSqlRows Then
      With RdoQty
         If Not IsNull(.Fields(0)) Then
            cRework = .Fields(0)
         End If
         If Not IsNull(.Fields(1)) Then
            cScrap = .Fields(1)
         End If
         .Cancel
      End With
   End If
   cReject = cRework + cScrap
   sSql = "UPDATE RunsTable SET RUNREMAININGQTY=RUNQTY-" _
          & cReject & ",RUNSCRAP=" & cScrap & ",RUNREWORK=" & cRework & " " _
          & "WHERE RUNREF='" & Compress(lblPrt) & "' AND " _
          & "RUNNO=" & Val(lblRun) & " "
   clsADOCon.ExecuteSQL sSql
   SysMsg "Operation Updated.", True
   Set RdoQty = Nothing
   Unload Me
   
End Sub
