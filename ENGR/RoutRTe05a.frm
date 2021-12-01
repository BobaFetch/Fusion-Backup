VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form RoutRTe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Routing Operation Library"
   ClientHeight    =   4770
   ClientLeft      =   1890
   ClientTop       =   1515
   ClientWidth     =   6180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   6012
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RoutRTe05a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1080
      Width           =   3075
   End
   Begin VB.ComboBox cmbOpr 
      Height          =   315
      Left            =   1650
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Add/Edit Operation"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5220
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtQdy 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1650
      TabIndex        =   4
      Tag             =   "1"
      Top             =   1890
      Width           =   825
   End
   Begin VB.TextBox txtMdy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3420
      TabIndex        =   5
      Tag             =   "1"
      Top             =   1890
      Width           =   825
   End
   Begin VB.TextBox txtSet 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1650
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2250
      Width           =   825
   End
   Begin VB.TextBox txtUnt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3420
      TabIndex        =   8
      Tag             =   "1"
      Top             =   2250
      Width           =   825
   End
   Begin VB.TextBox txtCmt 
      Enabled         =   0   'False
      Height          =   1545
      Left            =   1650
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Tag             =   "9"
      Text            =   "RoutRTe05a.frx":07AE
      ToolTipText     =   "Comment (5120 Chars Max)"
      Top             =   2610
      Width           =   4335
   End
   Begin VB.CheckBox optPck 
      Alignment       =   1  'Right Justify
      Caption         =   "Pick Op?"
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   1890
      Width           =   1185
   End
   Begin VB.CheckBox optSrv 
      Alignment       =   1  'Right Justify
      Caption         =   "Service Op?"
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Top             =   2250
      Width           =   1185
   End
   Begin VB.ComboBox cmbPrt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Tag             =   "3"
      Top             =   4230
      Width           =   3345
   End
   Begin VB.ComboBox cmbShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1650
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select Shop From List"
      Top             =   360
      Width           =   1815
   End
   Begin VB.ComboBox cmbWcn 
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1650
      TabIndex        =   3
      Tag             =   "8"
      ToolTipText     =   "Select Work Center From List"
      Top             =   1530
      Width           =   1815
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   4320
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4770
      FormDesignWidth =   6180
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   285
      Index           =   7
      Left            =   180
      TabIndex        =   22
      Top             =   2640
      Width           =   1185
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   21
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Part"
      Height          =   285
      Index           =   3
      Left            =   270
      TabIndex        =   20
      Top             =   4230
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Revise Operation Name"
      Height          =   375
      Index           =   2
      Left            =   180
      TabIndex        =   19
      Top             =   630
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Work Center"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   18
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Queue"
      Height          =   285
      Index           =   5
      Left            =   180
      TabIndex        =   16
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Move"
      Height          =   285
      Index           =   6
      Left            =   2790
      TabIndex        =   15
      Top             =   1890
      Width           =   645
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup"
      Height          =   285
      Index           =   8
      Left            =   180
      TabIndex        =   14
      Top             =   2250
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   285
      Index           =   9
      Left            =   2790
      TabIndex        =   13
      Top             =   2250
      Width           =   645
   End
End
Attribute VB_Name = "RoutRTe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/2/04 see Query_Unload. Added trap for blank Work Center
'12/13/05 Added ClearBoxes
Option Explicit
'Dim RdoStm As rdoQuery
Dim AdoCmdStm As ADODB.Command
Dim RdoLib As ADODB.Recordset

Dim bGoodOp As Byte
Dim bNewOp As Byte
Dim bOnLoad As Byte

Dim iService As Integer

Dim sOldCenter As String
Dim sOldOp As String
Dim sOldShop As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbOpr_LostFocus()
   cmbOpr = CheckLen(cmbOpr, 12)
   If Len(cmbOpr) = 0 Then
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   End If
   bGoodOp = GetOperation(True)
   If bGoodOp = 0 Then AddOperation
   
End Sub

Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBSERVPART = Compress(cmbPrt)
         .Update
      End With
   End If
   
End Sub

Private Sub cmbShp_Click()
   FillCenters
   FillOperations
   
End Sub

Private Sub cmbShp_LostFocus()
   cmbShp = CheckLen(cmbShp, 12)
   
End Sub

Private Sub cmbWcn_Click()
   sOldCenter = cmbWcn
   
End Sub

Private Sub cmbWcn_LostFocus()
   cmbWcn = CheckLen(cmbWcn, 12)
   sOldCenter = cmbWcn
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBCENTER = Compress(cmbWcn)
         .Update
      End With
   End If
   
End Sub

Private Sub cmdCan_Click()
   If Trim(cmbWcn) = "" Then
      If cmbWcn.ListCount > 0 Then
         MsgBox "You Must Include A Valid Work Center.", _
            vbExclamation, Caption
      Else
         MsgBox "There Is Not A Valid Work Center For The" & vbCrLf _
            & "Selected Shop. The Library Entry Will Not" & vbCrLf _
            & "Be Saved..", vbExclamation, Caption
         Form_Deactivate
      End If
   Else
      Form_Deactivate
   End If
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3105
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      ES_TimeFormat = GetTimeFormat()
      FillShops
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
   
End Sub


Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bNewOp = 0
   bGoodOp = 0
   
   cUR.CurrentShop = GetSetting("Esi2000", "Current", "Shop", cUR.CurrentShop)
   sSql = "SELECT * FROM RlbrTable WHERE LIBREF= ? AND LIBSHOP= ? "
   
   Set AdoCmdStm = New ADODB.Command
   AdoCmdStm.CommandText = sSql
   
   Dim prmLibRef As ADODB.Parameter
   Set prmLibRef = New ADODB.Parameter
   prmLibRef.Type = adChar
   prmLibRef.Size = 12
   AdoCmdStm.Parameters.Append prmLibRef
   
   Dim prmLibShop As ADODB.Parameter
   Set prmLibShop = New ADODB.Parameter
   prmLibShop.Type = adChar
   prmLibShop.Size = 12
   AdoCmdStm.Parameters.Append prmLibShop
   
   'Set RdoStm = RdoCon.CreateQuery("", sSql)
   'TODO RdoStm.MaxRows = 1
   bOnLoad = 1
   
End Sub


'11/2/04 Added to prevent no Work Center

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   sSql = "DELETE FROM RlbrTable WHERE LIBREF='" & Compress(cmbOpr) _
          & "' AND LIBCENTER=''"
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   RdoLib.Close
   Set AdoCmdStm = Nothing
   Set RdoLib = Nothing
   FormUnload
   Set RoutRTe05a = Nothing
   
End Sub


Private Sub optPck_Click()
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBPICKOP = optPck.value
         .Update
      End With
   End If
   
End Sub

Private Sub optPck_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub optSrv_Click()
   If optSrv.value = 1 Then
      cmbPrt.Enabled = True
   Else
      cmbPrt = "NONE"
      cmbPrt.Enabled = False
   End If
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBSERVICE = optSrv.value
         If optSrv.value = vbChecked Then
            !LIBSERVPART = Compress(cmbPrt)
         Else
            !LIBSERVPART = ""
         End If
         .Update
      End With
   End If
   
End Sub

Private Sub optSrv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 5120)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBCOMT = txtCmt
         .Update
      End With
   End If
   
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 30)
   txtDsc = StrCase(txtDsc)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBDESC = txtDsc
         .Update
      End With
      bNewOp = 0
   End If
   
End Sub

Private Sub txtMdy_LostFocus()
   txtMdy = CheckLen(txtMdy, 7)
   txtMdy = Format(Abs(Val(txtMdy)), ES_QuantityDataFormat)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBMHRS = Val(txtMdy)
         .Update
      End With
   End If
   
End Sub

Private Sub txtQdy_LostFocus()
   txtQdy = CheckLen(txtQdy, 7)
   txtQdy = Format(Abs(Val(txtQdy)), ES_QuantityDataFormat)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBQHRS = Val(txtQdy)
         .Update
      End With
      bNewOp = 0
   End If
   
End Sub

Private Sub txtSet_LostFocus()
   txtSet = CheckLen(txtSet, 7)
   txtSet = Format(Abs(Val(txtSet)), ES_QuantityDataFormat)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBSETUP = Val(txtSet)
         .Update
      End With
   End If
   
End Sub

Private Sub txtUnt_LostFocus()
   txtUnt = CheckLen(txtUnt, 8)
   txtUnt = Format(Abs(Val(txtUnt)), ES_TimeFormat)
   If bGoodOp Then
      With RdoLib
         '.Edit
         !LIBUNIT = Val(txtUnt)
         .Update
      End With
   End If
   
End Sub



Private Sub FillShops()
   On Error GoTo DiaErr1
   sSql = "Qry_FillShops"
   LoadComboBox cmbShp
   If bSqlRows Then
      If Len(cUR.CurrentShop) <> 0 Then
         cmbShp = cUR.CurrentShop
      Else
         cmbShp = cmbShp.List(0)
      End If
   End If
   AddComboStr cmbPrt.hwnd, "NONE"
   sSql = "Qry_FillRoutingPT7"
   LoadComboBox cmbPrt
   cmbPrt = "NONE"
   FillCenters
   FillOperations
   Exit Sub
   
DiaErr1:
   sProcName = "fillshops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillOperations()
   On Error GoTo DiaErr1
   cmbOpr.Clear
   sSql = "SELECT LIBREF,LIBNUM FROM RlbrTable WHERE LIBSHOP='" & Compress(cmbShp) & "'"
   LoadComboBox cmbOpr
   bGoodOp = False
   If cmbOpr.ListCount > 0 Then
      cmbOpr = cmbOpr.List(0)
      bNewOp = 0
      GetOperation (False)
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "filloper"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddOperation()
   Dim sCurrShop As String
   Dim sCurrOp As String
   Dim bResponse As Byte
   
   bResponse = MsgBox(cmbOpr & " Wasn't Found. Add It?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodOp = False
      On Error Resume Next
      Width = Width + 10
      cmbOpr = ""
      cmbOpr.SetFocus
      Exit Sub
   End If
   MouseCursor 11
   sCurrShop = Compress(cmbShp)
   sCurrOp = Compress(cmbOpr)
   
   sSql = "INSERT INTO RlbrTable (LIBREF,LIBNUM,LIBSHOP) " _
          & "VALUES('" & sCurrOp & "','" & Trim(cmbOpr) & "'," _
          & "'" & sCurrShop & "')"
   On Error GoTo RlibrAdd1
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   MouseCursor 0
   cmbOpr.AddItem cmbOpr
   AddComboStr cmbOpr.hwnd, cmbOpr
   SysMsg cmbOpr & " Added To Library.", True, Me
   bNewOp = 1
   bGoodOp = GetOperation(True)
   Exit Sub
   
RlibrAdd1:
   CurrError.Description = Err.Description
   Resume RlibrAdd2
RlibrAdd2:
   MouseCursor 0
   MsgBox CurrError.Description & vbCrLf & "Couldn't Add To Library.", vbExclamation, Caption
   
End Sub



Private Function GetOperation(bOpen As Byte) As Byte
   Dim sOperation As String
   Dim sThisShop As String
   
   sThisShop = Compress(cmbShp)
   sOperation = Compress(cmbOpr)
   GetOperation = 0
   On Error GoTo RlibGo1
   sSql = "SELECT * FROM RlbrTable WHERE LIBREF='" & sOperation & "' " _
          & "AND LIBSHOP='" & sThisShop & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLib, ES_KEYSET)
   If bSqlRows Then
      With RdoLib
         sOldShop = "" & Trim(!LIBSHOP)
         sOldOp = "" & Trim(!LIBREF)
         sOldCenter = "" & Trim(!LIBCENTER)
         cmbOpr = "" & Trim(!LIBNUM)
         txtDsc = "" & Trim(!LIBDESC)
         cmbWcn = "" & Trim(!LIBCENTER)
         txtQdy = Format(!LIBQHRS, ES_QuantityDataFormat)
         txtMdy = Format(!LIBMHRS, ES_QuantityDataFormat)
         txtSet = Format(!LIBSETUP, ES_QuantityDataFormat)
         txtUnt = Format(!LIBUNIT, ES_TimeFormat)
         txtCmt = "" & Trim(!LIBCOMT)
         optSrv.value = !LIBSERVICE
         optPck.value = !LIBPICKOP
         cmbPrt = "" & Trim(!LIBSERVPART)
         If cmbPrt = "" Then cmbPrt = "NONE"
         GetOperation = 1
      End With
      If bOpen Then
         On Error Resume Next
         cmbWcn.Enabled = True
         txtQdy.Enabled = True
         txtMdy.Enabled = True
         txtSet.Enabled = True
         txtUnt.Enabled = True
         txtCmt.Enabled = True
         optSrv.Enabled = True
         optPck.Enabled = True
         If optSrv.value = vbChecked Then cmbPrt.Enabled = True
         txtDsc.SetFocus
      Else
         cmbWcn.Enabled = False
         txtQdy.Enabled = False
         txtMdy.Enabled = False
         txtSet.Enabled = False
         txtUnt.Enabled = False
         txtCmt.Enabled = False
         optSrv.Enabled = False
         optPck.Enabled = False
         cmbPrt.Enabled = False
      End If
      GetCenter
   Else
      GetOperation = 0
      sOldShop = ""
      sOldOp = ""
      sOldCenter = ""
   End If
   On Error Resume Next
   Exit Function
   
RlibGo1:
   Resume RlibGo2
RlibGo2:
   GetOperation = False
   
End Function

Private Sub FillCenters()
   ClearBoxes
   cmbWcn.Clear
   On Error GoTo DiaErr1
   sSql = "Qry_FillWorkCenters '" & Compress(cmbShp) & "'"
   LoadComboBox cmbWcn
   If cmbWcn.ListCount > 0 Then cmbWcn = cmbWcn.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcenters"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetCenter()
   Dim RdoCnt As ADODB.Recordset
   Dim sThisCenter As String
   Dim cTime As Currency
   On Error GoTo DiaErr1
   sThisCenter = Compress(cmbWcn)
   sSql = "SELECT * FROM WcntTable WHERE WCNREF='" & sThisCenter & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCnt, ES_STATIC)
   If bSqlRows Then
      With RdoCnt
         cmbWcn = "" & Trim(!WCNNUM)
         If bNewOp = 1 Then
            cTime = !WCNQHRS + !WCNMHRS + !WCNSUHRS + !WCNUNITHRS
            If cTime = 0 Then
               txtQdy = Format(!WCNQHRS, ES_QuantityDataFormat)
               txtMdy = Format(!WCNMHRS, ES_QuantityDataFormat)
               txtSet = Format(!WCNSUHRS, ES_QuantityDataFormat)
               txtUnt = Format(!WCNUNITHRS, ES_TimeFormat)
            End If
            If !WCNSERVICE = 1 Then
               optSrv.value = 1
               cmbPrt.Enabled = True
            Else
               optSrv.value = 0
               cmbPrt.Enabled = False
            End If
         End If
         ClearResultSet RdoCnt
      End With
   End If
   Set RdoCnt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcenter"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Public Sub ClearBoxes()
   txtDsc = ""
   txtQdy = "0.000"
   txtMdy = "0.000"
   txtSet = "0.000"
   txtUnt = Format(0, ES_TimeFormat)
   txtMdy = "0.000"
   txtCmt = ""
   optPck.value = vbUnchecked
   optSrv.value = vbUnchecked
   cmbPrt = ""
   
End Sub
