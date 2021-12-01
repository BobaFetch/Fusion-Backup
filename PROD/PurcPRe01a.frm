VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form PurcPRe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Purchase Order"
   ClientHeight    =   2940
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6435
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   4301
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbTaxable 
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optReq 
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Open Next New Purchase Order With The Current Request By"
      Top             =   2520
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.ComboBox cmbReq 
      Height          =   288
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Contains Previous Table Entries (20 Char Max) Including Blanks"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdVew 
      DownPicture     =   "PurcPRe01a.frx":07AE
      Height          =   350
      Left            =   3120
      Picture         =   "PurcPRe01a.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Show Existing Purchase Orders"
      Top             =   1080
      Width           =   350
   End
   Begin VB.CheckBox optRev 
      Caption         =   "From Revise"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   5440
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Add This Purchase Order"
      Top             =   480
      Width           =   915
   End
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5760
      Top             =   1680
   End
   Begin VB.CheckBox optSrv 
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Is This A Services PO?"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtPon 
      Height          =   285
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "New Po Number"
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox cmbVnd 
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1555
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5760
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2940
      FormDesignWidth =   6435
   End
   Begin VB.Label Label1 
      Caption         =   "Taxable?"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Repeat Requested By"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Open Next New Purchase Order With The Current Request By"
      Top             =   2520
      Width           =   1800
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requested By"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Contains Previous Table Entries (20 Char Max) Including Blanks"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Purchase Order"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLst 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      ToolTipText     =   "Last Po Entered"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label txtRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   5520
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   3720
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Service PO?"
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Is This A Services PO?"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   252
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "PurcPRe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/14/04 Added Requested By
'11/24/04 Added trap for first PO (GetLastPo)
'1/20/05 Swap to add New Services Purchase Order
'7/15/05 Added COLASTPURCHASEORDER
'5/10/06 Added to Save Last Request By
'5/11/06 Corrected Help Call
Option Explicit


Dim bPoAdded As Byte
Dim bPoExists As Byte
Dim bGoodVnd As Byte
Dim bOnLoad As Byte

Dim cNetDays As Currency
Dim cDDays As Currency
Dim cDiscount As Currency
Dim cProxDt As Currency
Dim cProxdue As Currency

Dim sFob As String
Dim sBcontact As String
Dim sShipTo As String
Dim sBuyer As String
Dim sLastBuyer As String

'Dim FusionToolTip As New clsToolTip


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Sub cbTaxable_KeyPress(KeyAscii As Integer)
    KeyLock KeyAscii
End Sub

Private Sub cmbReq_LostFocus()
   cmbReq = CheckLen(cmbReq, 20)
   On Error Resume Next
   If Len(cmbReq) < 4 Then cmbReq = UCase$(cmbReq) _
          Else cmbReq = StrCase(cmbReq)
   
End Sub


Private Sub cmbVnd_Click()
   bGoodVnd = FindVendor(cmbVnd, lblNme)
   LoadVendorAddress
   
End Sub

Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 10)
   If Len(cmbVnd) > 0 Then
      bGoodVnd = FindVendor(cmbVnd, lblNme)
      
   Else
      bGoodVnd = False
   End If
   
End Sub

Private Sub cmdAdd_Click()
   Dim bResponse As Byte
   Dim strReason As String
   
   'BBS Added this logic for Ticket #25640
   If bGoodVnd Then
        If IsVendorApproved(cmbVnd.Text, 0, strReason) = 0 Then
            If MsgBox(strReason & " Continue?", vbInformation + vbYesNo, "Vendor Approval") = vbNo Then Exit Sub
        End If
    End If
   GetCompany
   If bGoodVnd Then
      If Left$(Caption, 5) = "New S" Then
         If optSrv.Value = vbUnchecked Then
            bResponse = MsgBox("Is This Meant To Be A Services Purchase Order?", _
                        ES_NOQUESTION, Caption)
            If bResponse = vbYes Then optSrv.Value = vbChecked
         End If
      End If
            
      tmr1.Enabled = False
      AddNewPo
   Else
      MsgBox "A Purchase Order Requires A Valid Vendor.", _
         vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   tmr1.Enabled = False
   Unload Me
   
End Sub



Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4301
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub


Private Sub cmdVew_Click()
   Dim iList As Integer
   Dim iCol As Integer
   Dim iRows As Integer
   Dim RdoVew As ADODB.Recordset
   On Error Resume Next
   iRows = 10
   With PurcPOview.Grd
      .Rows = iRows
      .ColAlignment(0) = 0
      .ColAlignment(1) = 0
      If Screen.Width > 9999 Then
         .ColWidth(0) = 1100 * 1.25
         .ColWidth(1) = 1550 * 1.25
         .ColWidth(2) = 1900 * 1.25
      Else
         .ColWidth(0) = 1100
         .ColWidth(1) = 1550
         .ColWidth(2) = 1900
      End If
   End With
   sSql = "SELECT PONUMBER,POVENDOR,PODATE,VEREF,VENICKNAME FROM " _
          & "PohdTable,VndrTable WHERE (POVENDOR=VEREF) " _
          & "ORDER BY PONUMBER DESC"
   MouseCursor 13
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVew)
   If bSqlRows Then
      With RdoVew
         Do Until .EOF
            iList = iList + 1
            iRows = iRows + 1
            PurcPOview.Grd.Rows = iRows
            PurcPOview.Grd.row = iRows - 11
            
            PurcPOview.Grd.Col = 0
            PurcPOview.Grd = "" & Format(!PONUMBER, "00000")
            
            PurcPOview.Grd.Col = 1
            PurcPOview.Grd = "" & Format(!PODATE, "mm/dd/yy")
            
            PurcPOview.Grd.Col = 2
            PurcPOview.Grd = Trim(!VENICKNAME)
            .MoveNext
         Loop
         ClearResultSet RdoVew
      End With
      If iList > 9 Then PurcPOview.Grd.Rows = iList + 1
   End If
   MouseCursor 0
   Set RdoVew = Nothing
   On Error GoTo 0
   PurcPOview.Show
   
End Sub




Private Sub Form_Activate()
   Dim b As Byte
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      GetLastPo True
      FillVendors Me
      FillReqBy
      If cUR.CurrentVendor <> "" Then
         cmbVnd = cUR.CurrentVendor
         LoadVendorAddress
         bGoodVnd = FindVendor(cmbVnd, lblNme)
         
         sJournalID = GetOpenJournal("PJ", Format$(ES_SYSDATE, "mm/dd/yy"))
         If Left(sJournalID, 4) = "None" Then
            sJournalID = ""
            b = 1
         Else
            If sJournalID = "" Then b = 0 Else b = 1
         End If
         If b = 0 Then
            MouseCursor 0
            MsgBox "There Is No Open Purchases Journal For This Period.", _
               vbExclamation, Caption
            Sleep 500
            Unload Me
            Exit Sub
         End If
      End If
      tmr1.Enabled = True
      bOnLoad = 0
   End If
   MouseCursor 0
   Call HookToolTips
   MVBBubble.MaxTipWidth = 600
End Sub

Private Sub Form_Load()
   If bPOCaption = 1 Then Caption = "New Services Purchase Order"
   FormLoad Me
   bPOCaption = 0
   FormatControls
   HelpContextID = 4301
   If iBarOnTop Then
      Move PurcPRe02a.Left + (MDISect.Left + 400), PurcPRe02a.Top + (MDISect.TopBar.Height + 1400)
   Else
      Move PurcPRe02a.Left + (MDISect.SideBar.Width + 400), PurcPRe02a.Top + (1100 + 400)
   End If
'   sSql = "SELECT PONUMBER,POVENDOR,VEREF,VENICKNAME, POTAXABLE FROM PohdTable," _
'          & "VndrTable WHERE POVENDOR=VEREF ORDER BY PONUMBER DESC"
'   Set AdoQry = New ADODB.Command
'   AdoQry.CommandText = sSql
   
   cUR.CurrentVendor = GetSetting("Esi2000", "Current", "Vendor", cUR.CurrentVendor)
   sLastBuyer = GetSetting("Esi2000", "EsiProd", "LastBuyer", sLastBuyer)
   MDISect.Enabled = False
   bOnLoad = 1
'   With FusionToolTip
'        Call .Create(Me)
'        .MaxTipWidth = 240
'        .DelayTime(ttDelayShow) = 20000
'        Call .AddTool(cmbVnd)
'    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "EsiProd", "PRe01a", cmbReq
   SaveSetting "Esi2000", "EsiProd", "PRe01aRb", Trim(str$(optReq))
   MDISect.Enabled = True
   PurcPRe02a.optNwl.Value = vbUnchecked
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   cUR.CurrentVendor = cmbVnd
   SaveCurrentSelections
   If optRev.Value = vbChecked Then
      bPoAdded = 1
      MDISect.lblBotPanel.Caption = "Revise A Purchase Order"
   End If
   If bPoAdded = 0 Then
      FormUnload
      Unload PurcPRe02a
   End If
   On Error Resume Next
   UnhookToolTips
'   Set AdoQry = Nothing
   Set PurcPRe01a = Nothing
 
End Sub






Private Sub optRev_Click()
   'never visible-loaded from PurcPRe02a
   
End Sub


Private Sub optSrv_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub tmr1_Timer()
   GetLastPo False
   
End Sub



Private Sub GetLastPo(bFillText As Boolean)
   Dim RdoCmn As ADODB.Recordset
   Static bNoPos As Byte
   Dim lOldPo As Long
   Static sOldLast As String
   
   On Error GoTo DiaErr1
   sSql = "SELECT COLASTPURCHASEORDER From ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmn, ES_FORWARD)
   If bSqlRows Then
      If RdoCmn!COLASTPURCHASEORDER > 0 Then lOldPo = RdoCmn!COLASTPURCHASEORDER
   End If
   If lOldPo = 0 Then
      sSql = "SELECT MAX(PONUMBER) AS LASTPO FROM PohdTable "
      Set RdoCmn = clsADOCon.GetRecordSet(sSql)
      If Not RdoCmn!LASTPO Then
         If RdoCmn!LASTPO > 0 Then lOldPo = RdoCmn!LASTPO
      Else
         '1/24/04
         lblLst = "000000"
         If bNoPos = 0 Then txtPon = "000001"
         bNoPos = 1
         tmr1.Enabled = False
      End If
   End If
   If lOldPo = 0 Then
      lblLst = "000000"
      If bNoPos = 0 Then txtPon = "000001"
   Else
      lblLst = Format(lOldPo, "000000")
      If bFillText Then
         If sOldLast <> lblLst Then txtPon = Format(lOldPo + 1, "000000")
      End If
   End If
   sOldLast = lblLst
   Set RdoCmn = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getlastpo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub txtPon_LostFocus()
   txtPon = CheckLen(txtPon, 6)
   txtPon = Format(Abs(Val(txtPon)), "000000")
   
End Sub


Private Sub GetVendorTerms()
   Dim RdoTrm As ADODB.Recordset
   Dim sVendRef As String
   sVendRef = Compress(cmbVnd)
   
   On Error GoTo DiaErr1
   sSql = "SELECT VEREF,VENETDAYS,VEDDAYS,VEDISCOUNT," _
          & "VEPROXDT,VEPROXDUE,VEFOB,VEBNAME,VEBADR,VEBCITY," _
          & "VEBSTATE,VEBZIP,VEBCONTACT,VEBUYER FROM VndrTable WHERE " _
          & "VEREF='" & sVendRef & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTrm)
   If bSqlRows Then
      With RdoTrm
         On Error Resume Next
         cNetDays = 0 + !VENETDAYS
         cDDays = 0 + !VEDDAYS
         cDiscount = 0 + !VEDISCOUNT
         cProxDt = 0 + !VEPROXDT
         cProxdue = 0 + !VEPROXDUE
         sFob = "" & Trim(!VEFOB)
         sBcontact = "" & Trim(!VEBCONTACT)
         If sLastBuyer = "" Then
            sBuyer = "" & Trim(!VEBUYER)
         Else
            sBuyer = sLastBuyer
         End If
      End With
   End If
   Set RdoTrm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getvendorterms"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddNewPo()
   Dim RdoNpo As ADODB.Recordset
   Dim sVendRef As String
   Dim dDate As Date
   bPoExists = CheckPo()
   GetShipTo
   If Len(txtPon) = 0 Then Exit Sub
   If Not bGoodVnd Then
      MsgBox "Requires A Valid Vendor.", vbExclamation, Caption
      On Error Resume Next
      cmbVnd.SetFocus
      Exit Sub
   End If
'   If bPoExists Then
'      MsgBox "That Po Number Has Been Used.", vbInformation, Caption
'      On Error Resume Next
'      txtPon.SetFocus
'      Exit Sub
'   End If
   
   'make sure no one is simultaneously creating the same PO number
   Dim po As Long
   po = CLng(txtPon)
   
   ' determine what to do if PO already exists
   Dim GetNextAvailablePo As Boolean
   Dim rs As ADODB.Recordset

   GetNextAvailablePo = False
   sSql = "SELECT GetNextAvailablePoNumber From Preferences "
   bSqlRows = clsADOCon.GetDataSet(sSql, rs, ES_FORWARD)
   If bSqlRows Then
      If rs!GetNextAvailablePoNumber > 0 Then GetNextAvailablePo = True
   End If

   clsADOCon.BeginTrans
   Do While True
      sSql = "select PONUMBER from PohdTable where PONUMBER = " & po
      If clsADOCon.GetDataSet(sSql, rs) = 0 Then Exit Do
      If Not GetNextAvailablePo Then
         MsgBox "That PO number is in use.  Please select another number."
         Exit Sub
      End If
      po = po + 1
      txtPon = po
   Loop

   sVendRef = Compress(cmbVnd)
   GetVendorTerms
   dDate = Format(ES_SYSDATE, "mm/dd/yy")
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM PohdTable"
   Set RdoNpo = clsADOCon.GetRecordSet(sSql, ES_FORWARD)
   RdoNpo.AddNew
   RdoNpo!PONUMBER = Val(txtPon)
   RdoNpo!POVENDOR = "" & sVendRef
   RdoNpo!PODATE = dDate
   RdoNpo!POREQBY = Trim(cmbReq)
   RdoNpo!POSERVICE = optSrv.Value
   RdoNpo!PONETDAYS = cNetDays
   RdoNpo!PODDAYS = cDDays
   RdoNpo!PODISCOUNT = cDiscount
   RdoNpo!POPROXDT = cProxDt
   RdoNpo!POPROXDUE = cProxdue
   RdoNpo!POFOB = "" & sFob
   RdoNpo!POBCONTACT = sBcontact
   RdoNpo!POBUYER = sBuyer
   RdoNpo!POSHIPTO = sShipTo
   If cbTaxable.Value = 1 Then RdoNpo!POTAXABLE = 1 Else RdoNpo!POTAXABLE = 0
   RdoNpo.Update
   sSql = "UPDATE ComnTable SET COLASTPURCHASEORDER=" & Val(txtPon) & " "
   clsADOCon.ExecuteSql sSql
   PurcPRe02a.optNew = vbChecked
   PurcPRe02a.cmbVnd = cmbVnd
   PurcPRe02a.txtFob = sFob
   PurcPRe02a.txtCnt = sBcontact
   PurcPRe02a.txtShp = sShipTo
   PurcPRe02a.cmbPon = txtPon
   If cbTaxable.Value = 1 Then PurcPRe02a.cbTaxable.Value = 1 Else PurcPRe02a.cbTaxable = 0
   PurcPRe02a.cmbPon.AddItem txtPon
   bPoAdded = 1
   sSql = "UPDATE ComnTable SET CURPONUMBER=" & Val(txtPon) & " "
   clsADOCon.ExecuteSql sSql
   clsADOCon.CommitTrans
   Set RdoNpo = Nothing
   Unload Me
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "addnewpo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function CheckPo() As Byte
   Dim RdoPon As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PONUMBER FROM PohdTable WHERE " _
          & "PONUMBER=" & Val(txtPon) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon)
   If bSqlRows Then
      CheckPo = True
   Else
      CheckPo = False
   End If
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkpo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub GetShipTo()
   On Error GoTo DiaErr1
   GetCompany True
   sShipTo = Co.Name & vbCrLf _
             & Co.Addr(1) & vbCrLf _
             & Co.Addr(2) & vbCrLf _
             & Co.Addr(3) & vbCrLf _
             & Co.Addr(4)
   Exit Sub
   
DiaErr1:
   sProcName = "getshipto"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub FillReqBy()
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT POREQBY FROM PohdTable WHERE PODATE> '" _
          & Format(ES_SYSDATE - 360, "mm/dd/yy") & "' ORDER BY POREQBY"
   LoadComboBox cmbReq, -1
   On Error Resume Next
   optReq.Value = Val(GetSetting("Esi2000", "EsiProd", "PRe01aRb", optReq))
   If optReq.Value = vbChecked Then cmbReq = GetSetting("Esi2000", "EsiProd", "PRe01a", cmbReq)
   Exit Sub
   
DiaErr1:
   sProcName = "fillreqby"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub LoadVendorAddress()
    Dim RdoVndr As ADODB.Recordset
    Dim sAddress As String
    sAddress = ""

    sSql = "SELECT VEBADR, VEBCITY, VEBSTATE, VEBZIP,VEBCONTACT, VEBPHONE, VEBEXT FROM VndrTable WHERE VEREF = '" & Compress(cmbVnd) & "'"
    bSqlRows = clsADOCon.GetDataSet(sSql, RdoVndr, ES_FORWARD)
    If bSqlRows Then
        If Len(Trim("" & RdoVndr!VEBCONTACT)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCONTACT) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBADR)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBADR) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBCITY)) > 0 Then sAddress = sAddress & Trim(RdoVndr!VEBCITY) & " ," & Trim(RdoVndr!VEBSTATE) & " " & Trim(RdoVndr!VEBZIP) & vbCrLf
        If Len(Trim("" & RdoVndr!VEBPHONE)) > 0 And Trim("" & RdoVndr!VEBPHONE) <> "___-___-____" Then sAddress = sAddress & "" & Trim(RdoVndr!VEBPHONE)
        If Len(Trim("" & Trim(RdoVndr!VEBEXT))) > 0 And Val("" & Trim(RdoVndr!VEBEXT)) > 0 Then sAddress = sAddress & " Ext: " & Trim(RdoVndr!VEBEXT)
    End If
    Set RdoVndr = Nothing
    If Len(sAddress) > 0 Then
        cmbVnd.ToolTipText = sAddress
'        FusionToolTip.ToolText(cmbVnd) = sAddress
    Else
        cmbVnd.ToolTipText = "Enter the Vendor"
    End If
End Sub
