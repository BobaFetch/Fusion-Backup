VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inspection Report"
   ClientHeight    =   5520
   ClientLeft      =   1845
   ClientTop       =   885
   ClientWidth     =   7530
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InspRTe01b.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe01b.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton optPrn 
      DownPicture     =   "InspRTe01b.frx":0AB8
      Height          =   320
      Left            =   6120
      Picture         =   "InspRTe01b.frx":104A
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      ToolTipText     =   "Print/Display Report"
      Top             =   480
      Width           =   350
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New"
      Height          =   255
      Left            =   5400
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDwg 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   17
      Tag             =   "3"
      Top             =   4200
      Width           =   2925
   End
   Begin VB.TextBox txtDwg 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   16
      Tag             =   "3"
      Top             =   3840
      Width           =   2925
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Tag             =   "3"
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox cmbDiv 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Select A Division (2 char)"
      Top             =   960
      Width           =   860
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "4"
      Top             =   555
      Width           =   1095
   End
   Begin VB.ComboBox cmbVnd 
      Height          =   315
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   8
      Tag             =   "3"
      ToolTipText     =   "Select Vendor From List"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox cmbIns 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "8"
      Top             =   600
      Width           =   1665
   End
   Begin VB.CheckBox optItm 
      Caption         =   "Items Open"
      Height          =   255
      Left            =   4080
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InspRTe01b.frx":15DC
      Left            =   1200
      List            =   "InspRTe01b.frx":15DE
      TabIndex        =   19
      Tag             =   "8"
      ToolTipText     =   "Select Item From List"
      Top             =   4920
      Width           =   1665
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6720
      Top             =   4680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5520
      FormDesignWidth =   7530
   End
   Begin VB.ComboBox txtShp 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5640
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "8"
      ToolTipText     =   "Select Shop From List"
      Top             =   2010
      Width           =   1815
   End
   Begin VB.TextBox txtRun 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   1320
      Width           =   1005
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Height          =   315
      Left            =   6600
      TabIndex        =   36
      ToolTipText     =   "Open Report Items For Revisions/Additions"
      Top             =   480
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.TextBox txtRel 
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Tag             =   "3"
      Top             =   4560
      Width           =   2925
   End
   Begin VB.TextBox txtDwg 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   15
      Tag             =   "3"
      Top             =   3480
      Width           =   2925
   End
   Begin VB.TextBox txtRid 
      Height          =   285
      Left            =   5640
      TabIndex        =   6
      Tag             =   "3"
      Top             =   1650
      Width           =   1095
   End
   Begin VB.TextBox txtSer 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Tag             =   "3"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtPon 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Tag             =   "3"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtAcc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   12
      Tag             =   "1"
      Top             =   2760
      Width           =   915
   End
   Begin VB.TextBox txtRej 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Tag             =   "1"
      Top             =   2760
      Width           =   915
   End
   Begin VB.TextBox txtRec 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Tag             =   "1"
      Top             =   2760
      Width           =   915
   End
   Begin VB.TextBox txtMdl 
      Height          =   285
      Left            =   5400
      TabIndex        =   20
      Tag             =   "3"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   7
      Tag             =   "3"
      Text            =   " "
      ToolTipText     =   "Select Customer From List"
      Top             =   2040
      Width           =   1555
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1320
      Width           =   3275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(3)"
      Height          =   285
      Index           =   20
      Left            =   4200
      TabIndex        =   48
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(2)"
      Height          =   285
      Index           =   19
      Left            =   4200
      TabIndex        =   47
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(1)"
      Height          =   285
      Index           =   18
      Left            =   4200
      TabIndex        =   46
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Tag"
      Height          =   285
      Index           =   17
      Left            =   2400
      TabIndex        =   45
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   285
      Index           =   16
      Left            =   120
      TabIndex        =   44
      Top             =   960
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Type"
      Height          =   285
      Index           =   15
      Left            =   60
      TabIndex        =   42
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run "
      Height          =   285
      Index           =   14
      Left            =   4680
      TabIndex        =   41
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector"
      Height          =   285
      Index           =   13
      Left            =   2400
      TabIndex        =   40
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblType 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   39
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label lblTag 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   1200
      TabIndex        =   38
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   37
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Related Part"
      Height          =   285
      Index           =   12
      Left            =   60
      TabIndex        =   34
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref Drawings/ Documents:"
      Height          =   645
      Index           =   11
      Left            =   60
      TabIndex        =   33
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rec ID"
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   32
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   285
      Index           =   9
      Left            =   60
      TabIndex        =   31
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Serial No"
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   30
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Accpt"
      Height          =   285
      Index           =   7
      Left            =   4680
      TabIndex        =   29
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Discrepant"
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   28
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Received"
      Height          =   285
      Index           =   5
      Left            =   60
      TabIndex        =   27
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
      Height          =   285
      Index           =   4
      Left            =   4560
      TabIndex        =   26
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shop"
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   25
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   555
      Width           =   1095
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "InspRTe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/18/06 Revisited - no changes
Option Explicit
Dim RdoTag As ADODB.Recordset

Dim bGoodIns As Byte
Dim bGoodPart As Byte
Dim bGoodTag As Byte
Dim bOnLoad As Byte
Dim bPrint As Byte

Dim cAccepted As Currency

Dim sOldPart As String
Dim sUsedOn As String

Dim sInspectors() As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Sub cmbCst_Click()
   GetCust
   
End Sub


Private Sub cmbCst_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbCst_LostFocus()
   cmbCst = CheckLen(cmbCst, 12)
   GetCust
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJCUST = "" & Compress(cmbCst)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub cmbDiv_LostFocus()
   cmbDiv = CheckLen(cmbDiv, 4)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJDIVISION = "" & Compress(cmbDiv)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbIns_Click()
   On Error Resume Next
   If cmbIns.ListIndex > -1 Then
      cmbIns.ToolTipText = sInspectors(cmbIns.ListIndex)
   End If
   
End Sub

Private Sub cmbIns_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbIns_LostFocus()
   cmbIns = CheckLen(cmbIns, 12)
   On Error GoTo DiaErr1
   If cmbIns.ListCount > 0 Then
      If Len(Trim(cmbIns)) = 0 Then cmbIns = cmbIns.List(0)
      If cmbIns.ListIndex > -1 Then
         cmbIns.ToolTipText = sInspectors(cmbIns.ListIndex)
         bGoodIns = 1
      Else
         bGoodIns = 0
      End If
   Else
      MsgBox "There Are No Inspectors.", _
         vbExclamation, Caption
      bGoodIns = 0
   End If
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJINSP = "" & Compress(cmbIns)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   Exit Sub
   
DiaErr1:
   bGoodIns = 0
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   bGoodPart = GetPart()
   If bGoodTag Then
      If bGoodPart = 1 Then
         On Error Resume Next
         RdoTag!REJPART = "" & Compress(cmbPrt)
         RdoTag.Update
         If Err > 0 Then ValidateEdit
      End If
   End If
   
End Sub

Private Sub cmbTyp_Click()
   Dim b As Byte
   Dim iList As Integer
   For iList = 0 To cmbTyp.ListCount - 1
      If cmbTyp = cmbTyp.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      cmbTyp = cmbTyp.List(0)
   End If
   sUsedOn = Left(cmbTyp, 1)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJUSEDON = sUsedOn
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub cmbTyp_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbTyp_LostFocus()
   sUsedOn = Left(cmbTyp, 1)
   If Trim(cmbTyp) = "" Then cmbTyp = cmbTyp.List(0)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJUSEDON = sUsedOn
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbVnd_Click()
   GetCust
   
End Sub


Private Sub cmbVnd_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub cmbVnd_LostFocus()
   cmbVnd = CheckLen(cmbVnd, 12)
   GetCust
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJVENDOR = "" & Compress(cmbVnd)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmdCan_Click()
   If optItm.Value = vbChecked Then
      If ES_CUSTOM = "JEVCO" Then
         Unload jevRTe01c
      Else
         Unload InspRTe01c
      End If
   End If
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCan_Click
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   MouseCursor 13
   optItm.Value = vbChecked
   If ES_CUSTOM = "JEVCO" Then
      jevRTe01c.lblTag = lblTag
      jevRTe01c.Show
   Else
      InspRTe01c.lblTag = lblTag
      InspRTe01c.Show
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If optItm.Value = vbChecked Then Unload InspRTe01c
   If Left(lblType, 1) = "V" Then z1(1) = "Vendor" Else z1(1) = "Customer"
   If bOnLoad Then
      If UCase(Left(Caption, 3)) = "NEW" Then optNew.Value = vbChecked
      FillDivisions
      FillShops
      FillRtCustomers
      FillRTParts
      cmbDiv = GetSetting("Esi2000", "Quality", "LastDiv", cmbDiv)
      txtShp = GetSetting("Esi2000", "Quality", "LastShop", txtShp)
      bOnLoad = 0
   End If
   If bGoodTag = 0 Then bGoodTag = GetTag()
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   cmbTyp = "Detail"
   AddComboStr cmbTyp.hwnd, "Detail"
   AddComboStr cmbTyp.hwnd, "Assembly"
   AddComboStr cmbTyp.hwnd, "Processing"
   AddComboStr cmbTyp.hwnd, "Installation"
   AddComboStr cmbTyp.hwnd, "Material"
   AddComboStr cmbTyp.hwnd, "Hardware"
   bOnLoad = 1
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bGoodTag = 1 And bDataHasChanged Then UpdateThisTag
   If Len(Trim(cmbDiv)) > 0 Then SaveSetting "Esi2000", "Quality", "LastDiv", cmbDiv
   If Len(Trim(txtShp)) > 0 Then SaveSetting "Esi2000", "Quality", "LastShop", txtShp
   If Len(cmbIns) Then bGoodIns = 1
   If bGoodIns = 0 Then
      MsgBox "This Tag Requires A Valid Inspector. " & vbCr _
         & "Reporting On This Tag May Not Be Accurate", _
         vbExclamation, Caption
   End If
   If bGoodPart = 0 Then
      MsgBox "This Tag Requires A Valid Part Number. " & vbCr _
         & "Reporting On This Tag May Not Be Accurate", _
         vbExclamation, Caption
   End If
   If bPrint = 0 Then
      If optNew.Value = vbChecked Then
         InspRTe01a.Show
         optNew.Value = vbUnchecked
      Else
         InspRTe02a.Show
      End If
   End If
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RdoTag = Nothing
   Set InspRTe01b = Nothing
   
End Sub



Private Sub FillRTParts()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
          & "WHERE PAPRODCODE<>'BID' AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPrt
   Exit Sub
   
DiaErr1:
   sProcName = "fillrtparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillRtCustomers()
   On Error GoTo DiaErr1
   AddComboStr cmbCst.hwnd, "NONE"
   If sLastType = "V" Then
      cmbCst.Visible = False
      cmbVnd.Visible = True
      cmbVnd.Left = cmbCst.Left
      cmbVnd.TabIndex = 7
      txtPon.Width = 915
      z1(1) = "Vendor"
      sSql = "SELECT VEREF,VENICKNAME,VEBNAME FROM VndrTable "
      LoadComboBox cmbVnd
      z1(1).Caption = "Vendor"
      cmbCst.ToolTipText = "Select Vendor From List"
   Else
      txtPon.Width = 1575
      sSql = "Qry_FillSortedCustomers"
      LoadComboBox cmbCst
      z1(1).Caption = "Customer"
      cmbCst.ToolTipText = "Select Customer From List"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillrtcustomers"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub GetCust()
   On Error GoTo DiaErr1
   Dim RdoTyp As ADODB.Recordset
   If sLastType = "V" Then
      If Trim(cmbVnd) = "" Or Trim(cmbVnd) = "NONE" Then
         lblCst = "Not Assigned"
         Exit Sub
      End If
      sSql = "SELECT VENICKNAME,VEBNAME FROM VndrTable WHERE " _
             & "VEREF='" & Compress(cmbVnd) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
      If bSqlRows Then
         With RdoTyp
            cmbVnd = "" & Trim(!VENICKNAME)
            lblCst = "" & Trim(!VEBNAME)
            ClearResultSet RdoTyp
         End With
      Else
         cmbVnd = "NONE"
         lblCst = "*** Vendor Wasn't Found ***"
      End If
   Else
      If Trim(cmbCst) = "" Or Trim(cmbCst) = "NONE" Then
         lblCst = "Not Assigned"
         Exit Sub
      End If
      sSql = "SELECT CUREF,CUNICKNAME,CUNAME FROM CustTable WHERE " _
             & "CUREF='" & Compress(cmbCst) & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoTyp, ES_FORWARD)
      If bSqlRows Then
         With RdoTyp
            cmbCst = "" & Trim(!CUNICKNAME)
            lblCst = "" & Trim(!CUNAME)
            ClearResultSet RdoTyp
         End With
      Else
         cmbCst = "NONE"
         lblCst = "*** Customer Wasn't Found ***"
      End If
   End If
   Set RdoTyp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcust"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PARTNUM,PADESC FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblPrt = "" & Trim(!PADESC)
         GetPart = 1
         ClearResultSet RdoPrt
      End With
   Else
      GetPart = 0
      If Len(Trim(cmbPrt)) > 0 Then
         lblPrt = "*** Part Number Wasn't Found ***"
      Else
         lblPrt = "*** No Valid Part Number Selected ***"
      End If
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   GetPart = 0
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function



Private Sub lblCst_Change()
   If Left(lblPrt, 4) = "*** " Then
      lblPrt.ForeColor = ES_RED
   Else
      lblPrt.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblPrt_Change()
   If Left(lblPrt, 8) = "*** Part" Or Left(lblPrt, 6) = "*** No" Then
      lblPrt.ForeColor = ES_RED
   Else
      lblPrt.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optItm_Click()
   'never visible - on if items are loaded
   
End Sub

Private Sub optNew_Click()
   'never visible-new tag?
   
End Sub

Private Sub optPrn_Click()
   If optPrn Then
      bPrint = 1
      Load InspRTp01a
      InspRTp01a.optFrm.Value = vbChecked
      InspRTp01a.Show
      optPrn = False
   End If
   
End Sub

Private Sub txtAcc_LostFocus()
   txtAcc = CheckLen(txtAcc, 10)
   txtAcc = Format(txtAcc, ES_QuantityDataFormat)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJACCT = 0 + Val(txtAcc)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJDATE = Format(txtDte, "mm/dd/yy")
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
   
End Sub

Private Sub txtDwg_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub

Private Sub txtDwg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtDwg_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCase KeyAscii
   
End Sub


Private Sub txtDwg_LostFocus(Index As Integer)
   txtDwg(Index) = CheckLen(txtDwg(Index), 30)
   If bGoodTag Then
      On Error Resume Next
      Select Case Index
         Case 1
            RdoTag!REJDRAWING2 = "" & txtDwg(Index)
         Case 2
            RdoTag!REJDRAWING3 = "" & txtDwg(Index)
         Case Else
            RdoTag!REJDRAWING1 = "" & txtDwg(Index)
      End Select
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMdl_LostFocus()
   txtMdl = CheckLen(txtMdl, 12)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJMODEL = "" & txtMdl
      RdoTag!REJUSEDON = Left(cmbTyp, 1)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtPon_LostFocus()
   txtPon = CheckLen(txtPon, 20)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJPON = "" & txtPon
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRec_LostFocus()
   txtRec = CheckLen(txtRec, 10)
   txtRec = Format(txtRec, ES_QuantityDataFormat)
   If Val(txtRej) > Val(txtRec) Then
      txtRej = Format(Val(txtRec), ES_QuantityDataFormat)
   End If
   cAccepted = Val(txtRec) - Val(txtRej)
   txtAcc = Format(cAccepted, ES_QuantityDataFormat)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJREC = 0 + Val(txtRec)
      RdoTag!REJACCT = cAccepted
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRej_LostFocus()
   txtRej = CheckLen(txtRej, 10)
   txtRej = Format(txtRej, ES_QuantityDataFormat)
   If Val(txtRej) > Val(txtRec) Then
      txtRej = Format(Val(txtRec), ES_QuantityDataFormat)
   End If
   cAccepted = Val(txtRec) - Val(txtRej)
   txtAcc = Format(cAccepted, ES_QuantityDataFormat)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJREJ = 0 + Val(txtRej)
      RdoTag!REJACCT = cAccepted
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRel_LostFocus()
   txtRel = CheckLen(txtRel, 30)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJRELPART = "" & txtRel
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRid_LostFocus()
   txtRid = CheckLen(txtRid, 12)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJRECID = "" & txtRid
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtRun_LostFocus()
   txtRun = CheckLen(txtRun, 5)
   txtRun = Format(Val(txtRun), "####0")
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJRUN = 0 + txtRun
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtSer_LostFocus()
   txtSer = CheckLen(txtSer, 20)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJSERNO = "" & txtSer
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub txtShp_Click()
   GetShop
   
End Sub

Private Sub txtShp_DropDown()
   bDataHasChanged = True
   
End Sub


Private Sub txtShp_LostFocus()
   Dim b As Byte
   Dim iList As Integer
   txtShp = CheckLen(txtShp, 12)
   For iList = 0 To txtShp.ListCount - 1
      If txtShp = txtShp.List(iList) Then b = 1
   Next
   If b = 0 Then
      Beep
      If txtShp.ListCount > 0 Then txtShp = txtShp.List(0)
   End If
   GetShop
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJSHOP = "" & Compress(txtShp)
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetTag() As Byte
   Dim sRevTag As String
   sRevTag = Compress(lblTag)
   If bGoodTag = 1 And bDataHasChanged Then UpdateThisTag
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "Select * FROM RjhdTable WHERE REJREF='" & sRevTag & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTag, ES_KEYSET)
   If bSqlRows Then
      With RdoTag
         lblTag = "" & Trim(!REJNUM)
         txtDte = "" & Format(Trim(!REJDATE), "mm/dd/yy")
         cmbCst = "" & Trim(!REJCUST)
         cmbVnd = "" & Trim(!REJVENDOR)
         cmbPrt = "" & Trim(!REJPART)
         If optNew.Value = vbUnchecked Then txtShp = "" & Trim(!REJSHOP)
         txtMdl = "" & Trim(!REJMODEL)
         txtRec = Format(!REJREC, ES_QuantityDataFormat)
         txtRej = Format(!REJREJ, ES_QuantityDataFormat)
         txtAcc = Format(!REJACCT, ES_QuantityDataFormat)
         txtPon = "" & Trim(!REJPON)
         txtSer = "" & Trim(!REJSERNO)
         txtRid = "" & Trim(!REJRECID)
         txtRel = "" & Trim(!REJRELPART)
         txtDwg(0) = "" & Trim(!REJDRAWING1)
         txtDwg(1) = "" & Trim(!REJDRAWING2)
         txtDwg(2) = "" & Trim(!REJDRAWING3)
         txtTag = "" & Trim(!REJCUSTTAG)
         txtPon = "" & Trim(!REJPON)
         cmbIns = "" & Trim(!REJINSP)
         txtRun = Format(0 + !REJRUN)
         If Len(cmbIns) Then bGoodIns = 1
         If optNew.Value = vbUnchecked Then cmbDiv = "" & Trim(!REJDIVISION)
         Select Case Trim(!REJUSEDON)
            Case "D"
               cmbTyp = "Detail"
            Case "A"
               cmbTyp = "Assembly"
            Case "P"
               cmbTyp = "Processing"
            Case "I"
               cmbTyp = "Installation"
            Case "M"
               cmbTyp = "Material"
            Case "H"
               cmbTyp = "Hardware"
         End Select
      End With
      bGoodPart = GetPart()
      GetCust
      GetShop
      GetTag = 1
   Else
      GetTag = 0
   End If
   bDataHasChanged = False
   sOldPart = cmbPrt
   MouseCursor 0
   Exit Function
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillShops()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   iList = -1
   On Error GoTo DiaErr1
   sSql = "SELECT SHPREF,SHPNUM FROM ShopTable "
   LoadComboBox txtShp
   
   sSql = "SELECT INSID,INSFIRST,INSMIDD,INSLAST " _
          & "FROM RinsTable WHERE INSACTIVE=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            iList = iList + 1
            ReDim Preserve sInspectors(iList)
            AddComboStr cmbIns.hwnd, "" & Trim(!INSID)
            sInspectors(iList) = Trim(!INSFIRST) & " " _
                        & Trim(!INSMIDD) & " " & Trim(!INSLAST)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbIns.ListCount > 0 Then
      If Trim(cmbIns) = "" Then
         cmbIns = cmbIns.List(0)
         cmbIns.ToolTipText = sInspectors(0)
      Else
         For iList = 0 To cmbIns.ListCount - 1
            If Trim(cmbIns) = Trim(cmbIns.List(iList)) Then
               cmbIns.ToolTipText = sInspectors(iList)
            End If
         Next
      End If
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillshops"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub txtTag_LostFocus()
   txtTag = CheckLen(txtTag, 12)
   If bGoodTag Then
      On Error Resume Next
      RdoTag!REJCUSTTAG = "" & txtTag
      RdoTag.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub GetShop()
   Dim RdoShop As ADODB.Recordset
   sSql = "SELECT SHPREF,SHPNUM FROM ShopTable " _
          & "WHERE SHPREF='" & Compress(txtShp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoShop, ES_FORWARD)
   If bSqlRows Then
      With RdoShop
         txtShp = "" & Trim(!SHPNUM)
         ClearResultSet RdoShop
      End With
   Else
      Beep
      txtShp = ""
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getshop"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub UpdateThisTag()
   On Error Resume Next
   With RdoTag
      !REJINSP = "" & Compress(cmbIns)
      !REJPART = "" & Compress(cmbPrt)
      !REJCUST = "" & Compress(cmbCst)
      !REJVENDOR = "" & Compress(cmbVnd)
      !REJSHOP = "" & Compress(txtShp)
      .Update
   End With
   bDataHasChanged = False
End Sub
