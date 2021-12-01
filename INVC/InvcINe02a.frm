VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Unit of Measure and Weight"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame z2 
      Height          =   30
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   1080
      Width           =   7212
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdLst 
      Caption         =   "<< &Last    "
      Enabled         =   0   'False
      Height          =   315
      Left            =   5580
      TabIndex        =   31
      ToolTipText     =   "Last Page (Page Up)"
      Top             =   3360
      Width           =   875
   End
   Begin VB.CommandButton cmdNxt 
      Caption         =   " &Next >>"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6480
      TabIndex        =   30
      ToolTipText     =   "Next Page (Page Down)"
      Top             =   3360
      Width           =   875
   End
   Begin VB.CommandButton cmbSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "Blank Or Leading Characters (Fills Up To 300 Part Numbers >= The Entry)"
      Top             =   640
      Width           =   735
   End
   Begin VB.ComboBox cmbLvl 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InvcINe02a.frx":07AE
      Left            =   1320
      List            =   "InvcINe02a.frx":07B0
      TabIndex        =   0
      Tag             =   "8"
      ToolTipText     =   "Select From List"
      Top             =   260
      Width           =   1335
   End
   Begin VB.ComboBox cmbPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Blank Or Leading Characters (Fills Up To 300 Part Numbers >= The Entry)"
      Top             =   640
      Width           =   3255
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4680
      Top             =   3840
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3780
      FormDesignWidth =   7440
   End
   Begin VB.TextBox txtPun 
      Height          =   285
      Left            =   6360
      TabIndex        =   7
      Tag             =   "3"
      ToolTipText     =   "Purchasing Units"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Purchasing Conversion"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtWht 
      Height          =   285
      Left            =   6360
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtLoc 
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Inventory Location"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtUom 
      Height          =   285
      Left            =   6360
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Units (EA,OZ,etc)"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   29
      Top             =   280
      Width           =   1335
   End
   Begin VB.Label z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   255
      Index           =   9
      Left            =   6720
      TabIndex        =   28
      Top             =   660
      Width           =   615
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6240
      TabIndex        =   27
      Top             =   645
      Width           =   510
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   26
      Top             =   660
      Width           =   615
   End
   Begin VB.Label LblPun 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   25
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purc Units"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblPrc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purc Conv"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label z1 
      Caption         =   "Type "
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
      Index           =   5
      Left            =   3240
      TabIndex        =   21
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                          "
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
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3240
      TabIndex        =   18
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label lblWht 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblLoc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   15
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Extended Part Description"
      Top             =   2010
      Width           =   4005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1725
      Width           =   3015
   End
   Begin VB.Label lblPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit of Meas"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number(s)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   1455
   End
End
Attribute VB_Name = "InvcINe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/8/05 Added ToolTips and changed to << Last Next >> buttons
'        Added the missing Class Tags
Option Explicit
Dim AdoPrt As ADODB.Command
Dim AdoParameter As ADODB.Parameter

Dim bOnLoad As Byte
Dim bUnLoad As Byte

Dim iLevel As Integer
Dim iIndex As Integer
Dim iTotalParts As Integer

Dim sOldPart As String
Dim Parts(302, 2) As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbLvl_Click()
   If Not bOnLoad Then sOldPart = ""
   
End Sub


Private Sub cmbLvl_LostFocus()
   If Trim(cmbLvl) = "" Then cmbLvl = cmbLvl.List(0)
   iLevel = cmbLvl.ListIndex
   If iLevel < 0 Then iLevel = 0
   
End Sub


Private Sub cmbPrt_LostFocus()
   
   cmbPrt = CheckLen(cmbPrt, 30)
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   
End Sub


Private Sub cmbSel_Click()
   If Not bUnLoad Then _
      If sOldPart <> cmbPrt Then FillParts
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bUnLoad = True
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext ("5102")
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdLst_Click()
   iIndex = iIndex - 1
   If iIndex < 1 Then iIndex = 1
   FillItems (iIndex)
   
End Sub

Private Sub cmdNxt_Click()
   iIndex = iIndex + 1
   If iIndex > iTotalParts Then iIndex = iTotalParts
   FillItems (iIndex)
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      GetBeginningParts
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALOCATION," _
          & "PAWEIGHT,PALEVEL,PAPUNITS,PAPURCONV,PAEXTDESC " _
          & "FROM PartTable WHERE PAINACTIVE = 0 AND PAOBSOLETE = 0 AND PARTREF = ? "
   
   Set AdoPrt = New ADODB.Command
   AdoPrt.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.Size = 30
   
   AdoPrt.Parameters.Append AdoParameter
   
   'Set RdoPrt = RdoCon.CreateQuery("", sSql)
   'RdoPrt.MaxRows = 1
   sOldPart = "0"
   cmbLvl.AddItem "ALL"
   cmbLvl.AddItem "1 - Top"
   cmbLvl.AddItem "2 - Mid"
   cmbLvl.AddItem "3 - Base"
   cmbLvl.AddItem "4 - Raw"
   cmbLvl.AddItem "7 - Service"
   cmbLvl.AddItem "8 - Project"
   cmbLvl = cmbLvl.List(0)
   bOnLoad = 1
   bUnLoad = False
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   FormUnload
   Set AdoParameter = Nothing
   Set AdoPrt = Nothing
   Set InvcINe02a = Nothing
   
End Sub



Private Sub FillParts()
   Dim RdoCmb As ADODB.Recordset
   Dim bLen As Byte
   Dim iRow As Integer
   Dim sPartNumber As String
   
   MouseCursor 13
   sPartNumber = Compress(cmbPrt)
   bLen = Len(sPartNumber) + 1
   cmbPrt.Clear
   Erase Parts
   iTotalParts = 0
   cmdLst.Enabled = False
   cmdNxt.Enabled = False
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "(LEFT(PARTREF," & bLen & ")>= '" & sPartNumber & "' "
   If iLevel > 0 Then sSql = sSql & "AND PALEVEL=" & iLevel & " "
   sSql = sSql & "AND PATOOL= 0) ORDER BY PARTREF "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         If Len(cmbPrt) = 0 Then cmbPrt = sPartNumber
         sOldPart = cmbPrt
         cmbPrt.AddItem sOldPart
         Do Until .EOF
            iRow = iRow + 1
            Parts(iRow, 0) = "" & Trim(!PartRef)
            Parts(iRow, 1) = "" & Trim(!PartNum)
            AddComboStr cmbPrt.hWnd, "" & Trim(!PartNum)
            .MoveNext
            If iRow > 300 Then Exit Do
         Loop
         ClearResultSet RdoCmb
      End With
      iIndex = 1
      iTotalParts = iRow
      lblNum = Format(iTotalParts, "##0")
   Else
      MouseCursor 0
      MsgBox "No Matching Items Found.", vbInformation, Caption
      On Error Resume Next
      iTotalParts = 0
      cmbPrt.SetFocus
      Exit Sub
   End If
   If iTotalParts > 0 Then
      cmdLst.Enabled = True
      cmdNxt.Enabled = True
      FillItems (iIndex)
   Else
      MsgBox "No Matching Parts Were Found.", vbInformation, Caption
      iIndex = 0
   End If
   MouseCursor 0
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub FillItems(iCurrIndex)
   Dim RdoItm As ADODB.Recordset
   On Error GoTo DiaErr1
   If iCurrIndex = 0 Then Exit Sub
   cmbPrt = Parts(iCurrIndex, 1)
   'RdoPrt.RowsetSize = 1
   'RdoPrt(0) = Parts(iCurrIndex, 0)
   AdoPrt.Parameters(0) = Parts(iCurrIndex, 0)
   bSqlRows = clsADOCon.GetQuerySet(RdoItm, AdoPrt, ES_FORWARD, False, 1)
   If bSqlRows Then
      With RdoItm
         sOldPart = cmbPrt
         lblPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblExt = "" & Trim(!PAEXTDESC)
         lblUom = "" & Trim(!PAUNITS)
         lblLoc = "" & Trim(!PALOCATION)
         lblWht = Format(!PAWEIGHT, ES_QuantityDataFormat)
         lblPrc = Format(!PAPURCONV, ES_QuantityDataFormat)
         LblPun = "" & Trim(!PAPUNITS)
         txtUom = "" & Trim(!PAUNITS)
         txtLoc = "" & Trim(!PALOCATION)
         txtWht = Format(!PAWEIGHT, ES_QuantityDataFormat)
         txtPrc = Format(!PAPURCONV, ES_QuantityDataFormat)
         txtPun = "" & Trim(!PAPUNITS)
         lblTyp = str(!PALEVEL)
         ClearResultSet RdoItm
      End With
   Else
      lblPrt = ""
      lblDsc = ""
      lblExt = ""
      lblUom = ""
      lblLoc = ""
      lblWht = ""
      lblPrc = ""
      LblPun = ""
      txtUom = ""
      txtLoc = ""
      txtWht = ""
      txtPrc = ""
      txtPun = ""
      lblTyp = ""
   End If
   Set RdoItm = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtLoc_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtLoc_LostFocus()
   txtLoc = CheckLen(txtLoc, 4)
   On Error GoTo txtLocErr1
   If iIndex > 0 Then
      sSql = "UPDATE PartTable SET PALOCATION='" & txtLoc & "' " _
             & "WHERE PARTREF='" & Parts(iIndex, 0) & "' "
      clsADOCon.ExecuteSQL sSql
      lblLoc = txtLoc
   End If
   Exit Sub
   
txtLocErr1:
   Resume txtLocErr2
txtLocErr2:
   
End Sub

Private Sub txtPrc_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub

Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 10)
   txtPrc = Format(Abs(Val(txtPrc)), ES_QuantityDataFormat)
   On Error GoTo txtPrcErr1
   If iIndex > 0 Then
      sSql = "UPDATE PartTable SET PAPURCONV=" & Format(Val(txtPrc), ES_QuantityDataFormat) & " " _
             & "WHERE PARTREF='" & Parts(iIndex, 0) & "' "
      clsADOCon.ExecuteSQL sSql
      lblPrc = txtPrc
   End If
   Exit Sub
   
txtPrcErr1:
   Resume txtPrcErr2
txtPrcErr2:
   
End Sub






Private Sub txtPun_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtPun_LostFocus()
   txtPun = CheckLen(txtPun, 2)
   If Len(txtPun) = 0 Then txtPun = txtUom
   If Len(txtPun) = 0 Then txtPun = "EA"
   On Error GoTo txtPunErr1
   If iIndex > 0 Then
      sSql = "UPDATE PartTable SET PAPUNITS='" & txtPun & "' " _
             & "WHERE PARTREF='" & Parts(iIndex, 0) & "' "
      clsADOCon.ExecuteSQL sSql
      LblPun = txtPun
   End If
   Exit Sub
   
txtPunErr1:
   Resume txtPunErr2
txtPunErr2:
   
End Sub

Private Sub txtUom_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtUom_LostFocus()
   txtUom = CheckLen(txtUom, 2)
   On Error GoTo txtUomErr1
   If iIndex > 0 Then
      sSql = "UPDATE PartTable SET PAUNITS='" & txtUom & "' " _
             & "WHERE PARTREF='" & Parts(iIndex, 0) & "' "
      clsADOCon.ExecuteSQL sSql
      lblUom = txtUom
   End If
   Exit Sub
   
txtUomErr1:
   Resume txtUomErr2
txtUomErr2:
   
End Sub

Private Sub txtWht_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdLst_Click
   If KeyCode = vbKeyPageDown Then cmdNxt_Click
   
End Sub


Private Sub txtWht_LostFocus()
   txtWht = CheckLen(txtWht, 10)
   txtWht = Format(Abs(Val(txtWht)), ES_QuantityDataFormat)
   On Error GoTo txtWhtErr1
   If iIndex > 0 Then
      sSql = "UPDATE PartTable SET PAWEIGHT=" & Format(txtWht, ES_QuantityDataFormat) & " " _
             & "WHERE PARTREF='" & Parts(iIndex, 0) & "' "
      clsADOCon.ExecuteSQL sSql
      lblWht = txtWht
   End If
   Exit Sub
   
txtWhtErr1:
   Resume txtWhtErr2
txtWhtErr2:
   
End Sub



Private Sub GetBeginningParts()
   Dim RdoBeg As ADODB.Recordset
   sSql = "SELECT PARTREF From PartTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBeg)
   If bSqlRows Then
      With RdoBeg
         cmbPrt = "" & Left(!PartRef, 2)
         sOldPart = cmbPrt
         ClearResultSet RdoBeg
      End With
   End If
   Set RdoBeg = Nothing
   FillParts
   Exit Sub
   
DiaErr1:
   sProcName = "getbegpar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
