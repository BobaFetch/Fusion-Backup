VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PickMCe01b 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pick Items (Only Items Not Picked)"
   ClientHeight    =   3585
   ClientLeft      =   1620
   ClientTop       =   750
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cbfrom1a 
      Caption         =   "From form 1a"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "PickMCe01b.frx":0000
      DownPicture     =   "PickMCe01b.frx":04F2
      Enabled         =   0   'False
      Height          =   372
      Left            =   6480
      MaskColor       =   &H00000000&
      Picture         =   "PickMCe01b.frx":09E4
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2640
      Width           =   400
   End
   Begin VB.CommandButton cmdDn 
      DisabledPicture =   "PickMCe01b.frx":0ED6
      DownPicture     =   "PickMCe01b.frx":13C8
      Enabled         =   0   'False
      Height          =   372
      Left            =   6480
      MaskColor       =   &H00000000&
      Picture         =   "PickMCe01b.frx":18BA
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3024
      Width           =   400
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PickMCe01b.frx":1DAC
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtTyp 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   19
      Tag             =   "1"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmbNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Add A New Item To The Pick List"
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   6120
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Remove The Current Item From The List"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtCmt 
      Height          =   1455
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Tag             =   "9"
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ComboBox cmbPrt 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Tag             =   "3"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2280
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3585
      FormDesignWidth =   7080
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   5
      Left            =   4200
      TabIndex        =   21
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   20
      Top             =   240
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Status (PL,PP)"
      Height          =   255
      Index           =   15
      Left            =   1920
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblStat 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3240
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Dsup 
      Height          =   300
      Left            =   360
      Picture         =   "PickMCe01b.frx":255A
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Enup 
      Height          =   300
      Left            =   120
      Picture         =   "PickMCe01b.frx":2A4C
      Top             =   3000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Endn 
      Height          =   300
      Left            =   360
      Picture         =   "PickMCe01b.frx":2F3E
      Top             =   3000
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Dsdn 
      Height          =   300
      Left            =   120
      Picture         =   "PickMCe01b.frx":3430
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uom     "
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
      Left            =   5400
      TabIndex        =   16
      Top             =   960
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity           "
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
      Left            =   4200
      TabIndex        =   15
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                 "
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
      TabIndex        =   14
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rec    "
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
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   840
      TabIndex        =   12
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblRec 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblMon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblRun 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "MO"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "PickMCe01b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/14/02 added PKRECORD for new index
Option Explicit
Dim bOnLoad As Byte

Dim iIndex As Integer
Dim iTotalItems As Integer

Dim sPartNumber As String

Dim vItems(300, 7)

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtTyp.BackColor = BackColor
   
End Sub

Private Sub cmbNew_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim iPkRecord As Integer
   
   If cmbNew.Value = True Then
      sMsg = "Add A New Part To The Pick List?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         Dim tempPart As String
         ' MM 10/4/2009 - Added Compress
         Dim strDate As Variant
         strDate = Format(ES_SYSDATE, "mm/dd/yy")
         tempPart = Compress(cmbPrt)
         cmbPrt.Enabled = True
         cmbPrt.BackColor = vbWindowBackground
         cmbPrt = ""
         lblDsc = "*** Part Number Not Yet Selected ***"
         lblLvl = ""
         txtQty = "0.000"
         txtCmt = ""
         iTotalItems = iTotalItems + 1
         iIndex = iTotalItems
         lblRec = iTotalItems
         iPkRecord = GetNextPickRecord(sPartNumber, Val(lblRun))
         
         
         sSql = "INSERT INTO MopkTable (PKPARTREF,PKMOPART,PKMORUN,PKRECORD," & vbCrLf _
                    & " PKCOMT,PKTYPE,PKPDATE,PKADATE)" & vbCrLf _
                & "VALUES('" & tempPart & "','" & sPartNumber & "'," & lblRun & vbCrLf _
                & "," & iPkRecord & ",'',9,'" & strDate & "','" & strDate & "')"
         clsADOCon.ExecuteSQL sSql
      Else
         CancelTrans
      End If
      cmbNew.Value = False
   End If
   
End Sub

Private Sub cmbPrt_Click()
   GetPlPart
   
End Sub


Private Sub cmbPrt_LostFocus()
   Dim iList As Integer
   cmbPrt = CheckLen(cmbPrt, 30)
   GetPlPart
   iList = iTotalItems
   If lblDsc.ForeColor <> ES_RED Then
      On Error Resume Next
      sSql = "UPDATE MopkTable SET PKPARTREF='" & Compress(cmbPrt) & "' " _
             & "WHERE (PKMOPART='" & sPartNumber & "' AND " _
             & "PKMORUN=" & lblRun & " AND PKRECORD=" & lblRec & ") "
      clsADOCon.ExecuteSQL sSql
      vItems(iList, 0) = cmbPrt
      vItems(iList, 1) = lblDsc
      vItems(iList, 2) = "0.000"
      vItems(iList, 3) = lblUom
      vItems(iList, 4) = ""
      vItems(iList, 5) = lblLvl
   End If
   
End Sub


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If Trim(cmbPrt) <> "" Then
      If Val(txtQty) = 0 Then
         sMsg = "The Quantity Is 0, Really Want To Quit?"
         bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
         If bResponse = vbNo Then Exit Sub
      Else
         On Error Resume Next
         sSql = "UPDATE MopkTable SET PKPQTY=" & txtQty & ",PKCOMT='" _
                & RTrim(txtCmt) & "' WHERE PKMOPART='" & sPartNumber & "' AND " _
                & "PKMORUN=" & Compress(lblRun) & " AND PKRECORD=" & lblRec & " "
         clsADOCon.ExecuteSQL sSql
      End If
   End If
   Unload Me
   
End Sub



Private Sub cmdDn_Click()
   iIndex = iIndex + 1
   If iIndex > iTotalItems Then iIndex = iTotalItems
   GetThisItem
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5201"
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdItm_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   If cmdItm.Value = True Then
      sMsg = "Do You Really Want To Cancel " & vbCr _
             & "Pick item " & cmbPrt & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         clsADOCon.ADOErrNum = 0
         sSql = "UPDATE MopkTable SET PKPQTY=0,PKTYPE=12 " _
                & "WHERE PKMOPART='" & sPartNumber & "' AND " _
                & "PKMORUN=" & lblRun & " AND PKRECORD=" & lblRec & " "
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.ADOErrNum = 0 Then
            MsgBox "Item Was Canceled.", vbInformation, Caption
            GetItems
         Else
            MsgBox "Unable To Cancel item.", vbInformation, Caption
            txtQty = vItems(Val(lblRec), 2)
         End If
      Else
         CancelTrans
         txtQty = vItems(Val(lblRec), 2)
      End If
      cmdItm.Value = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "cmditm_click"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmdUp_Click()
   iIndex = iIndex - 1
   If iIndex < 1 Then iIndex = 1
   GetThisItem
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then
      GetItems
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
'   lblMon = PickMCe01a.cmbPrt
'   lblRun = PickMCe01a.cmbRun
   cmdUp.Picture = Dsup
   cmdUp.Enabled = False
   cmdDn.Picture = Endn
   cmdDn.Enabled = True
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   sSql = "UPDATE MopkTable SET PKRECORD=0 WHERE PKRECORD>0 " _
          & "AND PKMOPART='" & sPartNumber & "' AND PKMORUN=" & lblRun & " "
   clsADOCon.ExecuteSQL sSql
   If cbfrom1a = vbChecked Then PickMCe01a.optItm.Value = vbUnchecked
   
   sSql = "DELETE FROM MopkTable WHERE (PKMOPART='" & sPartNumber _
          & "' AND PKMORUN=" & lblRun & " AND PKPARTREF='')"
   clsADOCon.ExecuteSQL sSql
   cbfrom1a = vbUnchecked
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set PickMCe01b = Nothing
   
End Sub






Private Sub GetItems()
   Dim RdoPck As ADODB.Recordset
   Dim iRow As Integer
   Dim iRecord As Integer
   Erase vItems
   
   On Error GoTo DiaErr1
   sPartNumber = Compress(lblMon)
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PAUNITS,PALEVEL," _
          & "PKPARTREF,PKMOPART,PKMORUN,PKREV,PKTYPE,PKPQTY," _
          & "PKAQTY,PKCOMT,PKRECORD FROM PartTable,MopkTable " _
          & "WHERE PARTREF=PKPARTREF AND PKMOPART='" & sPartNumber & "'" _
          & "AND PKMORUN=" & Val(lblRun) & " AND PKAQTY=0 AND " _
          & "(PKTYPE=9 or PKTYPE=23 ) ORDER BY PARTREF"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPck, ES_KEYSET)
   If bSqlRows Then
      With RdoPck
         lblRec = "1"
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         txtQty = Format(!PKPQTY, ES_QuantityDataFormat)
         lblUom = "" & Trim(!PAUNITS)
         txtCmt = "" & Trim(!PKCOMT)
         lblLvl = Format(!PALEVEL, "0")
         Do Until .EOF
            iRecord = GetNextPickRecord(sPartNumber, Val(lblRun))
            iRow = iRow + 1
            vItems(iRow, 0) = "" & Trim(!PartNum)
            vItems(iRow, 1) = "" & Trim(!PADESC)
            vItems(iRow, 2) = Format(!PKPQTY, ES_QuantityDataFormat)
            vItems(iRow, 3) = "" & Trim(!PAUNITS)
            vItems(iRow, 4) = "" & Trim(!PKCOMT)
            vItems(iRow, 5) = Format(!PALEVEL, "0")
            If IsNull(!PKRECORD) Or !PKRECORD = 0 Then _
                      !PKRECORD = iRecord Else _
                      iRecord = !PKRECORD
            .Update
            vItems(iRow, 6) = Format(iRecord)
            .MoveNext
         Loop
         ClearResultSet RdoPck
      End With
      iTotalItems = iRow
   Else
      iTotalItems = 0
   End If
   If iTotalItems = 0 Then
      lblRec = ""
      cmbPrt = ""
      lblDsc = ""
      txtQty = "0.000"
      lblUom = ""
      txtCmt = ""
      cmdItm.Enabled = False
      MouseCursor 0
      MsgBox "There Are No Pick Items.", vbInformation, Caption
      Exit Sub
   End If
   If iTotalItems < 2 Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   End If
   Set RdoPck = Nothing
   MouseCursor 0
   If iTotalItems > 0 Then
      iIndex = 1
      FillPlParts
      GetThisItem
      On Error Resume Next
      txtQty.SetFocus
   Else
      MsgBox "No Unpicked Items Found.", _
         vbInformation, Caption
      Unload Me
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "getitems"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetThisItem()
   'update old
   On Error GoTo DiaErr1
   cmbPrt.Enabled = False
   cmbPrt.BackColor = vbButtonFace
   If Val(txtQty) = 0 Then txtQty = vItems(Val(lblRec), 2)
   lblRec = vItems(iIndex, 6)
   cmbPrt = vItems(iIndex, 0)
   lblDsc = vItems(iIndex, 1)
   txtQty = vItems(iIndex, 2)
   lblUom = vItems(iIndex, 3)
   txtCmt = vItems(iIndex, 4)
   lblLvl = vItems(iIndex, 5)
   If iIndex = 1 Then
      cmdUp.Picture = Dsup
      cmdUp.Enabled = False
   Else
      cmdUp.Picture = Enup
      cmdUp.Enabled = True
   End If
   If iIndex = iTotalItems Then
      cmdDn.Picture = Dsdn
      cmdDn.Enabled = False
   Else
      cmdDn.Picture = Endn
      cmdDn.Enabled = True
   End If
   On Error Resume Next
   txtQty.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "getthisitem"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
   
End Sub

Private Sub txtCmt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   On Error Resume Next
   sSql = "UPDATE MopkTable SET PKCOMT='" & Trim(txtCmt) & "'" _
          & "WHERE PKMOPART='" & sPartNumber & "' AND " _
          & "PKMORUN=" & lblRun & " AND PKRECORD=" & lblRec & " "
   clsADOCon.ExecuteSQL sSql
   
End Sub


Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyPageUp Then cmdUp_Click
   If KeyCode = vbKeyPageDown Then cmdDn_Click
   
End Sub


Private Sub txtQty_LostFocus()
   Dim cQuantity As Currency
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   On Error Resume Next
   cQuantity = txtQty
   sSql = "UPDATE MopkTable SET PKPQTY=" & cQuantity & ",PKCOMT='" _
          & Trim(txtCmt) & "' WHERE PKMOPART='" & sPartNumber & "' AND " _
          & "PKMORUN=" & lblRun & " AND PKRECORD=" & lblRec & " "
   clsADOCon.ExecuteSQL sSql
   
End Sub



Private Sub FillPlParts()
   Dim b As Byte
   On Error GoTo DiaErr1
   cmbPrt.Clear
   If Val(txtTyp) > 0 Then
      b = Val(txtTyp) - 1
      sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable " _
             & "WHERE PALEVEL<6 AND PALEVEL>" & b
      LoadComboBox cmbPrt
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillplparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetPlPart()
   Dim RdoPrt As ADODB.Recordset
   Dim bByte As Byte
   bByte = Val(txtTyp) - 1
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS " _
          & "FROM PartTable WHERE PARTREF='" & Compress(cmbPrt) & "' " _
          & " AND (PALEVEL<6 AND PALEVEL>" & bByte & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblUom = "" & Trim(!PAUNITS)
         lblLvl = "" & Trim(!PALEVEL)
         ClearResultSet RdoPrt
      End With
   Else
      lblDsc = "*** Part Wasn't Found Or Wrong Type ***"
   End If
   If Trim(cmbPrt) = Trim(lblMon) Then
      lblDsc = "*** Part Cannot Be Used On Itself ***"
      cmbPrt = ""
      lblUom = ""
      lblLvl = ""
   End If
   Set RdoPrt = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getplpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
