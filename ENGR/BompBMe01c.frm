VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Begin VB.Form BompBMe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Parts List Item"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPickAt 
      Height          =   285
      Left            =   6480
      TabIndex        =   43
      Tag             =   "1"
      Text            =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   1920
      Width           =   555
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe01c.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6840
      TabIndex        =   36
      ToolTipText     =   "Update This Item And Apply Changes"
      Top             =   840
      Width           =   875
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "BompBMe01c.frx":07AE
      DownPicture     =   "BompBMe01c.frx":1120
      Height          =   350
      Left            =   6240
      Picture         =   "BompBMe01c.frx":1A92
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Standard Comments"
      Top             =   2400
      Width           =   350
   End
   Begin VB.TextBox txtMatbr 
      Height          =   285
      Left            =   4440
      TabIndex        =   14
      Tag             =   "1"
      ToolTipText     =   "Material Burden Percentage"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Tag             =   "1"
      ToolTipText     =   "Total Material Costs For This Level"
      Top             =   4440
      Width           =   1035
   End
   Begin VB.TextBox txtLabOh 
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Tag             =   "1"
      ToolTipText     =   "Factory Overhead Rate"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Total Accumulated Labor Cost For This Level"
      Top             =   4080
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   26
      Top             =   3600
      Width           =   6855
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Sort Sequence (Otherwise Part Number)"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtRev 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Parts List Revision of Part"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   5355
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Quantity Used"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtBum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6300
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Unit of Measure for Parts List"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtCvt 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Units Conversion (Feet to Inches = 12.000)"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtAdr 
      Height          =   285
      Left            =   4740
      TabIndex        =   6
      Tag             =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txtSup 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox optPhn 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4740
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtCmt 
      Height          =   1150
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Comments (2048 Chars Max)"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4800
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5370
      FormDesignWidth =   7770
   End
   Begin VB.Label z1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pick At:"
      Height          =   255
      Index           =   20
      Left            =   5760
      TabIndex        =   44
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   19
      Left            =   5640
      TabIndex        =   41
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   40
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   11
      Left            =   4440
      TabIndex        =   39
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5400
      TabIndex        =   38
      ToolTipText     =   "Revision"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label cmbPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   37
      Top             =   840
      Width           =   3315
   End
   Begin VB.Label lblAssy 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   35
      ToolTipText     =   "Used On Part"
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   34
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
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
      Index           =   12
      Left            =   4200
      TabIndex        =   33
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev      "
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
      Index           =   10
      Left            =   4680
      TabIndex        =   32
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Burden"
      Height          =   255
      Index           =   13
      Left            =   2640
      TabIndex        =   31
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Overhead"
      Height          =   255
      Index           =   15
      Left            =   2640
      TabIndex        =   29
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Cost"
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimating Costs For This Level.  Should Not Include Lower Level Costs"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seq     "
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
      TabIndex        =   25
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                       "
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
      Left            =   720
      TabIndex        =   24
      Top             =   600
      Width           =   3270
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity       "
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
      Left            =   5355
      TabIndex        =   23
      Top             =   600
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Um     "
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
      Left            =   6300
      TabIndex        =   22
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert:   "
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   21
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Units Wasted:"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   20
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Qty:"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phantom:  "
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Top             =   1200
      Width           =   3075
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "BompBMe01c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
Option Explicit
Dim bOnLoad As Byte
Dim bGoodRev As Byte
Dim bDataChg As Byte
Dim bSaved As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetRevision() As Byte
   Dim RdoRes2 As ADODB.Recordset
   Dim sCurrPart As String
   sCurrPart = Compress(cmbPrt)
   If Trim(txtRev) = "" Then
      GetRevision = True
      Exit Function
   End If
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable " _
          & "WHERE BMHREF='" & sCurrPart & "' AND " _
          & "BMHREV='" & Trim(txtRev) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRes2)
   If bSqlRows Then
      GetRevision = True
   Else
      txtRev = ""
      GetRevision = False
   End If
   Set RdoRes2 = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getrevis"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bDataChg = 1 Then
      bResponse = MsgBox("Exit Without Saving Changes?..", _
                  ES_NOQUESTION, Caption)
      If bResponse = vbYes Then Unload Me
   Else
      Unload Me
   End If
   
End Sub



Private Sub cmdComments_Click()
   If cmdComments Then
      'See List For Index
      txtCmt.SetFocus
      SysComments.lblListIndex = 9
      SysComments.Show
      cmdComments = False
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3202
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdUpd_Click()
   bSaved = 1
   UpdateThisPart
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      cmdComments.Enabled = True
      GetThisPart
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move BompBMe01a.Left + 400, BompBMe01a.Top + 1200
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrForm = "Bills Of Material"
   If bSaved = 1 Then BompBMe01a.optRefresh.Value = vbChecked
   BompBMe01a.cmdQuit.Enabled = True
   BompBMe01a.cmdAdd.Enabled = True
   BompBMe01a.cmdEdit.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCut.Enabled = True
   BompBMe01a.cmdCopy.Enabled = True
   BompBMe01a.cmdDelete.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set BompBMe01c = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtRev.BackColor = Es_FormBackColor
   
End Sub


Private Sub UpdateThisPart()
   Dim RdoAdd As ADODB.Recordset
   MouseCursor 13
   cmdUpd.Enabled = False
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM BmplTable WHERE (BMASSYPART='" & Compress(lblAssy) _
          & "' AND BMPARTREF='" & Compress(cmbPrt) & "' AND BMREV='" _
          & lblRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAdd, ES_KEYSET)
   With RdoAdd
      Err = 0
      clsADOCon.ADOErrNum = 0
      ' TODO: not needed
      '.Edit
      '!BMASSYPART = Compress(lblAssy)
      '!BMPARTREF = Compress(cmbPrt)
      !BMPARTNUM = cmbPrt
      !BMQTYREQD = Val(txtQty)
      !BMUNITS = txtBum
      !BMCONVERSION = Val(txtCvt)
      !BMSEQUENCE = Val(txtSeq)
      !BMADDER = Val(txtAdr)
      !BMSETUP = Val(txtSup)
      !BMPHANTOM = str$(optPhn.Value)
      '!BMREFERENCE = txtRef
      !BMCOMT = Trim(txtCmt)
      !BMESTLABOR = Val(txtLab)
      !BMESTLABOROH = Val(txtLabOh)
      !BMESTMATERIAL = Val(txtMat)
      !BMESTMATERIALBRD = Val(txtMatbr)
      !BMPICKAT = Val(IIf(txtPickAt = "", 1, txtPickAt))
      .Update
      sSql = "UPDATE BmhdTable SET BMHREVDATE='" _
             & Format(ES_SYSDATE, "mm/dd/yy") & "' WHERE " _
             & "BMHREF='" & Compress(BompBMe01a.cmbPls) & "' " _
             & "AND BMHREV='" & Trim(BompBMe01a.cmbRev) & "'"
      clsADOCon.ExecuteSql sSql ' rdExecDirect
   End With
   If clsADOCon.ADOErrNum = 0 Then
      bDataChg = 0
      Sleep 500
      SysMsg "The Item Was Updated.", True
   Else
      MsgBox Trim(Err.Descripton) & vbCrLf _
                  & "Couldn't update The Item.", _
                  vbExclamation, Caption
   End If
   Set RdoAdd = Nothing
   Unload Me
   
End Sub

Private Sub GetThisPart()
   Dim RdoBom As ADODB.Recordset
   Dim cPartCost As Currency
   On Error GoTo DiaErr1
   sSql = "SELECT ISNULL(BMPICKAT, 1) PickAt, * FROM BmplTable WHERE (BMASSYPART='" & Compress(lblAssy) _
          & "' AND BMPARTREF='" & Compress(cmbPrt) & "' AND BMREV='" _
          & lblRev & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_STATIC)
   If bSqlRows Then
      With RdoBom
         txtSeq = Format(!BMSEQUENCE, "##0")
         cmbPrt = "" & Trim(!BMPARTNUM)
         txtRev = "" & Trim(!BMPARTREV)
         txtQty = Format(!BMQTYREQD, ES_QuantityDataFormat)
         txtBum = "" & Trim(!BMUNITS)
         txtCvt = Format(!BMCONVERSION, ES_QuantityDataFormat)
         txtAdr = Format(!BMADDER, ES_QuantityDataFormat)
         txtSup = Format(!BMSETUP, ES_QuantityDataFormat)
         optPhn.Value = !BMPHANTOM
         ' txtRef = "" & Trim(!BMREFERENCE)
         txtCmt = "" & !BMCOMT
         txtLab = Format(!BMESTLABOR, ES_QuantityDataFormat)
         txtLabOh = Format(!BMESTLABOROH, ES_QuantityDataFormat)
         txtMat = Format(!BMESTMATERIAL, ES_QuantityDataFormat)
         txtMatbr = Format(!BMESTMATERIALBRD, ES_QuantityDataFormat)
         
         txtPickAt = "" & Trim(!PickAt)
         
         If txtBum = "" Then txtBum = "EA"
         If Val(txtMat) = 0 Then
            cPartCost = GetPartCost(!BMPARTNUM)
            txtMat = Format(cPartCost, ES_QuantityDataFormat)
         End If
         ClearResultSet RdoBom
         FindPart
         bDataChg = 0
      End With
   Else
      txtSeq = "0"
      '  cmbPrt = ""
      txtRev = ""
      txtQty = "0.000"
      txtBum = ""
      txtCvt = "0.000"
      txtAdr = "0.000"
      txtSup = "0.000"
      optPhn.Value = vbUnchecked
      txtCmt = ""
      txtPickAt = "1"
   End If
   Set RdoBom = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub


Private Sub lblAssy_Change()
   GetIndexHeader
   
End Sub

Private Sub txtAdr_Change()
   bDataChg = 1
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 9)
   txtAdr = Format(Abs(Val(txtAdr)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtBum_Change()
   bDataChg = 1
   
End Sub

Private Sub txtBum_LostFocus()
   txtBum = CheckLen(txtBum, 2)
   If txtBum = "" Then txtBum = "EA"
   
End Sub


Private Sub txtCmt_Change()
   bDataChg = 1
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtCvt_Change()
   bDataChg = 1
   
End Sub



Private Sub txtCvt_LostFocus()
   txtCvt = CheckLen(txtCvt, 9)
   txtCvt = Format(Abs(Val(txtCvt)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLab_Change()
   bDataChg = 1
   
End Sub

Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 9)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLabOh_Change()
   bDataChg = 1
   
End Sub

Private Sub txtLabOh_LostFocus()
   txtLabOh = CheckLen(txtLabOh, 9)
   txtLabOh = Format(Abs(Val(txtLabOh)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMat_Change()
   bDataChg = 1
   
End Sub

Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 9)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMatbr_Change()
   bDataChg = 1
   
End Sub

Private Sub txtMatbr_LostFocus()
   txtMatbr = CheckLen(txtMatbr, 9)
   txtMatbr = Format(Abs(Val(txtMatbr)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtQty_Change()
   bDataChg = 1
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 4)
   bGoodRev = GetRevision()
   If Not bGoodRev Then
      MsgBox "That Revision Wasn't Found.", vbExclamation, Caption
      txtRev = ""
   End If
   
End Sub


Private Sub txtSeq_Change()
   bDataChg = 1
   
End Sub

Private Sub txtSeq_LostFocus()
   txtSeq = CheckLen(txtSeq, 3)
   txtSeq = Format$(Abs(Val(txtSeq)), "##0")
   
End Sub


Private Sub txtSup_Change()
   bDataChg = 1
   
End Sub

Private Sub txtSup_LostFocus()
   txtSup = CheckLen(txtSup, 9)
   txtSup = Format(Abs(Val(txtSup)), ES_QuantityDataFormat)
   
End Sub



Private Sub GetIndexHeader()
   Dim RdoHdr As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PARTREF='" & Compress(lblAssy) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdr, ES_FORWARD)
   If bSqlRows Then
      With RdoHdr
         lblAssy = "" & Trim(.Fields(1))
         ClearResultSet RdoHdr
      End With
   End If
   Set RdoHdr = Nothing
   
End Sub

