VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe02g 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit A Parts List Item For An Estimate"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Default Or Entered Rate"
      Top             =   4440
      Width           =   795
   End
   Begin VB.TextBox txtLabOh 
      Height          =   285
      Left            =   4920
      TabIndex        =   14
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Factory Overhead"
      Top             =   4440
      Width           =   815
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Tag             =   "1"
      ToolTipText     =   "Standard Or  Entered Unit Cost"
      Top             =   4080
      Width           =   795
   End
   Begin VB.TextBox txtMatbr 
      Height          =   285
      Left            =   4920
      TabIndex        =   11
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Material Burden"
      Top             =   4080
      Width           =   795
   End
   Begin VB.TextBox txtLabHrs 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Tag             =   "1"
      Text            =   ".0000"
      ToolTipText     =   "Total Labor Hours For This Quantity"
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6600
      TabIndex        =   33
      ToolTipText     =   "Update This Item"
      Top             =   840
      Width           =   875
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "EstiESe02g.frx":0000
      DownPicture     =   "EstiESe02g.frx":0972
      Height          =   350
      Left            =   6240
      Picture         =   "EstiESe02g.frx":12E4
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Standard Comments"
      Top             =   2400
      Width           =   350
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   7212
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Tag             =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtRev 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Parts List Revision of Part"
      Top             =   4920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   4875
      TabIndex        =   3
      Tag             =   "1"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtBum 
      Height          =   285
      Left            =   5820
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4740
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCmt 
      Height          =   1150
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "9"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   4920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4980
      FormDesignWidth =   7545
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   3000
      TabIndex        =   48
      ToolTipText     =   "Quantity (After Conversion)"
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   255
      Index           =   22
      Left            =   2400
      TabIndex        =   47
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   46
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblBid 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   45
      ToolTipText     =   "This Estimate"
      Top             =   240
      Width           =   915
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6960
      TabIndex        =   44
      ToolTipText     =   "Revision"
      Top             =   4920
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   43
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead"
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   42
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   41
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Burden"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   40
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   39
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   38
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label lblLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6360
      TabIndex        =   37
      ToolTipText     =   "Total Labor Cost For This Quantity"
      Top             =   4440
      Width           =   915
   End
   Begin VB.Label lblMaterial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6360
      TabIndex        =   36
      ToolTipText     =   "Total Material Cost"
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Labor Hours"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   35
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label cmbPrt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   34
      Top             =   840
      Width           =   3315
   End
   Begin VB.Label lblAssy 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   32
      ToolTipText     =   "Used On Part"
      Top             =   240
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   31
      Top             =   240
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
      TabIndex        =   30
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
      Height          =   252
      Index           =   10
      Left            =   6240
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimating Costs For This Level.  Should Not Include Lower Level Costs"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   28
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
      TabIndex        =   26
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number                                                                  "
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
      TabIndex        =   25
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
      Left            =   4875
      TabIndex        =   24
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
      Left            =   5820
      TabIndex        =   23
      Top             =   600
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Convert:   "
      Height          =   255
      Index           =   5
      Left            =   720
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inv Units Wasted:"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   21
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Setup Qty:"
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phantom:  "
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Top             =   1200
      Width           =   3075
   End
   Begin VB.Label lblTyp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4200
      TabIndex        =   16
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "EstiESe02g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'10/12/04 corrected GetThisPart column
Option Explicit
Dim bOnLoad As Byte
Dim bDataChg As Byte
Dim bSaved As Byte

Dim MatBurden As Currency
Dim FacOverHead As Currency
Dim LabRate As Currency

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub UpdateMaterial()
   Dim cAdder As Currency
   Dim cBurden As Currency
   Dim cConvert As Currency
   Dim cCost As Currency
   Dim cQuantity As Currency
   Dim cSetup As Currency
   
   cAdder = Val(txtAdr)
   cBurden = (Val(txtMatbr) / 100) + 1
   cCost = Val(txtMat)
   cSetup = Val(txtSup)
   cConvert = Format(Val(txtCvt), ES_QuantityDataFormat)
   If cConvert = 0 Then cConvert = 1
   cQuantity = Format(Val(txtQty) + cAdder + cSetup, ES_QuantityDataFormat)
   cQuantity = Format(cQuantity / cConvert, ES_QuantityDataFormat)
   lblQty = Format(cQuantity, ES_QuantityDataFormat)
   cBurden = cBurden * cCost
   lblMaterial = Format(cBurden * cQuantity, ES_QuantityDataFormat)
   
End Sub

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

Private Sub cmdUpd_Click()
   bSaved = 1
   UpdateThisPart
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      If cmbPrt = EstiESe02c.cmbPls Then
         txtLabHrs.Enabled = False
         txtLab.Enabled = False
         txtLabOh.Enabled = False
      End If
      GetEstimatingDefaults MatBurden, FacOverHead, LabRate
      cmdComments.Enabled = True
      GetThisPart
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   Move EstiESe02c.Left + 400, EstiESe02c.Top + 1200
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrForm = "Bills Of Material"
   If bSaved = 1 Then EstiESe02c.optRefresh.Value = vbChecked
   EstiESe02c.cmdQuit.Enabled = True
   EstiESe02c.cmdAdd.Enabled = True
   EstiESe02c.cmdEdit.Enabled = True
   EstiESe02c.cmdCut.Enabled = True
   EstiESe02c.cmdCut.Enabled = True
   EstiESe02c.cmdCopy.Enabled = True
   EstiESe02c.cmdDelete.Enabled = True
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EstiESe02g = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtRev.BackColor = Es_FormBackColor
   txtLabOh = "0.000"
   txtLabHrs = "0.0000"
   txtMat = "0.000"
   txtMatbr = "0.000"
   lblLabor = "0.000"
   lblMaterial = "0.000"
   txtSeq = "0"
   txtBum = "EA"
   
End Sub


Private Sub UpdateThisPart()
   Dim RdoAdd As ADODB.Recordset
   MouseCursor 13
   cmdUpd.Enabled = False
   On Error Resume Next
   Err = 0
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT * FROM EsbmTable WHERE (BIDBOMASSYPART='" & Compress(lblAssy) _
          & "' AND BIDBOMPARTREF='" & Compress(cmbPrt) & "' AND BIDBOMREF=" _
          & Val(lblBid) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAdd, ES_KEYSET)
   With RdoAdd
      Err = 0
      clsADOCon.ADOErrNum = 0
      '.Edit
      '!BIDBOMQTYREQD = Format(Val(lblQty), ES_QuantityDataFormat)
      !BIDBOMQTYREQD = Format(Val(txtQty), ES_QuantityDataFormat)
      !BIDBOMUNITS = txtBum
      !BIDBOMCONVERSION = Format(Val(txtCvt), ES_QuantityDataFormat)
      !BIDBOMSEQUENCE = Val(txtSeq)
      !BIDBOMADDER = Format(Val(txtAdr), ES_QuantityDataFormat)
      !BIDBOMSETUP = Format(Val(txtSup), ES_QuantityDataFormat)
      !BIDBOMPHANTOM = str$(optPhn.Value)
      !BIDBOMCOMT = Trim(txtCmt)
      !BIDBOMLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
      !BIDBOMLABORHRS = Format(Val(txtLabHrs), ES_QuantityDataFormat)
      !BIDBOMLABOROH = Format(Val(txtLabOh), ES_QuantityDataFormat)
      !BIDBOMMATERIAL = Format(Val(txtMat), ES_QuantityDataFormat)
      !BIDBOMMATERIALBRD = Format(Val(txtMatbr), ES_QuantityDataFormat)
      !BIDBOMESTUNITCOST = Format(Val(lblMaterial) + Val(lblLabor), ES_QuantityDataFormat)
      .Update
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
   sSql = "SELECT * FROM EsbmTable WHERE (BIDBOMASSYPART='" & Compress(lblAssy) _
          & "' AND BIDBOMPARTREF='" & Compress(cmbPrt) & "' AND BIDBOMREF=" _
          & Val(lblBid) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBom, ES_FORWARD)
   If bSqlRows Then
      With RdoBom
         GetPartInfo !BIDBOMPARTREF
         txtSeq = Format(!BIDBOMSEQUENCE, "##0")
         cmbPrt = "" & Trim(!BIDBOMPARTREF)
         txtQty = Format(!BIDBOMQTYREQD, ES_QuantityDataFormat)
         txtBum = "" & Trim(!BIDBOMUNITS)
         txtCvt = Format(!BIDBOMCONVERSION, ES_QuantityDataFormat)
         txtAdr = Format(!BIDBOMADDER, ES_QuantityDataFormat)
         txtSup = Format(!BIDBOMSETUP, ES_QuantityDataFormat)
         txtCmt = "" & !BIDBOMCOMT
         txtLabHrs = Format(!BIDBOMLABORHRS, ES_QuantityDataFormat)
         txtLab = Format(!BIDBOMLABOR, ES_QuantityDataFormat)
         txtLabOh = Format(!BIDBOMLABOROH, ES_QuantityDataFormat)
         txtMat = Format(!BIDBOMMATERIAL, ES_QuantityDataFormat)
         txtMatbr = Format(!BIDBOMMATERIALBRD, ES_QuantityDataFormat)
         lblMaterial = Format((Val(txtMat) * (Val(txtMatbr) / 100) + Val(txtMat)), ES_QuantityDataFormat)
         lblLabor = Format(((txtLab)) * ((Val(txtLabOh) / 100) + 1) * Val(txtLabHrs), ES_QuantityDataFormat)
         If txtBum = "" Then txtBum = "EA"
         If (!BIDBOMMATERIAL + !BIDBOMMATERIALBRD) = 0 Then _
             txtMatbr = Format(MatBurden, ES_QuantityDataFormat)
         If (!BIDBOMLABORHRS + !BIDBOMLABOROH) = 0 Then
            txtLabOh = Format(FacOverHead, ES_QuantityDataFormat)
            txtLab = Format(LabRate, ES_QuantityDataFormat)
         End If
         If Val(txtMat) = 0 Then
            cPartCost = GetPartCost(!BIDBOMPARTREF)
            txtMat = Format(cPartCost, ES_QuantityDataFormat)
         End If
         ClearResultSet RdoBom
         FindPart
         UpdateMaterial
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
   End If
   Set RdoBom = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getthispar"
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
   UpdateMaterial
   
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
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtCvt_Change()
   bDataChg = 1
   
End Sub



Private Sub txtCvt_LostFocus()
   txtCvt = CheckLen(txtCvt, 9)
   txtCvt = Format(Abs(Val(txtCvt)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtLab_Change()
   bDataChg = 1
   
End Sub

Private Sub txtLab_LostFocus()
   txtLab = CheckLen(txtLab, 9)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   lblLabor = Format(((txtLab)) * ((Val(txtLabOh) / 100) + 1) * Val(txtLabHrs), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLabHrs_LostFocus()
   txtLabHrs = CheckLen(txtLabHrs, 9)
   txtLabHrs = Format(Abs(Val(txtLabHrs)), "####0.0000")
   lblLabor = Format(((txtLab)) * ((Val(txtLabOh) / 100) + 1) * Val(txtLabHrs), ES_QuantityDataFormat)
   
End Sub


Private Sub txtLabOh_Change()
   bDataChg = 1
   
End Sub

Private Sub txtLabOh_LostFocus()
   txtLabOh = CheckLen(txtLabOh, 9)
   txtLabOh = Format(Abs(Val(txtLabOh)), ES_QuantityDataFormat)
   lblLabor = Format(((txtLab)) * ((Val(txtLabOh) / 100) + 1) * Val(txtLabHrs), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMat_Change()
   bDataChg = 1
   
End Sub

Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 9)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtMatbr_Change()
   bDataChg = 1
   
End Sub

Private Sub txtMatbr_LostFocus()
   txtMatbr = CheckLen(txtMatbr, 9)
   txtMatbr = Format(Abs(Val(txtMatbr)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtQty_Change()
   bDataChg = 1
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 4)
   
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
   UpdateMaterial
   
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


Private Sub GetPartInfo(sPartNumber As String)
   Dim RdoInf As ADODB.Recordset
   sSql = "SELECT PARTREF,PARTNUM,PADESC FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInf, ES_FORWARD)
   If bSqlRows Then
      With RdoInf
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         ClearResultSet RdoInf
      End With
   End If
   
End Sub
