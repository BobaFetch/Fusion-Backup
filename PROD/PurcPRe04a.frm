VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Part Purchasing Information"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFnd 
      DownPicture     =   "PurcPRe04a.frx":0000
      Height          =   315
      Left            =   4200
      Picture         =   "PurcPRe04a.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Find A Part Number"
      Top             =   600
      Width           =   350
   End
   Begin VB.TextBox txtPrt 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "3"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRe04a.frx":0684
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtMft 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Tag             =   "1"
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtPlt 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtEoq 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Tag             =   "1"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtPrc 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Tag             =   "1"
      Text            =   "1.0000"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtPun 
      Height          =   285
      Left            =   5520
      TabIndex        =   12
      Top             =   4440
      Width           =   375
   End
   Begin VB.CheckBox optMin 
      Caption         =   "Manufacturing"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtPou 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtSsq 
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Tag             =   "1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtRrq 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "0"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtSfs 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Tag             =   "1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtOst 
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Tag             =   "1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtMbe 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "M, B or E"
      Top             =   2160
      Width           =   375
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter a New Part or Select From List (30 chars)"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   5400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5760
      FormDesignWidth =   6225
   End
   Begin VB.Label txtUom 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   40
      Top             =   600
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uom"
      Height          =   285
      Index           =   12
      Left            =   4680
      TabIndex        =   39
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   " (Days)"
      Height          =   285
      Index           =   11
      Left            =   3360
      TabIndex        =   38
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   " (Days)"
      Height          =   285
      Index           =   33
      Left            =   3360
      TabIndex        =   37
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Flow"
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   36
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label txtTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   35
      Top             =   960
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   34
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing Lead Time"
      Height          =   285
      Index           =   32
      Left            =   240
      TabIndex        =   33
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Economic Order Qty"
      Height          =   285
      Index           =   30
      Left            =   240
      TabIndex        =   32
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing Conversion"
      Height          =   285
      Index           =   13
      Left            =   240
      TabIndex        =   31
      Top             =   4440
      Width           =   1995
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing Unit of Meas"
      Height          =   285
      Index           =   28
      Left            =   3360
      TabIndex        =   30
      Top             =   4440
      Width           =   2115
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Yield For"
      Height          =   525
      Index           =   5
      Left            =   3360
      TabIndex        =   29
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Point Of Use Qty"
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ship Set Qty"
      Height          =   285
      Index           =   35
      Left            =   3360
      TabIndex        =   27
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recommended Run Qty"
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   3120
      Width           =   2355
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchased Parts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufactured Parts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Safety Stock (Min)"
      Height          =   285
      Index           =   29
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overstock (Max)"
      Height          =   285
      Index           =   23
      Left            =   3360
      TabIndex        =   22
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make, Buy or Either"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(M,B or E)"
      Height          =   285
      Index           =   8
      Left            =   2640
      TabIndex        =   20
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label lblExt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   1320
      TabIndex        =   19
      Top             =   1320
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1155
   End
End
Attribute VB_Name = "PurcPRe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim RdoPur As ADODB.Recordset

Dim bOnLoad As Byte
Dim bGoodPart As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If Len(cmbPrt) = 0 Then
      cmdCan.SetFocus
      Exit Sub
   End If
   bGoodPart = GetPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPrt = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4304
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
   
      Dim bPartSearch As Boolean
      
      bPartSearch = GetPartSearchOption
      SetPartSearchOption (bPartSearch)
      
      If (Not bPartSearch) Then FillCombo
   
      
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS,PAMAKEBUY,PASAFETY," _
          & "PAOVERSTOCK,PARRQ,PASHIPSET,PAOUQTY,PAMINYIELD,PAFLOWTIME,PAEOQ," _
          & "PALEADTIME,PAPURCONV,PAPUNITS,PAEXTDESC FROM PartTable WHERE PARTREF= ? "

   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.SIZE = 30
   
   AdoQry.Parameters.Append AdoParameter
   
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   SaveCurrentSelections
   On Error Resume Next
   Set RdoPur = Nothing
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set PurcPRe04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub optMin_Click()
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAMINYIELD = optMin.Value
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub

Private Sub optMin_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtEoq_LostFocus()
   txtEoq = CheckLen(txtEoq, 7)
   txtEoq = Format(Abs(Val(txtEoq)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAEOQ = 0 + Val(txtEoq)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMbe_LostFocus()
   Dim bByte As Byte
   txtMbe = CheckLen(txtMbe, 1)
   Select Case txtMbe
      Case "M", "B", "E"
         bByte = True
      Case Else
         bByte = False
   End Select
   If Not bByte Then
      MsgBox "Must Be M, B or E.", vbInformation, Caption
      txtMbe = "M"
   End If
   
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAMAKEBUY = "" & txtMbe
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtMft_LostFocus()
   txtMft = CheckLen(txtMft, 7)
   txtMft = Format(Abs(Val(txtMft)), "###0.00")
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAFLOWTIME = 0 + Val(txtMft)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtOst_LostFocus()
   txtOst = CheckLen(txtOst, 7)
   txtOst = Format(Abs(Val(txtOst)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAOVERSTOCK = 0 + Val(txtOst)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPlt_LostFocus()
   txtPlt = CheckLen(txtPlt, 7)
   txtPlt = Format(Abs(Val(txtPlt)), "###0.00")
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PALEADTIME = 0 + Val(txtPlt)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPou_LostFocus()
   txtPou = CheckLen(txtPou, 7)
   txtPou = Format(Abs(Val(txtPou)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAOUQTY = 0 + Val(txtPou)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPrc_LostFocus()
   txtPrc = CheckLen(txtPrc, 10)
   If Val(txtPrc) = 0 Then txtPrc = 1
   txtPrc = Format(Abs(Val(txtPrc)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAPURCONV = 0 + Val(txtPrc)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtPun_LostFocus()
   txtPun = CheckLen(txtPun, 2)
   If Len(txtPun) = 0 Then txtPun = txtUom
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PAPUNITS = "" & txtPun
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtRrq_LostFocus()
   txtRrq = CheckLen(txtRrq, 6)
   txtRrq = Format(Val(txtRrq), "#####0")
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PARRQ = 0 + Val(txtRrq)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtSfs_LostFocus()
   txtSfs = CheckLen(txtSfs, 7)
   txtSfs = Format(Abs(Val(txtSfs)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PASAFETY = 0 + Val(txtSfs)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetPart() As Byte
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   On Error GoTo DiaErr1
   bSqlRows = clsADOCon.GetQuerySet(RdoPur, AdoQry, ES_KEYSET, True, 1)
   If bSqlRows Then
      With RdoPur
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         txtTyp = 0 + Format(!PALEVEL, "#0")
         txtMbe = "" & Trim(!PAMAKEBUY)
         txtSfs = Format(!PASAFETY, ES_QuantityDataFormat)
         txtRrq = 0 + Format(!PARRQ, "#####0")
         txtPrc = Format(0 + !PAPURCONV, ES_QuantityDataFormat)
         txtUom = "" & Trim(!PAUNITS)
         txtPun = "" & Trim(!PAPUNITS)
         If Trim(txtPun) = "" Then txtPun = txtUom
         txtSfs = Format(0 + !PASAFETY, ES_QuantityDataFormat)
         txtOst = Format(0 + !PAOVERSTOCK, ES_QuantityDataFormat)
         txtPou = Format(0 + !PAOUQTY, ES_QuantityDataFormat)
         txtEoq = Format(0 + !PAEOQ, ES_QuantityDataFormat)
         txtSsq = Format(0 + !PASHIPSET, ES_QuantityDataFormat)
         txtMft = Format(0 + !PAFLOWTIME, "###0.00")
         txtPlt = Format(0 + !PALEADTIME, "###0.00")
         lblExt = "" & Trim(!PAEXTDESC)
         If !PALEVEL < 5 Then cUR.CurrentPart = cmbPrt
      End With
      GetPart = True
   Else
      MsgBox "That Part Wasn't Found.", vbInformation, Caption
      lblDsc = ""
      txtTyp = "0"
      lblExt = ""
      txtSfs = "0.000"
      txtRrq = "0"
      txtPrc = "0.0000"
      txtPun = ""
      txtSfs = "0.000"
      txtOst = "0.000"
      txtPou = "0.000"
      txtEoq = "0.000"
      txtSsq = "0.000"
      txtMft = "0.00"
      txtPlt = "0.00"
      GetPart = False
      On Error Resume Next
      cmbPrt.SetFocus
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtSsq_LostFocus()
   txtSsq = CheckLen(txtSsq, 7)
   txtSsq = Format(Abs(Val(txtSsq)), ES_QuantityDataFormat)
   If bGoodPart Then
      On Error Resume Next
      RdoPur!PASHIPSET = 0 + Val(txtSsq)
      RdoPur.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If bSqlRows Then
      sPassedPart = Trim(cUR.CurrentPart)
      If Len(sPassedPart) > 0 Then
         cmbPrt = cUR.CurrentPart
      Else
         If cmbPrt.ListCount > 0 Then cmbPrt = cmbPrt.List(0)
      End If
      bGoodPart = GetPart()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Function SetPartSearchOption(bPartSearch As Boolean)
   
   If (bPartSearch = True) Then
      cmbPrt.Visible = False
      txtPrt.Visible = True
      cmdFnd.Visible = True
   Else
      cmbPrt.Visible = True
      txtPrt.Visible = False
      cmdFnd.Visible = False
   End If
End Function

Private Sub txtPrt_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF4 Then
      ViewParts.lblControl = "TXTPRT"
      ViewParts.txtPrt = txtPrt
      ViewParts.Show
   End If
   
End Sub


Private Sub txtPrt_LostFocus()
   txtPrt = CheckLen(txtPrt, 30)
   If Trim(txtPrt) = "" Or Trim(txtPrt) = "ALL" Then txtPrt = "ALL"
   cmbPrt = txtPrt
End Sub


Private Sub cmdFnd_Click()
   If txtPrt.Visible Then
      cmbPrt = txtPrt
      ViewParts.lblControl = "TXTPRT"
   Else
      ViewParts.lblControl = "CMBPRT"
   End If
   ViewParts.txtPrt = cmbPrt
   ViewParts.Show
End Sub


