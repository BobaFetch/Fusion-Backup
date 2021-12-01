VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe02f 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add A Part To An Estimating Parts List"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFindPart 
      Height          =   375
      Left            =   4200
      Picture         =   "EstiESe02f.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtPrt 
      Height          =   285
      Left            =   720
      TabIndex        =   52
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtLabHrs 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Tag             =   "1"
      Text            =   ".0000"
      ToolTipText     =   "Total Labor Hours For This Quantity"
      Top             =   4440
      Width           =   795
   End
   Begin VB.ComboBox cmbRev 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Revision (Blank For Default)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6960
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Add This Part"
      Top             =   840
      Width           =   875
   End
   Begin VB.ListBox lstAssy 
      Height          =   840
      ItemData        =   "EstiESe02f.frx":043A
      Left            =   7320
      List            =   "EstiESe02f.frx":043C
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdComments 
      DisabledPicture =   "EstiESe02f.frx":043E
      DownPicture     =   "EstiESe02f.frx":0DB0
      Height          =   350
      Left            =   6240
      Picture         =   "EstiESe02f.frx":1722
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Standard Comments"
      Top             =   2400
      Width           =   350
   End
   Begin VB.TextBox txtMatbr 
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Material Burden"
      Top             =   4080
      Width           =   795
   End
   Begin VB.TextBox txtMat 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Standard Cost Or Entered Cost"
      Top             =   4080
      Width           =   795
   End
   Begin VB.TextBox txtLabOh 
      Height          =   285
      Left            =   4920
      TabIndex        =   13
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Factory Overhead"
      Top             =   4440
      Width           =   815
   End
   Begin VB.TextBox txtLab 
      Height          =   285
      Left            =   3000
      TabIndex        =   12
      Tag             =   "1"
      Text            =   "0.000"
      ToolTipText     =   "Default Or Entered Rate"
      Top             =   4440
      Width           =   795
   End
   Begin VB.Frame z2 
      Height          =   40
      Left            =   120
      TabIndex        =   27
      Top             =   3600
      Width           =   7680
   End
   Begin VB.TextBox txtSeq 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Tag             =   "1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtQty 
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Tag             =   "1"
      Top             =   840
      Width           =   915
   End
   Begin VB.TextBox txtBum 
      Height          =   285
      Left            =   6240
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Unit of Measure for Parts List"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtCvt 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Tag             =   "1"
      ToolTipText     =   "Units Conversion (Feet to Inches = 12.000)"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtAdr 
      Height          =   285
      Left            =   4740
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Wasted (cut off)"
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox txtSup 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
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
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCmt 
      Height          =   1150
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "9"
      Top             =   2400
      Width           =   4335
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6960
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   5040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5055
      FormDesignWidth =   7980
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "BOM Level"
      Height          =   195
      Index           =   23
      Left            =   3840
      TabIndex        =   51
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label lblBomLevel 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      Height          =   285
      Left            =   4800
      TabIndex        =   50
      ToolTipText     =   "This Estimate"
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   255
      Index           =   22
      Left            =   2400
      TabIndex        =   49
      Top             =   4080
      Width           =   495
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
      Caption         =   "Labor Hours"
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   47
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblMaterial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6360
      TabIndex        =   46
      ToolTipText     =   "Total Material Costs"
      Top             =   4080
      Width           =   915
   End
   Begin VB.Label lblLabor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.000"
      Height          =   285
      Left            =   6360
      TabIndex        =   45
      ToolTipText     =   "Total Labor Costs For This Quantity"
      Top             =   4440
      Width           =   915
   End
   Begin VB.Label lblBid 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   44
      ToolTipText     =   "This Estimate"
      Top             =   120
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimate"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   43
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   42
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   255
      Index           =   18
      Left            =   5760
      TabIndex        =   41
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision"
      Height          =   255
      Index           =   11
      Left            =   3360
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   7080
      TabIndex        =   39
      ToolTipText     =   "Revision"
      Top             =   5280
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblAssy 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   36
      ToolTipText     =   "Used On Part"
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assembly "
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   35
      Top             =   120
      Width           =   975
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
      Left            =   4680
      TabIndex        =   34
      Top             =   600
      Width           =   615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rev              "
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
      Left            =   4800
      TabIndex        =   33
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Burden"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   32
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Material Cost"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   31
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead"
      Height          =   255
      Index           =   15
      Left            =   3960
      TabIndex        =   30
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   255
      Index           =   16
      Left            =   2400
      TabIndex        =   29
      Top             =   4440
      Width           =   1095
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
      TabIndex        =   25
      Top             =   600
      Width           =   3735
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
      Left            =   5280
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
      Left            =   6240
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
      Left            =   4680
      TabIndex        =   16
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "EstiESe02f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'New 9/18/03
'9/1/04 omit tools
Option Explicit
Dim bOnLoad As Byte
Dim bChanged As Byte
Dim bSaved As Byte

Dim MatBurden As Currency
Dim FacOverHead As Currency
Dim LabRate As Currency

Dim sOldPart As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub GetIndexHeader()
   Dim RdoHdr As ADODB.Recordset
   On Error Resume Next
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PARTREF='" & Compress(lblAssy) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoHdr, ES_FORWARD)
   If bSqlRows Then
      With RdoHdr
         lblAssy = "" & Trim(!PartNum)
         ClearResultSet RdoHdr
      End With
   End If
   Set RdoHdr = Nothing
   
End Sub

'Private Sub cmbPrt_Click()
'   If sOldPart <> cmbPrt Then FillBomhRev cmbPrt
'
'End Sub

'Private Sub cmbPrt_LostFocus()
'   cmbPrt = CheckLen(cmbPrt, 30)
'   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
'   If lblDsc.ForeColor <> ES_RED Then
'      GetThisPart
'      If sOldPart <> cmbPrt Then FillBomhRev cmbPrt
'      sOldPart = cmbPrt
'      bChanged = 1
'   End If
'
'End Sub


Private Sub cmbRev_LostFocus()
   cmbRev = CheckLen(cmbRev, 4)
   
End Sub


Private Sub cmdAdd_Click()
   Dim b As Byte
   Dim iList As Integer
   
   If Val(txtQty) > 0 Then _
          If lblDsc.ForeColor <> ES_RED Then cmdAdd.Enabled = True
   
   If Val(txtQty) = 0 Then
      MsgBox "Requires A Valid Quantity.", _
         vbInformation, Caption
      Exit Sub
   End If
'   For iList = 0 To EstiESe02c.lstNodes.ListCount - 1
'      If Compress(cmbPrt) = EstiESe02c.lstNodes.List(iList) Then
'         b = 1
'         Exit For
'      End If
'   Next
'   If b = 1 Then
'      lblDsc = "*** Part Number Is In Use ***"
'      MsgBox "The Selected Part Is Used Higher " & vbCrLf _
'         & "And Cannot Be Used On This Assembly.", vbInformation, _
'         Caption
'      Exit Sub
'   End If
   
   
'   For iList = 0 To cmbPrt.ListCount - 1
'      If cmbPrt = cmbPrt.List(iList) Then
'         b = 1
'         Exit For
'      End If
'   Next
'   If b = 0 Then
'      lblDsc = "*** Part Number Is The Wrong Type ***"
'      MsgBox "The Selected Part Is The Wrong Part Type " & vbCrLf _
'         & "Cannot Be Used On This Assembly.", vbInformation, _
'         Caption
'      Exit Sub
'   End If
   If lblDsc.ForeColor = ES_RED Or txtPrt = "NONE" Then
      MsgBox "Requires A Valid Part Number.", vbInformation, _
         Caption
      Exit Sub
   End If
   b = 0
   If lstAssy.ListCount > 0 Then
      For iList = 0 To lstAssy.ListCount - 1
         If Compress(txtPrt) = lstAssy.List(iList) Then b = 1
      Next
   End If
   If b = 1 Then
      MsgBox "You May Not Use The Same Part Number Twice.", _
         vbInformation, Caption
   Else
      bSaved = 1
      AddThisPart
   End If
End Sub

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bSaved = 0 And bChanged = 1 Then
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

Private Sub cmdFindPart_Click()
   ViewParts.lblControl = "TXTPRT"
   ViewParts.txtPrt = txtPrt
'   ViewParts.lblWhereClause = "PALEVEL BETWEEN " & EstiESe02c.lblLvl & " AND 5 AND PALEVEL>0 AND PATOOL=0"
   ViewParts.lblWhereClause = "PALEVEL BETWEEN 2 AND 5 AND PALEVEL>0 AND PATOOL=0"
   ViewParts.Show
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   '12/02/03
   'If lblAssy = EstiESe02c.cmbPls Then
   '    txtLabHrs.Enabled = False
   '    txtLab.Enabled = False
   '    txtLabOh.Enabled = False
   'End If
   If bOnLoad Then
      sOldPart = ""
      GetEstimatingDefaults MatBurden, FacOverHead, LabRate
      txtMatbr = Format(MatBurden, "##0.00")
      txtLabOh = Format(FacOverHead, "##0.00")
      txtLab = Format(LabRate, "##0.00")
      cmdComments.Enabled = True
      FillList
      'FillParts
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   Dim b As Byte
   FormLoad Me, ES_DONTLIST
   Move EstiESe02c.Left + 400, EstiESe02c.Top + 1200
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
   Set EstiESe02f = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtQty = "0.000"
   txtCvt = "0.000"
   txtAdr = "0.000"
   txtSup = "0.000"
   txtLab = "0.000"
   txtLabOh = "0.000"
   txtLabHrs = "0.0000"
   txtMat = "0.000"
   txtMatbr = "0.000"
   txtSeq = "0"
   txtBum = "EA"
   lblQty = "0.000"
   
End Sub

Private Sub FillList()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT BIDBOMREF,BIDBOMASSYPART,BIDBOMPARTREF " _
          & "FROM EsbmTable WHERE (BIDBOMREF=" & Val(lblBid) _
          & " AND BIDBOMASSYPART='" & Compress(lblAssy) & "') "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         lstAssy.AddItem "" & Compress(lblAssy)
         Do Until .EOF
            lstAssy.AddItem "" & Trim(!BIDBOMPARTREF)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   EstiESe02c.cmdQuit.Enabled = True
   EstiESe02c.cmdAdd.Enabled = True
   EstiESe02c.cmdEdit.Enabled = True
   EstiESe02c.cmdCut.Enabled = True
   EstiESe02c.cmdCut.Enabled = True
   EstiESe02c.cmdCopy.Enabled = True
   EstiESe02c.cmdDelete.Enabled = True
   sProcName = "fillList"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


'Private Sub FillParts()
'   Dim RdoPrt As ADODB.Recordset
'   Dim iParts As Integer
'   Dim iNodes As Integer
'   MouseCursor 11
'   On Error GoTo DiaErr1
'   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PATOOL FROM PartTable " _
'          & "WHERE (PALEVEL >=" & Val(EstiESe02c.lblLvl) & " AND " _
'          & "PALEVEL<6 AND PALEVEL>1 AND PATOOL=0) ORDER BY PARTREF"
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoPrt)
'   If bSqlRows Then
'      With RdoPrt
'         Do Until .EOF
'            If "" & Trim(!PartRef) <> Compress(lblAssy) Then _
'               AddComboStr cmbPrt.hwnd, "" & Trim(!PartNum)
'            .MoveNext
'         Loop
'         ClearResultSet RdoPrt
'      End With
'   End If
'   If cmbPrt.ListCount > 0 And EstiESe02c.lstNodes.ListCount > 0 Then
'      For iParts = 0 To cmbPrt.ListCount - 1
'         For iNodes = 0 To EstiESe02c.lstNodes.ListCount - 1
'            If cmbPrt.List(iParts) = EstiESe02c.lstNodes.List(iNodes) Then
'               cmbPrt.RemoveItem iParts
'            End If
'         Next
'      Next
'   End If
'   Set RdoPrt = Nothing
'   bChanged = 0
'   Exit Sub
'
'DiaErr1:
'   EstiESe02c.cmdQuit.Enabled = True
'   EstiESe02c.cmdAdd.Enabled = True
'   EstiESe02c.cmdEdit.Enabled = True
'   EstiESe02c.cmdCut.Enabled = True
'   EstiESe02c.cmdCut.Enabled = True
'   EstiESe02c.cmdCopy.Enabled = True
'   EstiESe02c.cmdDelete.Enabled = True
'   sProcName = "fillpartco"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'
'
'End Sub

Private Sub AddThisPart()
   Dim RdoAdd As ADODB.Recordset
   MouseCursor 13
   cmdAdd.Enabled = False
   On Error Resume Next
   sSql = "SELECT * FROM EsbmTable WHERE BIDBOMASSYPART='" _
          & Compress(lblAssy) & "' AND BIDBOMREF=" & Val(lblBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoAdd, ES_DYNAMIC)
   With RdoAdd
      Err = 0
      clsADOCon.ADOErrNum = 0
      .AddNew
      !BIDBOMLEVEL = Me.lblBomLevel
      !BIDBOMREF = Val(lblBid)
      !BIDBOMASSYPART = Compress(lblAssy)
      !BIDBOMPARTREF = Compress(txtPrt)
      !BIDBOMQTYREQD = Format(Val(lblQty), ES_QuantityDataFormat)
      !BIDBOMUNITS = txtBum
      !BIDBOMCONVERSION = Format(Val(txtCvt), ES_QuantityDataFormat)
      !BIDBOMSEQUENCE = Val(txtSeq)
      !BIDBOMSETUP = Format(Val(txtSup), ES_QuantityDataFormat)
      !BIDBOMADDER = Format(Val(txtAdr), ES_QuantityDataFormat)
      !BIDBOMCOMT = Trim(txtCmt)
      !BIDBOMLABOR = Format(Val(txtLab), ES_QuantityDataFormat)
      !BIDBOMLABOROH = Format(Val(txtLabOh), ES_QuantityDataFormat)
      !BIDBOMLABORHRS = Format(Val(txtLabHrs), ES_QuantityDataFormat)
      !BIDBOMMATERIAL = Format(Val(txtMat), ES_QuantityDataFormat)
      !BIDBOMMATERIALBRD = Format(Val(txtMatbr), ES_QuantityDataFormat)
      !BIDBOMESTUNITCOST = Format(Val(lblMaterial) + Val(lblLabor), ES_QuantityDataFormat)
      .Update
   End With
   If clsADOCon.ADOErrNum = 0 Then
      EstiESe02c.optRefresh = vbChecked
      lstAssy.AddItem txtPrt
      txtPrt = ""
      txtCmt = ""
      txtQty = "0.000"
      txtCvt = "0.000"
      txtAdr = "0.000"
      txtSup = "0.000"
      txtLab = "0.000"
      txtLabOh = "0.000"
      txtMat = "0.000"
      txtMatbr = "0.000"
      txtBum = "EA"
      Sleep 500
      SysMsg "The Item Was Added.", True
   Else
      MsgBox Trim(Err.Descripton) & vbCrLf _
                  & "Couldn't Add The Item.", _
                  vbExclamation, Caption
   End If
   Set RdoAdd = Nothing
   Unload Me
   
End Sub

Private Sub GetThisPart()
   Dim RdoPrt As ADODB.Recordset
   Dim Units As String
   On Error Resume Next
   sSql = "SELECT PARTREF,PAUNITS,PASTDCOST,PATOOL FROM PartTable " _
          & "WHERE (PARTREF='" & Compress(txtPrt) & "' AND PATOOL=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt, ES_FORWARD)
   If bSqlRows Then
      With RdoPrt
         If Not IsNull(.Fields(1)) Then _
                       Units = "" & Trim(!PAUNITS) Else _
                       Units = "EA"
         If Not IsNull(.Fields(2)) Then _
                       txtMat = Format(!PASTDCOST, ES_QuantityDataFormat) Else _
                       txtMat = "0.000"
         ClearResultSet RdoPrt
      End With
   Else
      Units = "EA"
   End If
   If txtBum = "" Or txtBum = "EA" Then txtBum = Units
   Set RdoPrt = Nothing
   
End Sub

Private Sub lblAssy_Change()
   GetIndexHeader
   
End Sub

Private Sub lblDsc_Change()
   If Left(lblDsc, 12) = "*** Part Num" Then _
           lblDsc.ForeColor = ES_RED Else _
           lblDsc.ForeColor = Es_TextForeColor
   
End Sub

Private Sub txtAdr_LostFocus()
   txtAdr = CheckLen(txtAdr, 9)
   txtAdr = Format(Abs(Val(txtAdr)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtBum_Change()
   bChanged = 1
   
End Sub

Private Sub txtBum_LostFocus()
   txtBum = CheckLen(txtBum, 2)
   If txtBum = "" Then txtBum = "EA"
   
End Sub


Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   txtCmt = StrCase(txtCmt, ES_FIRSTWORD)
   
End Sub


Private Sub txtCvt_LostFocus()
   txtCvt = CheckLen(txtCvt, 9)
   txtCvt = Format(Abs(Val(txtCvt)), ES_QuantityDataFormat)
   UpdateMaterial
   
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


Private Sub txtLabOh_LostFocus()
   txtLabOh = CheckLen(txtLabOh, 9)
   txtLabOh = Format(Abs(Val(txtLabOh)), ES_QuantityDataFormat)
   txtLab = Format(Abs(Val(txtLab)), ES_QuantityDataFormat)
   lblLabor = Format(((txtLab)) * ((Val(txtLabOh) / 100) + 1) * Val(txtLabHrs), ES_QuantityDataFormat)
   
End Sub


Private Sub txtMat_LostFocus()
   txtMat = CheckLen(txtMat, 9)
   txtMat = Format(Abs(Val(txtMat)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub


Private Sub txtMatbr_LostFocus()
   txtMatbr = CheckLen(txtMatbr, 9)
   txtMatbr = Format(Abs(Val(txtMatbr)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub




Private Sub txtPrt_LostFocus()
    Dim iCurrPartType As Integer
    
    txtPrt = CheckLen(txtPrt, 30)
    txtPrt = GetCurrentPart(txtPrt, lblDsc)
    iCurrPartType = CurrentPartType(txtPrt)
    If PartOk(txtPrt) = 0 Then
        lblDsc.ForeColor = ES_RED
        lblDsc = "*** Part Number is Invalid ***"
    End If
'    If iCurrPartType < Val(EstiESe02c.lblLvl) Or iCurrPartType > 5 Then
    If iCurrPartType < 2 Or iCurrPartType > 5 Then
        lblDsc.ForeColor = ES_RED
        lblDsc = "*** Part Number is the Wrong Type ***"
    End If
    
    If lblDsc.ForeColor <> ES_RED Then
      GetThisPart
'      If sOldPart <> txtPrt Then FillBomhRev cmbPrt
      sOldPart = txtPrt
      bChanged = 1
   End If
End Sub

Private Sub txtQty_Change()
   bChanged = 1
   
End Sub

Private Sub txtQty_LostFocus()
   txtQty = CheckLen(txtQty, 9)
   txtQty = Format(Abs(Val(txtQty)), ES_QuantityDataFormat)
   If Val(txtQty) > 0 Then _
          If lblDsc.ForeColor <> ES_RED Then cmdAdd.Enabled = True
   UpdateMaterial
   
End Sub


Private Sub txtSeq_LostFocus()
   txtSeq = CheckLen(txtSeq, 3)
   txtSeq = Format$(Abs(Val(txtSeq)), "##0")
   
End Sub



Private Sub UpdateMaterial()
   Dim cAdder As Currency
   Dim cBurden As Currency
   Dim cConvert As Currency
   Dim cCost As Currency
   Dim cQuantity As Currency
   Dim cSetup As Currency
   
   On Error GoTo 0
   cAdder = Val(txtAdr)
   cBurden = (Val(txtMatbr) / 100) + 1
   cCost = Val(txtMat)
   cSetup = Val(txtSup)
   cConvert = Format(Val(txtCvt), ES_QuantityDataFormat)
   If cConvert = 0 Then cConvert = 1
   cQuantity = Format((Val(txtQty) + cAdder + cSetup), ES_QuantityDataFormat)
   cQuantity = cQuantity / cConvert
   lblQty = Format(cQuantity, ES_QuantityDataFormat)
   cBurden = cBurden * cCost
   lblMaterial = Format(cBurden * cQuantity, ES_QuantityDataFormat)
   
End Sub

Private Sub txtSup_LostFocus()
   txtSup = CheckLen(txtSup, 9)
   txtSup = Format(Abs(Val(txtSup)), ES_QuantityDataFormat)
   UpdateMaterial
   
End Sub

Private Function PartOk(ByVal sPartNum As String) As Byte
    Dim iNodes As Integer
    Dim rdoPrtTool As ADODB.Recordset
    
    PartOk = 1
    If Compress(sPartNum) = Compress(lblAssy) Then
        PartOk = 0
        Exit Function
    End If
    For iNodes = 0 To EstiESe02c.lstNodes.ListCount - 1
        If txtPrt = EstiESe02c.lstNodes.List(iNodes) Then
            PartOk = 0
            Exit Function
        End If
    Next iNodes
    
    sSql = "SELECT PARTREF FROM PartTable WHERE PartREF='" & Compress(sPartNum) & "' AND PATOOL=0"
    If clsADOCon.GetDataSet(sSql, rdoPrtTool, ES_FORWARD) <> 1 Then PartOk = 0
    ClearResultSet rdoPrtTool
    Set rdoPrtTool = Nothing
    
End Function




