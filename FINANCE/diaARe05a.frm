VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaARe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Invoice Comments"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCalledByCR 
      Caption         =   "chkCalledByCR"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtCmt 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   8
      Tag             =   "9"
      Top             =   1440
      Width           =   5295
   End
   Begin VB.ComboBox cmbInv 
      Height          =   315
      ItemData        =   "diaARe05a.frx":0000
      Left            =   1920
      List            =   "diaARe05a.frx":0002
      TabIndex        =   2
      Tag             =   "3"
      Top             =   240
      Width           =   1125
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Subject Help"
      Top             =   0
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
      PictureUp       =   "diaARe05a.frx":0004
      PictureDn       =   "diaARe05a.frx":014A
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   3480
      Top             =   360
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2910
      FormDesignWidth =   5670
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
      Height          =   285
      Index           =   2
      Left            =   270
      TabIndex        =   7
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label lblPre 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   1320
   End
End
Attribute VB_Name = "diaARe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of                     ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA          ***
'*** and is protected under US and International copyright             ***
'*** laws and treaties.                                                ***

'See the UpdateTables prodecure for database revisions

Option Explicit

'*************************************************************************************
' diaARe05a - Revise Invoice Comments
'
' Created: (cjs)
'
' Revisions:
'   06/14/04 (nth) Changed keyset cursor to static.
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bGoodInv As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Private Sub cmbInv_Click()
   bGoodInv = GetInvoice()
End Sub

Private Sub cmbInv_LostFocus()
   cmbInv = CheckLen(cmbInv, 6)
   cmbInv = Format(Abs(Val(cmbInv)), "000000")
   bGoodInv = GetInvoice()
End Sub

Private Sub cmdCan_Click()
   On Error Resume Next
   txtCmt_LostFocus
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Customer Invoice (Sales Order)"
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If chkCalledByCR = vbUnchecked Then FormUnload
   Set diaARe05a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT INVNO FROM CihdTable WHERE INVTYPE<>'TM' " _
          & "AND INVCANCELED=0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         cmbInv = Format(!invno, "000000")
         Do Until .EOF
            AddComboStr cmbInv.hwnd, Format$(!invno, "000000")
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbInv.ListCount > 0 Then
      cmbInv = cmbInv.List(0)
      bGoodInv = GetInvoice()
   End If
   Set RdoCmb = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetInvoice() As Byte
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT INVNO,INVPRE,INVCUST,INVCOMMENTS FROM CihdTable " _
          & "WHERE INVNO=" & Val(cmbInv) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_STATIC)
   If bSqlRows Then
      With RdoInv
         cmbInv = Format(!invno, "000000")
         lblPre = "" & Trim(!INVPRE)
         txtCmt = "" & Trim(!INVCOMMENTS)
         FindCustomer Me, "" & Trim(!INVCUST)
         .Cancel
      End With
      GetInvoice = True
   Else
      lblPre = ""
      txtCmt = ""
      lblCst = ""
      txtNme = "*** No Current Invoice ***"
      GetInvoice = False
   End If
   Set RdoInv = Nothing
   Exit Function
DiaErr1:
   sProcName = "getinvoice"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   lblCst = ""
   txtNme = "*** No Current Invoice ***"
   DoModuleErrors Me
End Function

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 2048)
   txtCmt = CheckComments(txtCmt)
   If bGoodInv Then
      sSql = "UPDATE CihdTable SET INVCOMMENTS='" _
             & txtCmt & "' WHERE INVNO=" & Val(cmbInv) & " "
      clsADOCon.ExecuteSQL sSql
   End If
End Sub

Private Sub txtNme_Change()
   If Left(txtNme, 3) = "***" Then
      txtNme.ForeColor = ES_RED
   Else
      txtNme.ForeColor = vbBlack
   End If
End Sub
