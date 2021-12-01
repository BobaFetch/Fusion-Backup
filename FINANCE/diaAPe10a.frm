VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaAPe10a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Check Memos"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMem 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "(40) Characters"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.ComboBox cmbChk 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "List Of Checks Not Printed"
      Top             =   720
      Width           =   1775
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
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
      PictureUp       =   "diaAPe10a.frx":0000
      PictureDn       =   "diaAPe10a.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5400
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3000
      FormDesignWidth =   6210
   End
   Begin VB.Label lblChkAcct 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   1200
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Account"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblAmt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   3465
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "diaAPe10a"
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
Dim RdoChk As ADODB.Recordset
Dim bOnLoad As Byte
Dim bCanceled As Boolean

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbChk_Click()
   GetCheckVendor
End Sub

Private Sub cmbChk_LostFocus()
   cmbChk = CheckLen(cmbChk, 12)
   If bCanceled Then Exit Sub
   If Len(Trim(cmbChk)) > 0 Then GetCheckVendor
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
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
   FormUnload
   Set RdoChk = Nothing
   Set diaAPe10a = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Public Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CHKNUMBER FROM ChksTable WHERE (CHKPRINTED=0 " _
          & "AND CHKVOID=0)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbChk.hwnd, "" & Trim(!CHKNUMBER)
            .MoveNext
         Loop
         .Cancel
      End With
   End If
   If cmbChk.ListCount > 0 Then
      cmbChk = cmbChk.List(0)
      GetCheckVendor
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub GetCheckVendor()
   Dim RdoVnd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT CHKNUMBER,CHKVENDOR,CHKACCT, VEREF,VEBNAME " _
          & "FROM ChksTable,VndrTable WHERE (CHKNUMBER='" _
          & Trim(cmbChk) & "' AND CHKVENDOR=VEREF)"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd, ES_FORWARD)
   If bSqlRows Then
      With RdoVnd
         lblVnd = "" & Trim(!VEBNAME)
         lblChkAcct = "" & Trim(!CHKACCT)
         .Cancel
      End With
   Else
      lblVnd = ""
   End If
   If lblVnd <> "" Then
      sSql = "SELECT CHKNUMBER,CHKPOSTDATE,CHKAMOUNT,CHKMEMO " _
             & "FROM ChksTable WHERE CHKNUMBER='" & Trim(cmbChk) _
             & "' AND CHKACCT = '" & Trim(lblChkAcct) & "'"
             
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoChk, ES_KEYSET)
      If bSqlRows Then
         With RdoChk
            txtMem = "" & Trim(!chkMemo)
            lblDte = Format(!CHKPOSTDATE, "mm/dd/yy")
            lblAmt = Format(!CHKAMOUNT, "########0.00")
            .Cancel
         End With
      End If
   Else
      lblVnd = "*** Check Was Not Found, Printed Or Void ***"
      txtMem = """"
      lblDte = ""
      lblAmt = ""
   End If
   Set RdoVnd = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getcheckve"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblVnd_Change()
   If Left(lblVnd, 6) = "*** Ch" Then
      lblVnd.ForeColor = ES_RED
   Else
      lblVnd.ForeColor = vbBlack
   End If
End Sub

Private Sub txtMem_LostFocus()
   txtMem = StrCase(CheckLen(txtMem, 40))
   On Error Resume Next
   RdoChk!chkMemo = "" & txtMem
   RdoChk.Update
   
End Sub
