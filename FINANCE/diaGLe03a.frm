VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaGLe03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Financial Statement Structure"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFdtxAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   17
      Tag             =   "3"
      Top             =   4920
      Width           =   1320
   End
   Begin VB.TextBox txtOexpAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   15
      Tag             =   "3"
      Top             =   4320
      Width           =   1320
   End
   Begin VB.TextBox txtOincAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   13
      Tag             =   "3"
      Top             =   3960
      Width           =   1320
   End
   Begin VB.TextBox txtExpnAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   11
      Tag             =   "3"
      Top             =   3240
      Width           =   1320
   End
   Begin VB.TextBox txtCOGSAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "3"
      Top             =   2520
      Width           =   1320
   End
   Begin VB.TextBox txtIncmAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "3"
      Top             =   2160
      Width           =   1320
   End
   Begin VB.TextBox txtEqtyAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   5
      Tag             =   "3"
      Top             =   1560
      Width           =   1320
   End
   Begin VB.TextBox txtLiabAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1080
      Width           =   1320
   End
   Begin VB.TextBox txtAsstAcct 
      Height          =   285
      Left            =   3840
      MaxLength       =   12
      TabIndex        =   1
      Tag             =   "3"
      Top             =   720
      Width           =   1320
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3360
      TabIndex        =   19
      ToolTipText     =   "Update Structure And Associated Entries"
      Top             =   120
      Visible         =   0   'False
      Width           =   875
   End
   Begin VB.TextBox txtFdtx 
      Height          =   285
      Left            =   720
      TabIndex        =   16
      Tag             =   "3"
      Top             =   4920
      Width           =   3000
   End
   Begin VB.TextBox txtOexp 
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Tag             =   "3"
      Top             =   4320
      Width           =   3000
   End
   Begin VB.TextBox txtOinc 
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Tag             =   "3"
      Top             =   3960
      Width           =   3000
   End
   Begin VB.TextBox txtExpn 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Tag             =   "3"
      Top             =   3240
      Width           =   3000
   End
   Begin VB.TextBox txtCogs 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Tag             =   "3"
      Top             =   2520
      Width           =   3000
   End
   Begin VB.TextBox txtIncm 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Tag             =   "3"
      Top             =   2160
      Width           =   3000
   End
   Begin VB.TextBox txtEqty 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox txtLiab 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1080
      Width           =   3000
   End
   Begin VB.TextBox txtAsst 
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Tag             =   "3"
      Top             =   720
      Width           =   3000
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   20
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
      PictureUp       =   "diaGLe03a.frx":0000
      PictureDn       =   "diaGLe03a.frx":0146
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Master Account"
      Height          =   255
      Index           =   19
      Left            =   3840
      TabIndex        =   50
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   49
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   48
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "+"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   47
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   46
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Operating Profit (With Cost of Goods and Other Expense)"
      Height          =   210
      Index           =   15
      Left            =   720
      TabIndex        =   45
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   44
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Profit (Pretax Net)"
      Height          =   210
      Index           =   13
      Left            =   720
      TabIndex        =   43
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   42
      Top             =   3650
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-------------------------------------"
      Height          =   210
      Index           =   11
      Left            =   720
      TabIndex        =   41
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Operating Profit (With Cost of Goods and Other Expense)"
      Height          =   210
      Index           =   10
      Left            =   720
      TabIndex        =   40
      Top             =   3650
      Width           =   5415
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   39
      Top             =   2900
      Width           =   375
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Profit (If Cost of Goods Sold Included)"
      Height          =   210
      Index           =   8
      Left            =   720
      TabIndex        =   38
      Top             =   2900
      Width           =   4215
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter/Revise Financial Statement Structure:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   37
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter/Revise Balance Sheet Structure:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   36
      Top             =   420
      Width           =   2835
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   34
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-------------------------------------"
      Height          =   210
      Index           =   3
      Left            =   720
      TabIndex        =   33
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-------------------------------------"
      Height          =   150
      Index           =   1
      Left            =   720
      TabIndex        =   31
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Z1 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblFdtx 
      BackStyle       =   0  'Transparent
      Caption         =   "Fed Inc Tax"
      Height          =   255
      Left            =   5280
      TabIndex        =   29
      Top             =   4980
      Width           =   1200
   End
   Begin VB.Label lblOexp 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Exp"
      Height          =   255
      Left            =   5280
      TabIndex        =   28
      Top             =   4380
      Width           =   1200
   End
   Begin VB.Label lblOinc 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Inc"
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   4020
      Width           =   1200
   End
   Begin VB.Label lblExpn 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense"
      Height          =   255
      Left            =   5280
      TabIndex        =   26
      Top             =   3300
      Width           =   1200
   End
   Begin VB.Label lblCogs 
      BackStyle       =   0  'Transparent
      Caption         =   "COGS"
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   2580
      Width           =   1200
   End
   Begin VB.Label lblIncm 
      BackStyle       =   0  'Transparent
      Caption         =   "Income"
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   2220
      Width           =   1200
   End
   Begin VB.Label lblEqty 
      BackStyle       =   0  'Transparent
      Caption         =   "Equity"
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   1620
      Width           =   1200
   End
   Begin VB.Label lblLiab 
      BackStyle       =   0  'Transparent
      Caption         =   "Liability"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   1140
      Width           =   1200
   End
   Begin VB.Label lblAsst 
      BackStyle       =   0  'Transparent
      Caption         =   "Assets"
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   780
      Width           =   1200
   End
End
Attribute VB_Name = "diaGLe03a"
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

'**************************************************************************************
' diaGLe03a - Create/revise financial statment structure
'
' Notes:
'
' Created: (cjs)
' Revisions:
'   02/01/02 (nth) Determine if the statment structure has never been setup.
'                  If so then create a new structure rather than update a existing one
'
'**************************************************************************************

Dim bOnLoad As Byte
Dim bUpdated As Byte

Dim sAccount(10, 3) As String
Const SACCT_InitialAcctNo = 0
Const SACCT_FinalAcctNo = 1
Const SACCT_CompressedFinalAcctNo = 2 ' no dashes etc.


Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'**************************************************************************************
Private Sub MakeEverythingReadOnly()
    Dim MyControl As Control ' Decalre Variable to hold the Controls on the Form
    For Each MyControl In Me.Controls  ' Loop through the Controls on the Form
        If TypeOf MyControl Is TextBox Then  ' If the Control is a TextBox
            MyControl.enabled = False
            MyControl.BackColor = &H80000016
        End If
    Next

End Sub





Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, "Financial Statement Structure"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   'Dim b As Byte
   Dim i As Integer
   bUpdated = True
   For i = 1 To 9
      If Len(sAccount(i, SACCT_FinalAcctNo)) = 0 Then
         MsgBox "You must define all master accounts even if you do not intend to use them"
         Exit Sub
      End If
      'If sAccount(i, SACCT_InitialAcctNo) <> sAccount(i, SACCT_FinalAcctNo) Then b = True
   Next
   '    If sAccount(i, SACCT_InitialAcctNo) <> sAccount(i, SACCT_FinalAcctNo) Then b = True
   '    If Not b Then
   '        MsgBox "The Accounts Haven't Changed.", _
   '            vbInformation, Caption
   '    Else
   '        UpdateAccounts
   '    End If
   
   'update regardless.  maybe just the description has changed
   UpdateAccounts
   
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      SetDefaultAccounts
      FillBoxes
      bOnLoad = False
   End If
   MouseCursor 0
   MakeEverythingReadOnly
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   sCurrForm = Caption
   bOnLoad = True
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim b As Byte
   Dim bResponse As Byte
   Dim i As Integer
   Dim sMsg As String
   For i = 1 To 9
      If sAccount(i, SACCT_InitialAcctNo) <> sAccount(i, SACCT_FinalAcctNo) Then b = True
   Next
   'If sAccount(i, SACCT_InitialAcctNo) <> sAccount(i, SACCT_FinalAcctNo) Then b = True
   If b And Not bUpdated Then
      sMsg = "The Structure Has Changed." & vbCrLf _
             & "Do You Want Exit Without Saving?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Cancel = True
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaGLe03a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Public Sub FillBoxes()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm)
   If bSqlRows Then
      With RdoGlm
         txtAsstAcct = "" & Trim(!COASSTACCT)
         txtAsst = "" & Trim(!COASSTDESC)
         sAccount(1, SACCT_InitialAcctNo) = txtAsstAcct
         
         txtLiabAcct = "" & Trim(!COLIABACCT)
         txtLiab = "" & Trim(!COLIABDESC)
         sAccount(2, SACCT_InitialAcctNo) = txtLiabAcct
         
         txtEqtyAcct = "" & Trim(!COEQTYACCT)
         txtEqty = "" & Trim(!COEQTYDESC)
         sAccount(3, SACCT_InitialAcctNo) = txtEqtyAcct
         
         txtIncmAcct = "" & Trim(!COINCMACCT)
         txtIncm = "" & Trim(!COINCMDESC)
         sAccount(4, SACCT_InitialAcctNo) = txtIncmAcct
         
         txtCOGSAcct = "" & Trim(!COCOGSACCT)
         txtCogs = "" & Trim(!COCOGSDESC)
         sAccount(5, SACCT_InitialAcctNo) = txtCOGSAcct
         
         txtExpnAcct = "" & Trim(!COEXPNACCT)
         txtExpn = "" & Trim(!COEXPNDESC)
         sAccount(6, SACCT_InitialAcctNo) = txtExpnAcct
         
         txtOincAcct = "" & Trim(!COOINCACCT)
         txtOinc = "" & Trim(!COOINCDESC)
         sAccount(7, SACCT_InitialAcctNo) = txtOincAcct
         
         txtOexpAcct = "" & Trim(!COOEXPACCT)
         txtOexp = "" & Trim(!COOEXPDESC)
         sAccount(8, SACCT_InitialAcctNo) = txtOexpAcct
         
         txtFdtxAcct = "" & Trim(!COFDTXACCT)
         txtFdtx = "" & Trim(!COFDTXDESC)
         sAccount(9, SACCT_InitialAcctNo) = txtFdtxAcct
         .Cancel
      End With
   End If
   For i = 1 To 9
      sAccount(i, SACCT_FinalAcctNo) = sAccount(i, SACCT_InitialAcctNo)
   Next
   sAccount(i, SACCT_FinalAcctNo) = sAccount(i, SACCT_InitialAcctNo)
   Set RdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "fillboxes"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub SetDefaultAccounts()
   Dim i As Integer
   Dim RdoGlm As ADODB.Recordset
   On Error GoTo DiaErr1
   MouseCursor 13
   sSql = "SELECT * FROM GlmsTable WHERE COACCTREC=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGlm, ES_KEYSET)
   Dim changed As Boolean
   changed = False
   If Not bSqlRows Then
      RdoGlm.AddNew 'no gl master record -- add it now
      RdoGlm!COACCTREC = 1
      changed = True
   End If
   
   With RdoGlm
      If Len("" & Trim(!COASSTACCT)) = 0 Then
         !COASSTACCT = "1"
         !COASSTREF = "1"
         !COASSTDESC = "Assets"
         changed = True
      End If
      
      If Len("" & Trim(!COLIABACCT)) = 0 Then
         !COLIABACCT = "2"
         !COLIABREF = "2"
         !COLIABDESC = "Liabilities"
         changed = True
      End If
      
      If Len("" & Trim(!COEQTYACCT)) = 0 Then
         !COEQTYACCT = "3"
         !COEQTYREF = "3"
         !COEQTYDESC = "Equity"
         changed = True
      End If
      
      If Len("" & Trim(!COINCMACCT)) = 0 Then
         !COINCMACCT = "4"
         !COINCMREF = "4"
         !COINCMDESC = "Income"
         changed = True
      End If
      
      If Len("" & Trim(!COCOGSACCT)) = 0 Then
         !COCOGSACCT = "5"
         !COCOGSREF = "5"
         !COCOGSDESC = "Cost of Goods Sold"
         changed = True
      End If
      
      If Len("" & Trim(!COEXPNACCT)) = 0 Then
         !COEXPNACCT = "6"
         !COEXPNREF = "6"
         !COEXPNDESC = "Expense"
         changed = True
      End If
      
      If Len("" & Trim(!COOINCACCT)) = 0 Then
         !COOINCACCT = "7"
         !COOINCREF = "7"
         !COOINCDESC = "Other Income"
         changed = True
      End If
      
      If Len("" & Trim(!COOEXPACCT)) = 0 Then
         !COOEXPACCT = "8"
         !COOEXPREF = "8"
         !COOEXPDESC = "Other Expense"
         changed = True
      End If
      
      If Len("" & Trim(!COFDTXACCT)) = 0 Then
         !COFDTXACCT = "9"
         !COFDTXREF = "9"
         !COFDTXDESC = "Federal Income Tax"
         changed = True
      End If
      
   End With
   
   If changed Then
      RdoGlm.Update
   Else
      RdoGlm.Cancel
   End If
   Set RdoGlm = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "SetDefaultAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub txtAsstAcct_LostFocus()
   '1
   'txtAsst = CheckLen(txtAsst, 12)
   '    If Len(txtAsst) = 0 Then
   '        Beep
   '        txtAsstAcct = sAccount(1, SACCT_InitialAcctNo)
   '    End If
   '    CheckName txtAsstAcct, 1
   sAccount(1, SACCT_FinalAcctNo) = txtAsstAcct
   
End Sub


Private Sub txtCogsAcct_LostFocus()
   '5
   'txtCogs = CheckLen(txtCogsAcct, 12)
   '    If Len(txtCogs) = 0 Then Beep
   '    If Len(txtCOGSAcct) Then CheckName txtCOGSAcct, 5
   sAccount(5, SACCT_FinalAcctNo) = txtCOGSAcct
   
End Sub


Private Sub txtEqtyAcct_LostFocus()
   '3
   'txtEqty = CheckLen(txtEqtyAcct, 12)
   '    If Len(txtEqtyAcct) = 0 Then
   '        Beep
   '        txtEqtyAcct = sAccount(3, SACCT_InitialAcctNo)
   '    End If
   '    CheckName txtEqtyAcct, 3
   sAccount(3, SACCT_FinalAcctNo) = txtEqtyAcct
   
End Sub


Private Sub txtExpnAcct_LostFocus()
   '6
   'txtExpnAcct = CheckLen(txtExpnAcct, 12)
   '    If Len(txtExpnAcct) = 0 Then
   '        Beep
   '        txtExpnAcct = sAccount(6, SACCT_InitialAcctNo)
   '    End If
   '    CheckName txtExpnAcct, 6
   sAccount(6, SACCT_FinalAcctNo) = txtExpnAcct
   
End Sub


Private Sub txtFdtxAcct_LostFocus()
   '9
   'txtFdtxAcct = CheckLen(txtFdtxAcct, 12)
   '    If Len(txtFdtxAcct) = 0 Then Beep
   '    If Len(txtFdtxAcct) Then CheckName txtFdtxAcct, 9
   sAccount(9, SACCT_FinalAcctNo) = txtFdtxAcct
   
End Sub

Private Sub txtIncmAcct_LostFocus()
   '4
   'txtIncmAcct = CheckLen(txtIncmAcct, 12)
   '    If Len(txtIncmAcct) = 0 Then
   '        Beep
   '        txtIncmAcct = sAccount(4, SACCT_InitialAcctNo)
   '    End If
   '    CheckName txtIncmAcct, 4
   sAccount(4, SACCT_FinalAcctNo) = txtIncmAcct
   
End Sub


Private Sub txtLiabAcct_LostFocus()
   '2
   'txtLiabAcct = CheckLen(txtLiabAcct, 12)
   '    If Len(txtLiabAcct) = 0 Then
   '        Beep
   '        txtLiabAcct = sAccount(2, SACCT_InitialAcctNo)
   '    End If
   '    CheckName txtLiabAcct, 2
   sAccount(2, SACCT_FinalAcctNo) = txtLiabAcct
   
End Sub


Private Sub txtOexpAcct_LostFocus()
   '8
   'txtOexpAcct = CheckLen(txtOexpAcct, 12)
   '    If Len(txtOexpAcct) = 0 Then Beep
   '    If Len(txtOexpAcct) Then CheckName txtOexpAcct, 8
   sAccount(8, SACCT_FinalAcctNo) = txtOexpAcct
   
End Sub


Private Sub txtOincAcct_LostFocus()
   '7
   'txtOincAcct = CheckLen(txtOincAcct, 12)
   '    If Len(txtOincAcct) = 0 Then Beep
   '    If Len(txtOincAcct) Then CheckName txtOincAcct, 7
   sAccount(7, SACCT_FinalAcctNo) = txtOincAcct
   
End Sub

Public Sub UpdateAccounts()
   '    Dim b         As Byte
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   MouseCursor 13
   cmdUpd.enabled = False
   For i = 1 To 9
      sAccount(i, SACCT_CompressedFinalAcctNo) = Compress(sAccount(i, SACCT_FinalAcctNo))
   Next
   MouseCursor 13
   
   sSql = "UPDATE GlmsTable SET " _
          & "COASSTREF='" & Compress(txtAsstAcct.Text) & "'," _
          & "COASSTACCT='" & txtAsstAcct.Text & "'," _
          & "COASSTDESC='" & txtAsst.Text & "'," _
          & "COLIABREF='" & Compress(txtLiabAcct.Text) & "'," _
          & "COLIABACCT='" & txtLiabAcct.Text & "'," _
          & "COLIABDESC='" & txtLiab.Text & "'," _
          & "COEQTYREF='" & Compress(txtEqtyAcct.Text) & "'," _
          & "COEQTYACCT='" & txtEqtyAcct.Text & "'," _
          & "COEQTYDESC='" & txtEqty.Text & "'," _
          & "COINCMREF='" & Compress(txtIncmAcct.Text) & "'," _
          & "COINCMACCT='" & txtIncmAcct.Text & "'," _
          & "COINCMDESC='" & txtIncm.Text & "',"
   sSql = sSql _
          & "COCOGSREF='" & Compress(txtCOGSAcct.Text) & "'," _
          & "COCOGSACCT='" & txtCOGSAcct.Text & "'," _
          & "COCOGSDESC='" & txtCogs.Text & "'," _
          & "COEXPNREF='" & Compress(txtExpnAcct.Text) & "'," _
          & "COEXPNACCT='" & txtExpnAcct.Text & "'," _
          & "COEXPNDESC='" & txtExpn.Text & "'," _
          & "COOINCREF='" & Compress(txtOincAcct.Text) & "'," _
          & "COOINCACCT='" & txtOincAcct.Text & "'," _
          & "COOINCDESC='" & txtOinc.Text & "'," _
          & "COOEXPREF='" & Compress(txtOexpAcct.Text) & "'," _
          & "COOEXPACCT='" & txtOexpAcct.Text & "'," _
          & "COOEXPDESC='" & txtOexp & "'," _
          & "COFDTXREF='" & Compress(txtFdtxAcct.Text) & "'," _
          & "COFDTXACCT='" & txtFdtxAcct.Text & "'," _
          & "COFDTXDESC='" & txtFdtx & "' " _
          & "WHERE COACCTREC=1 "
   
   clsADOCon.ExecuteSQL sSql
   If clsADOCon.RowsAffected > 0 Then
      For i = 1 To 9
         If sAccount(i, SACCT_InitialAcctNo) <> sAccount(i, SACCT_FinalAcctNo) Then
            'If sAccount(i, SACCT_FinalAcctNo) = "" Then b = 1 Else b = 0
            '                sSql = "UPDATE GlacTable SET " _
            '                    & "GLMASTER='" & sAccount(i, SACCT_CompressedFinalAcctNo) & "'," _
            '                    & "GLINACTIVE=" & b & " " _
            '                    & "WHERE GLMASTER='" & sAccount(i, SACCT_InitialAcctNo) & "' "
            sSql = "UPDATE GlacTable SET " _
                   & "GLMASTER='" & sAccount(i, SACCT_CompressedFinalAcctNo) & "'" & vbCrLf _
                   & "WHERE rtrim(GLMASTER)='" & sAccount(i, SACCT_InitialAcctNo) & "' "
            Debug.Print sSql
            clsADOCon.ExecuteSQL sSql
         End If
      Next
      MouseCursor 0
      bUpdated = True
      MsgBox "Structure Successfully Updated.", _
         vbInformation, Caption
      Unload Me
   Else
      MouseCursor 0
      MsgBox "Couldn't Update Structure.", _
         vbExclamation, Caption
      bUpdated = False
      cmdUpd.enabled = True
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "UpdateAccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Sub CheckName(sAccounts As String, Index As Integer)
   Dim b As Byte
   Dim i As Integer
   On Error Resume Next
   For i = 0 To 8
      If i <> Index Then
         If sAccount(i, SACCT_FinalAcctNo) = sAccounts Then b = 1
      End If
   Next
   If b = 1 Then
      Beep
      Select Case Index
         Case 1
            txtAsst = sAccount(1, SACCT_FinalAcctNo)
         Case 2
            txtLiab = sAccount(2, SACCT_FinalAcctNo)
         Case 3
            txtEqty = sAccount(3, SACCT_FinalAcctNo)
         Case 4
            txtIncm = sAccount(4, SACCT_FinalAcctNo)
         Case 5
            txtCogs = sAccount(5, SACCT_FinalAcctNo)
         Case 6
            txtExpn = sAccount(6, SACCT_FinalAcctNo)
         Case 7
            txtOinc = sAccount(7, SACCT_FinalAcctNo)
         Case 8
            txtOexp = sAccount(8, SACCT_FinalAcctNo)
         Case Else
            txtFdtx = sAccount(8, SACCT_FinalAcctNo)
      End Select
   End If
   
End Sub
