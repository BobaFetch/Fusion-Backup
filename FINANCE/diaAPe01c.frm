VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form diaAPe01c 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add GL Account Distribution"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox cmbRun 
      Height          =   315
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "9"
      ToolTipText     =   "Select Run From List"
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox cmbMon 
      Height          =   315
      Left            =   600
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "9"
      ToolTipText     =   "Blank, Enter Run or Select "
      Top             =   2280
      Width           =   2775
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4320
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2880
      FormDesignWidth =   5895
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtCom 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbAct 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      ToolTipText     =   "Add Item To Invoice"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4920
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Close Without Posting"
      Top             =   120
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   10
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
      PictureUp       =   "diaAPe01c.frx":0000
      PictureDn       =   "diaAPe01c.frx":0146
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "M.O."
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      Height          =   255
      Index           =   19
      Left            =   3600
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account                        "
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
      Index           =   6
      Left            =   3000
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount               "
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
      Index           =   8
      Left            =   4560
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description                                                                "
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
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
   End
End
Attribute VB_Name = "diaAPe01c"
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
'
' diaAPe01c - Add AP invoice item not on PO
'
' Created: 08/01/02 (nth)
' Revisions:
'   10/18/02 (nth) Completely reworked
'   01/24/03 (nth) Set focus back to diaAPe01b on form unload
'   05/20/04 (nth) Added run allocations per ANDELE
'   05/20/04 (nth) removed gl description requirement per DAP
'
'*************************************************************************************

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodRun As Byte
Dim bPosted As Boolean

Dim sNote As String
Dim sAccount As String
Dim sApAcct As String
Dim sFrAcct As String
Dim sTxAcct As String
Dim sMsg As String
Private whereClause As String       'where clause for MO list

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'*************************************************************************************

Public Sub FillAccounts()
   Dim rdoAct As ADODB.Recordset
   Dim RdoVnd As ADODB.Recordset
   Dim b As Byte
   Dim i As Integer
   
   On Error GoTo DiaErr1
   
   sJournalID = GetOpenJournal("PJ", Format(diaAPe01a.txtPdt, "mm/dd/yyyy"))
   If Left(sJournalID, 4) <> "None" And sJournalID = "" Then
      MsgBox "There Is No Open Purchases Journal For The Period.", _
         vbInformation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   If sJournalID <> "" Then
      b = GetDBAccounts(sApAcct, sFrAcct, sTxAcct)
      If b = 0 Then
         MsgBox "One Or More Of The AP Accounts Required Is Not Installed." & vbCr _
            & "Please Install All Accounts In The AP Tab Of Company Setup.", _
            vbInformation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
   End If
   
   sSql = "Qry_FillLowAccounts"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoAct, ES_FORWARD)
   If bSqlRows Then
      With rdoAct
         While Not .EOF
            AddComboStr cmbAct.hWnd, "" & Trim(!GLACCTNO)
            .MoveNext
         Wend
         .Cancel
      End With
   End If
   Set rdoAct = Nothing
   
   ' Get the default vendor account
   sSql = "SELECT GLACCTNO FROM VndrTable,GlacTable WHERE VEACCOUNT=" _
          & "GLACCTREF AND VEREF='" & Compress(lblVnd) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoVnd)
   If bSqlRows Then
      With RdoVnd
         sAccount = "" & Trim(.Fields(0))
         .Cancel
      End With
   End If
   Set RdoVnd = Nothing
   If sAccount = "" Then
      sAccount = GetSetting("Esi2000", "Fina", "LastAccount", sAccount)
   End If
   cmbAct.Text = sAccount
   lblDsc = UpdateActDesc(cmbAct)
   MouseCursor 0
   Exit Sub
DiaErr1:
   sProcName = "fillaccou"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Public Function GetDBAccounts(ApAcct As String, FrAcct As String, TxAcct As String) As Byte
   Dim RdoCdm As ADODB.Recordset
   Dim b As Byte
   Dim i As Integer
   
   On Error GoTo DiaErr1
   sSql = "SELECT COAPACCT,COPJTAXACCT,COPJTFRTACCT FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCdm, ES_FORWARD)
   If bSqlRows Then
      With RdoCdm
         For i = 0 To 2
            If Not IsNull(.Fields(i)) Then
               If Trim(.Fields(i)) = "" Then b = 1
            Else
               b = 1
            End If
         Next
         ApAcct = "" & Trim(!COAPACCT)
         TxAcct = "" & Trim(!COPJTAXACCT)
         FrAcct = "" & Trim(!COPJTFRTACCT)
         .Cancel
      End With
      If b = 0 Then GetDBAccounts = 1
   Else
      ApAcct = ""
      FrAcct = ""
      TxAcct = ""
      GetDBAccounts = 0
   End If
   Set RdoCdm = Nothing
   Exit Function
DiaErr1:
   sProcName = "getdbaccounts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub cmbAct_Click()
   lblDsc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmbAct_DropDown()
   'Accounts.Show
   'ShowAccounts Me
End Sub

Private Sub cmbAct_LostFocus()
   lblDsc = UpdateActDesc(cmbAct)
End Sub

Private Sub cmdCan_Click()
   Unload Me
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      SelectHelpTopic Me, Caption
      cmdHlp = False
      MouseCursor 0
   End If
End Sub

Private Sub cmdUpdate_Click()
   sMsg = ""
   If Left(lblDsc, 3) = "***" Then
      sMsg = "Valid GL Account Required."
   Else
      diaAPe01b.AddGLDis
   End If
   If Trim(sMsg) <> "" Then
      MsgBox sMsg, vbInformation, Caption
   End If
End Sub

Private Sub Form_Activate()
    ' do not allow return to calling form until this form is closed
    diaAPe01b.enabled = False
    
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      'FillCombo
      whereClause = "where RUNSTATUS <>'CA' and  RUNSTATUS <>'CL'"
      FillMoPartCombo Me.cmbMon, Me.cmbRun, whereClause, True
      cmbMon.Text = "<NONE>"
      FillAccounts
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Load()
   FormLoad Me
   Move diaAPe01b.Left + 200, diaAPe01b.Top + 200
   FormatControls
   sNote = GetSetting("Esi2000", "Fina", "LastNote", sNote)
   lblVnd = diaAPe01b.lblVnd
   lblInv = diaAPe01b.lblInv
   txtAmt = "0.00"
   bOnLoad = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   SaveSetting "Esi2000", "Fina", "LastNote", txtCom
   SaveSetting "Esi2000", "Fina", "LastAccount", cmbAct
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   FormUnload 1
   diaAPe01b.enabled = True

   diaAPe01b.SetFocus
   Set diaAPe01c = Nothing
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
End Sub

Private Sub txtAmt_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtAmt_LostFocus()
   'txtAmt = Format(txtAmt, CURRENCYMASK)
   CheckCurrencyTextBox txtAmt, False
End Sub

Private Sub txtCom_GotFocus()
   SelectFormat Me
End Sub

Private Sub txtCom_LostFocus()
   txtCom = StrCase(CheckComments(CheckLen(txtCom, 30)))
End Sub

'Private Function GetRuns() As Byte
'   Dim RdoRns As ADODB.RecordSet
'   Dim SPartRef As String
'   cmbRun.Clear
'   SPartRef = Compress(cmbMon)
'   If Len(SPartRef) > 0 Then
'      On Error GoTo DiaErr1
'      sSql = "SELECT PARTREF,PARTNUM,PADESC,RUNREF,RUNSTATUS," _
'             & "RUNNO FROM PartTable,RunsTable WHERE PARTREF='" _
'             & SPartRef & "' AND PARTREF=RUNREF AND RUNSTATUS<>'CA'"
'      bSqlRows = clsAdoCon.GetDataSet(sSql,RdoRns)
'      If bSqlRows Then
'         With RdoRns
'            cmbRun = Format(0 + !RunNo, "####0")
'            Do Until .EOF
'               cmbRun.AddItem Format(0 + !RunNo, "####0")
'               .MoveNext
'            Loop
'            .Cancel
'         End With
'         GetRuns = True
'      Else
'         SPartRef = ""
'         GetRuns = False
'      End If
'   End If
'   Set RdoRns = Nothing
'   Exit Function
'DiaErr1:
'   sProcName = "getruns"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Function
'
'Public Sub FillMOs()
'   Dim RdoFrn As ADODB.RecordSet
'   Dim b As Byte
'   On Error GoTo DiaErr1
'   '    sSql = "Qry_RunsNotCanceled"       'shouldn't get closed runs
'   sSql = "SELECT DISTINCT PARTREF,PARTNUM,RUNREF" & vbCrLf _
'          & "From PartTable join RunsTable on PARTREF = RUNREF" & vbCrLf _
'          & "WHERE RUNSTATUS <>'CA' AND  RUNSTATUS <>'CL'" & vbCrLf _
'          & "ORDER BY PARTREF"
'
'   bSqlRows = clsAdoCon.GetDataSet(sSql,RdoFrn, ES_FORWARD)
'   If bSqlRows Then
'      With RdoFrn
'         While Not .EOF
'            AddComboStr cmbMon.hWnd, "" & Trim(!PARTNUM)
'            .MoveNext
'         Wend
'         .Cancel
'      End With
'   End If
'   On Error Resume Next
'   Set RdoFrn = Nothing
'   Exit Sub
'DiaErr1:
'   sProcName = "FillMOs"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub
'
'Public Sub FillCombo()
'   On Error GoTo DiaErr1
'   FillAccounts
'   FillMOs
'   Exit Sub
'DiaErr1:
'   sProcName = "FillCombo"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   DoModuleErrors Me
'End Sub


Private Sub cmbMon_Click()
   'GetRuns
   FillMoPartInfo cmbMon, lblDsc
   FillMoRunCombo cmbMon, cmbRun, whereClause
End Sub

'Private Sub cmbMon_LostFocus()
'   cmbMon = CheckLen(cmbMon, 30)
'   If Len(cmbMon) Then
'      GetRuns
'   Else
'      cmbRun.Clear
'   End If
'End Sub
