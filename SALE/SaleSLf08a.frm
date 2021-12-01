VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Price Book ID"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "New ID 12 Chars Max"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CheckBox opthlp 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "C&opy"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Create A New Price Book Identical To The Listed Price Book"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbPrb 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter Price Book ID (12 Char Max)"
      Top             =   960
      Width           =   1800
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "(40) Char Max"
      Top             =   1320
      Width           =   3475
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2040
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   6225
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Price Book ID"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price Book ID"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1635
   End
End
Attribute VB_Name = "SaleSLf08a"
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
Dim bOnLoad As Byte
Dim bGoodBook As Byte
Dim bBookExists As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd


Private Function GetPriceBookId() As Byte
   Dim RdoBok As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetPriceBook '" & Compress(cmbPrb) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
   If bSqlRows Then
      With RdoBok
         cmbPrb = "" & Trim(!PBHID)
         txtDsc = "" & Trim(!PBHDESC)
         ClearResultSet RdoBok
         GetPriceBookId = 1
      End With
   Else
      txtDsc = "*** Price Book Wasn't Found ***"
      GetPriceBookId = 0
   End If
   Set RdoBok = Nothing
   
   Exit Function
   
DiaErr1:
   sProcName = "getpricebkid"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbPrb_Click()
   bGoodBook = GetPriceBookId()
   
End Sub


Private Sub cmbPrb_Validate(Cancel As Boolean)
   cmbPrb = CheckLen(cmbPrb, 12)
   bGoodBook = GetPriceBookId()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Len(Trim(txtNew)) = 0 Then
      MsgBox "Requires A Valid Price Book ID.", _
         vbInformation, Caption
      Exit Sub
   End If
   
   If bGoodBook = 0 Then
      MsgBox "Requires A Valid Price Book.", _
         vbInformation, Caption
   Else
      bBookExists = TestNewBook()
      If bBookExists = 1 Then
         MsgBox "The New Price ID " & txtNew & " Is In Use...   ", _
            vbInformation, Caption
         Exit Sub
      End If
      
      bResponse = IllegalCharacters(txtNew)
      On Error Resume Next
      If bResponse > 0 Then
         MsgBox "The New Price Book ID Contains An Illegal " & Chr$(bResponse) & ".", _
            vbInformation, Caption
         Exit Sub
      End If
      sMsg = "This Copies The Contents Of The Selected" & vbCrLf _
             & "Price Book To The New Price Book ID." & vbCrLf _
             & "Continue To Copy The Price Book " & cmbPrb & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         CopyPriceBook
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2158
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf08a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc.BackColor = Es_FormBackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   cmbPrb.Clear
   sSql = "Qry_FillPriceBooks"
   LoadComboBox cmbPrb
   txtNew = " "
   If cmbPrb.ListCount > 0 Then
      cmbPrb = cmbPrb.List(0)
      'bGoodBook = GetPriceBookId()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub CopyPriceBook()
   Dim RdoCpy As ADODB.Recordset
   Dim sNewBook As String
   Dim sOldBook As String
   
   sNewBook = Compress(txtNew)
   sOldBook = Compress(cmbPrb)
   On Error Resume Next
   'Begin
   clsADOCon.BeginTrans
   clsADOCon.ADOErrNum = 0
   sSql = "Qry_GetPriceBook '" & sOldBook & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_FORWARD)
   If bSqlRows Then
      With RdoCpy
         sSql = "INSERT INTO PbhdTable (PBHREF,PBHID," _
                & "PBHDESC,PBHSTARTDATE,PBHENDDATE," _
                & "PBHDISCOUNT,PBHCLASS) " _
                & "VALUES('" & sNewBook & "','" _
                & Trim(txtNew) & "','" _
                & Trim(!PBHDESC) & "','" _
                & Format(!PBHSTARTDATE, "mm/dd/yy") & "','" _
                & Format(!PBHENDDATE, "mm/dd/yy") & "'," _
                & !PBHDISCOUNT & ",'" _
                & Trim(!PBHCLASS) & "')"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
         ClearResultSet RdoCpy
      End With
      sSql = "SELECT * FROM PbitTable WHERE PBIREF='" _
             & sOldBook & "' "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoCpy, ES_STATIC)
      If bSqlRows Then
         With RdoCpy
            Do Until .EOF
               sSql = "INSERT INTO PbitTable (PBIREF,PBIPARTREF,PBIPRICE) " _
                      & "VALUES('" & sNewBook & "','" _
                      & Trim(!PBIPARTREF) & "'," _
                      & !PBIPRICE & ")"
               clsADOCon.ExecuteSQL sSql 'rdExecDirect
               .MoveNext
            Loop
            ClearResultSet RdoCpy
         End With
      End If
   End If
   If clsADOCon.ADOErrNum = 0 Then
      clsADOCon.CommitTrans
      SysMsg "Price Book Copied", True
      txtNew = ""
      FillCombo
   Else
      clsADOCon.RollbackTrans
      MsgBox "Could Not Copy The Price Book.", _
         vbInformation, Caption
   End If
   
End Sub


Private Sub txtDsc_Change()
   If txtDsc.Text = "*** Price Book Wasn't Found ***" Then
      txtDsc.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      txtDsc.ForeColor = Es_TextForeColor
      cmdDel.Enabled = True
   End If
   
End Sub


Private Sub txtNew_Validate(Cancel As Boolean)
   txtNew = CheckLen(txtNew, 12)
   
End Sub



Private Function TestNewBook() As Byte
   Dim RdoBok As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetPriceBook '" & Compress(txtNew) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBok, ES_FORWARD)
   If bSqlRows Then
      TestNewBook = 1
   Else
      TestNewBook = 0
   End If
   Set RdoBok = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "testnewbo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
   
End Function
