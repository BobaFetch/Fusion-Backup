VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPC Processes"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   6301
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optTem 
      Caption         =   "Teams"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUon 
      Caption         =   "Go"
      Height          =   315
      Left            =   6720
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Fill Used On Assemblies"
      Top             =   1680
      Width           =   495
   End
   Begin VB.ComboBox cmbUon 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "3"
      ToolTipText     =   "Used On Assemblies (Press Go To Fill)"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox optDsc 
      Caption         =   "Sources"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDat 
      Caption         =   "&Sources"
      Height          =   315
      Left            =   6360
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Add/Revise Data Sources Of Information"
      Top             =   840
      Width           =   875
   End
   Begin VB.CommandButton cmdTem 
      Caption         =   "&Team"
      Height          =   315
      Left            =   6360
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Add/Revise Team Members"
      Top             =   480
      Width           =   875
   End
   Begin VB.TextBox txtCmt 
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "9"
      ToolTipText     =   "Notes: (255) Characters"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.ComboBox cmbPrc 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Tag             =   "3"
      ToolTipText     =   "Enter Process ID Or Select From List"
      Top             =   2760
      Width           =   1875
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Tag             =   "4"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox cmbRsc 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Tag             =   "3"
      ToolTipText     =   "Enter Code Or Select From List"
      Top             =   2400
      Width           =   1875
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Tag             =   "2"
      ToolTipText     =   "Key Description (60 Char)"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.ComboBox cmbKey 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   2
      Tag             =   "3"
      ToolTipText     =   "Enter New Or Select Key For Part Number"
      Top             =   2040
      Width           =   1875
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter Part Number Or Select From List"
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6960
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   4860
      FormDesignWidth =   7350
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Used On"
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   26
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label Editing 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label lblPrc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   20
      Top             =   2760
      Width           =   3480
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Process ID"
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label lblRsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   2400
      Width           =   3480
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason Code"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Key Dimension"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label lblFam 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5640
      TabIndex        =   13
      Top             =   960
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   12
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Type"
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   11
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1350
      TabIndex        =   10
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1275
   End
End
Attribute VB_Name = "StatSPe01a"
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
Dim AdoKey As ADODB.Recordset

Dim bCanceled As Byte
Dim bGoodKey As Byte
Dim bGoodPart As Byte
Dim bGoodProc As Byte
Dim bGoodRes As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetReason() As Byte
   Dim RdoRcd As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT RCOREF,RCOID,RCODESC FROM " _
          & "RjrcTable WHERE RCOREF='" & Compress(cmbRsc) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRcd, ES_FORWARD)
   If bSqlRows Then
      With RdoRcd
         cmbRsc = "" & Trim(!RCOID)
         lblRsc = "" & Trim(!RCODESC)
         ClearResultSet RdoRcd
         GetReason = 1
      End With
   Else
      GetReason = 0
      lblRsc = "*** No Valid Reasoning Code Selected ***"
   End If
   Set RdoRcd = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getreason"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbKey_Click()
   bGoodKey = GetKey()
   
End Sub

Private Sub cmbKey_LostFocus()
   cmbKey = CheckLen(cmbKey, 15)
   If bCanceled Then Exit Sub
   bGoodKey = GetKey()
   If bGoodKey = 0 Then AddKey
   
End Sub


Private Sub cmbPrc_Click()
   bGoodProc = GetProcessId()
   
End Sub


Private Sub cmbPrc_LostFocus()
   cmbPrc = CheckLen(cmbPrc, 15)
   bGoodProc = GetProcessId()
   If bGoodKey Then
      On Error Resume Next
      AdoKey!KEYPROCESS = "" & Compress(cmbPrc)
      AdoKey.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbPrt_Click()
   bGoodPart = GetPartNumber()
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   If bCanceled = 0 Then bGoodPart = GetPartNumber()
   
End Sub


Private Sub cmbRsc_Click()
   bGoodRes = GetReason()
   
End Sub


Private Sub cmbRsc_LostFocus()
   cmbRsc = CheckLen(cmbRsc, 15)
   bGoodRes = GetReason()
   If bGoodKey Then
      On Error Resume Next
      AdoKey!KEYREASON = "" & Compress(cmbRsc)
      AdoKey.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub cmbUon_Click()
   If cmbUon.ListCount > 0 Then
      cmbPrt = cmbUon
      bGoodPart = GetPartNumber()
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCanceled = True
   
End Sub


Private Sub cmdDat_Click()
   optDsc.Value = vbChecked
   StatSPe01c.lblPrt = cmbPrt
   StatSPe01c.lblKey = cmbKey
   StatSPe01c.Show
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6301
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdTem_Click()
   optTem.Value = vbChecked
   StatSPe01b.lblPrt = cmbPrt
   StatSPe01b.lblKey = cmbKey
   StatSPe01b.Show
   
End Sub

Private Sub cmdUon_Click()
   FillUsedOn
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   If optDsc.Value = vbChecked Then Unload StatSPe01c
   If optTem.Value = vbChecked Then Unload StatSPe01b
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "SELECT TOP 1 PARTREF,PARTNUM,PALEVEL,PADESC,PAFAMILY," _
          & "FAMREF,FAMID FROM PartTable,RjfmTable WHERE " _
          & "FAMREF=*PAFAMILY AND PARTREF= ? "
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Size = 30
   AdoParameter.Type = adChar

   AdoQry.Parameters.Append AdoParameter
          

   txtDte = Format(ES_SYSDATE, "mm/dd/yy")
   Editing.ForeColor = ES_RED
   Editing = "No Current Key Characteristic."
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set AdoParameter = Nothing
   Set AdoQry = Nothing
   Set AdoKey = Nothing
   Set StatSPe01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable WHERE " _
          & "(PALEVEL<>5 AND PALEVEL<>6 AND PATOOL=0 AND PAINACTIVE = 0 And PAOBSOLETE = 0) ORDER BY PARTREF"
   LoadComboBox cmbPrt
   
   'Reason Codes
   sSql = "Qry_FillSPReasoning"
   LoadComboBox cmbRsc
   
   'Process Id's
   sSql = "Qry_FillSPProcessID"
   LoadComboBox cmbPrc
   
   'Get the Part
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      bGoodPart = GetPartNumber()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPartNumber() As Byte
   Dim RdoPrt As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbUon.Clear
   cmbUon.Enabled = False
   'RdoQry(0) = Compress(cmbPrt)
   AdoQry.Parameters(0).Value = Compress(cmbPrt)
   bSqlRows = clsADOCon.GetQuerySet(RdoPrt, AdoQry, ES_STATIC)
   If bSqlRows Then
      With RdoPrt
         cmbPrt = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblLvl = Format(!PALEVEL, "#")
         lblFam = "" & Trim(!FAMID)
         ClearResultSet RdoPrt
      End With
      If lblFam = "" Then lblFam = "*** No Family ***"
      GetPartNumber = 1
      FillKeys
   Else
      GetPartNumber = 0
      lblLvl = ""
      lblFam = ""
      lblDsc = "*** No Current Part Number ***"
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpartnu"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** No C" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblFam_Change()
   If Left(lblFam, 8) = "*** No F" Then
      lblFam.ForeColor = ES_RED
   Else
      lblFam.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblPrc_Change()
   If Left(lblPrc, 8) = "*** No V" Then
      lblPrc.ForeColor = ES_RED
   Else
      lblPrc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub lblRsc_Change()
   If Left(lblRsc, 8) = "*** No V" Then
      lblRsc.ForeColor = ES_RED
   Else
      lblRsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub optDsc_Click()
   'never visible-StatSPe01c is showing
   
End Sub

Private Sub txtCmt_LostFocus()
   txtCmt = CheckLen(txtCmt, 255)
   If bGoodKey Then
      On Error Resume Next
      AdoKey!KEYNOTES = "" & txtCmt
      AdoKey.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 60)
   If bGoodKey Then
      On Error Resume Next
      AdoKey!KEYDESC = "" & txtDsc
      AdoKey.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   If bGoodKey Then
      On Error Resume Next
      AdoKey!KeyDate = "" & Format(txtDte, "mm/dd/yy")
      AdoKey.Update
      If Err > 0 Then ValidateEdit
   End If
   
End Sub



Private Function GetProcessId() As Byte
   Dim RdoPrc As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PROREF,PROID,PRODESC,PRONOTES FROM " _
          & "RjprTable WHERE PROREF='" & Compress(cmbPrc) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrc, ES_FORWARD)
   If bSqlRows Then
      With RdoPrc
         cmbPrc = "" & Trim(!PROID)
         lblPrc = "" & Trim(!PRODESC)
         ClearResultSet RdoPrc
         GetProcessId = 1
      End With
   Else
      GetProcessId = 0
      lblPrc = "*** No Valid Process ID Selected ***"
   End If
   Set RdoPrc = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getprocess"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetKey() As Byte
   On Error Resume Next
   'RdoKey.Close
   On Error GoTo DiaErr1
   sSql = "SELECT KEYREF,KEYDIM,KEYDESC,KEYDATE,KEYREASON," _
          & "KEYPROCESS,KEYNOTES FROM RjkyTable WHERE " _
          & "(KEYREF='" & Compress(cmbPrt) & "' AND " _
          & "KEYDIM='" & cmbKey & "')"
   bSqlRows = clsADOCon.GetDataSet(sSql, AdoKey, ES_KEYSET)
   If bSqlRows Then
      With AdoKey
         cmbKey = "" & Trim(!KEYDIM)
         txtDsc = "" & Trim(!KEYDESC)
         txtDte = "" & Format(!KeyDate, "mm/dd/yy")
         cmbRsc = "" & Trim(!KEYREASON)
         cmbPrc = "" & Trim(!KEYPROCESS)
         txtCmt = "" & Trim(!KEYNOTES)
      End With
      bGoodRes = GetReason()
      bGoodProc = GetProcessId()
      Editing.ForeColor = ES_BLUE
      Editing = "Editing " & cmbKey & "."
      GetKey = 1
   Else
      txtDsc = ""
      txtDte = Format(ES_SYSDATE, "mm/dd/yy")
      cmbRsc = ""
      cmbPrc = ""
      lblRsc = ""
      lblPrc = ""
      txtCmt = ""
      Editing.ForeColor = ES_RED
      Editing = "No Current Key Characteristic."
      GetKey = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getkey "
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillKeys()
   cmbKey.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT KEYREF,KEYDIM FROM RjkyTable WHERE " _
          & "KEYREF='" & Compress(cmbPrt) & "' "
   LoadComboBox cmbKey
   If cmbKey.ListCount > 0 Then
      cmbKey = cmbKey.List(0)
      bGoodKey = GetKey()
   Else
      txtDsc = ""
      txtDte = Format(ES_SYSDATE, "mm/dd/yy")
      cmbRsc = ""
      cmbPrc = ""
      lblRsc = ""
      lblPrc = ""
      txtCmt = ""
      Editing.ForeColor = ES_RED
      Editing = "No Current Key Characteristic."
      bGoodKey = 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillkeys"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub AddKey()
   Dim b As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sPart As String
   
   On Error GoTo DiaErr1
   cmbKey = Trim(cmbKey)
   If Len(cmbKey) < 5 Then
      If Len(cmbKey) = 0 Then
         Exit Sub
      Else
         MsgBox "Please Make Your Key Dimensions At " & vbCr _
            & "Least Five (5) Characters Long.", _
            vbInformation, Caption
         Exit Sub
      End If
   End If
   sMsg = "That Key Dimension Doesn't Exist.    " & vbCr _
          & "Record " & Trim(cmbKey) & " Now?.."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      b = SearchKeyString(cmbKey)
      If b > 0 Then Exit Sub
      sPart = Compress(cmbPrt)
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      'Both tables...
      sSql = "INSERT INTO RjkyTable (KEYREF,KEYDIM) " _
             & "VALUES('" & sPart & "','" & cmbKey & "')"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "INSERT INTO RjdtTable (DATREF,DATKEY) " _
             & "VALUES('" & sPart & "','" & cmbKey & "')"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         AddComboStr cmbKey.hwnd, cmbKey
         MsgBox "Key Dimension Was Successfully Added.", _
            vbInformation, Caption
         bGoodKey = GetKey()
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "Unable To Successfully Add The Key Dimension.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "addkey"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function SearchKeyString(sKeyString As String) As Byte
   Dim b As Integer
   Dim iList As Integer
   On Error Resume Next
   b = Len(Trim(sKeyString))
   iList = InStr(sKeyString, Chr(34))
   If iList > 0 Then
      MsgBox "Please Use INCH Or IN Instead Of '' For Inches.", _
         vbInformation, Caption
      SearchKeyString = 1
      Exit Function
   Else
      SearchKeyString = 0
   End If
   iList = InStr(sKeyString, Chr(39))
   If iList > 0 Then
      MsgBox "Please Use FEET Or FT Instead Of ' For Feet.", _
         vbInformation, Caption
      SearchKeyString = 1
   Else
      SearchKeyString = 0
   End If
   
   
End Function

Private Sub FillUsedOn()
   cmbUon.Clear
   sSql = "SELECT DISTINCT BMPARTREF,BMPARTNUM " _
          & "FROM BmplTable WHERE BMASSYPART='" & Compress(cmbPrt) & "' "
   LoadComboBox cmbUon
   If cmbUon.ListCount > 0 Then
      cmbUon = cmbUon.List(0)
      cmbUon.Enabled = True
   Else
      cmbUon.Enabled = False
      cmbUon = "No Used On Assemblies Found"
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillusedon"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
