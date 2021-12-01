VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form InvcINf03a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Part Number"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   2280
      Width           =   3852
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINf03a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.ComboBox cmbPrt 
      DataSource      =   "rDt1"
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter A Part Number or Select From List "
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCpy 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Change This Part Number"
      Top             =   840
      Width           =   875
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
      Left            =   5760
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2700
      FormDesignWidth =   6240
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblWrn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Close All Other Sections Before Proceeding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Number"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "InvcINf03a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'11/4/04 Tools
'6/24/05 Changed to allow Part Number change if PARTREF is the same.
'   (allows for forgotten spaces, dashes).
'5/26/06 VnapTable, BuypTable
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim bEstiTable As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

'7/9/01 See if EsiTable is here

Private Function CheckEstTable() As Byte
   On Error Resume Next
   clsADOCon.ADOErrNum = 0
   sSql = "SELECT BIDREF FROM EstiTable where BIDREF=1"
   clsADOCon.ExecuteSql sSql
   If clsADOCon.ADOErrNum = 0 Then CheckEstTable = 1 Else CheckEstTable = 0
   
End Function

Private Function CheckWindows() As Byte
   Dim b As Byte
   b = Val(GetSetting("Esi2000", "Sections", "admn", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "prod", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "engr", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "sale", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "fina", 0))
   b = b + Val(GetSetting("Esi2000", "Sections", "qual", 0))
   If b > 0 Then
      lblWrn(0) = sSysCaption & " Has Determined " & b & " Other Open Section(s)"
      lblWrn(0).Visible = True
      lblWrn(1).Visible = True
      cmdCpy.Enabled = False
   End If
   CheckWindows = b
   
End Function

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   If bCancel = 1 Then Exit Sub
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodOld = GetOldPart()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdCpy_Click()
   If Trim(txtNew) = "" Then
      MsgBox "Requires A New Part Number.", _
         vbInformation, Caption
   Else
      If bGoodOld And bGoodNew Then ChangePartNumber
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5152"
      cmdHlp = False
      bCancel = 0
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdHlp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub Form_Activate()
   Dim b As Byte
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      b = CheckWindows()
      bEstiTable = CheckEstTable()
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   FormatControls
   lblWrn(0).ForeColor = ES_RED
   lblWrn(1).ForeColor = ES_RED
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set InvcINf03a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub FillCombo()
   Dim b As Integer
   On Error GoTo DiaErr1
   sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Inventory Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   End If
   cmbPrt.Clear
   sSql = "Qry_FillSortedParts"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      FindPart cmbPrt, 0
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetOldPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sGetPart As String
   sGetPart = Compress(cmbPrt)
   On Error GoTo DiaErr1
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS,PATOOL " _
             & "FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         With RdoPrt
            On Error Resume Next
            cmbPrt = "" & Trim(!PartNum)
            lblDsc = "" & !PADESC
            lblTyp = Format(0 + !PALEVEL, "0")
            ClearResultSet RdoPrt
         End With
         GetOldPart = 1
      Else
         cmbPrt = ""
         lblDsc = "*** Invalid Part ***"
         lblTyp = ""
         GetOldPart = 0
      End If
      Set RdoPrt = Nothing
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getoldpart"
   GetOldPart = 0
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetNewPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim sGetPart As String
   sGetPart = Compress(txtNew)
   On Error GoTo DiaErr1
   If Compress(txtNew) = Compress(cmbPrt) Then
      GetNewPart = 1
      Exit Function
   End If
   If Len(sGetPart) > 0 Then
      sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         With RdoPrt
            On Error Resume Next
            txtNew = "" & Trim(!PartNum)
            ClearResultSet RdoPrt
         End With
         MsgBox "That Part Already Exists In The Database.", _
            vbInformation, Caption
         GetNewPart = False
      Else
         GetNewPart = True
      End If
      Set RdoPrt = Nothing
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getnewpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub ChangePartNumber()
   Dim bResponse As Byte
   Dim bByte As Byte
   Dim sMsg As String
   Dim sOldPart As String
   Dim sNewPart As String
   
   sOldPart = Compress(cmbPrt)
   sNewPart = Compress(txtNew)
   
   sMsg = "It Is Not A Good Idea To Change A Part Number " & vbCr _
          & "If There Is Any Chance That It Is In Use Right Now."
   MsgBox sMsg, vbInformation, Caption
   
   bByte = IllegalCharacters(txtNew)
   If bByte > 0 Then
      MsgBox "The Part Number Contains An Illegal " & Chr$(bByte) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   If sOldPart = sNewPart Then
      Err.Clear
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "UPDATE PartTable SET PARTNUM='" & Trim(txtNew) & "' " _
             & "WHERE PARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE BmhdTable SET BMHPARTNO='" & Trim(txtNew) & "' " _
             & "WHERE BMHREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE BmplTable SET BMPARTNUM='" & Trim(txtNew) & "' " _
             & "WHERE BMPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "The Part Number Was Changed.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         MsgBox "Couldn't Change That Part Number.", _
            vbExclamation, Caption
      End If
      FillCombo
      cmbPrt = txtNew
      txtNew = ""
      Exit Sub
   End If
   
   sMsg = "This Operation Will Change Part " & cmbPrt & vbCr _
          & "To Revised Part Number " & txtNew & "."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      prg1.Visible = True
      
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      'Bids 7/9/01
      'First, the PartTable
      
      'Disable constraints
      
        sSql = "ALTER TABLE Invatable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE MrplTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
            
        sSql = "ALTER TABLE MopkTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE PartTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE RnopTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE RunsTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE BmhdTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE PsitTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE LoitTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE LohdTable NOCHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql

      
      '            sSql = "ALTER TABLE PartTable NOCHECK Constraint PK_PartTable_PARTREF"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE RunsTable NOCHECK Constraint FK_RunsTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE RnopTable NOCHECK Constraint FK_RnopTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmhdTable NOCHECK Constraint FK_BmhdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmplTable NOCHECK Constraint FK_BmplTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE PsitTable NOCHECK Constraint FK_PsitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LohdTable NOCHECK Constraint FK_LohdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LoitTable NOCHECK Constraint FK_LoitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      
      
      If bEstiTable = 1 Then
         sSql = "UPDATE EstiTable SET BIDPART='" & sNewPart & "' " _
                & "WHERE BIDPART='" & sOldPart & "'"
         clsADOCon.ExecuteSql sSql
         prg1.Value = 5
      End If
      
      sSql = "UPDATE BmhdTable SET BMHREF='" & sNewPart & "'," _
             & "BMHPARTNO='" & txtNew & "',BMHPART='" & sNewPart _
             & "' WHERE BMHREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
         
      sSql = "UPDATE MrplTable SET MRP_PARTREF='" & sNewPart & "'," _
            & " MRP_PARTNUM='" & txtNew & "' " _
         & " WHERE MRP_PARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'Parts Lists/Picks
      sSql = "UPDATE BmhdTable SET BMHREF='" & sNewPart & "'," _
             & "BMHPARTNO='" & txtNew & "',BMHPART='" & sNewPart _
             & "' WHERE BMHREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 5
      
      sSql = "UPDATE BmplTable SET BMASSYPART='" & sNewPart & "' " _
             & "WHERE BMASSYPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 10
      
      sSql = "UPDATE BmplTable SET BMPARTREF='" & sNewPart & "'," _
             & "BMPARTNUM='" & txtNew & "' WHERE " _
             & "BMPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 15
      
      sSql = "UPDATE MopkTable SET PKMOPART='" & sNewPart & "' " _
             & "WHERE PKMOPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 20
      
      sSql = "UPDATE MopkTable SET PKPARTREF='" & sNewPart & "' " _
             & "WHERE PKPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 25
      
      'Po's/Ps
      sSql = "UPDATE PoitTable SET PIPART='" & sNewPart & "' " _
             & "WHERE PIPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 30
      
      sSql = "UPDATE PsitTable SET PIPART='" & sNewPart & "' " _
             & "WHERE PIPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 35
      
      'Rejections
      sSql = "UPDATE RjhdTable SET REJPART='" & sNewPart & "' " _
             & "WHERE REJPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 40
      
      sSql = "UPDATE RjkyTable SET KEYREF='" & sNewPart & "' " _
             & "WHERE KEYREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 45
      
      'Sales
      sSql = "UPDATE SoitTable SET ITPART='" & sNewPart & "' " _
             & "WHERE ITPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 50
      
      sSql = "UPDATE ViitTable SET VITMO='" & sNewPart & "' " _
             & "WHERE VITMO='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 55
      
      'Mo's/Routings
      sSql = "UPDATE RtopTable SET OPSERVPART='" & sNewPart & "' " _
             & "WHERE OPSERVPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 60
      
      sSql = "UPDATE RnopTable SET OPSERVPART='" & sNewPart & "' " _
             & "WHERE OPSERVPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 65
      
      sSql = "UPDATE RnopTable SET OPREF='" & sNewPart & "' " _
             & "WHERE OPREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 70
      
      sSql = "UPDATE RunsTable SET RUNREF='" & sNewPart & "' " _
             & "WHERE RUNREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 75
      
      sSql = "UPDATE RnalTable SET RAREF='" & sNewPart & "' " _
             & "WHERE RAREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 80
      
      'Inventory
      sSql = "UPDATE InvaTable SET INPART='" & sNewPart & "' " _
             & "WHERE INPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE InvaTable SET INMOPART='" & sNewPart & "' " _
             & "WHERE INMOPART='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 85
      
      'Documents
      sSql = "UPDATE DlstTable SET DLSREF='" & sNewPart & "' " _
             & "WHERE DLSREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      prg1.Value = 90
      
      'Time Cards
      sSql = "UPDATE TcitTable SET TCPARTREF='" & sNewPart & "' " _
             & "WHERE TCPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'Journals
      sSql = "UPDATE JritTable SET DCPARTNO='" & sNewPart & "' " _
             & "WHERE DCPARTNO='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'V 2/21/02
      'Price Books
      sSql = "UPDATE PbdtTable SET PBDPARTREF='" & sNewPart & "' " _
             & "WHERE PBDPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE PbitTable SET PBIPARTREF='" & sNewPart & "' " _
             & "WHERE PBIPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'Alias
      sSql = "UPDATE PaalTable SET ALPARTREF='" & sNewPart & "' " _
             & "WHERE ALPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE PaalTable SET ALALIASREF='" & sNewPart & "' " _
             & "WHERE ALALIASREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'Lots
      sSql = "UPDATE LohdTable SET LOTPARTREF='" & sNewPart & "' " _
             & "WHERE LOTPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE LohdTable SET LOTMOPARTREF='" & sNewPart & "' " _
             & "WHERE LOTMOPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE LoitTable SET LOIPARTREF='" & sNewPart & "' " _
             & "WHERE LOIPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE LoitTable SET LOIMOPARTREF='" & sNewPart & "' " _
             & "WHERE LOIMOPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      'Tools
      sSql = "UPDATE TohdTable SET TOOL_PARTREF='" & sNewPart & "'," _
             & "TOOL_NUM='" & txtNew & "' WHERE TOOL_PARTREF='" _
             & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE TlitTable SET TOOLLISTIT_TOOLREF='" & sNewPart & "' " _
             & "WHERE TOOLLISTIT_TOOLREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      '5/26/06
      sSql = "UPDATE VnapTable SET AVPARTREF='" & sNewPart & "' " _
             & "WHERE AVPARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      '5/26/06
      sSql = "UPDATE BuypTable SET BYPARTNUMBER='" & sNewPart & "' " _
             & "WHERE BYPARTNUMBER='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      sSql = "UPDATE PartTable SET PARTREF='" & sNewPart & "'," _
             & "PARTNUM='" & txtNew & "' " _
             & "WHERE PARTREF='" & sOldPart & "'"
      clsADOCon.ExecuteSql sSql
      
      prg1.Value = 95
      
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         '11/04/04 MRP (little priority)
         sSql = "UPDATE MrplTable SET MRP_PARTREF='" & sNewPart & "'," _
                & "MRP_PARTNUM='" & txtNew & "' WHERE MRP_PARTREF='" & sOldPart & "'"
         clsADOCon.ExecuteSql sSql
         
         prg1.Value = 100
         SysMsg "Part Number Was Successfully Changed", True
      Else
         clsADOCon.RollbackTrans
         clsADOCon.ADOErrNum = 0
         
         MsgBox "Could Not Change The Part Number.", _
            vbInformation, Caption
         prg1.Visible = False
      End If
      '            sSql = "ALTER TABLE PartTable CHECK Constraint PK_PartTable_PARTREF"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE RunsTable CHECK Constraint FK_RunsTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE RnopTable CHECK Constraint FK_RnopTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmhdTable CHECK Constraint FK_BmhdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE BmplTable CHECK Constraint FK_BmplTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE PsitTable CHECK Constraint FK_PsitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LohdTable CHECK Constraint FK_LohdTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      '
      '            sSql = "ALTER TABLE LoitTable CHECK Constraint FK_LoitTable_PartTable"
      '            clsADOCon.ExecuteSQL sSql
      
        sSql = "ALTER TABLE MopkTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE PartTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE RnopTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE RunsTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE BmhdTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE PsitTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE LoitTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
        
        sSql = "ALTER TABLE LohdTable WITH CHECK CHECK CONSTRAINT ALL"
        clsADOCon.ExecuteSql sSql
      
      txtNew = ""
      FillCombo
      prg1.Visible = False
      cmbPrt.SetFocus
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "changepartnu"
   CurrError.Number = Err
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Sub lblDsc_Change()
   If Left(lblDsc, 8) = "*** Part" Then
      lblDsc.ForeColor = ES_RED
   Else
      lblDsc.ForeColor = vbBlack
   End If
   
End Sub

Private Sub txtNew_LostFocus()
   txtNew = CheckLen(txtNew, 30)
   If Len(txtNew) Then
      bGoodNew = GetNewPart()
   Else
      bGoodNew = False
   End If
   
End Sub
