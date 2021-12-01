VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form StatSPf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Family ID By Product Code"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "StatSPf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1920
      Width           =   3475
   End
   Begin VB.TextBox txtDsc 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1200
      Width           =   3475
   End
   Begin VB.CommandButton cmdUpd 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      ToolTipText     =   "Update All Parts With This Product Code To The Selected Family ID"
      Top             =   520
      Width           =   875
   End
   Begin VB.ComboBox cmbFam 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter New Family ID (15 Char) Or Select From List"
      Top             =   1560
      Width           =   1875
   End
   Begin VB.ComboBox cmbCde 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Enter New Team Member  (15 Char) Or Select From List"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   120
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2520
      FormDesignWidth =   5775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Family ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "StatSPf04a"
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
Dim bGoodCde As Byte
Dim bGoodFam As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetCode() As Byte
   Dim RdoCde As ADODB.Recordset
   sSql = "SELECT PCREF,PCCODE,PCDESC FROM PcodTable " _
          & "WHERE PCREF='" & Compress(cmbCde) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCde, ES_FORWARD)
   If bSqlRows Then
      With RdoCde
         cmbCde = "" & Trim(!PCCODE)
         txtDsc(0) = "" & Trim(!PCDESC)
         ClearResultSet RdoCde
      End With
      GetCode = 1
   Else
      GetCode = 0
      txtDsc(0) = "*** No Valid Code Selected ***"
   End If
   Set RdoCde = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcode"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetFamily() As Byte
   Dim RdoFam As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT FAMREF,FAMID,FAMDESC,FAMNOTES FROM " _
          & "RjfmTable WHERE FAMREF='" & Compress(cmbFam) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoFam, ES_FORWARD)
   If bSqlRows Then
      With RdoFam
         cmbFam = "" & Trim(!FAMID)
         txtDsc(1) = "" & Trim(!FAMDESC)
         ClearResultSet RdoFam
         GetFamily = 1
      End With
   Else
      GetFamily = 0
      txtDsc(1) = "*** No Valid Family ID ***"
   End If
   Set RdoFam = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getfamily"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub cmbCde_Click()
   bGoodCde = GetCode()
   
End Sub


Private Sub cmbCde_LostFocus()
   cmbCde = CheckLen(cmbCde, 6)
   bGoodCde = GetCode()
   
End Sub


Private Sub cmbFam_Click()
   bGoodFam = GetFamily
   
End Sub


Private Sub cmbFam_LostFocus()
   cmbFam = CheckLen(cmbFam, 15)
   bGoodFam = GetFamily
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6353
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdUpd_Click()
   On Error Resume Next
   If bGoodCde = 0 Then
      MsgBox "That Product Code Wasn't Found.", _
         vbExclamation, Caption
      cmbCde.SetFocus
      Exit Sub
   End If
   If bGoodFam = 0 Then
      MsgBox "That Family ID Wasn't Found.", _
         vbExclamation, Caption
      cmbFam.SetFocus
      Exit Sub
   End If
   UpdateParts
   
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
   Set StatSPf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDsc(0).BackColor = BackColor
   txtDsc(1).BackColor = BackColor
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   FillProductCodes
   If cmbCde.ListCount > 0 Then
      cmbCde = cmbCde.List(0)
      bGoodCde = GetCode()
   End If
   sSql = "Qry_FillSPFamily"
   LoadComboBox cmbFam
   If cmbFam.ListCount > 0 Then cmbFam = cmbFam.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtDsc_Change(Index As Integer)
   If Left(txtDsc(Index), 8) = "*** No V" Then
      txtDsc(Index).ForeColor = ES_RED
   Else
      txtDsc(Index).ForeColor = vbBlack
   End If
   
End Sub



Private Sub UpdateParts()
   Dim bResponse As Byte
   Dim l As Long
   Dim sCode As String
   Dim sFam As String
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "You Have Chosen To Set The Family ID Of All " & vbCr _
          & "Parts To " & cmbFam & " With Product Code " & vbCr _
          & cmbCde & " Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      sFam = Compress(cmbFam)
      sCode = Compress(cmbCde)
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      clsADOCon.BeginTrans
      
      sSql = "UPDATE PartTable SET PAFAMILY='" _
             & sFam & "' WHERE PAPRODCODE='" _
             & sCode & "' "
      clsADOCon.ExecuteSQL sSql
      l = clsADOCon.RowsAffected
      If l = 0 Or clsADOCon.ADOErrNum > 0 Then
         clsADOCon.RollbackTrans
         MsgBox "No Parts Will Be Affected.", _
            vbInformation
      Else
         sMsg = l & " Part Numbers Will Be Updated.  " & vbCr _
                & "Do You Still Want To Continue? "
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            clsADOCon.CommitTrans
            MsgBox "Selected Parts Were Updated.", _
               vbInformation
         Else
            clsADOCon.RollbackTrans
            clsADOCon.ADOErrNum = 0
            
            CancelTrans
         End If
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "updatepar"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
