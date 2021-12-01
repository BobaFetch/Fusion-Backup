VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InvcINf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy A Part Number"
   ClientHeight    =   2595
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
   ScaleHeight     =   2595
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InvcINf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
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
      ToolTipText     =   "Enter Part Number or Select From List "
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCpy 
      Caption         =   "&Copy"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      ToolTipText     =   "Copy this Part Number To A New Part Number"
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
      Left            =   5400
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2595
      FormDesignWidth =   6240
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
Attribute VB_Name = "InvcINf02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***

Option Explicit
Dim bOnLoad As Byte
Dim bGoodOld As Byte
Dim bGoodNew As Byte
Dim bShowParts As Byte

Dim sTableDef As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPrt_Click()
   cmbPrt = GetCurrentPart(cmbPrt, lblDsc)
   
End Sub


Private Sub cmbPrt_LostFocus()
   cmbPrt = CheckLen(cmbPrt, 30)
   
   If (Not ValidPartNumber(cmbPrt.Text)) Then
      MsgBox "Can't Select The Part Number Which Is Obsolete or Inactive. ", _
         vbInformation, Caption
      cmbPrt = ""
      Exit Sub
   End If
   
   bGoodOld = GetOldPart()
   If bGoodOld Then cmdCpy.Enabled = True
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdCpy_Click()
   If lblDsc.ForeColor <> ES_RED Then
      If bGoodOld = 1 And bGoodNew = 1 Then
         CopyNewPart
      End If
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5151"
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
   FormLoad Me, ES_DONTLIST
   FormatControls
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   'It is redundant to try to drop the Temp table, but
   'may reduce clutter
   sSql = "DROP TABLE " & sTableDef
   clsADOCon.ExecuteSQL sSql
   If bShowParts = 0 Then
      FormUnload
   Else
      InvcINe01a.cmbPrt = txtNew
   End If
   Set InvcINf02a = Nothing
   
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
   On Error Resume Next
   sTableDef = "#" & Trim(sInitials) & "NewPart"
   cmbPrt.Clear
   sSql = "SELECT PARTREF,PARTNUM FROM PartTable WHERE " _
          & "PATOOL=0 AND PAINACTIVE = 0 AND PAOBSOLETE = 0 ORDER BY PARTREF"
   LoadComboBox cmbPrt
   If cmbPrt.ListCount > 0 Then
      cmbPrt = cmbPrt.List(0)
      FindPart cmbPrt
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
      sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PAUNITS FROM PartTable WHERE PARTREF='" & sGetPart & "'"
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
      If bSqlRows Then
         With RdoPrt
            On Error Resume Next
            cmbPrt = "" & Trim(!PartNum)
            lblDsc = "" & !PADESC
            lblTyp = Format(!PALEVEL, "0")
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
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Function GetNewPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   Dim bByte As Byte
   Dim sGetPart As String
   
   sGetPart = Compress(txtNew)
   bByte = IllegalCharacters(txtNew)
   If bByte > 0 Then
      MsgBox "The Part Number Contains An Illegal " & Chr$(bByte) & ".", _
         vbExclamation, Caption
      GetNewPart = 0
      Exit Function
   End If
   On Error GoTo DiaErr1
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
            vbExclamation, Caption
         GetNewPart = 0
      Else
         GetNewPart = 1
      End If
      Set RdoPrt = Nothing
   End If
   Exit Function
   
DiaErr1:
   GetNewPart = 0
   sProcName = "getnewpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CopyNewPart()
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sPart As String
   Dim sNewPart As String, UOM As String
   
   On Error GoTo DiaErr1
   
   sPart = Compress(cmbPrt)
   sNewPart = Compress(txtNew)
   sMsg = "This Operation Will Copy Part " & cmbPrt & ". Do " & vbCr _
          & "You Want To Copy The Contents To " & txtNew & "."
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      cmdCpy.Enabled = False
      On Error Resume Next
      MouseCursor ccHourglass
      sSql = "DROP TABLE " & sTableDef
      clsADOCon.ExecuteSQL sSql
      
      Err.Clear
      sSql = "SELECT * INTO " & sTableDef & " FROM PartTable WHERE PARTREF='" _
             & sPart & "'"
      clsADOCon.ExecuteSQL sSql
      
      If Err > 0 Then
         MouseCursor 0
         MsgBox "Table Is In Use, Come Back In A Few Minutes.", _
            vbInformation, Caption
         
         'RdoCon.Close
         Set clsADOCon = Nothing
         OpenDBServer True
         SysMsg "Table Reset.", True
         FillCombo
         Exit Sub
      End If
      
      On Error GoTo DiaErr1
      
      sSql = "UPDATE " & sTableDef & " SET PARTREF='" & sNewPart & "'," & vbCrLf _
             & "PARTNUM='" & txtNew & "',PAQOH=0,PALOTQTYREMAINING=0," & vbCrLf _
             & "PAAVGCOST=0,PADOCLISTREF='',PADOCLISTREV='', PARUN=0"
      clsADOCon.ExecuteSQL sSql
      
      sSql = "select PAUNITS from " & sTableDef
      Dim rdo As ADODB.Recordset

      If clsADOCon.GetDataSet(sSql, rdo) Then
         UOM = rdo!PAUNITS
      End If
      Set rdo = Nothing
      
      'copy the part
      clsADOCon.BeginTrans
      
      sSql = "INSERT INTO PartTable SELECT * FROM " & sTableDef & ""
      clsADOCon.ExecuteSQL sSql
      
      'note that a BmhdTable row is required for each part, so a row is inserted here
      sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHPART)" & vbCrLf _
        & "VALUES('" & sNewPart & "','" & txtNew & "','" & sNewPart & "')"
      clsADOCon.ExecuteSQL sSql
      
      clsADOCon.CommitTrans
      
      cUR.CurrentPart = txtNew
      SaveSetting "Esi2000", "Current", "Part", cUR.CurrentPart
      sPassedPart = txtNew
      
      sSql = "DROP TABLE " & sTableDef
      clsADOCon.ExecuteSQL sSql
      
      'create the corresponding lot structure
      Dim part As New ClassPart
      part.CreateInitialLot sNewPart, "EA"
      
      clsADOCon.CommitTrans
      
      MouseCursor ccDefault
      sMsg = "The Part Number Was Successfully Copied.  Edit The New Part Number Now?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         bShowParts = 1
         InvcINe01a.cmbPrt = txtNew
         InvcINe01a.txtPrt = txtNew
         Dim b As Byte
         b = InvcINe01a.GetPart(1)
         InvcINe01a.Show
         Unload Me
         Exit Sub
      Else
         txtNew = ""
         FillCombo
      End If
      clsADOCon.ExecuteSQL sSql
   Else
      CancelTrans
   End If
   'On Error Resume Next
   sSql = "DROP TABLE " & sTableDef
   clsADOCon.ExecuteSQL sSql
   Exit Sub
   
DiaErr1:
   clsADOCon.RollbackTrans
   sProcName = "CopyNewPart"
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
