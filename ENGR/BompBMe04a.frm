VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form BompBMe04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Parts List Revisions For Type 4's"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmbUpd 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Height          =   312
      Left            =   6000
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Create The Part Type 4 Revision"
      Top             =   1680
      Width           =   915
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "BompBMe04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtRev 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Enter New Revision"
      Top             =   1680
      Width           =   735
   End
   Begin VB.ComboBox cmbPls 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Select Part Number"
      Top             =   960
      Width           =   3345
   End
   Begin VB.CommandButton cmdCan 
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2520
      FormDesignWidth =   7005
   End
   Begin VB.Label lblRev 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   6120
      TabIndex        =   9
      Top             =   996
      Width           =   732
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Rev"
      Height          =   252
      Index           =   3
      Left            =   5040
      TabIndex        =   8
      Top             =   1000
      Width           =   1092
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Revision"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblDsc 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1000
      Width           =   1335
   End
End
Attribute VB_Name = "BompBMe04a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'6/19/06 Added Apply Button
Option Explicit
Dim bGoodPart As Byte
Dim bHeaderIs As Byte
Dim bOnLoad As Byte

Dim sHeader As String
Dim sPartNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub cmbPls_Click()
   bGoodPart = GetPart()
   
End Sub


Private Sub cmbPls_LostFocus()
   cmbPls = CheckLen(cmbPls, 30)
   bGoodPart = GetPart()
   If Not bGoodPart Then
      If Len(cmbPls) > 0 Then MsgBox "Part Wasn't Found or Isn't Type 4.", vbExclamation, Caption
   End If
   
End Sub

Private Sub cmbUpd_Click()
   If Len(txtRev) > 0 Then
      bHeaderIs = GetHeader()
   Else
      Exit Sub
   End If
   If Not bHeaderIs Then
      AddHeader
   Else
      MsgBox "That Revision Is Already Recorded.", vbInformation, Caption
   End If
   
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3204
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Command1_Click()
   
End Sub

Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillParts
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
   On Error Resume Next
   FormUnload
   Set BompBMe04a = Nothing
   
End Sub



Private Sub FillParts()
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PALEVEL FROM PartTable " _
          & "WHERE (PALEVEL=4 AND PAPRODCODE<>'BID' AND PAINACTIVE = 0 AND PAOBSOLETE = 0) " _
          & "ORDER BY PARTREF"
   LoadComboBox cmbPls
   If bSqlRows Then
      cmbPls = cmbPls.List(0)
      bGoodPart = GetPart()
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillparts"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPart() As Byte
   Dim RdoPrt As ADODB.Recordset
   sPartNumber = Compress(cmbPls)
   On Error GoTo DiaErr1
   sSql = "SELECT PARTREF,PARTNUM,PADESC,PALEVEL,PABOMREV FROM PartTable " _
          & "WHERE PARTREF='" & sPartNumber & "' AND PALEVEL=4"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPrt)
   If bSqlRows Then
      With RdoPrt
         cmbPls = "" & Trim(!PartNum)
         lblDsc = "" & Trim(!PADESC)
         lblRev = "" & Trim(!PABOMREV)
         ClearResultSet RdoPrt
      End With
      GetPart = True
   Else
      lblDsc = ""
      lblRev = ""
      sPartNumber = ""
      GetPart = False
   End If
   Set RdoPrt = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpart"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtRev_LostFocus()
   txtRev = CheckLen(txtRev, 4)
   
End Sub



Private Function GetHeader() As Byte
   Dim RdoBls As ADODB.Recordset
   sPartNumber = Compress(cmbPls)
   sPartNumber = cmbPls
   txtRev = Compress(txtRev)
   sHeader = txtRev
   If Not bGoodPart Then Exit Function
   On Error GoTo DiaErr1
   sSql = "SELECT BMHREF,BMHREV FROM BmhdTable " _
          & "WHERE BMHREF='" & sPartNumber & "' AND BMHREV='" & sHeader & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoBls)
   If bSqlRows Then
      txtRev = "" & Trim(RdoBls!BMHREV)
      GetHeader = True
   Else
      sHeader = ""
      GetHeader = False
   End If
   Set RdoBls = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getheader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddHeader()
   Dim bSuccess As Byte
   Dim bResponse As Byte
   Dim sMsg As String
   Dim sRevision As String
   sRevision = Compress(txtRev)
   sMsg = "Create Revison " & sRevision & " For " & cmbPls & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   On Error GoTo DiaErr1
   If bResponse = vbYes Then
      sSql = "INSERT INTO BmhdTable (BMHREF,BMHPARTNO,BMHREV) " _
             & "VALUES('" & sPartNumber & "','" & cmbPls & "','" & sRevision & "')"
      clsADOCon.ExecuteSQL sSql ' rdExecDirect
      If clsADOCon.RowsAffected > 0 Then
         bSuccess = True
         sMsg = "Revision Created, Assign To Part?"
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            sSql = "UPDATE PartTable SET PABOMREV='" & sRevision & "' " _
                   & "WHERE PARTREF='" & sPartNumber & "' "
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            If clsADOCon.RowsAffected > 0 Then SysMsg "Revision Assigned.", True, Me
         End If
      Else
         MsgBox "Couldn't Create Revision.", vbExclamation, Caption
         bSuccess = False
      End If
   Else
      CancelTrans
   End If
   On Error Resume Next
   If bSuccess Then
      lblRev = sHeader
      txtRev = ""
   End If
   cmbPls.SetFocus
   Exit Sub
   
DiaErr1:
   sProcName = "addheader"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
