VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reactivate An Inspector"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      ToolTipText     =   "Reactivate This Inspector"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbIns 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Inspector ID From List"
      Top             =   840
      Width           =   1665
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1920
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2130
      FormDesignWidth =   5520
   End
   Begin VB.Label lblDiv 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   400
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stamp"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblStp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspector Id"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "InspRTf04a"
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
Dim bGoodIns As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetInspector() As Byte
   Dim RdoIns As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT INSID,INSFIRST,INSMIDD,INSLAST,INSSTAMP,INSDIVISION " _
          & "FROM RinsTable WHERE INSID='" & Compress(cmbIns) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoIns, ES_FORWARD)
   If bSqlRows Then
      With RdoIns
         cmbIns = "" & Trim(!INSID)
         lblNme = "" & Trim(!INSFIRST) _
                  & " " & Trim(!INSMIDD) _
                  & " " & Trim(!INSLAST)
         lblStp = Format(0 + !INSSTAMP, "####0")
         lblDiv = "" & Trim(!INSDIVISION)
      End With
      ClearResultSet RdoIns
      GetInspector = True
   Else
      lblNme = "*** Inspector Wasn't Found ***"
      lblStp = ""
      GetInspector = False
   End If
   Set RdoIns = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getinspect"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub FillInspectors()
   On Error GoTo DiaErr1
   sSql = "SELECT INSID FROM RinsTable WHERE INSACTIVE=0"
   LoadComboBox cmbIns, -1
   If cmbIns.ListCount > 0 Then
      cmbIns = cmbIns.List(0)
   Else
      cmdDel.Enabled = False
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillinspe"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub cmbIns_Click()
   bGoodIns = GetInspector()
   
End Sub


Private Sub cmbIns_LostFocus()
   cmbIns = CheckLen(cmbIns, 12)
   cmbIns = Compress(cmbIns)
   If Len(cmbIns) Then
      bGoodIns = GetInspector()
   Else
      lblNme = ""
      lblStp = ""
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If lblNme.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Inspector.", _
         vbExclamation, Caption
   Else
      ReactivateInspector
   End If
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6153
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillInspectors
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
   FormUnload
   Set InspRTf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub lblNme_Change()
   If Left(lblNme, 8) = "*** Insp" Then
      lblNme.ForeColor = ES_RED
   Else
      lblNme.ForeColor = vbBlack
   End If
   
End Sub


Private Sub ReactivateInspector()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "You Are Requesting To Reactivate This " & vbCr _
          & "Inspector. Do You Wish To Continue?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE RinsTable SET INSACTIVE=1 " _
             & "WHERE INSID='" & Compress(cmbIns) & "'"
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox cmbIns & " Inspector Is Active.", _
            vbInformation, Caption
         lblNme = ""
         lblStp = ""
         lblDiv = ""
         cmbIns.Clear
         FillInspectors
      Else
         MsgBox "Could Not Reactivate The Inspector.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
