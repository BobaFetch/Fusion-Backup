VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTf02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change An Inspection Report Type Flag"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTf02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Change To New Tag Type"
      Top             =   480
      Width           =   875
   End
   Begin VB.ComboBox cmbTag 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      Text            =   " "
      ToolTipText     =   "Select Or Enter Tag Number"
      Top             =   720
      Width           =   1905
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InspRTf02a.frx":07AE
      Left            =   1920
      List            =   "InspRTf02a.frx":07BE
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Tag Type From List"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4200
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2115
      FormDesignWidth =   5160
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tag Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revise Tag Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Tag Type"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1545
   End
End
Attribute VB_Name = "InspRTf02a"
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
Dim bGoodTag As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbTag_Click()
   bGoodTag = GetTag()
   
End Sub


Private Sub cmbTag_LostFocus()
   cmbTag = CheckLen(cmbTag, 12)
   bGoodTag = GetTag()
   
End Sub


Private Sub cmbTyp_LostFocus()
   If Trim(cmbTyp) = "" Then cmbTyp = cmbTyp.List(0)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdChg_Click()
   If lblTyp.ForeColor = ES_RED Then
      MsgBox "Requires A Valid Inspection Report.", _
         vbExclamation, Caption
   Else
      If lblTyp = cmbTyp Then
         MsgBox "The New And Old Types Are The Same.", _
            vbExclamation, Caption
      Else
         ChangeType
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6151
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      cmbTyp = cmbTyp.List(0)
      FillTags
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
   Set InspRTf02a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillTags()
   'Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_FillRejectionTags"
   LoadComboBox cmbTag
   If cmbTag.ListCount > 0 Then cmbTag = cmbTag.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "filltags"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetTag() As Byte
   Dim RdoTag As ADODB.Recordset
   Dim sType As String
   On Error GoTo DiaErr1
   sSql = "Qry_GetRejectionTag '" & Compress(cmbTag) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTag, ES_FORWARD)
   If bSqlRows Then
      With RdoTag
         cmbTag = "" & Trim(!REJNUM)
         sType = "" & Trim(!REJTYPE)
         ClearResultSet RdoTag
      End With
      Select Case sType
         Case "I" 'Internal
            lblTyp = "Internal"
         Case "V" 'Vendor
            lblTyp = "Vendor"
         Case "M" 'MRB
            lblTyp = "MRB"
         Case Else 'Customer
            lblTyp = "Customer"
      End Select
      lblTyp.Width = (cmbTyp.Width * 0.85)
      GetTag = 1
   Else
      lblTyp.Width = cmbTyp.Width
      lblTyp = "***Tag Wasn't Found***"
      GetTag = 0
   End If
   Set RdoTag = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblTyp_Change()
   If Left(lblTyp, 8) = "***Tag W" Then
      lblTyp.ForeColor = ES_RED
   Else
      lblTyp.ForeColor = vbBlack
   End If
   
End Sub


Private Sub ChangeType()
   Dim bResponse As Byte
   Dim sMsg As String
   
   sMsg = "You Are About To Change The Type " & vbCr _
          & "Flag Of This Inspection Report. Continue?...  "
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE RjhdTable SET REJTYPE='" _
             & Left(cmbTyp, 1) & "' WHERE REJREF='" _
             & Compress(cmbTag) & "' "
      clsADOCon.ExecuteSQL sSql
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "Inspection Report Type Was Successfully Changed.", _
            vbInformation, Caption
         bGoodTag = GetTag()
      Else
         MsgBox "Could Not Change Inspection Report Type.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
