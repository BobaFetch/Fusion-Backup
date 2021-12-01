VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete An Inspection Report"
   ClientHeight    =   1995
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
   ScaleHeight     =   1995
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbTag 
      Height          =   315
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "8"
      Text            =   " "
      ToolTipText     =   "Select Report Number To Be Deleted (Contains Only Qualifying Reports)"
      Top             =   840
      Width           =   1905
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Remove This Rejection Tag"
      Top             =   600
      Width           =   875
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   4440
      Top             =   1440
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   1995
      FormDesignWidth =   5160
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rejection Tag Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tag Type"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label lblTyp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "InspRTf01a"
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
      lblTyp.Width = (cmbTag.Width * 0.85)
      GetTag = 1
   Else
      lblTyp.Width = cmbTag.Width
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

Private Sub cmbTag_Click()
   bGoodTag = GetTag()
   
End Sub


Private Sub cmbTag_LostFocus()
   cmbTag = CheckLen(cmbTag, 12)
   bGoodTag = GetTag()
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   If Trim(cmbTag) <> "" Then
      If lblTyp.ForeColor = ES_RED Then
         MsgBox "Requires A Valid Inspection Report.", _
            vbExclamation, Caption
      Else
         DeleteTag
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6150
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
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
   Set InspRTf01a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillTags()
   On Error GoTo DiaErr1
   sSql = "SELECT RjhdTable.REJREF,RjhdTable.REJNUM FROM " _
          & "RjhdTable LEFT JOIN RjitTable ON RjhdTable.REJREF=" _
          & "RjitTable.RITREF WHERE (RjitTable.RITREF Is Null)"
   LoadComboBox cmbTag
   If cmbTag.ListCount > 0 Then cmbTag = cmbTag.List(0)
   Exit Sub
   
DiaErr1:
   sProcName = "filltags"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblTyp_Change()
   If Left(lblTyp, 8) = "***Tag W" Then
      lblTyp.ForeColor = ES_RED
   Else
      lblTyp.ForeColor = vbBlack
   End If
   
End Sub


Private Sub DeleteTag()
   Dim bTagHasItems As Boolean
   Dim bResponse As Byte
   Dim sMsg As String
   
   On Error GoTo DiaErr1
   sMsg = "You Are About To Permanently Delete All" & vbCr _
          & "Record Of This Inspection Report. Continue?...  "
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      bTagHasItems = CheckTag()
      If bTagHasItems Then
         MsgBox "This Tag Has Rejection Items.     " & vbCr _
            & "Delete All Tag Items First...    ", _
            vbExclamation
      Else
         sMsg = "No Inspection Report Items. Continue To Delete?...  "
         bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            On Error Resume Next
            clsADOCon.ADOErrNum = 0
            sSql = "DELETE FROM RjhdTable WHERE " _
                   & "REJREF='" & Compress(cmbTag) & "'"
            clsADOCon.ExecuteSQL sSql
            If clsADOCon.ADOErrNum = 0 Then
               MsgBox "The Tag Was Successfully Deleted.", _
                  vbInformation, Caption
               cmbTag.Clear
               FillTags
            Else
               MsgBox "Could Not Successfully Delete The Tag.", _
                  vbExclamation, Caption
            End If
         Else
            CancelTrans
         End If
         
      End If
   Else
      CancelTrans
   End If
   
   Exit Sub
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function CheckTag() As Boolean
   Dim RdoItm As ADODB.Recordset
   sSql = "SELECT RITREF FRom RjitTable WHERE " _
          & "RITREF='" & Compress(cmbTag) & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      ClearResultSet RdoItm
      CheckTag = True
   Else
      CheckTag = False
   End If
   Set RdoItm = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checktag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function
