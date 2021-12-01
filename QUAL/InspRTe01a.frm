VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Inspection Report"
   ClientHeight    =   1845
   ClientLeft      =   2625
   ClientTop       =   1395
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   HelpContextID   =   6101
   Icon            =   "InspRTe01a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1845
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe01a.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optNew 
      Caption         =   "New"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InspRTe01a.frx":0AB8
      Left            =   1680
      List            =   "InspRTe01a.frx":0AC8
      TabIndex        =   1
      Tag             =   "8"
      ToolTipText     =   "Select Tag Type From List"
      Top             =   1080
      Width           =   1455
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
      FormDesignHeight=   1845
      FormDesignWidth =   5160
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4230
      TabIndex        =   4
      ToolTipText     =   "Add The New Inspection Report"
      Top             =   600
      Width           =   875
   End
   Begin VB.TextBox txtTag 
      Height          =   285
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   0
      Tag             =   "3"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   4230
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label lblTyp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label lblTag 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Report"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Report"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1635
   End
End
Attribute VB_Name = "InspRTe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/02/05 Changed the KeySet
'4/18/06 Revisited and changed messages
Option Explicit
Dim RdoNew As ADODB.Recordset

Dim bGoodRecord As Boolean
Dim bOnLoad As Byte
Dim bGoodIns As Byte

Dim sTagType As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbTyp_Click()
   Select Case cmbTyp
      Case "Vendor"
         sLastType = "V"
         sTagType = "Vendor Tag"
      Case "MRB"
         sLastType = "M"
         sTagType = "MRB Tag"
      Case "Customer"
         sLastType = "C"
         sTagType = "Customer Tag"
      Case Else
         cmbTyp = "Internal"
         sLastType = "I"
         sTagType = "Internal Tag"
   End Select
   
End Sub


Private Sub cmbTyp_LostFocus()
   Select Case cmbTyp
      Case "Vendor"
         sLastType = "V"
         sTagType = "Vendor Tag"
      Case "MRB"
         sLastType = "M"
         sTagType = "MRB Tag"
      Case "Customer"
         sLastType = "C"
         sTagType = "Customer Tag"
      Case Else
         cmbTyp = "Internal"
         sLastType = "I"
         sTagType = "Internal Tag"
   End Select
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6101
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOk_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   If Len(Trim(txtTag)) = 0 Then Exit Sub
   bGoodRecord = GetTag()
   If bGoodRecord Then
      MouseCursor 0
      sMsg = "That Tag Is Already Recorded." & vbCr & "Do You Want To Revise It?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbYes Then
         InspRTe02a.optFrom.Value = vbChecked
         InspRTe02a.Show
      Else
         CancelTrans
         txtTag = ""
      End If
   Else
      AddTag
   End If
   
End Sub

Private Sub Form_Activate()
   MouseCursor 0
   optNew.Value = vbUnchecked
   If bOnLoad Then
      GetLastTag
      bGoodIns = CheckInspectors()
      bOnLoad = 0
   End If
   If Not bGoodIns Then MsgBox "Please Add At Least One Inspector.", _
      vbExclamation, Caption
   MdiSect.lblBotPanel = Caption
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   On Error Resume Next
   sSql = "SELECT REJLASTTAG FROM ComnTable WHERE COREF=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew, ES_FORWARD)
   If bSqlRows Then
      lblTag = "" & Trim(RdoNew!REJLASTTAG)
   End If
   'RdoNew.Close
   sLastType = GetSetting("Esi2000", "Quality", "LastType", sLastType)
   Select Case sLastType
      Case "C"
         cmbTyp = "Customer"
         sTagType = "Customer Tag"
      Case "I"
         cmbTyp = "Internal"
         sTagType = "Internal Tag"
      Case "M"
         cmbTyp = "MRB"
         sTagType = "MRB Tag"
      Case "V"
         cmbTyp = "Vendor  "
         sTagType = "Vendor Tag"
   End Select
   bOnLoad = 1
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sLastType = Left(cmbTyp, 1)
   SaveSetting "Esi2000", "Quality", "LastType", sLastType
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If optNew.Value = vbUnchecked Then FormUnload
   Set RdoNew = Nothing
   Set InspRTe01a = Nothing
   
End Sub








Private Sub AddTag()
   Dim bResponse As Byte
   Dim sNewTag As String
   bResponse = IllegalCharacters(txtTag)
   If bResponse > 0 Then
      MsgBox "The Report ID Contains An Illegal " & Chr$(bResponse) & ".", _
         vbExclamation, Caption
      Exit Sub
   End If
   bResponse = MsgBox("Add Inspection Report " & txtTag & "?", ES_YESQUESTION, Caption)
   If bResponse = vbNo Then
      bGoodRecord = False
      MouseCursor 0
      txtTag = ""
      On Error Resume Next
      cmdCan.SetFocus
      Exit Sub
   End If
   MouseCursor 13
   sNewTag = Compress(txtTag)
   On Error Resume Next
   'RdoNew.Close
   On Error GoTo DiaErr1
   sSql = "Select * FROM RjhdTable"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew, ES_KEYSET)
   RdoNew.AddNew
   RdoNew!REJREF = "" & sNewTag
   RdoNew!REJNUM = "" & txtTag
   RdoNew!REJTYPE = "" & Left(sTagType, 1)
   RdoNew!REJUSEDON = "D"
   RdoNew.Update
   MouseCursor 0
   SysMsg "Report " & txtTag & " Added.", True
   On Error Resume Next
   If Left(sTagType, 1) = "I" Or Left(sTagType, 1) = "V" Then
      clsADOCon.ExecuteSQL "UPDATE ComnTable SET REJLASTTAG='" & txtTag & "'"
   End If
   sSql = "Select * FROM RjhdTable WHERE REJREF='" & sNewTag & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoNew, ES_KEYSET)
   bGoodRecord = True
   optNew.Value = vbChecked
   InspRTe01b.optNew.Value = vbChecked
   InspRTe01b.Caption = "New " & InspRTe01b.Caption
   InspRTe01b.lblTag = txtTag
   InspRTe01b.lblType = sTagType
   InspRTe01b.Show
   Unload Me
   Exit Sub
   
DiaErr1:
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   bGoodRecord = False
   MouseCursor 0
   MsgBox CurrError.Description & vbCr & "Couldn't Add Tag.", vbExclamation, Caption
   
End Sub

Private Function GetTag() As Boolean
   Dim RdoTag As ADODB.Recordset
   
   On Error GoTo DiaErr1
   sSql = "Qry_GetRejectionTag '" & Compress(txtTag) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTag, ES_FORWARD)
   If bSqlRows Then
      With RdoTag
         txtTag = "" & Trim(!REJNUM)
         lblTyp = "" & Trim(!REJTYPE)
         GetTag = True
         ClearResultSet RdoTag
      End With
   Else
      GetTag = False
   End If
   Set RdoTag = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function




Private Sub txtTag_LostFocus()
   txtTag = CheckLen(txtTag, 12)
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

'see if there are inspectors, if not warn them

Private Function CheckInspectors() As Byte
   Dim RdoCmb As ADODB.Recordset

   On Error GoTo DiaErr1
   sSql = "Qry_FillInspectorsAll"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         CheckInspectors = True
         ClearResultSet RdoCmb
      End With
   Else
      CheckInspectors = False
   End If
   Set RdoCmb = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkinsp"
   CurrError.Description = Err.Description
   Resume DiaErr2
DiaErr2:
   DoModuleErrors Me
   
End Function

Private Sub GetLastTag()
   Dim RdoLst As ADODB.Recordset

   sSql = "SELECT MAX(REJDATE),MAX(CONVERT(INT, REJNUM)) FROM RjhdTable "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoLst, ES_FORWARD)
   If bSqlRows Then
      With RdoLst
         If Not IsNull(.Fields(1)) Then
            lblTag = Trim(.Fields(1))
            If ((lblTag <> "") And IsNumeric(lblTag)) Then
               Dim totLen As Long
               Dim iNewNum As Long
               Dim strNewNum As String
               
               totLen = Len(Trim(lblTag))
               iNewNum = Val(lblTag) + 1
               strNewNum = Format(CStr(iNewNum), String(totLen, "0"))
               txtTag.SelText = strNewNum
            Else
               txtTag.SelText = lblTag
            End If
            
         End If
         ClearResultSet RdoLst
      End With
   End If
   
   Set RdoLst = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
