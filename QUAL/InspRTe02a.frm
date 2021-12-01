VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form InspRTe02a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise An Inspection Report"
   ClientHeight    =   2115
   ClientLeft      =   2685
   ClientTop       =   1425
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   Icon            =   "InspRTe02a.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2115
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "InspRTe02a.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "From New"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox optNew 
      Caption         =   "optNew"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbTyp 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "InspRTe02a.frx":0AB8
      Left            =   1800
      List            =   "InspRTe02a.frx":0AC8
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select Tag Type From List"
      Top             =   960
      Width           =   1455
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
      FormDesignHeight=   2115
      FormDesignWidth =   5175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "R&evise"
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Open To Revise An Inspection Report"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbTag 
      Height          =   315
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      Text            =   "cmbTag"
      Top             =   1320
      Width           =   1905
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Index           =   0
      Left            =   4200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tag Type"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revise Tag Number"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1545
   End
End
Attribute VB_Name = "InspRTe02a"
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
Dim AdoRej As ADODB.Command
Dim AdoParameter As ADODB.Parameter
Dim bGoodRecord As Byte
Dim bOnLoad As Byte

Dim iTagType As Integer
Dim sTagType As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub cmbTag_LostFocus()
   cmbTag = CheckLen(cmbTag, 12)
   
End Sub


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
   FillCombo sLastType
   
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


Private Sub cmdCan_Click(Index As Integer)
   Unload Me
   
End Sub


Private Sub cmdCan_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbTag = ""
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 6102
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdOk_Click(Index As Integer)
    GetTag
   If bGoodRecord = 0 Then
      MsgBox "Tag " & cmbTag & " Wasn't Found.", vbExclamation, Caption
   Else
      MouseCursor 13
      optNew.value = vbChecked
      InspRTe01b.optNew.value = vbUnchecked
      InspRTe01b.Caption = "Revise " & InspRTe01b.Caption
      InspRTe01b.lblTag = cmbTag
      InspRTe01b.lblType = sTagType
      InspRTe01b.Show
      Unload Me
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      bOnLoad = 0
      FillCombo sLastType
      If optFrom.value = vbChecked Then
         cmbTag = InspRTe01a.txtTag
         optFrom.value = vbUnchecked
         Unload InspRTe01a
         GetTag
      End If
   End If
   optNew.value = vbUnchecked
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   sSql = "Select REJREF,REJNUM,REJTYPE FROM RjhdTable WHERE REJTYPE= ?"
   Set AdoRej = New ADODB.Command
   AdoRej.CommandText = sSql
   
   Set AdoParameter = New ADODB.Parameter
   AdoParameter.Type = adChar
   AdoParameter.Size = 1
   
   AdoRej.Parameters.Append AdoParameter
   
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
         cmbTyp = "Vendor"
         sTagType = "Vendor Tag"
   End Select
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   If optNew.value = vbUnchecked Then FormUnload
   Set AdoParameter = Nothing
   Set AdoRej = Nothing
   Set InspRTe02a = Nothing

End Sub



Private Sub FillCombo(sType As String)
   Dim AdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbTag.Clear
   'RdoRej(0) = sType
   AdoRej.Parameters(0).value = sType
   bSqlRows = clsADOCon.GetQuerySet(AdoCmb, AdoRej)
   If bSqlRows Then
      With AdoCmb
         cmbTag = "" & Trim(!REJNUM)
         Do Until .EOF
            AddComboStr cmbTag.hwnd, "" & Trim(!REJNUM)
            .MoveNext
         Loop
         ClearResultSet AdoCmb
      End With
   End If
   Set AdoCmb = Nothing
   GetTag
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub


Private Function GetTag()
   Dim RdoTag As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_GetRejectionTag '" & Compress(cmbTag) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoTag, ES_FORWARD)
   If bSqlRows Then
      With RdoTag
         cmbTag = "" & Trim(!REJNUM)
         Select Case !REJTYPE
            Case "C" 'Customer
               cmbTyp = "Customer"
               sTagType = "Customer Tag"
            Case "I" 'Internal
               cmbTyp = "Internal"
               sTagType = "Internal Tag"
            Case "V" 'Vendor
               cmbTyp = "Vendor"
               sTagType = "Vendor Tag"
            Case "M" 'MRB
               cmbTyp = "MRB"
               sTagType = "MRB Tag"
         End Select
         ' Set the last type
         sLastType = !REJTYPE
         ClearResultSet RdoTag
      End With
      bGoodRecord = 1
   Else
      bGoodRecord = 0
   End If
   Set RdoTag = Nothing
   Exit Function
   
DiaErr1:
   bGoodRecord = 0
   sProcName = "gettag"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub optFrom_Click()
   'never visible - from new tag?
   
End Sub
