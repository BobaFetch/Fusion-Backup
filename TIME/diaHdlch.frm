VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form diaHdlch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Daily Time Charge"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbEmp 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select From List Or Enter Number"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   4125
      TabIndex        =   1
      Tag             =   "4"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   875
   End
   Begin Threed.SSRibbon cmdHlp 
      Height          =   225
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Subject Help"
      Top             =   0
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   397
      _StockProps     =   65
      BackColor       =   12632256
      Autosize        =   2
      RoundedCorners  =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      PictureUp       =   "diaHdlch.frx":0000
      PictureDn       =   "diaHdlch.frx":0146
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   1680
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2160
      FormDesignWidth =   6345
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblSsn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4125
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name/SSN"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Card Date"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "diaHdlch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2005) is the property of            ***
'*** ESI Software Engineering, Inc, Stanwood, Washington, USA ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
Option Explicit

Dim bOnLoad As Byte
Dim bGoodCard As Byte
Dim sCardNumber As String

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbEmp_Click()
   GetEmployee
   
End Sub


Private Sub cmbEmp_KeyUp(KeyCode As Integer, Shift As Integer)
   cmbEmp = CheckLen(cmbEmp, 6)
   If Len(cmbEmp) Then
      cmbEmp = Format(cmbEmp, "000000")
      GetEmployee
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdDel_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   
   bGoodCard = GetCard()
   If bGoodCard Then
      sMsg = "Do You Really Want To Delete The Entry Of " & vbCrLf _
             & "This Time Card For " & lblNme & "?"
      bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
      If bResponse = vbYes Then
         On Error Resume Next
         sSql = "DELETE FROM TcitTable WHERE TCCARD='" & sCardNumber & "'"
         clsADOCon.ExecuteSQL sSql
         sSql = "DELETE FROM TchdTable WHERE TMCARD='" & sCardNumber & "'"
         clsADOCon.ExecuteSQL sSql
         If clsADOCon.RowsAffected > 0 Then
            MsgBox "Time Card Deleted..", vbInformation, Caption
         Else
            MsgBox "Couldn't Remove Time Card.", vbExclamation, Caption
         End If
      Else
         CancelTrans
      End If
   End If
   
End Sub

Private Sub cmdHlp_Click(Value As Integer)
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "hs1550"
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillCombo
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   
   txtDte = Format(Now - 1, "mm/dd/yy")
   If sCurrDate = "" Then
      If Format(txtDte, "w") = 1 Then
         txtDte = Format(Now - 2, "mm/dd/yy")
      End If
   Else
      txtDte = sCurrDate
   End If
   bOnLoad = 1
   Show
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   sCurrDate = txtDte
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set diaHdlch = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   On Error GoTo DiaErr1
   sSql = "Qry_FillEmployees"
   LoadNumComboBox cmbEmp, "000000"
   If cmbEmp.ListCount > 0 Then
      If Trim(sCurrEmployee) = "" Then
         cmbEmp = cmbEmp.List(0)
      Else
         cmbEmp = sCurrEmployee
      End If
      GetEmployee
   End If
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub GetEmployee()
   Dim RdoEmp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "Qry_EmployeeName " & Val(cmbEmp)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoEmp)
   If bSqlRows Then
      With RdoEmp
         cmbEmp = Format(!PREMNUMBER, "000000")
         lblNme = "" & Trim(!PREMLSTNAME) & ", " _
                  & Trim(!PREMFSTNAME) & " " _
                  & Trim(!PREMMINIT)
         lblSsn = "" & Trim(!PREMSOCSEC)
         .Cancel
         sCurrEmployee = cmbEmp
      End With
   Else
      MsgBox "Employee Wasn't Found.", vbExclamation, Caption
      lblNme = "No Current Employee"
      lblSsn = ""
   End If
   Set RdoEmp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getemploy"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetCard() As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TMCARD,TMEMP,TMDAY FROM TchdTable WHERE " _
          & "TMEMP=" & Val(cmbEmp) & " AND TMDAY='" & txtDte & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      GetCard = True
      sCardNumber = Trim(RdoGet!TMCARD)
      On Error Resume Next
      cmbEmp.SetFocus
   Else
      sCardNumber = ""
      sSql = "There Is No Time Card Recorded " & vbCrLf _
             & "For " & Trim(lblNme) & " On " & txtDte & "."
      'Beep
      MsgBox sSql, vbInformation, Caption
      GetCard = False
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getcard"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub txtDte_DropDown()
   ShowCalendar Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDate(txtDte)
   
End Sub
