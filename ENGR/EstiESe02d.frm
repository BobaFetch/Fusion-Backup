VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe02d 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estimate Outside Services"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe02d.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Index           =   5
      Left            =   4800
      TabIndex        =   19
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   2400
      Width           =   915
   End
   Begin VB.ComboBox cmbOsp 
      Height          =   315
      Index           =   5
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   17
      Tag             =   "3"
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   16
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   15
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   2040
      Width           =   915
   End
   Begin VB.ComboBox cmbOsp 
      Height          =   315
      Index           =   4
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   13
      Tag             =   "3"
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   11
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   1680
      Width           =   915
   End
   Begin VB.ComboBox cmbOsp 
      Height          =   315
      Index           =   3
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   9
      Tag             =   "3"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   7
      Tag             =   "1"
      ToolTipText     =   "Use for Operation Testing"
      Top             =   1320
      Width           =   915
   End
   Begin VB.ComboBox cmbOsp 
      Height          =   315
      Index           =   2
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "3"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CheckBox optLot 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      ToolTipText     =   "Check For Lot Charge"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtCst 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Tag             =   "1"
      ToolTipText     =   "Cost Unit Or Lot"
      Top             =   960
      Width           =   915
   End
   Begin VB.ComboBox cmbOsp 
      Height          =   315
      Index           =   1
      Left            =   960
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Description (Up To 60 Chars)"
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   0
      Top             =   3000
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3225
      FormDesignWidth =   6585
   End
   Begin VB.Label lblUpd 
      BackStyle       =   0  'Transparent
      Caption         =   "Updating Rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   2880
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label lblBid 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Total Services"
      Top             =   360
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Lot           "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   23
      Top             =   720
      Width           =   735
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   22
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Service                                                                                       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   21
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Item 5"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Item 4"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Item 3"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Item 2"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label It 
      BackStyle       =   0  'Transparent
      Caption         =   "Item 1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "EstiESe02d"
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
Dim bBadRow As Byte



Private Sub cmbOsp_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub cmbOsp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub

Private Sub cmbOsp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyCheck KeyAscii
   
End Sub


Private Sub cmbOsp_LostFocus(Index As Integer)
   cmbOsp(Index) = CheckLen(cmbOsp(Index), 60)
   cmbOsp(Index) = StrCase(cmbOsp(Index))
   
End Sub

Private Sub cmdCan_Click()
   UpdateRows
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3511
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      Caption = Caption & " - Estimate " & lblBid
      FillCombo
      GetServices
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim b As Byte
   Move 1000, 1000
   bOnLoad = 1
   For b = 1 To 4
      txtCst(b) = "0.000"
   Next
   txtCst(b) = "0.000"
   
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim bResponse As Byte
   Dim sMsg As String
   If bBadRow = 1 Then
      sMsg = "One Or More Entries Contains An Empty " & vbCrLf _
             & "Service With A Value And Won't Be Saved. " & vbCrLf _
             & "Continue Closing Anyway?"
      bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
      If bResponse = vbNo Then
         Cancel = True
      Else
         bResponse = GetBidServices(EstiESe02a, CCur("0" & EstiESe02a.txtQty))
      End If
   Else
      bResponse = GetBidServices(EstiESe02a, CCur("0" & EstiESe02a.txtQty))
   End If
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set EstiESe02d = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   'b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   sSql = "SELECT DISTINCT BIDOSSERVICE FROM EsosTable " _
          & "WHERE BIDOSSERVICE<>'' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            For b = 1 To 5
               AddComboStr cmbOsp(b).hwnd, "" & Trim(!BIDOSSERVICE)
            Next
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub optLot_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtCst_GotFocus(Index As Integer)
   SelectFormat Me
   
End Sub


Private Sub txtCst_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   CheckKeys KeyCode
   
End Sub


Private Sub txtCst_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyValue KeyAscii
   
End Sub



Private Sub GetServices()
   Dim RdoOsp As ADODB.Recordset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   sSql = "SELECT * FROM EsosTable WHERE BIDOSREF=" _
          & Val(lblBid) & " "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoOsp, ES_KEYSET)
   If bSqlRows Then
      With RdoOsp
         Do Until .EOF
            cmbOsp(!BIDOSROW) = "" & Trim(!BIDOSSERVICE)
            txtCst(!BIDOSROW) = Format(!BIDOSTOTALCOST, ES_QuantityDataFormat)
            optLot(!BIDOSROW).value = !BIDOSLOT
            .MoveNext
         Loop
         ClearResultSet RdoOsp
      End With
   Else
      For b = 1 To 5
         RdoOsp.AddNew
         RdoOsp!BIDOSREF = Val(lblBid)
         RdoOsp!BIDOSROW = b
         RdoOsp.Update
      Next
   End If
   Set RdoOsp = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "getservices"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub txtCst_LostFocus(Index As Integer)
   txtCst(Index) = CheckLen(txtCst(Index), 9)
   txtCst(Index) = Format(Abs(Val(txtCst(Index))), ES_QuantityDataFormat)
   
End Sub



'trying keysets again

Private Sub UpdateRows()
   Dim RdoOsp As ADODB.Recordset
   Dim b As Byte
   
   On Error GoTo DiaErr1
   MouseCursor 13
   lblUpd.Visible = True
   lblUpd.Refresh
   For b = 1 To 5
      sSql = "SELECT * FROM EsosTable WHERE " _
             & "(BIDOSREF=" & Val(lblBid) & " AND " _
             & "BIDOSROW=" & b & ") "
      bSqlRows = clsADOCon.GetDataSet(sSql, RdoOsp, ES_KEYSET)
      If bSqlRows Then
         With RdoOsp
            If Len(Trim(cmbOsp(b))) = 0 And Val(txtCst(b)) > 0 Then
               bBadRow = 1
               '.Edit
               !BIDOSSERVICE = ""
               !BIDOSTOTALCOST = 0
               !BIDOSLOT = 0
               .Update
            Else
               '.Edit
               !BIDOSSERVICE = cmbOsp(b)
               !BIDOSTOTALCOST = Val(txtCst(b))
               !BIDOSLOT = optLot(b).value
               .Update
            End If
         End With
         ClearResultSet RdoOsp
      End If
   Next
   Sleep 1000
   MouseCursor 0
   Set RdoOsp = Nothing
   lblUpd.Visible = False
   lblUpd.Refresh
   Exit Sub
   
DiaErr1:
   sProcName = "updaterows"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
