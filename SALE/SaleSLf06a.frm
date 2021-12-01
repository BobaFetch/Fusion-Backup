VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form SaleSLf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revise Booked Dates"
   ClientHeight    =   2880
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   2160
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select or Enter Sales Order Number (Contains 300 Max)"
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "4"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox optItm 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Change Item Booked Dates"
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   6000
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Change Sales Order Number"
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Enter Sales Order"
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6240
      Top             =   2400
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2880
      FormDesignWidth =   6975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revise Item Dates"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revised Booked Date"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Booked Date"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1785
   End
   Begin VB.Label lblCst 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   1212
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   3240
      TabIndex        =   7
      Top             =   1200
      Width           =   3372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order Number"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label lblOld 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "SaleSLf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'10/17/05 Added ComboBox
Option Explicit
Dim bOldExists As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   Dim sYear As String
   
   On Error GoTo DiaErr1
   MouseCursor 13
   cmbSon.Clear
   'iList = Format(Now, "yyyy")
   'iList = iList - 2
   'sYear = Trim$(iList) & "-" & Format(Now, "mm-dd")
   'sSql = "Qry_FillSalesOrders '" & sYear & "'"
   sSql = "Qry_FillSalesOrders '" & DateAdd("yyyy", -2, Now) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      iList = -1
      With RdoCmb
         lblOld = "" & Trim(!SOTYPE)
         cmbSon = Format(!SoNumber, SO_NUM_FORMAT)
         Do Until .EOF
            iList = iList + 1
            If iList > 999 Then Exit Do
            AddComboStr cmbSon.hWnd, Format$(!SoNumber, SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MouseCursor 0
      MsgBox "No Sales Orders Where Found.", vbInformation, Caption
      Exit Sub
   End If
   Set RdoCmb = Nothing
   txtOld = cmbSon
   If cmbSon.ListCount > 0 Then bOldExists = GetSalesOrder(txtOld)
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub

Private Sub cmbSon_Click()
   txtOld = cmbSon
   bOldExists = GetSalesOrder(txtOld)
   
End Sub

Private Sub cmbSon_LostFocus()
   cmbSon = CheckLen(cmbSon, SO_NUM_SIZE)
   cmbSon = Format(Abs(Val(cmbSon)), SO_NUM_FORMAT)
   txtOld = cmbSon
   bOldExists = GetSalesOrder(txtOld)
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdChg_Click()
   If bOldExists Then
      UpdateSalesOrder
   Else
      MsgBox "Requires A Valid Sales Order Number.", vbExclamation, Caption
      
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2156
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad = 1 Then FillCombo
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
   bOnLoad = 1
   
End Sub


Function GetSalesOrder(lSalesOrder As Variant) As Byte
   Dim RdoGet As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SONUMBER,SOTYPE,SOCUST,SODATE FROM SohdTable " _
          & "WHERE SONUMBER=" & Trim(str(lSalesOrder)) & " " _
          & "AND SOCANCELED=0 "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoGet)
   If bSqlRows Then
      With RdoGet
         lblOld = "" & Trim(!SOTYPE)
         lblDte = "" & !SODATE
         FindCustomer Me, !SOCUST, False
         ClearResultSet RdoGet
         GetSalesOrder = True
      End With
   Else
      Beep
      lblCst = "****No Such "
      lblNme = "Sales Order Or Sales Order Canceled****"
      GetSalesOrder = False
   End If
   Set RdoGet = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getsaleso"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf06a = Nothing
   
End Sub

Private Sub optItm_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub

Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   
End Sub

Private Sub txtOld_LostFocus()
   txtOld = CheckLen(txtOld, SO_NUM_SIZE)
   txtOld = Format(Abs(Val(txtOld)), SO_NUM_FORMAT)
   If Val(txtOld) > 0 Then
      bOldExists = GetSalesOrder(txtOld)
   Else
      bOldExists = False
      lblCst = ""
      lblNme = ""
   End If
   
End Sub



Private Sub UpdateSalesOrder()
   Dim bResponse As Integer
   Dim sMsg As String
   
   sMsg = "Are You Sure That You Wish To Revise " & vbCrLf _
          & "The Sales Order " & lblOld & txtOld & " Booking Date ?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbNo Then Exit Sub
   MouseCursor 13
   cmdChg.Enabled = False
   On Error GoTo DiaErr1
   sSql = "UPDATE SohdTable SET SODATE='" & txtDte & "' WHERE " _
          & "SONUMBER=" & txtOld & " "
   clsADOCon.ExecuteSQL sSql 'rdExecDirect
   If clsADOCon.RowsAffected > 0 Then
      If optItm Then
         sSql = "UPDATE SoitTable SET ITBOOKDATE='" & txtDte & "' WHERE " _
                & "ITSO=" & txtOld & " AND ITCANCELED=0"
         clsADOCon.ExecuteSQL sSql 'rdExecDirect
      End If
      MouseCursor 0
      lblDte = txtDte
      MsgBox "Date Changed.", vbInformation, Caption
   Else
      MouseCursor 0
      MsgBox "Unable To Change Dates.", vbExclamation, Caption
   End If
   cmdChg.Enabled = True
   On Error Resume Next
   txtOld.SetFocus
   Exit Sub
   
DiaErr1:
   Resume DiaErr2
DiaErr2:
   On Error Resume Next
   MouseCursor 0
   MsgBox "Couldn't Change The Booked Date.", vbExclamation, Caption
   
End Sub
