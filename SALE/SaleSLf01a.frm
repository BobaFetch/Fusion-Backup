VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SaleSLf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Sales Order"
   ClientHeight    =   2550
   ClientLeft      =   1845
   ClientTop       =   1065
   ClientWidth     =   6465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   4932
      _ExtentX        =   8705
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "SaleSLf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox lblCst 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   360
      Top             =   0
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   6465
   End
   Begin VB.CommandButton cmdCnc 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5520
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Cancel This Sales Order"
      Top             =   480
      Width           =   915
   End
   Begin VB.ComboBox cmbSon 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1740
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Select Sales Order Number From List"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   915
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Order"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1320
      Width           =   3375
   End
End
Attribute VB_Name = "SaleSLf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'Added ITINVOICE
'10/8/99 New
'4/26/05 Removed Combo
'10/6/06 Streamlined GetItems and added RnalTable (MO Allocations) to
'       CancelSalesOrder
Option Explicit
Dim bOnLoad As Byte
Dim bUnload As Boolean
Dim bGoodSO As Byte
Dim bGoodIt As Byte

Private txtKeyPress As New EsiKeyBd

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder()
End Sub

Private Sub cmbSon_LostFocus()
   bGoodSO = GetSalesOrder()
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub


'Private Sub cmdCnc_Click()
'   Dim bShippedInvoiced
'   bShippedInvoiced = GetItems(cmbSon)
'   If bShippedInvoiced = 0 Then
'      CancelSalesOrder
'   Else
'      MsgBox "This Sales Order Has Shipped Or Invoiced Items" & vbcrlf _
'         & "And Cannot Be Canceled. Suggest Checking Sales " & vbcrlf _
'         & "Order Items And Canceling Unwanted Entries.", vbExclamation, Caption
'   End If
'
'End Sub
'

Private Sub cmdCnc_Click()
   CancelSalesOrder
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2150
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      FillSalesOrders
      lblCst.BackColor = Me.BackColor
      bOnLoad = 0
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me, ES_DONTLIST
   bOnLoad = 1
   
End Sub

Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   FormUnload
   Set SaleSLf01a = Nothing
   
End Sub


Private Function GetSalesOrder() As Byte
   Dim rdo As ADODB.Recordset
   
   On Error GoTo whoops
   sSql = "SELECT SONUMBER, SOTYPE, SOCUST, CUNAME FROM SohdTable " & vbCrLf _
      & "join CustTable on SOCUST = CUREF" & vbCrLf _
      & "where SONUMBER = " & cmbSon & " "
   If clsADOCon.GetDataSet(sSql, rdo) Then
      With rdo
         lblPre = "" & Trim(!SOTYPE)
         lblCst = Trim(!SOCUST)
         lblNme = !CUNAME
         cmdCnc.Enabled = True
         GetSalesOrder = 1
      End With
   Else
      GetSalesOrder = 0
      cmdCnc.Enabled = False
      lblPre = ""
      lblCst = ""
      lblNme = ""
   End If
   
   Exit Function
   
whoops:
   cmdCnc.Enabled = False
   lblPre = ""
   lblCst = ""
   lblNme = ""
   sProcName = "GetSalesOrder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me

End Function

Private Sub CancelSalesOrder()
   Dim bResponse As Byte
   Dim lSoNumber As Long
   Dim sMsg As String
   lSoNumber = Val(cmbSon)
   sMsg = "This function is final.  It removes MO allocations." & vbCrLf _
          & "Are You Sure That You Wish To Cancel Sales Order  " & vbCrLf _
          & Trim(lblPre) & cmbSon & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   
   On Error GoTo whoops
   If bResponse = vbYes Then
      cmdCnc.Enabled = False
      MouseCursor ccHourglass
      'Sleep 1000
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE SoitTable SET ITCANCELDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
             & "ITCANCELED=1 WHERE ITSO=" & lSoNumber & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "UPDATE SohdTable SET SOCANDATE='" & Format(ES_SYSDATE, "mm/dd/yy") & "'," _
             & "SOCANCELED=1 WHERE SONUMBER=" & lSoNumber & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
      sSql = "DELETE FROM RnalTable WHERE RASO=" & lSoNumber & " "
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      
'      sMsg = "Last Chance. Are You Sure That You Wish " _
'             & "To Cancel Sales Order " & Trim(lblPre) & cmbSon & "."
'      MouseCursor 0
'      If bResponse = vbYes Then
         clsADOCon.CommitTrans
'      Else
'         RdoCon.RollbackTrans
'         CancelTrans
'         Exit Sub
'      End If
'      If RdoCon.RowsAffected > 0 And Err = 0 Then
         MsgBox "Sales order " & Trim(lblPre) & cmbSon & " canceled.", vbInformation, Caption
'      Else
'         MsgBox "Couldn't Cancel The Sales Order.", vbExclamation, Caption
'      End If
'      Sleep 1000
'      Unload Me
'   Else
'      CancelTrans
      MouseCursor ccDefault
   End If
   FillSalesOrders
   Exit Sub
   
whoops:
'   MouseCursor ccDefault
'   sProcName = "CancelSalesOrder"
'   CurrError.Number = Err.Number
'   CurrError.Description = Err.Description
'   RdoCon.RollbackTrans
'   DoModuleErrors Me

   ProcessError "CancelSalesOrder"
End Sub

'Private Sub txtSon_KeyPress(KeyAscii As Integer)
'   KeyValue KeyAscii
'
'End Sub
'
'Private Sub txtSon_LostFocus()
'   txtSon = Format(Abs(Val(txtSon)), "00000")
'   cmbSon = txtSon
'   cmbSon_Click
'
'End Sub
'
Private Sub FillSalesOrders()

   Dim RdoCmb As ADODB.Recordset
   Dim iList As Integer
   
   On Error GoTo DiaErr1
   MouseCursor 13
   
   sSql = "select SONUMBER from SohdTable" & vbCrLf _
      & "where SONUMBER not in (select ITSO from SoitTable where ITACTUAL is not null)" & vbCrLf _
      & "and SOCANCELED=0" & vbCrLf _
      & "order by SONUMBER"
'   LoadComboBox cmbSon, -1

   cmbSon.Clear
      
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb, ES_FORWARD)
   If bSqlRows Then
      iList = -1
      With RdoCmb
         Do Until .EOF
            iList = iList + 1
            'If iList > 999 Then Exit Do
            AddComboStr cmbSon.hWnd, "" & Format$(Trim(!SoNumber), SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   Else
      MouseCursor 0
      Exit Sub
   End If
   If cmbSon.ListCount <> 0 Then
      bSqlRows = 1
      cmbSon.ListIndex = 0
   Else
      bSqlRows = 0
   End If
   
   Set RdoCmb = Nothing
   MouseCursor 0
   Exit Sub
   
DiaErr1:
   sProcName = "FillSalesOrders"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
