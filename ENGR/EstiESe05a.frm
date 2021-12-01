VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form EstiESe05a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requests For Quotation"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbCompleted 
      Height          =   315
      Left            =   2040
      TabIndex        =   19
      Top             =   2850
      Width           =   1095
   End
   Begin VB.ComboBox cmbRFQDueAmPm 
      Height          =   315
      ItemData        =   "EstiESe05a.frx":0000
      Left            =   5880
      List            =   "EstiESe05a.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.ComboBox cmbRFQDateAMPM 
      Height          =   315
      ItemData        =   "EstiESe05a.frx":0016
      Left            =   3000
      List            =   "EstiESe05a.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "EstiESe05a.frx":002C
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optCom 
      Caption         =   "____"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox txtBuy 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Tag             =   "2"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox txtDue 
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Tag             =   "4"
      Top             =   2160
      Width           =   1250
   End
   Begin VB.ComboBox txtDte 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Tag             =   "4"
      Top             =   2160
      Width           =   1250
   End
   Begin VB.ComboBox cmbCst 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select A Customer"
      Top             =   720
      Width           =   1555
   End
   Begin VB.TextBox txtDsc 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1800
      Width           =   4635
   End
   Begin VB.ComboBox cmbRfq 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "3"
      ToolTipText     =   "Select Or Enter RFQ Number (14 Chars Max)"
      Top             =   1440
      Width           =   2040
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2880
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   3300
      FormDesignWidth =   6615
   End
   Begin VB.Label txtNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   15
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ By"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ Due"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   12
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ Date"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "RFQ Number"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Enter Or Select An RFQ"
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "EstiESe05a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'4/13/06 Fixed default Customer (cmbCst)
Option Explicit
Dim RdoRfq As ADODB.Recordset

Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodCst As Byte
Dim bGoodRfq As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd





Private Sub cmbCompleted_DropDown()
   ShowCalendarEx Me
End Sub

Private Sub cmbCompleted_LostFocus()
   txtDte = CheckDateEx(cmbCompleted)
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQCOMPLETED = Format(cmbCompleted, "mm/dd/yyyy")
      RdoRfq.Update
   End If
End Sub

Private Sub cmbCst_Click()
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst
   bGoodRfq = GetThisRfq()
   
End Sub


Private Sub cmbCst_LostFocus()
   Dim b As Byte
   cmbCst = CheckLen(cmbCst, 10)
   FindCustomer Me, cmbCst
   FillCustomerRFQs Me, cmbCst
   bGoodRfq = GetThisRfq()
   
End Sub


Private Sub cmbRfq_Click()
   bGoodRfq = GetThisRfq()
   
   
End Sub


Private Sub cmbRfq_LostFocus()
   cmbRfq = CheckLen(cmbRfq, 14)
   If bCancel = 1 Then Exit Sub
   If Len(Trim(cmbRfq)) > 3 Then
      bGoodRfq = GetThisRfq()
   Else
      MsgBox "4 Or More Characters Please.", _
         vbInformation, Caption
      bGoodRfq = 0
      Exit Sub
   End If
   If bGoodRfq = 2 Then
      MsgBox "Customer/RFQ Mismatch.", _
         vbInformation, Caption
   Else
      If bGoodRfq = 0 Then AddNewRfq
   End If
   
End Sub



Private Sub cmbRFQDateAMPM_LostFocus()
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQDATEAMPM = cmbRFQDateAMPM.ListIndex
      RdoRfq.Update
   End If
End Sub


Private Sub cmbRFQDueAmPm_LostFocus()
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQDUEAMPM = cmbRFQDueAmPm.ListIndex
      RdoRfq.Update
   End If
End Sub

Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 3505
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      'kill a problem with a default edit out after 12/06/01
      sSql = "UPDATE RfqsTable SET RFQCOMPLETE=0 WHERE " _
             & "RFQCOMPLETE IS NULL"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      FillCustomers
      If cmbCst.ListCount > 0 Then
         cmbCst = cmbCst.List(0)
         FindCustomer Me, cmbCst
         FillCustomerRFQs Me, cmbCst
         If cmbRfq.ListCount > 0 Then bGoodRfq = GetThisRfq()
      End If
      bOnLoad = 0
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
   FormUnload
   Set RdoRfq = Nothing
   Set EstiESe05a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   txtDte = Format(ES_SYSDATE, "mm/dd/yyyy")
   txtDue = Format(ES_SYSDATE, "mm/dd/yyyy")
   
End Sub


Private Sub optCom_Click()
   On Error Resume Next
   If optCom.Value = vbChecked Then
      If bGoodRfq Then
         'RdoRfq.Edit
         RdoRfq!RFQCOMPLETE = 1
         RdoRfq!RFQCOMPLETED = Format(ES_SYSDATE, "mm/dd/yyyy")
         RdoRfq.Update
    End If
        cmbCompleted.Enabled = True
        cmbCompleted.Visible = True

   Else
      If bGoodRfq Then
         'RdoRfq.Edit
         RdoRfq!RFQCOMPLETE = 0
         RdoRfq!RFQCOMPLETED = Null
         RdoRfq.Update
      End If
        cmbCompleted.Enabled = False
        cmbCompleted.Visible = False
   End If
End Sub

Private Sub optCom_KeyPress(KeyAscii As Integer)
   KeyLock KeyAscii
   
End Sub


Private Sub txtBuy_LostFocus()
   txtBuy = CheckLen(txtBuy, 20)
   txtBuy = StrCase(txtBuy)
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQBY = txtBuy
      RdoRfq.Update
   End If
   
End Sub


Private Sub txtDsc_LostFocus()
   txtDsc = CheckLen(txtDsc, 40)
   txtDsc = StrCase(txtDsc)
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQDESC = txtDsc
      RdoRfq.Update
   End If
   
End Sub


Private Sub txtDte_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtDte_LostFocus()
   txtDte = CheckDateEx(txtDte)
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQDATE = Format(txtDte, "mm/dd/yyyy")
      RdoRfq.Update
   End If
   
End Sub


Private Sub txtDue_DropDown()
   ShowCalendarEx Me
   
End Sub




Private Sub txtDue_LostFocus()
   txtDue = CheckDateEx(txtDue)
   If bGoodRfq Then
      On Error Resume Next
      'RdoRfq.Edit
      RdoRfq!RFQDUE = Format(txtDte, "mm/dd/yyyy")
      RdoRfq.Update
   End If
   
End Sub

Private Sub txtNme_Change()
   If Left(txtNme, 8) = "*** Cust" Then
      txtNme.ForeColor = ES_RED
      bGoodCst = 0
   Else
      txtNme.ForeColor = Es_TextForeColor
      bGoodCst = 1
   End If
   
End Sub


Private Function GetThisRfq() As Byte
   Dim sCust As String
   Dim sAmPm As String
   
   sCust = Compress(cmbCst)
   sSql = "SELECT * FROM RfqsTable WHERE RFQREF='" & cmbRfq & "' " _
          & "AND RFQREF<>'NONE' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRfq, ES_KEYSET)
   If bSqlRows Then
      With RdoRfq
         If sCust <> "" & Trim(!RFQCUST) Then
            MsgBox "That RFQ Belongs To Another Customer", vbInformation, Caption
            GetThisRfq = 2
         Else
            txtDte = Format(!RFQDATE, "mm/dd/yyyy")
            txtDue = Format(!RFQDUE, "mm/dd/yyyy")
            txtBuy = "" & Trim(!RFQBY)
            txtDsc = "" & Trim(!RFQDESC)
            optCom.Value = Format(!RFQCOMPLETE, "0")
            cmbRFQDateAMPM.ListIndex = Format("0" & !RFQDATEAMPM, "0")
            cmbRFQDueAmPm.ListIndex = Format("0" & !RFQDUEAMPM, "0")
            cmbCompleted = Format(!RFQCOMPLETED, "mm/dd/yyyy")
            GetThisRfq = 1
         End If
      End With
   Else
      txtDsc = ""
      optCom.Value = vbUnchecked
      txtBuy = ""
      GetThisRfq = 0
   End If
   Exit Function
   
DiaErr1:
   sProcName = "getthisrfq"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub AddNewRfq()
   Dim bResponse As Byte
   Dim sMsg As String
   sMsg = "Do You Wish To Add RFQ " & cmbRfq & " For " _
          & "Customer " & cmbCst & "?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.ADOErrNum = 0
      sSql = "INSERT INTO RfqsTable (RFQREF,RFQCUST,RFQCOMPLETE) " _
             & "VALUES('" & cmbRfq & "','" & Compress(cmbCst) _
             & "',0)"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         MsgBox "The RFQ Was Successfully Added.", _
            vbInformation, Caption
         cmbRfq.AddItem cmbRfq
         bGoodRfq = GetThisRfq()
         txtDsc.SetFocus
      Else
         MsgBox "Could Not Successfully Add The RFQ.", _
            vbExclamation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub
