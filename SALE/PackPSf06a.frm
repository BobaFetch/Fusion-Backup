VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PackPSf06a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Company Transfer"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PackPSf06a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5040
      TabIndex        =   13
      ToolTipText     =   "Cancel This Packing Slip Company Transfer"
      Top             =   960
      Width           =   915
   End
   Begin VB.CheckBox optTransfers 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox cmbPsl 
      Height          =   315
      Left            =   1455
      Sorted          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Only Qualifying Packing Slips.  Select From List"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6600
      Top             =   4200
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2730
      FormDesignWidth =   6015
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Function Is To Reset Company Transfers Only. See Help"
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   4785
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed"
      Height          =   255
      Index           =   23
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   3795
   End
   Begin VB.Label lblDte 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblPrn 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfers Are Active"
      Height          =   285
      Index           =   8
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   2265
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing Slip"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "PackPSf06a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007)) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
'8/3/05 New
Option Explicit
Dim bCancel As Byte
Dim bOnLoad As Byte
Dim bGoodPs As Byte
Dim lTransfer As Long

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Function GetPackingSlip() As Byte
   Dim RdoPsl As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT PSNUMBER,PSCUST,PSDATE,PSPRINTED " _
          & "FROM PshdTable WHERE PSNUMBER='" & Trim(cmbPsl) _
          & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPsl, ES_FORWARD)
   If bSqlRows Then
      With RdoPsl
         lblCst = "" & Trim(!PSCUST)
         lblDte = Format(!PSDATE, "mm/dd/yyyy")
         lblPrn = Format(!PSPRINTED, "mm/dd/yyyy")
         GetPackingSlip = 1
         ClearResultSet RdoPsl
      End With
      FindCustomer Me, lblCst
   Else
      lblCst = ""
      lblDte = ""
      lblPrn = ""
      GetPackingSlip = 0
   End If
   Set RdoPsl = Nothing
   Exit Function
   
DiaErr1:
   GetPackingSlip = 0
   sProcName = "getpackingslip"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub CheckTransferInvoice()
   Dim RdoInv As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT TransferInvoice,AllowTransfers FROM Preferences " _
          & "WHERE PreRecord=1"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoInv, ES_FORWARD)
   If bSqlRows Then
      With RdoInv
         lTransfer = !TransferInvoice
         optTransfers.Value = !AllowTransfers
         ClearResultSet RdoInv
      End With
   End If
   Set RdoInv = Nothing
   Exit Sub
   
DiaErr1:
   lTransfer = 0
   
End Sub

Private Sub cmbPsl_Click()
   bGoodPs = GetPackingSlip()
   
End Sub


Private Sub cmbPsl_LostFocus()
   Dim bByte As Byte
   Dim iList As Integer
   If bCancel = 1 Then Exit Sub
   'cmbPsl = CheckLen(cmbPsl, 6)
   'cmbPsl = Format(Abs(Val(cmbPsl)), "000000")
   If cmbPsl.ListCount > 0 Then
      If Trim(cmbPsl) = "" Then
         cmbPsl = cmbPsl.List(0)
      End If
   End If
   For iList = 0 To cmbPsl.ListCount - 1
      If cmbPsl = cmbPsl.List(iList) Then bByte = 1
   Next
   If bByte = 1 Then
      bGoodPs = GetPackingSlip()
   Else
      bGoodPs = 0
      MsgBox "Select Or Enter The Packing Slip Number From One On The List.", _
         vbInformation, Caption
      If cmbPsl.ListCount > 0 Then cmbPsl = cmbPsl.List(0)
   End If
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub



Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2254
      MouseCursor 0
      cmdHlp = False
   End If
   bCancel = 0
   
End Sub

Private Sub cmdHlp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   bCancel = 1
   
End Sub


Private Sub cmdItm_Click()
   Dim bResponse As Byte
   Dim sMsg As String
   If bGoodPs = 0 Then
      MsgBox "Select A Qualifying Packing Slip From The List.", _
         vbInformation, Caption
      Exit Sub
   End If
   sMsg = "This Function Sets The PS Invoice " & Format(lTransfer, "000000") & " To " _
          & vbCrLf & "(0) Zero And Allow The Packing Slip To Be Canceled." & vbCrLf _
          & "The Action Cannot Be Reversed.  Continue Anyway?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      sSql = "UPDATE SoitTable SET ITINVOICE=0 WHERE " _
             & "ITPSNUMBER='" & cmbPsl & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      sSql = "UPDATE PshdTable SET PSINVOICE=0 WHERE " _
             & "PSNUMBER='" & cmbPsl & "'"
      clsADOCon.ExecuteSQL sSql 'rdExecDirect
      If clsADOCon.ADOErrNum = 0 Then
         clsADOCon.CommitTrans
         MsgBox "The Process On PS" & cmbPsl & " Was Completed.", _
            vbInformation, Caption
         FillCombo
      Else
         clsADOCon.RollbackTrans
         MsgBox "The Process On PS" & cmbPsl & " Was No Successfully Completed.", _
            vbInformation, Caption
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If bOnLoad Then
      CheckTransferInvoice
      FillCombo
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
   Set PackPSf06a = Nothing
   
End Sub

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   On Error GoTo DiaErr1
   cmbPsl.Clear
   sSql = "SELECT DISTINCT PSNUMBER,PSINVOICE FROM " _
          & "PshdTable WHERE PSINVOICE=" & lTransfer & " "
'   bSqlRows = clsADOCon.GetDataSet(sSql,RdoCmb, ES_FORWARD)
'   If bSqlRows Then
'      With RdoCmb
'         Do Until .EOF
'            'AddComboStr cmbPsl.hWnd, Trim$(Right$(!PsNumber, 6))
'            AddComboStr cmbPsl.hWnd, Trim$(!PsNumber)
'            .MoveNext
'         Loop
'         ClearResultSet RdoCmb
'      End With
'   End If
   LoadComboBoxAndSelect cmbPsl
   If cmbPsl.ListCount > 0 Then
      'cmbPsl = cmbPsl.List(0)
      bGoodPs = GetPackingSlip()
   Else
      lblNme = "*** No Matching Pack Slips Found ***"
      lblNme.ForeColor = ES_RED
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub
