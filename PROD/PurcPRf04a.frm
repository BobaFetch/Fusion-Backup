VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PurcPRf04a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change A Buyer ID"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf04a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.TextBox txtByr 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "3"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton cmdChg 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Change The Current Buyer's ID"
      Top             =   600
      Width           =   875
   End
   Begin VB.ComboBox cmbByr 
      Height          =   315
      Left            =   1680
      Sorted          =   -1  'True
      TabIndex        =   0
      Tag             =   "3"
      ToolTipText     =   "Select A Buyer "
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5280
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5880
      Top             =   2160
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2550
      FormDesignWidth =   6225
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   300
      Left            =   1680
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   3732
      _ExtentX        =   6588
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Buyer ID"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblByr 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Buyer ID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "PurcPRf04a"
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
Dim bBuyerExits As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbByr_Click()
   If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
   
End Sub


Private Sub cmbByr_LostFocus()
   cmbByr = CheckLen(cmbByr, 20)
   If Len(Trim(cmbByr)) > 0 Then GetCurrentBuyer cmbByr
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdChg_Click()
   Dim b As Byte
   If Len(Trim(txtByr)) = 0 Then Exit Sub
   b = TestBuyer()
   If b = 1 Then
      MsgBox "That Buyer ID Is Currently In Use.", _
         vbInformation, Caption
   Else
      UpdateBuyer
   End If
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4353
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then
      FillBuyers
      If cmbByr.ListCount > 0 Then
         cmbByr = cmbByr.List(0)
         GetCurrentBuyer cmbByr
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
   Set PurcPRf04a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub


Private Sub lblByr_Change()
   If Left(lblByr, 8) = "*** Buye" Then
      lblByr.ForeColor = ES_RED
   Else
      lblByr.ForeColor = Es_TextForeColor
   End If
   
End Sub


Private Function TestBuyer() As Byte
   Dim RdoByr As ADODB.Recordset
   Dim sBuyer As String
   sBuyer = UCase$(Compress(txtByr))
   sSql = "SELECT BYREF,BYNUMBER FROM BuyrTable WHERE BYREF='" & sBuyer & "' "
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoByr, ES_FORWARD)
   If bSqlRows Then
      With RdoByr
         txtByr = "" & Trim(!BYNUMBER)
         TestBuyer = 1
         ClearResultSet RdoByr
      End With
   Else
      TestBuyer = 0
   End If
   Set RdoByr = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "testbuyer"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub UpdateBuyer()
   Dim bResponse As Byte
   Dim sNBuyer As String
   Dim sOBuyer As String
   Dim sMsg As String
   
   sMsg = "Dou Wish To Change The Buyer ID?"
   bResponse = MsgBox(sMsg, ES_YESQUESTION, Caption)
   If bResponse = vbYes Then
      On Error Resume Next
      MouseCursor 11
      sOBuyer = Compress(cmbByr)
      sNBuyer = Compress(txtByr)
      If sNBuyer = "ALL" Then
         MsgBox "ALL Is A Reverved Word And Cannot " & vbCr _
            & "Be Used.  Please Select Another Number.", _
            vbInformation, Caption
         cmbByr = ""
         Err = 0
         txtByr.SetFocus
      Else
         prg1.Visible = True
         prg1.Value = 10
         
         clsADOCon.BeginTrans
         clsADOCon.ADOErrNum = 0
         
         sSql = "UPDATE BuyrTable SET BYREF='" & sNBuyer _
                & "',BYNUMBER='" & txtByr & "' WHERE BYREF='" _
                & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 20
         sSql = "UPDATE PohdTable SET POBUYER='" & sNBuyer _
                & "' WHERE POBUYER='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 30
         sSql = "UPDATE VndrTable SET VEBUYER='" & sNBuyer _
                & "' WHERE VEBUYER='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 40
         sSql = "UPDATE MrplTable SET MRP_POBUYER='" & sNBuyer _
                & "' WHERE MRP_POBUYER='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 50
         sSql = "UPDATE PartTable SET PABUYER='" & sNBuyer _
                & "' WHERE PABUYER='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 60
         sSql = "UPDATE PcodTable SET PCBUYER='" & sNBuyer _
                & "' WHERE PCBUYER='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 70
         sSql = "UPDATE BuycTable SET BYREF='" & sNBuyer _
                & "' WHERE BYREF='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 80
         sSql = "UPDATE BuypTable SET BYREF='" & sNBuyer _
                & "' WHERE BYREF='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         
         prg1.Value = 90
         sSql = "UPDATE BuyvTable SET BYREF='" & sNBuyer _
                & "' WHERE BYREF='" & sOBuyer & "' "
         clsADOCon.ExecuteSQL sSql
         MouseCursor 0
         If clsADOCon.ADOErrNum = 0 Then
            prg1.Value = 100
            clsADOCon.CommitTrans
            MsgBox "The Buyer Has Been Updated.", _
               vbInformation, Caption
            cmbByr.Clear
            FillBuyers
            cmbByr = txtByr
            txtByr = ""
            cmbByr.SetFocus
         Else
            clsADOCon.RollbackTrans
            MsgBox "Couldn't Update That Buyer.", _
               vbInformation, Caption
            cmbByr.SetFocus
         End If
         prg1.Visible = False
      End If
   Else
      CancelTrans
   End If
   
End Sub

Private Sub txtByr_LostFocus()
   txtByr = CheckLen(txtByr, 20)
   
End Sub
