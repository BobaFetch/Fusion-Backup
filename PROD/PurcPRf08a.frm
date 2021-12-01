VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf08a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open A Canceled Purchase Order"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      ToolTipText     =   "Reopen The Selected Purchase Order"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf08a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox cmbSrt 
      Height          =   288
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.ComboBox cmbPon 
      Height          =   288
      Left            =   1320
      TabIndex        =   0
      Tag             =   "1"
      Top             =   840
      Width           =   1095
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   5280
      Top             =   1560
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   2160
      FormDesignWidth =   6195
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
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Visible - For Sorting"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2172
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1560
      Width           =   3720
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label txtRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "PurcPRf08a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/14/05 New
Option Explicit
Dim bOnLoad As Byte
Dim bGoodPo As Byte
Dim bValidPo As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder()
   
End Sub

Private Sub cmbPon_LostFocus()
   If Len(cmbPon) > 0 Then
      cmbPon = CheckLen(cmbPon, 6)
      cmbPon = Format(Abs(Val(cmbPon)), "000000")
      bGoodPo = GetPurchaseOrder()
   Else
      bGoodPo = False
   End If
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmbPon = ""
   
End Sub


Private Sub cmdDel_Click()
   Dim bResponse As Byte
   If bGoodPo Then
      bValidPo = CheckPurchaseOrder()
      If bValidPo = 1 Then
         bResponse = MsgBox("Re-Open The Selected Purchase Order?", _
                     ES_YESQUESTION, Caption)
         If bResponse = vbYes Then
            On Error Resume Next
            clsADOCon.ADOErrNum = 0
            
            sSql = "UPDATE PohdTable SET POCAN=0,POCANCELED=Null " _
                   & "WHERE PONUMBER=" & Val(cmbPon) & ""
            clsADOCon.ExecuteSQL sSql
            
            ' Set to Open PO if we have not received the items.
            sSql = "UPDATE PoitTable SET PITYPE = " & IATYPE_PoOpenItem & vbCrLf _
               & "WHERE PITYPE = " & IATYPE_PoCanceledItem & vbCrLf _
               & " AND PINUMBER=" & Val(cmbPon) & " AND PIADATE IS NULL"
            
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            
            ' Set to Open PO if we have not received the items.
            sSql = "UPDATE PoitTable SET PITYPE = " & IATYPE_PoReceipt & vbCrLf _
               & "WHERE PITYPE = " & IATYPE_PoCanceledItem & vbCrLf _
               & " AND PINUMBER=" & Val(cmbPon) & " AND PIADATE IS NOT NULL"
            
            clsADOCon.ExecuteSQL sSql ' rdExecDirect
            
            If clsADOCon.ADOErrNum = 0 Then
               SysMsg cmbPon & " Was Reopened.", True
               FillCombo
            Else
               MsgBox "Re-Opening The Purchase Order Failed.", _
                  vbInformation, Caption
            End If
         Else
            CancelTrans
         End If
      Else
         MsgBox "PO " & cmbPon & " Is Open And Need Not Be Reopened.", vbInformation, Caption
      End If
   Else
      MsgBox "That Purchase Order Wasn't Found Or Is Open.", vbInformation, Caption
   End If
   bValidPo = 0
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4357
      cmdHlp = False
      MouseCursor 0
   End If
   
End Sub


Private Sub Form_Activate()
   MDISect.lblBotPanel = Caption
   If bOnLoad Then FillCombo
   bOnLoad = 0
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
   On Error Resume Next
   FormUnload
   Set PurcPRf08a = Nothing
   
End Sub



Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub




Private Function GetPurchaseOrder() As Byte
   Dim RdoPon As ADODB.Recordset
   MouseCursor 13
   sSql = "SELECT PONUMBER,POVENDOR,POCAN,VEREF," _
          & "VENICKNAME,VEBNAME FROM PohdTable,VndrTable WHERE " _
          & "(POVENDOR=VEREF AND PONUMBER=" & Val(cmbPon) & " AND " _
          & "POCAN=1)"
   On Error GoTo DiaErr1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPon)
   If bSqlRows Then
      With RdoPon
         cmbPon = Format(!PONUMBER, "000000")
         lblVnd = "" & Trim(!VENICKNAME)
         lblNme = "" & Trim(!VEBNAME)
         ClearResultSet RdoPon
      End With
      GetPurchaseOrder = True
   Else
      lblVnd = "** Wasn't Found **"
      txtRel = ""
      lblNme = ""
      GetPurchaseOrder = False
   End If
   On Error Resume Next
   MouseCursor 0
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurchaseor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   GetPurchaseOrder = False
   DoModuleErrors Me
   
End Function

Private Function CheckPurchaseOrder() As Byte
   Dim RdoPit As ADODB.Recordset
   sSql = "SELECT PONUMBER FROM PohdTable WHERE POCAN=1 "
   On Error GoTo DiaErr1
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoPit, ES_FORWARD)
   If bSqlRows Then
      CheckPurchaseOrder = 1
   Else
      CheckPurchaseOrder = 0
   End If
   Set RdoPit = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "checkpurchaseorder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   CheckPurchaseOrder = False
   DoModuleErrors Me
   
End Function


Private Sub FillCombo()
   Dim iList As Integer
   cmbPon.Clear
   cmbSrt.Clear
   On Error GoTo DiaErr1
   sSql = "SELECT PONUMBER FROM PohdTable WHERE POCAN=1 ORDER BY PONUMBER DESC"
   LoadNumComboBox cmbSrt, "000000"
   If cmbSrt.ListCount > 0 Then
      For iList = cmbSrt.ListCount - 1 To 0 Step -1
         cmbPon.AddItem cmbSrt.List(iList)
      Next
      cmbPon = cmbPon.List(0)
      'bGoodPo = GetPurchaseOrder()
   End If
   cmbPon.ToolTipText = "Qualifying PO's. " & cmbSrt.ToolTipText
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Sub lblVnd_Change()
   If Left$(lblVnd, 4) = "** W" Then
      lblVnd.ForeColor = ES_RED
      cmdDel.Enabled = False
   Else
      lblVnd.ForeColor = vbBlack
      cmdDel.Enabled = True
   End If
   
End Sub
