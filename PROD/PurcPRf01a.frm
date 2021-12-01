VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form PurcPRf01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancel A Purchase Order"
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
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "PurcPRf01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
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
      TabIndex        =   9
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
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Apply"
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Cancel The Selected Purchase Order"
      Top             =   600
      Width           =   855
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
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   3720
   End
   Begin VB.Label lblVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1320
      TabIndex        =   6
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
Attribute VB_Name = "PurcPRf01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Seattle, Washington USA    ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'12/14/05 Revised selection criteria
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

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   cmbPon = ""
   
End Sub


Private Sub cmdDel_Click()
   If bGoodPo Then
      If OkToCancel() Then
         CancelPurchaseOrder
      Else
         MsgBox "PO " & cmbPon & " Has Open Or Received Items and Can't Be Canceled.", vbInformation, Caption
      End If
   Else
      MsgBox "That Purchase Order Wasn't Found Or Previously Canceled.", vbInformation, Caption
   End If
   bValidPo = 0
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 4350
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
   Set PurcPRf01a = Nothing
   
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
          & "POCAN=0)"
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

Private Function OkToCancel() As Boolean
   Dim RdoPit As ADODB.Recordset
'   sSql = "SELECT PINUMBER,PIAQTY,PITYPE FROM " _
'          & "PoitTable WHERE (PINUMBER=" & Val(cmbPon) & " " _
'          & "AND PIAQTY=0 AND PITYPE=16)"
   sSql = "SELECT TOP 1 PINUMBER FROM PoitTable" & vbCrLf _
      & "WHERE PINUMBER=" & Val(cmbPon) & vbCrLf _
      & "AND PITYPE IN (" & IATYPE_PoReceipt & "," & IATYPE_PoInvoiced & ")"
   If clsADOCon.GetDataSet(sSql, RdoPit, ES_FORWARD) Then
      OkToCancel = False
   Else
      OkToCancel = True
   End If
   
   Set RdoPit = Nothing
   
End Function

Private Sub CancelPurchaseOrder()
   Dim bResponse As Byte
   Dim iRows As Integer
   Dim sMsg As String
   
   'if PO already canceled, say so
   Dim rdo As ADODB.Recordset
   sSql = "select POCAN from PohdTable where PONUMBER = " & cmbPon
   If clsADOCon.GetDataSet(sSql, rdo) Then
      If rdo!POCAN = 1 Then
         MsgBox "PO " & cmbPon & " already canceled."
         Exit Sub
      End If
   Else
      MsgBox "No such PO"
      Exit Sub
   End If
   
   Set rdo = Nothing
   
   'On Error Resume Next
   sMsg = "Are You Sure That You Want To Cancel PO " & cmbPon & "?"
   bResponse = MsgBox(sMsg, ES_NOQUESTION, Caption)
   If bResponse = vbYes Then
      clsADOCon.BeginTrans
      clsADOCon.ADOErrNum = 0
      
      sSql = "UPDATE PohdTable SET POCAN=1,POCANBY='" & sInitials & "',POCANCELED=Getdate() " _
             & "WHERE PONUMBER=" & Val(cmbPon) & ""
      clsADOCon.ExecuteSQL sSql
      
      If clsADOCon.RowsAffected > 0 Then
'         sSql = "UPDATE PoitTable SET PIPQTY=0,PITYPE=16 WHERE " _
'                & "PINUMBER=" & Val(cmbPon) & ""
         sSql = "UPDATE PoitTable SET PITYPE = " & IATYPE_PoCanceledItem & vbCrLf _
            & "WHERE PINUMBER=" & Val(cmbPon) & ""
         clsADOCon.ExecuteSQL sSql
         'MsgBox cmbPon & " Successfully Canceled.", vbInformation, Caption
         cmbPon = ""
         txtRel = ""
         lblVnd = ""
         lblNme = ""
         clsADOCon.CommitTrans
         FillCombo
         MsgBox "Purchase Order Was Successfully Canceled.", _
            vbInformation, Caption
      Else
         clsADOCon.RollbackTrans
         MsgBox "Could Not Cancel That Purchase Order.", _
            vbExclamation, 48, Caption
      End If
   Else
      CancelTrans
   End If
   Exit Sub
   
End Sub

Private Sub FillCombo()
   Dim iList As Integer
   cmbPon.Clear
   cmbSrt.Clear
   On Error GoTo DiaErr1
'   sSql = "SELECT PohdTable.PONUMBER FROM PohdTable LEFT JOIN " _
'          & "PoitTable ON PohdTable.PONUMBER=PoitTable.PINUMBER " _
'          & "WHERE PoitTable.PINUMBER Is Null ORDER BY PohdTable.PONUMBER DESC"
'   LoadNumComboBox cmbSrt, "000000"
   
'   sSql = "SELECT DISTINCT PINUMBER FROM PoitTable WHERE " _
'          & "PIAQTY=0 AND (PITYPE=16 AND PITYPE<>17 AND PITYPE<>15 " _
'          & "AND PITYPE<>18) ORDER BY PINUMBER DESC "

   'only display PO's with no received or invoiced items
   sSql = "SELECT PONUMBER FROM PohdTable" & vbCrLf _
      & "WHERE NOT EXISTS (SELECT PINUMBER FROM PoitTable" & vbCrLf _
      & "  WHERE PINUMBER = PONUMBER" & vbCrLf _
      & "  AND PITYPE IN (" & IATYPE_PoReceipt & "," & IATYPE_PoInvoiced & "))" & vbCrLf _
      & "AND POCAN = 0" & vbCrLf _
      & "ORDER BY PONUMBER DESC "
      
   LoadNumComboBox cmbSrt, "000000"
   If cmbSrt.ListCount > 0 Then
      For iList = cmbSrt.ListCount - 1 To 0 Step -1
         cmbPon.AddItem cmbSrt.List(iList)
      Next
      cmbPon = cmbPon.List(0)
      bGoodPo = GetPurchaseOrder()
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
   If Left$(lblVnd, 4) = "** W" Then lblVnd.ForeColor = ES_RED _
            Else lblVnd.ForeColor = vbBlack
   
End Sub
