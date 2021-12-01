VERSION 5.00
Begin VB.Form RecvRVe01a 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Order Receipt"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "RecvRVe01a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.ComboBox txtRcd 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Tag             =   "4"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox optRcv 
      Caption         =   "Received"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox optItm 
      Caption         =   "items"
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdItm 
      Caption         =   "&Items"
      Height          =   315
      Left            =   5520
      TabIndex        =   2
      ToolTipText     =   "Show PO Items To Receive"
      Top             =   600
      Width           =   915
   End
   Begin VB.ComboBox cmbPon 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Select Or Enter PO"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5520
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin VB.PictureBox ReSize1 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Received Date"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   14
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label txtNme 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   3720
   End
   Begin VB.Label cmbVnd 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Date"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Number"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label txtRel 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPdt 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rel"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "RecvRVe01a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables procedure for database revisions
'1/03/05 Revised date checking to allow year change "before" (Larry H)
Option Explicit
Dim AdoQry As ADODB.Command
Dim AdoParameter1 As ADODB.Parameter
Dim AdoParameter2 As ADODB.Parameter

Dim bGoodPo As Byte
Dim bOnLoad As Byte

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   If txtRcd = "" Then
      txtRcd = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
End Sub


Private Sub cmbPon_Click()
   bGoodPo = GetPurchaseOrder
   
End Sub


Private Sub cmbPon_LostFocus()
   cmbPon = CheckLen(cmbPon, 6)
   cmbPon = Format(Abs(Val(cmbPon)), "000000")
   bGoodPo = GetPurchaseOrder
   
End Sub


Private Sub cmdCan_Click()
   Unload Me
   
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext "5301"
      MouseCursor 0
   End If
   
End Sub


Private Sub cmdItm_Click()
   If bGoodPo Then
      MouseCursor 13
      optItm.Value = vbChecked
      RecvRVe01b.Show
   Else
      MsgBox "Requires A PO With Unreceived Items.", vbInformation, Caption
   End If
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   If optItm.Value = vbChecked Then
      Unload RecvRVe01b
      optItm.Value = vbUnchecked
   End If
   If bOnLoad Then
      bOnLoad = 0
      FillCombo
   End If
   MouseCursor 0
   
End Sub

Private Sub Form_Load()
   FormLoad Me
   
   FormatControls
   sSql = "SELECT PONUMBER,PORELEASE,POVENDOR,PODATE,PINUMBER," _
          & "PIRELEASE,PITYPE FROM PohdTable,PoitTable WHERE " _
          & "(PONUMBER=PINUMBER AND PORELEASE=PIRELEASE) AND " _
          & "PITYPE=14 AND (PONUMBER= ? AND PORELEASE = ?)"
   Set AdoQry = New ADODB.Command
   AdoQry.CommandText = sSql
   Set AdoParameter1 = New ADODB.Parameter
   AdoParameter1.Type = adInteger
   Set AdoParameter2 = New ADODB.Parameter
   AdoParameter2.Type = adSmallInt
   
   AdoQry.Parameters.Append AdoParameter1
   AdoQry.Parameters.Append AdoParameter2
   
   
   'Set RdoQry = RdoCon.CreateQuery("", sSql)
   'RdoQry.MaxRows = 1
   bOnLoad = 1
   
End Sub


Private Sub Form_Resize()
   Refresh
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If optItm.Value = vbUnchecked Then FormUnload
   Set AdoParameter1 = Nothing
   Set AdoParameter2 = Nothing
   Set AdoQry = Nothing
   Set RecvRVe01a = Nothing
   
End Sub


Private Sub FillCombo()
   Dim RdoCmb As ADODB.Recordset
   Dim b As Byte
   On Error GoTo DiaErr1
   
   sJournalID = GetOpenJournal("PJ", Format$(ES_SYSDATE, "mm/dd/yy"))
   If Left(sJournalID, 4) = "None" Then
      sJournalID = ""
      b = 1
   Else
      If sJournalID = "" Then b = 0 Else b = 1
   End If
   If b = 0 Then
      MsgBox "There Is No Open Purchases Journal For This Period.", _
         vbExclamation, Caption
      Sleep 500
      Unload Me
      Exit Sub
   Else
      sJournalID = GetOpenJournal("IJ", Format$(ES_SYSDATE, "mm/dd/yy"))
      If Left(sJournalID, 4) = "None" Then
         sJournalID = ""
         b = 1
      Else
         If sJournalID = "" Then b = 0 Else b = 1
      End If
      If b = 0 Then
         MsgBox "There Is No Open Inventory Journal For This Period.", _
            vbExclamation, Caption
         Sleep 500
         Unload Me
         Exit Sub
      End If
   End If
   sProcName = "fillcombo"
   
   cmbPon.Clear
   sSql = "SELECT DISTINCT PONUMBER,POVENDOR,PORELEASE,POCAN,PINUMBER,PITYPE " _
          & "FROM PohdTable,PoitTable WHERE PONUMBER=PINUMBER " _
          & "AND PITYPE=14 and PIPQTY > 0"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCmb)
   If bSqlRows Then
      With RdoCmb
         Do Until .EOF
            AddComboStr cmbPon.hWnd, Format$(!PONUMBER, "000000")
            .MoveNext
         Loop
         ClearResultSet RdoCmb
      End With
   End If
   If cmbPon.ListCount > 0 Then
      cmbPon = cmbPon.List(0)
      bGoodPo = GetPurchaseOrder()
   End If
   Set RdoCmb = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillcombo"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Sub

Private Function GetPurchaseOrder() As Byte
   Dim RdoPon As ADODB.Recordset
   Dim bByte As Byte
   
   On Error GoTo DiaErr1
'   RdoQry.RowsetSize = 1
'   RdoQry(0) = Val(cmbPon)
'   RdoQry(1) = Val(txtRel)
   AdoQry.Parameters(0).Value = Val(cmbPon)
   AdoQry.Parameters(1).Value = Val(txtRel)
   bSqlRows = clsADOCon.GetQuerySet(RdoPon, AdoQry, ES_KEYSET, False, 1)
   If bSqlRows Then
      With RdoPon
         GetPurchaseOrder = 1
         txtRel = !PoRelease
         lblPdt = Format(!PODATE, "mm/dd/yyyy")
         cmbVnd = "" & Trim(!POVENDOR)
         bByte = FindVendor()
         ClearResultSet RdoPon
      End With
   Else
      GetPurchaseOrder = 0
      If Not bOnLoad Then MsgBox "Open Purchase Order " & cmbPon & " Wasn't Found.", vbInformation, Caption
   End If
   bOnLoad = 0
   Set RdoPon = Nothing
   Exit Function
   
DiaErr1:
   sProcName = "getpurchaseorder"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub lblNme_Click()
   
End Sub


Private Sub optItm_Click()
   'never visible-items are loaded
   
End Sub


Private Sub optRcv_Click()
   FillCombo
   
End Sub


Private Sub txtRcd_DropDown()
   ShowCalendarEx Me
   
End Sub


Private Sub txtRcd_LostFocus()
   Dim vToday As Variant
   Dim vPost As Variant
   txtRcd = CheckDateEx(txtRcd)
   vPost = Format(txtRcd, "yyyy,mm,dd")
   vToday = Format(ES_SYSDATE, "yyyy,mm,dd")
   On Error Resume Next
   If vPost > vToday Then
      Beep
      txtRcd = Format(ES_SYSDATE, "mm/dd/yyyy")
   End If
   
End Sub
