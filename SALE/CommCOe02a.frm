VERSION 5.00
Object = "{A964BDA3-3E93-11CF-9A0F-9E6261DACD1C}#3.7#0"; "Resize32.ocx"
Begin VB.Form CommCOe02a 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revise Sales Order Commissions"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdHlp 
      Appearance      =   0  'Flat
      Height          =   250
      Left            =   0
      Picture         =   "CommCOe02a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      ToolTipText     =   "Subject Help"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   250
   End
   Begin VB.CheckBox optFrom 
      Caption         =   "From Soit"
      Height          =   255
      Left            =   3600
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox optSoItems 
      Height          =   195
      Left            =   2520
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame z2 
      Height          =   615
      Left            =   1560
      TabIndex        =   39
      Top             =   4200
      Width           =   4095
      Begin VB.OptionButton optGro 
         Caption         =   "&Gross Margin"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "((Unit Cost - Standard Cost) * Commission %) + Base"
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optUnt 
         Caption         =   "&Unit Price"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "(Unit Cost * Commission %) + Base"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Select A New Sales Order"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "&Select"
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      ToolTipText     =   "Select Sales Order Items"
      Top             =   720
      Width           =   875
   End
   Begin VB.ComboBox cmbSon 
      Height          =   288
      Left            =   1800
      TabIndex        =   0
      Tag             =   "1"
      ToolTipText     =   "Contains Sales Orders With Commissionable Line Items.  "
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdAsg 
      Caption         =   "&Update"
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Assign Default Sales Person To All Sales Order Items"
      Top             =   2160
      Width           =   875
   End
   Begin VB.CheckBox optSlp 
      Height          =   195
      Left            =   1560
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame z3 
      Height          =   30
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   6495
   End
   Begin VB.CommandButton cmdDft 
      Caption         =   "&Default"
      Enabled         =   0   'False
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Restore Default Commission"
      Top             =   4920
      Width           =   875
   End
   Begin VB.CommandButton cmdChg 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2640
      Picture         =   "CommCOe02a.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Add/Remove Sales Persons"
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.ComboBox cmbSlp 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Percentage"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtFlt 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Tag             =   "1"
      ToolTipText     =   "Base"
      Top             =   4920
      Width           =   975
   End
   Begin VB.ComboBox cmbItm 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Tag             =   "8"
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdCan 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   435
      Left            =   5880
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   875
   End
   Begin ResizeLibCtl.ReSize ReSize1 
      Left            =   6480
      Top             =   4080
      _Version        =   196615
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      Enabled         =   -1  'True
      FormMinWidth    =   0
      FormMinHeight   =   0
      FormDesignHeight=   5625
      FormDesignWidth =   6810
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculate Commision on "
      Height          =   495
      Index           =   12
      Left            =   240
      TabIndex        =   40
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblStd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   38
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Cost"
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   37
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Commission"
      Height          =   495
      Index           =   6
      Left            =   240
      TabIndex        =   35
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblNme 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3000
      TabIndex        =   34
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblCst 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   33
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblPre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   288
      Left            =   1560
      TabIndex        =   31
      Top             =   360
      Width           =   255
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Orders"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label lblSln 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   29
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblSlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2280
      TabIndex        =   28
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Default Sales Person"
      Height          =   615
      Index           =   10
      Left            =   240
      TabIndex        =   27
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "%  ="
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   26
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Base)  + "
      Height          =   255
      Index           =   7
      Left            =   2640
      TabIndex        =   25
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Person(s)"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   24
      Top             =   3480
      Width           =   1812
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ext. Price"
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quanity"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label z1 
      BackStyle       =   0  'Transparent
      Caption         =   "Commissionable Items"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4680
      TabIndex        =   18
      ToolTipText     =   "(Estimated Commission)"
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblExt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5280
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblUnt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblISlp 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lblprt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   2220
      Width           =   2775
   End
End
Attribute VB_Name = "CommCOe02a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** ES/2000 (ES/2001 - ES/2007) is the property of           ***
'*** ESI Software Engineering Inc, Stanwood, Washington, USA  ***
'*** and is protected under US and International copyright    ***
'*** laws and treaties.                                       ***
'See the UpdateTables prodecure for database revisions
' CommCOe02a - Revise Sales Order Item Commissions
'
' Created 08/26/03 (JCW)
'
' Revisions
' 08/26/03 (nth) Revised and updated
' 12/16/03 (nth) Changed SO combo sort to DESC
' 01/24/05 (nth) Added gross margin option per AUBCOR.
' 8/26/05 optFrom and bSaved
'11/11/05 (cjs) Error trapped Object error (KeySet) and formula error in UpdateTotals
Option Explicit
Dim RdoCom As ADODB.Recordset

Dim bOnLoad As Byte
Dim bCancel As Byte
Dim bGoodCom As Byte
Public bGoodSO As Byte
Dim bSaved As Byte


Dim iItem As Integer
Dim sRev As String
Dim sMsg As String


'Const CURRENCYMASK = "#,###,###,##0.00"

Private txtKeyPress() As New EsiKeyBd
Private txtGotFocus() As New EsiKeyBd
Private txtKeyDown() As New EsiKeyBd

Private Sub cmbItm_Click()
   GetThisItem
   
End Sub

Private Sub cmbItm_LostFocus()
   If Not bCancel Then GetThisItem
   
End Sub

Private Sub cmbSlp_Click()
   bGoodCom = GetCommission()
End Sub

Private Sub cmbSlp_LostFocus()
   If Not bCancel Then bGoodCom = GetCommission()
   
End Sub

Private Sub cmbSon_Click()
   bGoodSO = GetSalesOrder
End Sub

Private Sub cmbSon_LostFocus()
   If Not bCancel Then bGoodSO = GetSalesOrder()
   
End Sub

Private Sub cmdAsg_Click()
   If bGoodSO Then ApplyDefaultSP
   
End Sub

Private Sub cmdCan_MouseDown(Button As Integer, Shift As Integer, _
                             x As Single, y As Single)
   bCancel = True
End Sub

Private Sub cmdChg_Click()
   optSlp.Value = vbChecked
   CommCOe02b.Show
End Sub

Private Sub cmdCnl_Click()
   Set RdoCom = Nothing
   cmbItm.Clear
   cmbSlp.Clear
   lblTot = ""
   lblprt = ""
   lblISlp = ""
   lblUnt = ""
   lblExt = ""
   lblQty = ""
   txtFlt = ""
   txtPer = ""
   ManageBoxs False
   cmbSon.SetFocus
   
End Sub

Private Sub cmdDft_Click()
   SetDefaultComm
End Sub

Private Sub cmdHlp_Click()
   If cmdHlp Then
      MouseCursor 13
      OpenHelpContext 2402
      MouseCursor 0
      cmdHlp = False
   End If
   
End Sub

Private Sub cmdSel_Click()
   If bGoodSO Then
      ManageBoxs True
      GetCommItems
      bGoodCom = GetCommission()
      cmdChg.Enabled = True
   End If
End Sub

Private Sub Form_Load()
   FormLoad Me
   FormatControls
'   On Error Resume Next
'   sSql = "ALTER TABLE SpcoTable ADD SMCOGM tinyint default(0) NULL"
'   RdoCon.Execute sSql, rdExecDirect
   bOnLoad = True
End Sub

Private Sub FillCombo()
   MouseCursor 13
   If bOnLoad Then FillSon
   bOnLoad = 0
   MouseCursor 0
   
End Sub

Public Sub FormatControls()
   Dim b As Byte
   b = AutoFormatControls(Me, txtKeyPress(), txtGotFocus(), txtKeyDown())
   
End Sub

Private Sub Form_Activate()
   MdiSect.lblBotPanel = Caption
   ' 9/3/2009 added Val(cmbSon) so that the combo box is not filled
   ' when called from Sales Item.
   If bOnLoad And Val(cmbSon) = 0 Then
      FillCombo
      bSaved = 0
      bOnLoad = False
   End If
   MouseCursor 0
End Sub

Private Sub Form_Resize()
   Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If optFrom.Value = vbUnchecked Then FormUnload
   Set RdoCom = Nothing
   Set CommCOe02a = Nothing
End Sub

Private Sub cmdCan_Click()
   Dim bResponse As Byte
   If bSaved = 0 Then
      bResponse = MsgBox("Any Changes Have Note been Saved (Updated) " & vbCrLf _
                  & "Continue To Exit Anyway?", ES_NOQUESTION, Caption)
      If bResponse = vbNo Then Exit Sub
   End If
   Unload Me
   
End Sub

Private Sub optGro_Click()
   If Not bCancel Then
      If optGro Then
         If bGoodCom = 1 Then
            On Error Resume Next
            With RdoCom
               '.Edit
               !SMCOGM = 1
               .Update
            End With
         End If
         If Err > 0 Then
            ValidateEdit
         End If
         UpdateTotals
      End If
   End If
End Sub

Private Sub optSlp_Click()
   If optSlp.Value = vbUnchecked Then
      FillItemSalesPersons
   End If
End Sub

Private Sub FillSon()
   Dim RdoSon As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
'   sSql = "SELECT DISTINCT SONUMBER,INVPAY FROM SohdTable,SoitTable," _
'          & "PartTable,CihdTable WHERE SONUMBER=ITSO AND ITPART=PARTREF " _
'          & "AND ITINVOICE*=INVNO AND PACOMMISSION=1 " _
'          & "ORDER BY SONUMBER DESC"
   sSql = "SELECT DISTINCT SONUMBER,INVPAY" & vbCrLf _
      & "FROM SohdTable" & vbCrLf _
      & "JOIN SoitTable ON SONUMBER=ITSO" & vbCrLf _
      & "JOIN PartTable ON ITPART=PARTREF" & vbCrLf _
      & "LEFT JOIN CihdTable ON ITINVOICE=INVNO" & vbCrLf _
      & "WHERE PACOMMISSION=1" & vbCrLf _
      & "ORDER BY SONUMBER DESC"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon, ES_FORWARD)
   If bSqlRows Then
      With RdoSon
         Do Until .EOF
            AddComboStr cmbSon.hWnd, Format(.Fields(0), SO_NUM_FORMAT)
            .MoveNext
         Loop
         ClearResultSet RdoSon
      End With
      If cmbSon.ListCount > 0 Then
         cmbSon.ListIndex = 0
         bGoodSO = GetSalesOrder
      End If
   End If
   Set RdoSon = Nothing
   Exit Sub
   
DiaErr1:
   sProcName = "fillson"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetSalesOrder() As Byte
   Dim RdoSon As ADODB.Recordset
   On Error GoTo DiaErr1
'   sSql = "SELECT SOTYPE,CUNICKNAME,CUNAME,SPNUMBER,SPLAST,SPFIRST " _
'          & "FROM SohdTable,CustTable,SprsTable WHERE SOCUST = CUREF AND " _
'          & "SOSALESMAN*=SPNUMBER AND SONUMBER = " & Val(cmbSon)
   sSql = "SELECT SOTYPE,CUNICKNAME,CUNAME,SPNUMBER,SPLAST,SPFIRST " _
          & "FROM SohdTable" & vbCrLf _
          & "JOIN CustTable ON SOCUST=CUREF" & vbCrLf _
          & "LEFT JOIN SprsTable ON SOSALESMAN=SPNUMBER" & vbCrLf _
          & "WHERE SONUMBER = " & Val(cmbSon)
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoSon)
   If bSqlRows Then
      With RdoSon
         lblPre = "" & Trim(!SOTYPE)
         lblCst = "" & Trim(!CUNICKNAME)
         lblNme = "" & Trim(!CUNAME)
         lblSlp = "" & Trim(!SPFIRST) & " " & Trim(!SPLAST)
         lblSln = "" & Trim(!SPNumber)
         If IsNull(!SPNumber) Then
            cmdAsg.Enabled = False
         Else
            cmdAsg.Enabled = True
         End If
         ClearResultSet RdoSon
         cmdSel.Enabled = True
         GetSalesOrder = True
      End With
   Else
      GetSalesOrder = False
      cmdSel.Enabled = False
      cmdAsg.Enabled = False
      lblPre = ""
      lblCst = ""
      lblNme = ""
   End If
   Set RdoSon = Nothing
   Exit Function
DiaErr1:
   sProcName = "getsalesor"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Function

Private Sub GetCommItems()
   ' Get commissionable items for this sales order
   Dim RdoItm As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT ITNUMBER,ITREV FROM SoitTable INNER JOIN " _
          & "PartTable ON SoitTable.ITPART = PartTable.PARTREF " _
          & "WHERE (PACOMMISSION = 1) And (ITSO = " _
          & Val(cmbSon) & ")"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   If bSqlRows Then
      cmdAsg.Enabled = True
      With RdoItm
         Do Until .EOF
            AddComboStr cmbItm.hWnd, .Fields(0) & Trim(.Fields(1))
            .MoveNext
         Loop
         ClearResultSet RdoItm
      End With
      If cmbItm.ListCount > 0 Then
         cmbItm.ListIndex = 0
      End If
   End If
   Set RdoItm = Nothing
   Exit Sub
DiaErr1:
   sProcName = "getcommit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Public Sub GetThisItem()
   Dim RdoItm As ADODB.Recordset
   
   On Error GoTo DiaErr1
   
   If Not IsNumeric(Right(cmbItm, 1)) Then
      sRev = Right(cmbItm, 1)
      iItem = Left(cmbItm, Len(cmbItm) - 1)
   Else
      iItem = cmbItm
      sRev = ""
   End If
   
   ' Get detail items part price qty ect.
   sSql = "SELECT PARTNUM,ITQTY,ITDOLLARS,PASTDCOST FROM SoitTable INNER JOIN " _
          & "PartTable ON ITPART = PARTREF WHERE ITSO = " & Val(cmbSon) _
          & " AND ITNUMBER = " & iItem & " AND ITREV = '" & sRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm)
   If bSqlRows Then
      With RdoItm
         lblprt = Trim(.Fields(0))
         lblQty = Format(.Fields(1), "0.000")
         lblUnt = Format(.Fields(2), "#,##0.00#")
         lblExt = Format((.Fields(1) * .Fields(2)), CURRENCYMASK)
         lblStd = Format(.Fields(3), "#,##0.00#")
      End With
   End If
   Set RdoItm = Nothing
   FillItemSalesPersons
   Exit Sub
DiaErr1:
   sProcName = "getthisit"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub FillItemSalesPersons()
   Dim rdoSlp As ADODB.Recordset
   cmbSlp.Clear
   sSql = "SELECT SPNUMBER FROM SpcoTable INNER JOIN " _
          & "SprsTable ON SMCOSM = SPNUMBER WHERE SMCOSO = " _
          & Val(cmbSon) & " AND SMCOSOIT = " & iItem & "" _
          & " AND SMCOITREV = '" & sRev & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp, ES_FORWARD)
   If bSqlRows Then
      With rdoSlp
         Do Until .EOF
            AddComboStr cmbSlp.hWnd, .Fields(0)
            .MoveNext
         Loop
         ClearResultSet rdoSlp
      End With
      If cmbSlp.ListCount > 0 Then
         cmbSlp.ListIndex = 0
      End If
   Else
      lblISlp = "*** No Sales Persons Assigned ***"
   End If
   Set rdoSlp = Nothing
   Exit Sub
DiaErr1:
   sProcName = "fillitemsa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Function GetCommission() As Byte
   Dim rdoSlp As ADODB.Recordset
   On Error GoTo DiaErr1
   sSql = "SELECT SPFIRST,SPLAST FROM SprsTable WHERE SPNUMBER = '" _
          & Trim(cmbSlp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, rdoSlp)
   If bSqlRows Then
      With rdoSlp
         lblISlp = Trim(.Fields(0)) & " " _
                   & Trim(.Fields(1))
      End With
   Else
      lblISlp = "*** Sales Person Not Found ***"
   End If
   
   Set rdoSlp = Nothing
   bGoodCom = 0
   sSql = "SELECT SMCOPCT,SMCOAMT,SMCOREVISED,SMCOGM FROM SpcoTable " _
          & "WHERE SMCOSO=" & Val(cmbSon) & " AND SMCOSOIT=" & iItem _
          & " AND SMCOITREV='" & sRev & "' AND SMCOSM='" & Trim(cmbSlp) & "'"
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoCom, ES_KEYSET)
   If bSqlRows Then
      With RdoCom
         cmdDft.Enabled = True
         txtFlt.Enabled = True
         txtPer.Enabled = True
         
         txtFlt = Format(.Fields(1), CURRENCYMASK)
         txtPer = Format(.Fields(0), "0.00#")
         
         If .Fields(3) = 1 Then
            optGro = vbChecked
         Else
            optUnt = vbChecked
         End If
         UpdateTotals
         
         bGoodCom = 1
         GetCommission = bGoodCom
         On Error Resume Next
         '.Edit
         !SMCOREVISED = Format(ES_SYSDATE, "mm/dd/yy")
         .Update
      End With
   Else
      cmdDft.Enabled = False
      txtFlt.Enabled = False
      txtPer.Enabled = False
      GetCommission = 0
   End If
   
   Exit Function
DiaErr1:
   sProcName = "getcommis"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
   
End Function

Private Sub optUnt_Click()
   If Not bCancel Then
      If optUnt Then
         If bGoodCom = 1 Then
            On Error Resume Next
            With RdoCom
               '.Edit
               !SMCOGM = 0
               .Update
            End With
         End If
         If Err > 0 Then
            ValidateEdit
         End If
         UpdateTotals
      End If
   End If
End Sub

Private Sub txtFlt_LostFocus()
   txtFlt = CheckLen(txtFlt, 10)
   txtFlt = Format(txtFlt, CURRENCYMASK)
   On Error Resume Next
   If bGoodCom = 1 Then
      With RdoCom
         '.Edit
         !SMCOAMT = txtFlt
         .Update
      End With
   End If
   If Err > 0 Then
      ValidateEdit
   End If
   UpdateTotals

End Sub

Private Sub txtPer_LostFocus()
   txtPer = CheckLen(Format(txtPer, "###0.000"), 10)
   On Error Resume Next
   If bGoodCom = 1 Then
      With RdoCom
         '.Edit
         !SMCOPCT = CSng(txtPer)
         .Update
      End With
   End If
   If Err > 0 Then
      ValidateEdit
   End If
   UpdateTotals
   
End Sub

Private Sub UpdateTotals()
   Dim cExt As Currency
   On Error Resume Next
   If optGro Then
      cExt = Val(Format(lblExt, "0000000.000")) - (Val(lblStd) * Val(Format(lblQty, "0000000.000")))
   Else
      cExt = Val(Format(lblExt, "0000000.000"))
   End If
   lblTot = Format(Val(txtFlt) + _
            (cExt * Val(txtPer) / 100), CURRENCYMASK)
   
   
End Sub

Private Sub ManageBoxs(bOn As Byte)
   txtPer.Enabled = bOn
   txtFlt.Enabled = bOn
   cmbItm.Enabled = bOn
   cmbSlp.Enabled = bOn
   cmdCnl.Enabled = bOn
   cmdChg.Enabled = bOn
   cmdDft.Enabled = bOn
   optUnt.Enabled = bOn
   optGro.Enabled = bOn
   
   cmbSon.Enabled = Not bOn
   cmdSel.Enabled = Not bOn
   cmdAsg.Enabled = Not bOn
   
End Sub

Public Sub SetDefaultComm(Optional sSP As String, Optional cAmt As Currency)
   ' CommCOe02b calls this sub to fill default commission values
   ' when adding a sales person
   ' If so sSP equals the sales person number
   
   Dim RdoRte As ADODB.Recordset
   Dim iList As Integer
   Dim cTotal As Currency
   
   On Error GoTo DiaErr1
   
   If cAmt > 0 Then
      cTotal = cAmt
   Else
      If lblExt <> "" Then
         cTotal = CCur(lblExt)
      Else
         cTotal = 0
      End If
   End If
   
   sSql = "SELECT SPFROM1,SPTHRU1,SPBASE1,SPPERC1,SPFROM2,SPTHRU2," _
          & "SPBASE2,SPPERC2,SPFROM3,SPTHRU3,SPBASE3,SPPERC3,SPFROM4," _
          & "SPTHRU4,SPBASE4,SPPERC4,SPFROM5,SPTHRU5,SPBASE5,SPPERC5," _
          & "SPFROM6,SPTHRU6,SPBASE6,SPPERC6,SPFROM7,SPTHRU7,SPBASE7," _
          & "SPPERC7,SPFROM8,SPTHRU8,SPBASE8,SPPERC8,SPFROM9,SPTHRU9," _
          & "SPBASE9,SPPERC9,SPFROM10,SPTHRU10,SPBASE10,SPPERC10 " _
          & "FROM SprsTable WHERE SPNUMBER = '"
   If sSP <> "" Then
      sSql = sSql & sSP & "'"
   Else
      sSql = sSql & Trim(cmbSlp) & "'"
   End If
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoRte)
   If bSqlRows Then
      With RdoRte
         For iList = 0 To 39 Step 4
            If CCur(.Fields(iList)) <= cTotal And CCur(.Fields(iList + 1)) >= cTotal Then
               If sSP <> "" Then
                  sSql = "UPDATE SpcoTable SET SMCOAMT = " & .Fields(iList + 2) _
                         & ",SMCOPCT = " & (.Fields(iList + 3)) _
                         & " WHERE SMCOSO = " & Val(cmbSon) _
                         & " AND SMCOSOIT = " & iItem _
                         & " AND SMCOITREV = '" & sRev _
                         & "' AND SMCOSM = '" & sSP & "'"
                 clsADOCon.ExecuteSQL sSql 'rdExecDirect
               Else
                  txtFlt = Format(.Fields(iList + 2), CURRENCYMASK)
                  txtPer = Format(.Fields(iList + 3), "0.00#")
                  On Error Resume Next
                  'RdoCom.Edit
                  RdoCom!SMCOAMT = txtFlt
                  RdoCom!SMCOPCT = txtPer
                  RdoCom.Update
                  If Err > 0 Then
                     ValidateEdit
                  End If
                  On Error GoTo DiaErr1
                  UpdateTotals
               End If
               Exit For
            End If
         Next
      End With
   End If
   Set RdoRte = Nothing
   Exit Sub
DiaErr1:
   sProcName = "setdefaul"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub

Private Sub ApplyDefaultSP()
   Dim RdoItm As ADODB.Recordset
   Dim sSalesPerson As String
   Dim sSoItem As String
   
   On Error GoTo DiaErr1
   
   sSoItem = cmbItm
   sSalesPerson = Me.cmbSlp      'Trim(lblSln)
   sSql = "SELECT ITNUMBER,ITREV,ITDOLLARS,ITQTY FROM SoitTable INNER JOIN " _
          & "PartTable ON SoitTable.ITPART=PartTable.PARTREF " _
          & "WHERE (PACOMMISSION=1) And (ITSO = " & Val(cmbSon) & ")" _
          & " AND ITNUMBER = '" & sSoItem & "'"
          
   bSqlRows = clsADOCon.GetDataSet(sSql, RdoItm, ES_FORWARD)
   Dim Item As New ClassSoItem
   If bSqlRows Then
      
      With RdoItm
         On Error Resume Next
         Do Until .EOF

            Dim rdo As ADODB.Recordset
            sSql = "select SMCOSO from SpcoTable" & vbCrLf _
               & "where SMCOSO = " & Val(cmbSon) & vbCrLf _
               & "and SMCOSOIT = " & !ITNUMBER & vbCrLf _
               & "and SMCOITREV = '" & Trim(!ITREV) & "'" & vbCrLf _
               & "and SMCOSM = '" & sSalesPerson & "'"
            
            If Not clsADOCon.GetDataSet(sSql, rdo) Then
                Item.InsertCommission Val(cmbSon), !ITNUMBER, Trim(!ITREV), sSalesPerson
                Item.UpdateCommissions Val(cmbSon), !ITNUMBER, Trim(!ITREV)
            End If
            .MoveNext
            
         Loop
         ClearResultSet RdoItm
      End With
      sMsg = "Sales Person " & sSalesPerson & " Applied To " & sSoItem _
             & vbCrLf & " Commissionable Sales Items."
      MsgBox sMsg, vbInformation, Caption
   Else
      sMsg = "No Commissionable Items Found."
      MsgBox sMsg, vbExclamation, Caption
   End If
   bSaved = 1
   Set RdoItm = Nothing
   Exit Sub
DiaErr1:
   sProcName = "applydefa"
   CurrError.Number = Err.Number
   CurrError.Description = Err.Description
   DoModuleErrors Me
End Sub
